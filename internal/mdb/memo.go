package mdb

import (
	"encoding/binary"
	"fmt"
)

// LVAL storage types (bitmask byte in the 12-byte memo reference).
const (
	LvalInline    = 0x80 // data stored inline in the row
	LvalSingle    = 0x40 // data in a single LVAL page record
	LvalMultiPage = 0x00 // data spans multiple LVAL page records
)


// ResolveMemo resolves a MEMO/OLE field reference to its full data.
// The raw bytes from the variable column area contain a 12-byte reference (Jet4).
func (db *Database) ResolveMemo(raw []byte) ([]byte, error) {
	if len(raw) == 0 {
		return nil, nil
	}

	// Inline data: raw IS the data (no 12-byte header).
	// Short inline data (< 12 bytes) is just the value itself.
	if len(raw) < 12 {
		result := make([]byte, len(raw))
		copy(result, raw)
		return result, nil
	}

	// 12-byte memo reference:
	// bytes 0-2: memo_len (3 bytes, little-endian)
	// byte 3: bitmask/type
	// bytes 4-7: lval_dp (page:row pointer)
	// bytes 8-11: unknown
	memoLen := int(raw[0]) | int(raw[1])<<8 | int(raw[2])<<16
	bitmask := raw[3]

	switch bitmask {
	case LvalInline:
		// Data follows immediately after the 12-byte header.
		data := raw[12:]
		if len(data) > memoLen {
			data = data[:memoLen]
		}
		result := make([]byte, len(data))
		copy(result, data)
		return result, nil

	case LvalSingle:
		// Data is in a single LVAL page record.
		lvalPage := binary.LittleEndian.Uint32(raw[4:8])
		rowID := int(lvalPage & 0xFF)        // low byte = row ID
		pageNum := int64(lvalPage >> 8)       // upper 3 bytes = page number
		if pageNum == 0 {
			// Alternative encoding: full uint32 page, separate row
			pageNum = int64(binary.LittleEndian.Uint32(raw[4:8]))
			rowID = 0
		}
		return db.readLvalRecord(pageNum, rowID, memoLen)

	case LvalMultiPage:
		// Data spans multiple LVAL pages.
		lvalPage := binary.LittleEndian.Uint32(raw[4:8])
		rowID := int(lvalPage & 0xFF)
		pageNum := int64(lvalPage >> 8)
		if pageNum == 0 {
			pageNum = int64(binary.LittleEndian.Uint32(raw[4:8]))
			rowID = 0
		}
		return db.readLvalChain(pageNum, rowID, memoLen)

	default:
		// Unknown bitmask — try treating as inline.
		if memoLen > 0 && memoLen <= len(raw)-12 {
			data := raw[12 : 12+memoLen]
			result := make([]byte, len(data))
			copy(result, data)
			return result, nil
		}
		return nil, fmt.Errorf("mdb: unknown LVAL bitmask %#x", bitmask)
	}
}

// readLvalRecord reads a single record from an LVAL page.
func (db *Database) readLvalRecord(pageNum int64, rowID int, maxLen int) ([]byte, error) {
	page, err := db.ReadPage(pageNum)
	if err != nil {
		return nil, fmt.Errorf("mdb: LVAL page %d: %w", pageNum, err)
	}

	// Verify LVAL signature at offset 4.
	if page[0] != PageTypeData {
		return nil, fmt.Errorf("mdb: page %d type %#x is not data/LVAL", pageNum, page[0])
	}

	// LVAL pages use the same row structure as data pages.
	numRows := int(binary.LittleEndian.Uint16(page[dataNumRows:]))
	if rowID >= numRows {
		return nil, fmt.Errorf("mdb: LVAL page %d row %d out of range (%d rows)", pageNum, rowID, numRows)
	}

	rowOff := dataRowTable + rowID*2
	offVal := binary.LittleEndian.Uint16(page[rowOff:])
	offset := int(offVal & rowOffsetMask)

	var rowEnd int
	if rowID == 0 {
		rowEnd = PageSize
	} else {
		prevOff := binary.LittleEndian.Uint16(page[rowOff-2:])
		rowEnd = int(prevOff & rowOffsetMask)
	}

	if offset >= rowEnd || rowEnd > PageSize {
		return nil, fmt.Errorf("mdb: LVAL page %d row %d invalid bounds", pageNum, rowID)
	}

	data := page[offset:rowEnd]
	if len(data) > maxLen && maxLen > 0 {
		data = data[:maxLen]
	}

	result := make([]byte, len(data))
	copy(result, data)
	return result, nil
}

// ReadLvalChain reads an LVAL multi-page chain starting at the given page and row.
// maxLen caps the amount of data returned; pass 0 for no limit.
func (db *Database) ReadLvalChain(pageNum int64, rowID int, maxLen int) ([]byte, error) {
	effective := maxLen
	if effective <= 0 {
		effective = 1<<31 - 1 // 2GB cap
	}
	return db.readLvalChain(pageNum, rowID, effective)
}

// readLvalChain reads a multi-page LVAL chain.
func (db *Database) readLvalChain(pageNum int64, rowID int, totalLen int) ([]byte, error) {
	var result []byte
	currentPage := pageNum
	currentRow := rowID

	for currentPage != 0 && len(result) < totalLen {
		page, err := db.ReadPage(currentPage)
		if err != nil {
			return nil, fmt.Errorf("mdb: LVAL chain page %d: %w", currentPage, err)
		}

		numRows := int(binary.LittleEndian.Uint16(page[dataNumRows:]))
		if currentRow >= numRows {
			break
		}

		rowOff := dataRowTable + currentRow*2
		offVal := binary.LittleEndian.Uint16(page[rowOff:])
		offset := int(offVal & rowOffsetMask)

		var rowEnd int
		if currentRow == 0 {
			rowEnd = PageSize
		} else {
			prevOff := binary.LittleEndian.Uint16(page[rowOff-2:])
			rowEnd = int(prevOff & rowOffsetMask)
		}

		if offset >= rowEnd || rowEnd > PageSize {
			break
		}

		recordData := page[offset:rowEnd]

		// The first 4 bytes encode the next record pointer using the same
		// format as the initial LVAL reference: (pageNum << 8) | rowID.
		// Zero means end of chain.
		if len(recordData) < 4 {
			break
		}
		nextPtr := binary.LittleEndian.Uint32(recordData[0:4])
		nextPage := int64(nextPtr >> 8)
		nextRow := int(nextPtr & 0xFF)
		chunk := recordData[4:]

		result = append(result, chunk...)

		currentPage = nextPage
		currentRow = nextRow
	}

	if len(result) > totalLen && totalLen > 0 {
		result = result[:totalLen]
	}

	return result, nil
}
