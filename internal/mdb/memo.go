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
	if db.IsJet3() {
		return db.resolveMemoJet3(raw)
	}

	return db.resolveMemoWithLayout(raw, dataNumRows, dataRowTable)
}

func (db *Database) resolveMemoJet3(raw []byte) ([]byte, error) {
	return db.resolveMemoWithLayout(raw, dataNumRowsJet3, dataRowTableJet3)
}

func (db *Database) resolveMemoWithLayout(raw []byte, numRowsOff, rowTableOff int) ([]byte, error) {
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
	pageNum, rowID := decodeLvalPageAndRow(raw[4:8])

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
		return db.readLvalRecordWithLayout(pageNum, rowID, memoLen, numRowsOff, rowTableOff)

	case LvalMultiPage:
		// Data spans multiple LVAL pages.
		return db.readLvalChainWithLayout(pageNum, rowID, memoLen, numRowsOff, rowTableOff)

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

func decodeLvalPageAndRow(raw []byte) (int64, int) {
	lvalPage := binary.LittleEndian.Uint32(raw)
	rowID := int(lvalPage & 0xFF) // low byte = row ID

	pageNum := int64(lvalPage >> 8) // upper 3 bytes = page number
	if pageNum == 0 {
		// Alternative encoding: full uint32 page, separate row
		pageNum = int64(binary.LittleEndian.Uint32(raw))
		rowID = 0
	}

	return pageNum, rowID
}

func (db *Database) readLvalRecordWithLayout(pageNum int64, rowID int, maxLen int, numRowsOff, rowTableOff int) ([]byte, error) {
	page, err := db.ReadPage(pageNum)
	if err != nil {
		return nil, fmt.Errorf("mdb: LVAL page %d: %w", pageNum, err)
	}

	// Verify LVAL signature at offset 4.
	if page[0] != PageTypeData {
		return nil, fmt.Errorf("mdb: page %d type %#x is not data/LVAL", pageNum, page[0])
	}

	// LVAL pages use the same row structure as data pages.
	numRows := int(binary.LittleEndian.Uint16(page[numRowsOff:]))
	if rowID >= numRows {
		return nil, fmt.Errorf("mdb: LVAL page %d row %d out of range (%d rows)", pageNum, rowID, numRows)
	}

	rowOff := rowTableOff + rowID*2
	if rowOff+2 > len(page) {
		return nil, fmt.Errorf("mdb: LVAL page %d row %d offset table out of range", pageNum, rowID)
	}

	offVal := binary.LittleEndian.Uint16(page[rowOff:])
	offset := int(offVal & rowOffsetMask)

	var rowEnd int
	if rowID == 0 {
		rowEnd = len(page)
	} else {
		prevOff := binary.LittleEndian.Uint16(page[rowOff-2:])
		rowEnd = int(prevOff & rowOffsetMask)
	}

	if offset >= rowEnd || rowEnd > len(page) {
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

	numRowsOff, rowTableOff := dataPageLayoutOffsets(db)

	return db.readLvalChainWithLayout(pageNum, rowID, effective, numRowsOff, rowTableOff)
}

func (db *Database) readLvalChainWithLayout(pageNum int64, rowID int, totalLen int, numRowsOff, rowTableOff int) ([]byte, error) {
	var result []byte
	currentPage := pageNum
	currentRow := rowID

	for currentPage != 0 && len(result) < totalLen {
		page, err := db.ReadPage(currentPage)
		if err != nil {
			return nil, fmt.Errorf("mdb: LVAL chain page %d: %w", currentPage, err)
		}

		if page[0] != PageTypeData {
			return nil, fmt.Errorf("mdb: LVAL chain page %d type %#x is not data/LVAL", currentPage, page[0])
		}

		numRows := int(binary.LittleEndian.Uint16(page[numRowsOff:]))
		if numRows == 0 {
			return nil, fmt.Errorf("mdb: LVAL chain page %d has no rows", currentPage)
		}

		if currentRow >= numRows {
			return nil, fmt.Errorf("mdb: LVAL chain page %d row %d out of range (%d rows)", currentPage, currentRow, numRows)
		}

		rowOff := rowTableOff + currentRow*2
		if rowOff+2 > len(page) {
			return nil, fmt.Errorf("mdb: LVAL chain page %d row %d offset table out of range", currentPage, currentRow)
		}

		offVal := binary.LittleEndian.Uint16(page[rowOff:])
		offset := int(offVal & rowOffsetMask)

		var rowEnd int
		if currentRow == 0 {
			rowEnd = len(page)
		} else {
			prevOff := binary.LittleEndian.Uint16(page[rowOff-2:])
			rowEnd = int(prevOff & rowOffsetMask)
		}

		if offset >= rowEnd || rowEnd > len(page) {
			return nil, fmt.Errorf("mdb: LVAL chain page %d row %d invalid bounds", currentPage, currentRow)
		}

		recordData := page[offset:rowEnd]

		// The first 4 bytes encode the next record pointer using the same
		// format as the initial LVAL reference: (pageNum << 8) | rowID.
		// Zero means end of chain.
		if len(recordData) < 4 {
			return nil, fmt.Errorf("mdb: LVAL chain page %d row %d record too short", currentPage, currentRow)
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
