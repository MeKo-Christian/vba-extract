package mdb

import (
	"encoding/binary"
	"errors"
	"fmt"
	"math"
	"sort"
)

// Data page header offsets (Jet4).
const (
	dataFreeSpace = 0x02
	dataTDefPage  = 0x04
	dataNumRows   = 0x0C
	dataRowTable  = 0x0E // row offset table starts here (Jet4: 2 bytes per entry)

	// Row offset flags.
	rowDeleteFlag = 0x8000
	rowLookupFlag = 0x4000
	rowOffsetMask = 0x1FFF
)

// Row represents a single parsed table row. Values are keyed by column name.
type Row map[string]any

// DataPages returns the list of data page numbers for this table.
// It scans all database pages to find data pages belonging to this table.
func (td *TableDef) DataPages() ([]int64, error) {
	var pages []int64
	buf := make([]byte, 8) // only need first few bytes per page

	for i := int64(1); i < td.db.pageCount; i++ {
		_, err := td.db.f.ReadAt(buf, i*td.db.pageSize)
		if err != nil {
			continue
		}

		if buf[0] != PageTypeData {
			continue
		}

		tdefPg := binary.LittleEndian.Uint32(buf[dataTDefPage:])
		if int64(tdefPg) == td.DefPage {
			pages = append(pages, i)
		}
	}

	return pages, nil
}

// ReadRows reads all non-deleted rows from this table.
func (td *TableDef) ReadRows() ([]Row, error) {
	dataPages, err := td.DataPages()
	if err != nil {
		return nil, fmt.Errorf("mdb: ReadRows: %w", err)
	}

	// Sort columns by ColNum for proper ordering.
	sortedCols := make([]*Column, len(td.Columns))
	copy(sortedCols, td.Columns)
	sort.Slice(sortedCols, func(i, j int) bool {
		return sortedCols[i].ColNum < sortedCols[j].ColNum
	})

	var rows []Row

	for _, pageNum := range dataPages {
		page, err := td.db.ReadPage(pageNum)
		if err != nil {
			return nil, fmt.Errorf("mdb: ReadRows page %d: %w", pageNum, err)
		}

		if PageType(page) != PageTypeData {
			continue
		}

		// Verify this data page belongs to our table.
		pageTDEF := binary.LittleEndian.Uint32(page[dataTDefPage:])
		if int64(pageTDEF) != td.DefPage {
			continue
		}

		numRows := int(binary.LittleEndian.Uint16(page[dataNumRows:]))

		for rowIdx := range numRows {
			rowOff := dataRowTable + rowIdx*2
			if rowOff+2 > len(page) {
				break
			}

			offVal := binary.LittleEndian.Uint16(page[rowOff:])

			// Check delete/lookup flags.
			if offVal&rowDeleteFlag != 0 {
				continue
			}

			if offVal&rowLookupFlag != 0 {
				// Overflow row pointer — skip for now.
				continue
			}

			offset := int(offVal & rowOffsetMask)
			if offset >= len(page) {
				continue
			}

			// Determine row end: previous row's start offset, or page end for first row.
			var rowEnd int
			if rowIdx == 0 {
				rowEnd = len(page)
			} else {
				prevOff := binary.LittleEndian.Uint16(page[rowOff-2:])
				rowEnd = int(prevOff & rowOffsetMask)
			}

			if offset >= rowEnd || rowEnd > len(page) {
				continue
			}

			rowData := page[offset:rowEnd]

			row, err := td.parseRow(rowData, sortedCols)
			if err != nil {
				continue // skip malformed rows
			}

			rows = append(rows, row)
		}
	}

	return rows, nil
}

// parseRow parses a single row from raw row bytes (Jet4 format).
//
// Jet4 row layout:
//
//	START: [num_cols (2)] [fixed_data] [var_col_data...]
//	END:   [...] [var_offset_table ((nvc+1)*2)] [num_var_cols (2)] [null_mask (ceil(num_cols/8))]
func (td *TableDef) parseRow(data []byte, sortedCols []*Column) (Row, error) {
	if td.db != nil && td.db.IsJet3() {
		return td.parseRowJet3(data, sortedCols)
	}

	if len(data) < 4 {
		return nil, fmt.Errorf("mdb: row too short (%d bytes)", len(data))
	}

	// num_cols at the START of the row (2 bytes, Jet4).
	numCols := int(binary.LittleEndian.Uint16(data[0:2]))
	if numCols <= 0 {
		return nil, errors.New("mdb: row has 0 columns")
	}

	// null_mask at the END of the row (ceil(numCols/8) bytes).
	nullMaskLen := (numCols + 7) / 8

	pos := len(data) - nullMaskLen
	if pos < 2 {
		return nil, errors.New("mdb: row too short for null mask")
	}

	nullMask := data[pos : pos+nullMaskLen]

	// num_var_cols (2 bytes) immediately before null_mask.
	pos -= 2
	if pos < 2 {
		return nil, errors.New("mdb: row too short for num_var_cols")
	}

	numVarCols := int(binary.LittleEndian.Uint16(data[pos:]))

	// var_offset_table: (numVarCols+1) entries of 2 bytes each, before num_var_cols.
	// Stored in reverse order: last entry = offset of first var column start.
	numVarOffsets := numVarCols + 1

	pos -= numVarOffsets * 2
	if pos < 2 {
		return nil, errors.New("mdb: row too short for var_offset_table")
	}
	// Read the offset table and reverse it so index 0 = first var column boundary.
	varOffsets := make([]int, numVarOffsets)
	for i := range numVarOffsets {
		raw := binary.LittleEndian.Uint16(data[pos+i*2:])
		varOffsets[numVarOffsets-1-i] = int(raw)
	}

	row := make(Row)

	for _, col := range sortedCols {
		colIdx := int(col.ColNum)

		// Check null bit. Bit=1 means NOT NULL.
		if colIdx < numCols {
			byteIdx := colIdx / 8

			bitMask := byte(1 << (colIdx % 8))
			if byteIdx < len(nullMask) && nullMask[byteIdx]&bitMask == 0 {
				row[col.Name] = nil
				continue
			}
		}

		if col.IsFixed() {
			// Fixed columns: offset is relative to after num_cols (2 bytes).
			val := readFixedColumn(data, col, 2)
			row[col.Name] = val
		} else {
			val := readVarColumn(data, col, varOffsets, numVarCols)
			row[col.Name] = val
		}
	}

	return row, nil
}

func (td *TableDef) parseRowJet3(_ []byte, _ []*Column) (Row, error) {
	return nil, ErrJet3RowLayoutUnsupported
}

// readFixedColumn reads a fixed-length column value from row data.
// baseOff is the offset where fixed data begins (2 for Jet4, after num_cols).
func readFixedColumn(data []byte, col *Column, baseOff int) any {
	off := baseOff + int(col.OffsetFix)
	if off >= len(data) {
		return nil
	}

	switch col.Type {
	case ColTypeBool:
		return data[off] != 0
	case ColTypeByte:
		return data[off]
	case ColTypeInt:
		if off+2 > len(data) {
			return nil
		}

		return readLEInt16(data[off], data[off+1])
	case ColTypeLong:
		if off+4 > len(data) {
			return nil
		}

		return readLEInt32(data[off : off+4])
	case ColTypeFloat:
		if off+4 > len(data) {
			return nil
		}

		return math.Float32frombits(binary.LittleEndian.Uint32(data[off:]))
	case ColTypeDouble, ColTypeDatetime, ColTypeMoney:
		if off+8 > len(data) {
			return nil
		}

		return math.Float64frombits(binary.LittleEndian.Uint64(data[off:]))
	default:
		end := min(off+int(col.Length), len(data))

		result := make([]byte, end-off)
		copy(result, data[off:end])

		return result
	}
}

// readVarColumn reads a variable-length column value from row data.
// varOffsets is the reversed offset table (index 0 = first var column boundary).
// numVarCols is the number of variable columns stored in this row.
func readVarColumn(data []byte, col *Column, varOffsets []int, numVarCols int) any {
	idx := int(col.OffsetVar)
	if idx >= numVarCols || idx+1 >= len(varOffsets) {
		return nil
	}

	start := varOffsets[idx]
	end := varOffsets[idx+1]

	if start >= end || start >= len(data) || end > len(data) {
		return nil
	}

	raw := data[start:end]

	switch col.Type {
	case ColTypeText:
		return decodeUCS2(raw)
	case ColTypeBool:
		if len(raw) >= 1 {
			return raw[0] != 0
		}

		return nil
	case ColTypeByte:
		if len(raw) >= 1 {
			return raw[0]
		}

		return nil
	case ColTypeInt:
		if len(raw) >= 2 {
			return readLEInt16(raw[0], raw[1])
		}

		return nil
	case ColTypeLong:
		if len(raw) >= 4 {
			return readLEInt32(raw)
		}

		return nil
	case ColTypeFloat:
		if len(raw) >= 4 {
			return math.Float32frombits(binary.LittleEndian.Uint32(raw))
		}

		return nil
	case ColTypeDouble, ColTypeDatetime, ColTypeMoney:
		if len(raw) >= 8 {
			return math.Float64frombits(binary.LittleEndian.Uint64(raw))
		}

		return nil
	case ColTypeMemo, ColTypeOLE:
		result := make([]byte, len(raw))
		copy(result, raw)

		return result
	default:
		result := make([]byte, len(raw))
		copy(result, raw)

		return result
	}
}

func readLEInt16(b0, b1 byte) int16 {
	return int16(b0) | int16(b1)<<8
}

func readLEInt32(raw []byte) int32 {
	return int32(raw[0]) | int32(raw[1])<<8 | int32(raw[2])<<16 | int32(raw[3])<<24
}
