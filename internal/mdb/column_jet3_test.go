package mdb

import (
	"encoding/binary"
	"os"
	"testing"
)

func writeSyntheticJet3WithOneDataRow(t *testing.T) string {
	t.Helper()

	path := writeSyntheticJet3WithTDEF(t)

	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}

	pageStart := 1 * 2048
	page := data[pageStart : pageStart+2048]
	page[0] = PageTypeData
	binary.LittleEndian.PutUint32(page[dataTDefPage:], 2)
	binary.LittleEndian.PutUint16(page[dataNumRowsJet3:], 1)

	// Jet3 row with:
	// - rowCols=2
	// - fixed Id=42
	// - var Name=\"AB\"
	// - var offsets [5,7]
	// - rowVarCols=1
	// - null mask: both non-null
	row := []byte{
		2,           // rowCols
		42, 0, 0, 0, // fixed long Id
		'A', 'B', // variable text bytes
		7,    // varOffsets[1] end/eod
		5,    // varOffsets[0] start
		1,    // rowVarCols
		0x03, // null mask bits for col0,col1
	}

	rowStart := len(page) - len(row)
	copy(page[rowStart:], row)
	binary.LittleEndian.PutUint16(page[dataRowTableJet3:], uint16(rowStart))

	err = os.WriteFile(path, data, 0o600)
	if err != nil {
		t.Fatalf("WriteFile: %v", err)
	}

	return path
}

func TestReadRowsJet3Synthetic(t *testing.T) {
	path := writeSyntheticJet3WithOneDataRow(t)

	db, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	t.Cleanup(func() { _ = db.Close() })

	if !db.IsJet3() {
		t.Fatal("expected Jet3 page layout")
	}

	td, err := db.ReadTableDef(2)
	if err != nil {
		t.Fatalf("ReadTableDef: %v", err)
	}

	rows, err := td.ReadRows()
	if err != nil {
		t.Fatalf("ReadRows: %v", err)
	}

	if len(rows) != 1 {
		t.Fatalf("rows = %d, want 1", len(rows))
	}

	id, ok := rows[0]["Id"].(int32)
	if !ok || id != 42 {
		t.Fatalf("Id = %#v, want int32(42)", rows[0]["Id"])
	}

	name, ok := rows[0]["Name"].(string)
	if !ok || name != "AB" {
		t.Fatalf("Name = %#v, want %q", rows[0]["Name"], "AB")
	}
}
