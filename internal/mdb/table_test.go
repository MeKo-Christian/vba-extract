package mdb

import (
	"testing"
)

func TestReadTableDef_MSysObjects(t *testing.T) {
	db := testDB(t)

	td, err := db.ReadTableDef(2) // MSysObjects is always at page 2
	if err != nil {
		t.Fatalf("ReadTableDef(2): %v", err)
	}

	if td.TableType != TableTypeSystem {
		t.Errorf("TableType = %#x, want %#x (system)", td.TableType, TableTypeSystem)
	}

	if len(td.Columns) == 0 {
		t.Fatal("no columns parsed")
	}

	// MSysObjects should have well-known columns.
	wantCols := map[string]byte{
		"Id":         ColTypeLong,
		"Name":       ColTypeText,
		"Type":       ColTypeInt,
		"DateCreate": ColTypeDatetime,
		"DateUpdate": ColTypeDatetime,
		"Flags":      ColTypeLong,
		"ParentId":   ColTypeLong,
	}

	found := make(map[string]bool)
	for _, col := range td.Columns {
		if wantType, ok := wantCols[col.Name]; ok {
			found[col.Name] = true
			if col.Type != wantType {
				t.Errorf("column %q type = %s, want %s", col.Name, ColTypeName(col.Type), ColTypeName(wantType))
			}
		}
	}

	for name := range wantCols {
		if !found[name] {
			t.Errorf("expected column %q not found", name)
		}
	}

	t.Logf("MSysObjects: %d rows, %d columns:", td.NumRows, len(td.Columns))
	for _, col := range td.Columns {
		t.Logf("  %-20s %-10s len=%-5d fixed=%v", col.Name, ColTypeName(col.Type), col.Length, col.IsFixed())
	}
}
