package mdb

import (
	"testing"
)

func TestCatalog(t *testing.T) {
	db := testDB(t)

	entries, err := db.Catalog()
	if err != nil {
		t.Fatalf("Catalog: %v", err)
	}

	if len(entries) == 0 {
		t.Fatal("no catalog entries")
	}

	t.Logf("Catalog has %d entries", len(entries))

	// Look for MSysAccessStorage.
	found := false
	for _, e := range entries {
		if e.Name == "MSysAccessStorage" {
			found = true
			t.Logf("  MSysAccessStorage: ID=%d, Type=%d", e.ID, e.Type)
		}
	}
	if !found {
		t.Error("MSysAccessStorage not found in catalog")
	}
}

func TestTableNames(t *testing.T) {
	db := testDB(t)

	names, err := db.TableNames()
	if err != nil {
		t.Fatalf("TableNames: %v", err)
	}

	t.Logf("Tables: %v", names)

	// Start.mdb should have these user tables.
	wantTables := []string{"Module", "SYStabEinstellungen"}
	for _, want := range wantTables {
		found := false
		for _, name := range names {
			if name == want {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("expected table %q not found", want)
		}
	}
}

func TestFindTable(t *testing.T) {
	db := testDB(t)

	td, err := db.FindTable("MSysAccessStorage")
	if err != nil {
		t.Fatalf("FindTable: %v", err)
	}

	t.Logf("MSysAccessStorage: %d rows, %d columns", td.NumRows, len(td.Columns))
	for _, col := range td.Columns {
		t.Logf("  %-20s %-10s len=%-5d fixed=%v", col.Name, ColTypeName(col.Type), col.Length, col.IsFixed())
	}
}

func TestReadMSysAccessStorage(t *testing.T) {
	db := testDB(t)

	td, err := db.FindTable("MSysAccessStorage")
	if err != nil {
		t.Fatalf("FindTable: %v", err)
	}

	rows, err := td.ReadRows()
	if err != nil {
		t.Fatalf("ReadRows: %v", err)
	}

	t.Logf("MSysAccessStorage: read %d rows", len(rows))

	// Look for VBA-related entries.
	for _, row := range rows {
		name, _ := row["Name"].(string)
		id, _ := row["Id"].(int32)
		parentId, _ := row["ParentId"].(int32)
		typ, _ := row["Type"].(int32)
		if name == "VBAProject" || name == "VBA" || name == "PROJECT" || name == "dir" {
			t.Logf("  Id=%d ParentId=%d Name=%q Type=%d", id, parentId, name, typ)
		}
	}
}
