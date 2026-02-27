package mdb

import (
	"testing"
)

func TestReadRows_MSysObjects(t *testing.T) {
	db := testDB(t)

	td, err := db.ReadTableDef(2) // MSysObjects
	if err != nil {
		t.Fatalf("ReadTableDef: %v", err)
	}

	rows, err := td.ReadRows()
	if err != nil {
		t.Fatalf("ReadRows: %v", err)
	}

	if len(rows) == 0 {
		t.Fatal("no rows read from MSysObjects")
	}

	t.Logf("Read %d rows from MSysObjects (expected ~%d)", len(rows), td.NumRows)

	// Check that we can find known tables.
	wantTables := map[string]bool{
		"MSysAccessStorage": false,
		"MSysObjects":       false,
	}

	for _, row := range rows {
		name, _ := row["Name"].(string)
		if _, ok := wantTables[name]; ok {
			wantTables[name] = true
			typ, _ := row["Type"].(int16)
			id, _ := row["Id"].(int32)
			t.Logf("  Found %q: Type=%d, Id=%d", name, typ, id)
		}
	}

	for name, found := range wantTables {
		if !found {
			t.Errorf("table %q not found in MSysObjects", name)
		}
	}
}

func TestDataPages_MSysObjects(t *testing.T) {
	db := testDB(t)

	td, err := db.ReadTableDef(2)
	if err != nil {
		t.Fatalf("ReadTableDef: %v", err)
	}

	pages, err := td.DataPages()
	if err != nil {
		t.Fatalf("DataPages: %v", err)
	}

	if len(pages) == 0 {
		t.Fatal("no data pages found for MSysObjects")
	}

	t.Logf("MSysObjects has %d data pages: %v", len(pages), pages)

	// Verify pages are data pages.
	for _, p := range pages {
		page, err := db.ReadPage(p)
		if err != nil {
			t.Errorf("ReadPage(%d): %v", p, err)
			continue
		}
		if PageType(page) != PageTypeData {
			t.Errorf("page %d type = %#x, want data (%#x)", p, PageType(page), PageTypeData)
		}
	}
}
