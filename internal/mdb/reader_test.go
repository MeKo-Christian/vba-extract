package mdb

import (
	"os"
	"testing"
)

const testMDB = "../../testdata/sample.mdb"

func testDB(t *testing.T) *Database {
	t.Helper()

	_, err := os.Stat(testMDB)
	if os.IsNotExist(err) {
		t.Skip("testdata/sample.mdb not available")
	}

	db, err := Open(testMDB)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	t.Cleanup(func() { db.Close() })

	return db
}

const startMDB = "../../testdata/Start.mdb"

func startDB(t *testing.T) *Database {
	t.Helper()
	_, err := os.Stat(startMDB)
	if os.IsNotExist(err) {
		t.Skip("testdata/Start.mdb not available (proprietary fixture)")
	}
	db, err := Open(startMDB)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	t.Cleanup(func() { db.Close() })
	return db
}

func TestOpenAndHeader(t *testing.T) {
	db := testDB(t)

	if !db.IsJet4() {
		t.Error("IsJet4() = false, want true")
	}

	// sample.mdb: 71 pages (290816 / 4096).
	if db.PageCount() != 71 {
		t.Errorf("PageCount = %d, want 71", db.PageCount())
	}
}

func TestReadPage(t *testing.T) {
	db := testDB(t)

	// Page 0 should be DB definition page.
	page0, err := db.ReadPage(0)
	if err != nil {
		t.Fatalf("ReadPage(0): %v", err)
	}

	if PageType(page0) != PageTypeDB {
		t.Errorf("Page 0 type = %#x, want %#x", PageType(page0), PageTypeDB)
	}

	// Page 2 should be TDEF (MSysObjects).
	page2, err := db.ReadPage(2)
	if err != nil {
		t.Fatalf("ReadPage(2): %v", err)
	}

	if PageType(page2) != PageTypeTDEF {
		t.Errorf("Page 2 type = %#x, want %#x (TDEF)", PageType(page2), PageTypeTDEF)
	}

	// Out of range.
	_, err = db.ReadPage(10000)
	if err == nil {
		t.Error("ReadPage(10000) should fail")
	}
}

func TestOpenInvalidFile(t *testing.T) {
	// Non-existent file.
	_, err := Open("/tmp/nonexistent.mdb")
	if err == nil {
		t.Error("Open non-existent should fail")
	}

	// Create a too-small file.
	tmp, err := os.CreateTemp("", "tiny*.mdb")
	if err != nil {
		t.Fatal(err)
	}
	defer os.Remove(tmp.Name())

	_, err = tmp.WriteString("tiny")
	if err != nil {
		t.Fatal(err)
	}

	err = tmp.Close()
	if err != nil {
		t.Fatal(err)
	}

	_, err = Open(tmp.Name())
	if err == nil {
		t.Error("Open tiny file should fail")
	}
}
