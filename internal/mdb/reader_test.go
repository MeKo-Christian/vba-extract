package mdb

import (
	"os"
	"testing"
)

const testMDB = "../../testdata/Start.mdb"

func testDB(t *testing.T) *Database {
	t.Helper()
	if _, err := os.Stat(testMDB); os.IsNotExist(err) {
		t.Skip("testdata/Start.mdb not available")
	}
	db, err := Open(testMDB)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	t.Cleanup(func() { db.Close() })
	return db
}

func TestOpenAndHeader(t *testing.T) {
	db := testDB(t)

	if db.Header.DBName != "Standard ACE DB" {
		t.Errorf("DBName = %q, want %q", db.Header.DBName, "Standard ACE DB")
	}

	// Start.mdb is Access 2007+ format (version 3 = ACE12).
	if db.Header.JetVersion != JetVersionACE {
		t.Errorf("JetVersion = %d, want %d (ACE12)", db.Header.JetVersion, JetVersionACE)
	}

	if !db.IsJet4() {
		t.Error("IsJet4() = false, want true")
	}

	// 907 pages (3715072 / 4096).
	if db.PageCount() != 907 {
		t.Errorf("PageCount = %d, want 907", db.PageCount())
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
	tmp.Write([]byte("tiny"))
	tmp.Close()

	_, err = Open(tmp.Name())
	if err == nil {
		t.Error("Open tiny file should fail")
	}
}
