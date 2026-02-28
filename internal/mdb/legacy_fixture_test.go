package mdb

import (
	"os"
	"path/filepath"
	"testing"
)

func TestLegacyFixtureFromSQLServerBackup(t *testing.T) {
	path := filepath.Join("..", "..", "testdata", "jet35", "st990426.mdb")
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Skipf("legacy fixture not present: %s", path)
	}

	db, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer db.Close()

	// This fixture is an older-format production backup and is used as a
	// regression file for version/page-size detection.
	if db.PageSize() != PageSizeJet4 {
		t.Fatalf("PageSize = %d, want %d", db.PageSize(), PageSizeJet4)
	}

	names, err := db.TableNames()
	if err != nil {
		t.Fatalf("TableNames: %v", err)
	}

	if len(names) == 0 {
		t.Fatal("expected at least one table in legacy fixture")
	}
}
