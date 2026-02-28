package cmd

import (
	"os"
	"testing"
)

const testMDB = "../testdata/Start.mdb"

func skipIfNoFixture(t *testing.T) {
	t.Helper()

	if _, err := os.Stat(testMDB); os.IsNotExist(err) {
		t.Skip("testdata/Start.mdb not available")
	}
}

// loadModules

func TestLoadModules_returnsModules(t *testing.T) {
	skipIfNoFixture(t)

	modules, err := loadModules(testMDB)
	if err != nil {
		t.Fatalf("loadModules: %v", err)
	}

	if len(modules) == 0 {
		t.Fatal("expected at least one module from Start.mdb")
	}
}

func TestLoadModules_missingFile(t *testing.T) {
	_, err := loadModules("/nonexistent/path.mdb")
	if err == nil {
		t.Error("expected error for missing file")
	}
}

// loadSchema

func TestLoadSchema_returnsSchema(t *testing.T) {
	skipIfNoFixture(t)

	schema, err := loadSchema(testMDB)
	if err != nil {
		t.Fatalf("loadSchema: %v", err)
	}

	if schema == nil {
		t.Fatal("expected non-nil schema")
	}

	if len(schema.Tables) == 0 {
		t.Fatal("expected at least one table in schema")
	}
}

func TestLoadSchema_missingFile(t *testing.T) {
	_, err := loadSchema("/nonexistent/path.mdb")
	if err == nil {
		t.Error("expected error for missing file")
	}
}

func TestLoadSchema_legacyFixture(t *testing.T) {
	path := "../testdata/jet35/st990426.mdb"
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Skip("legacy fixture not available")
	}

	schema, err := loadSchema(path)
	if err != nil {
		t.Fatalf("loadSchema legacy fixture: %v", err)
	}

	if schema == nil {
		t.Fatal("expected non-nil schema from legacy fixture")
	}

	if len(schema.Tables) == 0 {
		t.Fatal("expected at least one table in legacy fixture schema")
	}
}
