package cmd

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/MeKo-Christian/accessdump/internal/mdb"
)

// writeSchema

func TestWriteSchema_createsFiles(t *testing.T) {
	dir := t.TempDir()
	s := &mdb.Schema{
		Tables: []mdb.TableSchema{
			{Name: "Users", Columns: []mdb.ColumnDef{{Name: "ID", SQLType: "INTEGER"}}},
		},
	}

	err := writeSchema(dir, filepath.Join("/data", "mydb.accdb"), s, false, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	sqlFile := filepath.Join(dir, "mydb", "mydb.schema.sql")
	mdFile := filepath.Join(dir, "mydb", "mydb.schema.md")

	_, err = os.Stat(sqlFile)
	if err != nil {
		t.Errorf("expected .sql file: %v", err)
	}

	_, err = os.Stat(mdFile)
	if err != nil {
		t.Errorf("expected .md file: %v", err)
	}
}

func TestWriteSchema_flatMode(t *testing.T) {
	dir := t.TempDir()
	s := &mdb.Schema{}

	err := writeSchema(dir, "/data/mydb.mdb", s, true, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// In flat mode files go directly into dir, not a subdirectory
	sqlFile := filepath.Join(dir, "mydb.schema.sql")

	_, err = os.Stat(sqlFile)
	if err != nil {
		t.Errorf("expected .sql directly in base dir: %v", err)
	}
}

func TestWriteSchema_sqlFileContainsDDL(t *testing.T) {
	dir := t.TempDir()
	s := &mdb.Schema{
		Tables: []mdb.TableSchema{
			{Name: "Products", Columns: []mdb.ColumnDef{{Name: "Name", SQLType: "VARCHAR(100)"}}},
		},
	}

	err := writeSchema(dir, "/db/shop.mdb", s, true, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	data, err := os.ReadFile(filepath.Join(dir, "shop.schema.sql"))
	if err != nil {
		t.Fatalf("reading sql file: %v", err)
	}

	if !strings.Contains(string(data), "CREATE TABLE") {
		t.Errorf("expected CREATE TABLE in sql file, got: %q", string(data))
	}
}

func TestWriteSchema_mdFileContainsMarkdown(t *testing.T) {
	dir := t.TempDir()
	s := &mdb.Schema{
		Tables: []mdb.TableSchema{
			{Name: "Orders", Columns: []mdb.ColumnDef{{Name: "ID", SQLType: "INTEGER"}}},
		},
	}

	err := writeSchema(dir, "/db/orders.mdb", s, true, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	data, err := os.ReadFile(filepath.Join(dir, "orders.schema.md"))
	if err != nil {
		t.Fatalf("reading md file: %v", err)
	}

	if !strings.Contains(string(data), "## Tables") {
		t.Errorf("expected ## Tables in md file, got: %q", string(data))
	}
}
