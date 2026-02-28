package mdb

import (
	"path/filepath"
	"strings"
	"testing"
)

func TestReadSchemaFromStartMDB(t *testing.T) {
	db, err := Open(filepath.Join("..", "..", "testdata", "Start.mdb"))
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer db.Close()

	s, err := db.ReadSchema()
	if err != nil {
		t.Fatalf("ReadSchema: %v", err)
	}

	if len(s.Tables) == 0 {
		t.Fatal("expected at least one user table")
	}

	// No MSys* tables should appear.
	for _, table := range s.Tables {
		if strings.HasPrefix(table.Name, "MSys") {
			t.Errorf("system table leaked into schema: %q", table.Name)
		}

		if len(table.Columns) == 0 {
			t.Errorf("table %q has no columns", table.Name)
		}
	}

	// Every column must have a non-empty SQL type.
	for _, table := range s.Tables {
		for _, col := range table.Columns {
			if col.SQLType == "" {
				t.Errorf("table %q column %q: empty SQLType", table.Name, col.Name)
			}

			if col.Name == "" {
				t.Errorf("table %q: column with empty name", table.Name)
			}
		}
	}

	// Relationships must have non-empty table names.
	for _, rel := range s.Relationships {
		if rel.FromTable == "" || rel.ToTable == "" {
			t.Errorf("relationship %q: missing table name", rel.Name)
		}

		if len(rel.FromColumns) == 0 || len(rel.ToColumns) == 0 {
			t.Errorf("relationship %q: missing column list", rel.Name)
		}
	}

	// Query names must be non-empty.
	for _, q := range s.Queries {
		if q.Name == "" {
			t.Error("query with empty name")
		}
	}
}

func TestJetTypeToSQL(t *testing.T) {
	cases := []struct {
		typ       byte
		length    int
		scale     byte
		precision byte
		want      string
	}{
		{ColTypeBool, 0, 0, 0, "BOOLEAN"},
		{ColTypeByte, 0, 0, 0, "TINYINT"},
		{ColTypeInt, 0, 0, 0, "SMALLINT"},
		{ColTypeLong, 0, 0, 0, "INTEGER"},
		{ColTypeMoney, 0, 0, 0, "DECIMAL(19,4)"},
		{ColTypeFloat, 0, 0, 0, "REAL"},
		{ColTypeDouble, 0, 0, 0, "DOUBLE PRECISION"},
		{ColTypeDatetime, 0, 0, 0, "DATETIME"},
		{ColTypeText, 100, 0, 0, "VARCHAR(50)"}, // 100 bytes / 2 = 50 chars
		{ColTypeText, 0, 0, 0, "VARCHAR(255)"},  // zero length → default 255
		{ColTypeMemo, 0, 0, 0, "TEXT"},
		{ColTypeOLE, 0, 0, 0, "OLE"},
		{ColTypeGUID, 0, 0, 0, "CHAR(38)"},
		{ColTypeNumeric, 0, 2, 10, "DECIMAL(10,2)"},
		{ColTypeNumeric, 0, 0, 0, "DECIMAL"},
		{0xFF, 0, 0, 0, "TYPE_0xFF"},
	}

	for _, tc := range cases {
		got := jetTypeToSQL(tc.typ, tc.length, tc.scale, tc.precision)
		if got != tc.want {
			t.Errorf("jetTypeToSQL(%#x, len=%d, scale=%d, prec=%d) = %q, want %q",
				tc.typ, tc.length, tc.scale, tc.precision, got, tc.want)
		}
	}
}
