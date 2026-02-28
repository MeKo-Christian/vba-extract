package mdb

import (
	"strings"
	"testing"
)

func TestResolveMemo_PROJECT(t *testing.T) {
	db := testDB(t)

	td, err := db.FindTable("MSysAccessStorage")
	if err != nil {
		t.Fatalf("FindTable: %v", err)
	}

	rows, err := td.ReadRows()
	if err != nil {
		t.Fatalf("ReadRows: %v", err)
	}

	// Find the PROJECT stream (Id=4847).
	for _, row := range rows {
		name, _ := row["Name"].(string)
		if name != "PROJECT" {
			continue
		}

		lvRaw, ok := row["Lv"].([]byte)
		if !ok || len(lvRaw) == 0 {
			t.Fatal("PROJECT stream has no Lv data")
		}

		t.Logf("PROJECT Lv raw: %d bytes, first 16: %x", len(lvRaw), lvRaw[:min(16, len(lvRaw))])

		data, err := db.ResolveMemo(lvRaw)
		if err != nil {
			t.Fatalf("ResolveMemo: %v", err)
		}

		if len(data) == 0 {
			t.Fatal("PROJECT stream data is empty after resolution")
		}

		// PROJECT stream should be latin-1 text containing module names.
		text := string(data)
		t.Logf("PROJECT stream: %d bytes", len(data))
		t.Logf("First 200 chars: %s", text[:min(200, len(text))])

		// Verify expected content.
		for _, want := range []string{"Module=Inidatei", "Module=Modul1", "Module=SQL", "Name=\"Start\""} {
			if !strings.Contains(text, want) {
				t.Errorf("PROJECT stream missing %q", want)
			}
		}

		return
	}

	t.Fatal("PROJECT entry not found in MSysAccessStorage")
}
