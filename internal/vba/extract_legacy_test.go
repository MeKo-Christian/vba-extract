package vba

import (
	"log/slog"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/MeKo-Christian/accessdump/internal/mdb"
)

func TestExtractAllModulesFromLegacyFixture(t *testing.T) {
	path := filepath.Join("..", "..", "testdata", "jet35", "st990426.mdb")

	_, err := os.Stat(path)

	if os.IsNotExist(err) {
		t.Skipf("legacy fixture not present: %s", path)
	}

	db, err := mdb.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer db.Close()

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	required, err := st.RequiredStreams()
	if err != nil {
		t.Fatalf("RequiredStreams: %v", err)
	}

	for _, name := range []string{"PROJECT", "dir"} {
		if required[name] == nil {
			t.Fatalf("required stream %q not found", name)
		}
	}

	modules, err := ExtractAllModules(st, slog.New(slog.DiscardHandler))
	if err != nil {
		t.Fatalf("ExtractAllModules: %v", err)
	}

	// This fixture currently yields 312 modules in CLI output.
	// Keep lower bound to avoid brittle failures if fixture changes.
	if len(modules) < 250 {
		t.Fatalf("module count = %d, want at least 250", len(modules))
	}

	hasNamed := false
	hasVBName := false

	for _, m := range modules {
		if strings.TrimSpace(m.Name) != "" {
			hasNamed = true
		}

		if strings.Contains(strings.ToLower(m.Text), "attribute vb_name") {
			hasVBName = true
		}
	}

	if !hasNamed {
		t.Fatal("no named modules extracted from legacy fixture")
	}

	if !hasVBName {
		t.Fatal("no module text contains Attribute VB_Name in legacy fixture")
	}
}
