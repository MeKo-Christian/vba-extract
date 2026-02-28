package vba

import (
	"bufio"
	"log/slog"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"testing"

	"github.com/MeKo-Tech/vba-extract/internal/mdb"
)

func TestCleanupVBA(t *testing.T) {
	in := "Attribute VB_Name = \"Mod1\"\r\nOption Explicit\r\n\x00\x00"
	out := cleanupVBA(in)

	if strings.Contains(out, "\r") {
		t.Fatalf("cleanupVBA should normalize CRLF to LF: %q", out)
	}
	if !strings.HasSuffix(out, "\n") {
		t.Fatalf("cleanupVBA should end with newline: %q", out)
	}
	if strings.Contains(out, "\x00") {
		t.Fatalf("cleanupVBA should trim trailing NUL bytes: %q", out)
	}
}

func TestRecoverPartialText(t *testing.T) {
	raw := []byte("\x00\x01garbage\nSub Test()\nMsgBox \"x\"\nmore\n")
	partial, ok := recoverPartialText(raw)
	if !ok {
		t.Fatal("expected partial recovery")
	}
	if !strings.Contains(partial, "[PARTIAL - reconstructed from p-code tokens]") {
		t.Fatalf("missing partial header: %q", partial)
	}
	if !strings.Contains(partial, "Sub Test()") {
		t.Fatalf("expected recovered VBA fragment: %q", partial)
	}
}

func TestExtractAllModulesFromStartMDB(t *testing.T) {
	db := testDB(t)
	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	modules, err := ExtractAllModules(st, slog.New(slog.DiscardHandler))
	if err != nil {
		t.Fatalf("ExtractAllModules: %v", err)
	}
	if len(modules) == 0 {
		t.Fatal("no modules extracted")
	}

	hasNamed := false
	hasSource := false
	for _, module := range modules {
		if strings.TrimSpace(module.Name) != "" {
			hasNamed = true
		}
		if strings.Contains(strings.ToLower(module.Text), "attribute vb_name") {
			hasSource = true
			break
		}
	}

	if !hasNamed {
		t.Fatal("extracted modules have no names")
	}
	if !hasSource {
		t.Fatal("no extracted module contains Attribute VB_Name")
	}
}

func TestExtractAllModulesMatchesExpectedManifest(t *testing.T) {
	db := testDB(t)
	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	modules, err := ExtractAllModules(st, slog.New(slog.DiscardHandler))
	if err != nil {
		t.Fatalf("ExtractAllModules: %v", err)
	}

	expectedPath := filepath.Join("..", "..", "testdata", "Start.expected.modules.txt")
	expected, err := readExpectedModuleNames(expectedPath)
	if err != nil {
		t.Fatalf("readExpectedModuleNames: %v", err)
	}

	actualSet := map[string]struct{}{}
	for _, module := range modules {
		name := strings.TrimSpace(module.Name)
		if name == "" {
			continue
		}
		actualSet[name] = struct{}{}
	}

	var missing []string
	for _, name := range expected {
		if _, ok := actualSet[name]; !ok {
			missing = append(missing, name)
		}
	}

	if len(missing) > 0 {
		sort.Strings(missing)
		t.Fatalf("missing expected modules: %v", missing)
	}
}

func TestOptionalExtractionFixtures(t *testing.T) {
	optional := []string{"AT990426.mdb", "PPS.mdb"}
	for _, fixtureName := range optional {
		fixtureName := fixtureName
		t.Run(fixtureName, func(t *testing.T) {
			path, ok := findOptionalFixture(fixtureName)
			if !ok {
				t.Skipf("optional fixture %q not found", fixtureName)
			}

			db, err := mdb.Open(path)
			if err != nil {
				t.Fatalf("open fixture: %v", err)
			}
			defer db.Close()

			st, err := LoadStorageTree(db)
			if err != nil {
				t.Fatalf("LoadStorageTree: %v", err)
			}

			modules, err := ExtractAllModules(st, slog.New(slog.DiscardHandler))
			if err != nil {
				t.Fatalf("ExtractAllModules: %v", err)
			}
			if len(modules) == 0 {
				t.Fatal("no modules extracted from optional fixture")
			}
		})
	}
}

func TestStartMDBSpotCheckCountsAndContent(t *testing.T) {
	db := testDB(t)
	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	modules, err := ExtractAllModules(st, slog.New(slog.DiscardHandler))
	if err != nil {
		t.Fatalf("ExtractAllModules: %v", err)
	}

	expected, err := readExpectedModuleNames(filepath.Join("..", "..", "testdata", "Start.expected.modules.txt"))
	if err != nil {
		t.Fatalf("readExpectedModuleNames: %v", err)
	}

	if len(modules) < len(expected) {
		t.Fatalf("module count = %d, expected at least %d", len(modules), len(expected))
	}

	moduleByName := make(map[string]ExtractedModule, len(modules))
	for _, m := range modules {
		if strings.TrimSpace(m.Name) == "" {
			continue
		}
		moduleByName[m.Name] = m
	}

	spotChecks := []string{"Inidatei", "SQL", "mod_api_Functions", "Form_Module"}
	for _, name := range spotChecks {
		m, ok := moduleByName[name]
		if !ok {
			t.Fatalf("spot-check module %q not found", name)
		}
		if !strings.Contains(strings.ToLower(m.Text), "attribute vb_name") {
			t.Fatalf("module %q does not contain Attribute VB_Name", name)
		}
		lines := strings.Count(m.Text, "\n")
		if lines < 2 {
			t.Fatalf("module %q has too few lines: %d", name, lines)
		}
	}
}

func readExpectedModuleNames(path string) ([]string, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	var names []string
	scanner := bufio.NewScanner(f)
	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line == "" || strings.HasPrefix(line, "#") {
			continue
		}
		names = append(names, line)
	}
	if err := scanner.Err(); err != nil {
		return nil, err
	}

	return names, nil
}

// findOptionalFixture looks for a named fixture file in:
//  1. The directory given by the VBA_FIXTURE_DIR environment variable (if set)
//  2. ../../testdata/ relative to this package
func findOptionalFixture(name string) (string, bool) {
	var candidates []string

	if dir := os.Getenv("VBA_FIXTURE_DIR"); dir != "" {
		candidates = append(candidates, filepath.Join(dir, name))
	}
	candidates = append(candidates, filepath.Join("..", "..", "testdata", name))

	for _, candidate := range candidates {
		if _, err := os.Stat(candidate); err == nil {
			return candidate, true
		}
	}

	return "", false
}
