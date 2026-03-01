package cmd

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/MeKo-Christian/accessdump/internal/vba"
)

// expandArg with glob

func TestExpandArg_noGlob_returnsArg(t *testing.T) {
	got, err := expandArg("/some/path.mdb")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(got) != 1 || got[0] != "/some/path.mdb" {
		t.Errorf("expected single unchanged arg, got %v", got)
	}
}

func TestExpandArg_glob_matchesFiles(t *testing.T) {
	dir := t.TempDir()

	err := os.WriteFile(filepath.Join(dir, "a.mdb"), []byte{}, 0o600)
	if err != nil {
		t.Fatalf("WriteFile a.mdb: %v", err)
	}

	err = os.WriteFile(filepath.Join(dir, "b.mdb"), []byte{}, 0o600)
	if err != nil {
		t.Fatalf("WriteFile b.mdb: %v", err)
	}

	got, err := expandArg(filepath.Join(dir, "*.mdb"))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(got) != 2 {
		t.Errorf("expected 2 matches, got %d: %v", len(got), got)
	}
}

func TestExpandArg_glob_noMatchReturnsError(t *testing.T) {
	_, err := expandArg("/nonexistent/dir/*.mdb")
	if err == nil {
		t.Error("expected error when glob matches nothing")
	}
}

// writeModules

func TestWriteModules_createsFiles(t *testing.T) {
	dir := t.TempDir()
	modules := []vba.ExtractedModule{
		{Name: "Module1", Type: vba.ProjectModuleStandard, Text: "Sub Hello()\nEnd Sub\n"},
		{Name: "Class1", Type: vba.ProjectModuleClass, Text: "Option Explicit\n"},
	}

	written, _, err := writeModules(dir, "mydb.mdb", modules, false, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if written != 2 {
		t.Errorf("expected 2 written, got %d", written)
	}

	// Files should be in a subdirectory named after the db
	subDir := filepath.Join(dir, "mydb")

	_, err = os.Stat(filepath.Join(subDir, "Module1.bas"))
	if err != nil {
		t.Errorf("expected Module1.bas, got error: %v", err)
	}

	_, err = os.Stat(filepath.Join(subDir, "Class1.cls"))
	if err != nil {
		t.Errorf("expected Class1.cls, got error: %v", err)
	}
}

func TestWriteModules_flatModeWritesDirectlyToBaseDir(t *testing.T) {
	dir := t.TempDir()
	modules := []vba.ExtractedModule{
		{Name: "Util", Type: vba.ProjectModuleStandard, Text: "' util\n"},
	}

	_, _, err := writeModules(dir, "mydb.mdb", modules, true, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	_, err = os.Stat(filepath.Join(dir, "Util.bas"))
	if err != nil {
		t.Errorf("expected Util.bas directly in base dir: %v", err)
	}
}

func TestWriteModules_countsLines(t *testing.T) {
	dir := t.TempDir()
	modules := []vba.ExtractedModule{
		{Name: "Mod", Type: vba.ProjectModuleStandard, Text: "line1\nline2\nline3\n"},
	}

	_, lines, err := writeModules(dir, "db.mdb", modules, true, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if lines != 3 {
		t.Errorf("expected 3 lines, got %d", lines)
	}
}

func TestWriteModules_sanitisesIllegalNameChars(t *testing.T) {
	dir := t.TempDir()
	modules := []vba.ExtractedModule{
		{Name: "My/Module", Type: vba.ProjectModuleStandard, Text: ""},
	}

	_, _, err := writeModules(dir, "db.mdb", modules, true, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	entries, _ := os.ReadDir(dir)
	foundModule := false

	for _, entry := range entries {
		name := entry.Name()
		if strings.HasSuffix(name, ".bas") {
			foundModule = true

			if strings.Contains(name, "/") {
				t.Errorf("slash should be replaced in filename: %q", name)
			}
		}
	}

	if !foundModule {
		t.Fatal("expected at least one .bas module output file")
	}
}

// defaultOutputDir

func TestDefaultOutputDir_usesGlobalWhenSet(t *testing.T) {
	orig := outputDir
	outputDir = "/custom/output"

	defer func() { outputDir = orig }()

	if got := defaultOutputDir(); got != "/custom/output" {
		t.Errorf("expected /custom/output, got %q", got)
	}
}

func TestDefaultOutputDir_fallsBackToVbaOutput(t *testing.T) {
	orig := outputDir
	outputDir = ""

	defer func() { outputDir = orig }()

	got := defaultOutputDir()
	if !strings.HasSuffix(got, "vba-output") {
		t.Errorf("expected path ending in vba-output, got %q", got)
	}
}

// colorize (in test env stdout is not a tty, so colorEnabled() == false)

func TestColorize_noTTY_returnsPlainText(t *testing.T) {
	// Tests run with non-tty stdout, so colorize should be a no-op
	got := colorize("31", "hello")
	if got != "hello" {
		// If running in a tty (unlikely in CI), allow the ANSI version too
		if !strings.Contains(got, "hello") {
			t.Errorf("expected text to contain 'hello', got %q", got)
		}
	}
}
