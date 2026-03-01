package cmd

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/MeKo-Christian/accessdump/internal/vba"
)

// hasGlob

func TestHasGlob_withAsterisk(t *testing.T) {
	if !hasGlob("*.mdb") {
		t.Error("expected true for pattern with *")
	}
}

func TestHasGlob_withQuestionMark(t *testing.T) {
	if !hasGlob("file?.mdb") {
		t.Error("expected true for pattern with ?")
	}
}

func TestHasGlob_withBracket(t *testing.T) {
	if !hasGlob("file[0-9].mdb") {
		t.Error("expected true for pattern with [")
	}
}

func TestHasGlob_plainPath(t *testing.T) {
	if hasGlob("/path/to/file.mdb") {
		t.Error("expected false for plain path")
	}
}

// isAccessFile

func TestIsAccessFile_mdb(t *testing.T) {
	if !isAccessFile("database.mdb") {
		t.Error("expected true for .mdb")
	}
}

func TestIsAccessFile_accdb(t *testing.T) {
	if !isAccessFile("database.accdb") {
		t.Error("expected true for .accdb")
	}
}

func TestIsAccessFile_uppercaseExtension(t *testing.T) {
	if !isAccessFile("database.MDB") {
		t.Error("expected true for .MDB (case-insensitive)")
	}
}

func TestIsAccessFile_wrongExtension(t *testing.T) {
	if isAccessFile("database.xlsx") {
		t.Error("expected false for .xlsx")
	}
}

func TestIsAccessFile_noExtension(t *testing.T) {
	if isAccessFile("database") {
		t.Error("expected false for file without extension")
	}
}

// safeModuleName

func TestSafeModuleName_plain(t *testing.T) {
	if got := safeModuleName("Module1"); got != "Module1" {
		t.Errorf("expected Module1, got %q", got)
	}
}

func TestSafeModuleName_empty(t *testing.T) {
	if got := safeModuleName(""); got != "unnamed" {
		t.Errorf("expected unnamed, got %q", got)
	}
}

func TestSafeModuleName_whitespaceOnly(t *testing.T) {
	if got := safeModuleName("   "); got != "unnamed" {
		t.Errorf("expected unnamed, got %q", got)
	}
}

func TestSafeModuleName_trailingSpaces(t *testing.T) {
	if got := safeModuleName("  Module1  "); got != "Module1" {
		t.Errorf("expected Module1, got %q", got)
	}
}

func TestSafeModuleName_illegalChars(t *testing.T) {
	illegal := []string{"/", "\\", ":", "*", "?", "\"", "<", ">", "|"}
	for _, ch := range illegal {
		got := safeModuleName("Module" + ch + "1")
		if strings.Contains(got, ch) {
			t.Errorf("char %q not replaced in %q", ch, got)
		}

		if !strings.Contains(got, "_") {
			t.Errorf("expected underscore replacement for %q, got %q", ch, got)
		}
	}
}

// moduleExt

func TestModuleExt_standardModule(t *testing.T) {
	if got := moduleExt(vba.ProjectModuleStandard); got != ".bas" {
		t.Errorf("expected .bas, got %q", got)
	}
}

func TestModuleExt_classModule(t *testing.T) {
	if got := moduleExt(vba.ProjectModuleClass); got != ".cls" {
		t.Errorf("expected .cls, got %q", got)
	}
}

func TestModuleExt_documentModule(t *testing.T) {
	if got := moduleExt(vba.ProjectModuleDocument); got != ".cls" {
		t.Errorf("expected .cls, got %q", got)
	}
}

// addUniquePath

func TestAddUniquePath_addsNewPath(t *testing.T) {
	seen := map[string]struct{}{}
	var files []string

	addUniquePath("testdata/a.mdb", seen, &files)

	if len(files) != 1 {
		t.Errorf("expected 1 file, got %d", len(files))
	}
}

func TestAddUniquePath_deduplicates(t *testing.T) {
	seen := map[string]struct{}{}
	var files []string

	addUniquePath("testdata/a.mdb", seen, &files)
	addUniquePath("testdata/a.mdb", seen, &files)

	if len(files) != 1 {
		t.Errorf("expected 1 file after dedup, got %d", len(files))
	}
}

func TestAddUniquePath_storesAbsPath(t *testing.T) {
	seen := map[string]struct{}{}
	var files []string

	addUniquePath("testdata/a.mdb", seen, &files)

	if len(files) == 1 && !filepath.IsAbs(files[0]) {
		t.Errorf("expected absolute path, got %q", files[0])
	}
}

// printListTable

func TestPrintListTable_header(t *testing.T) {
	var buf bytes.Buffer
	printListTable(&buf, nil)

	out := buf.String()
	if !strings.Contains(out, "MODULE") || !strings.Contains(out, "TYPE") {
		t.Errorf("header missing expected columns, got: %q", out)
	}
}

func TestPrintListTable_entry(t *testing.T) {
	var buf bytes.Buffer
	entries := []listEntry{
		{Name: "MyModule", Type: "standard", Stream: "MyStream", SizeBytes: 1024, Partial: false},
	}

	printListTable(&buf, entries)

	out := buf.String()
	if !strings.Contains(out, "MyModule") || !strings.Contains(out, "1024") {
		t.Errorf("expected entry in output, got: %q", out)
	}
}

// printListJSON

func TestPrintListJSON_emptyList(t *testing.T) {
	var buf bytes.Buffer

	err := printListJSON(&buf, []listEntry{})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	out := buf.String()
	if !strings.Contains(out, "[]") {
		t.Errorf("expected empty JSON array, got: %q", out)
	}
}

func TestPrintListJSON_entry(t *testing.T) {
	var buf bytes.Buffer
	entries := []listEntry{
		{Name: "Mod1", Type: "standard", SizeBytes: 42},
	}

	err := printListJSON(&buf, entries)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	out := buf.String()
	if !strings.Contains(out, `"name"`) || !strings.Contains(out, "Mod1") {
		t.Errorf("expected JSON with name field, got: %q", out)
	}
}

// computeFileHash

func TestComputeFileHash_returnsHex(t *testing.T) {
	f, err := os.CreateTemp(t.TempDir(), "*.mdb")
	if err != nil {
		t.Fatal(err)
	}

	_, err = f.WriteString("test content")
	if err != nil {
		t.Fatal(err)
	}

	err = f.Close()
	if err != nil {
		t.Fatal(err)
	}

	hash, err := computeFileHash(f.Name())
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(hash) != 64 {
		t.Errorf("expected 64-char hex hash (SHA-256), got %d chars: %q", len(hash), hash)
	}
}

func TestComputeFileHash_deterministicForSameContent(t *testing.T) {
	dir := t.TempDir()

	writeFile := func(name, content string) string {
		path := filepath.Join(dir, name)

		err := os.WriteFile(path, []byte(content), 0o600)
		if err != nil {
			t.Fatal(err)
		}

		return path
	}

	a := writeFile("a.mdb", "same content")
	b := writeFile("b.mdb", "same content")

	hashA, _ := computeFileHash(a)
	hashB, _ := computeFileHash(b)

	if hashA != hashB {
		t.Errorf("same content should produce same hash: %q vs %q", hashA, hashB)
	}
}

func TestComputeFileHash_differsForDifferentContent(t *testing.T) {
	dir := t.TempDir()

	writeFile := func(name, content string) string {
		path := filepath.Join(dir, name)

		err := os.WriteFile(path, []byte(content), 0o600)
		if err != nil {
			t.Fatal(err)
		}

		return path
	}

	a := writeFile("a.mdb", "content A")
	b := writeFile("b.mdb", "content B")

	hashA, _ := computeFileHash(a)
	hashB, _ := computeFileHash(b)

	if hashA == hashB {
		t.Error("different content should produce different hashes")
	}
}

func TestComputeFileHash_missingFile(t *testing.T) {
	_, err := computeFileHash("/nonexistent/path/file.mdb")
	if err == nil {
		t.Error("expected error for missing file")
	}
}

// discoverInputFiles

func TestDiscoverInputFiles_singleFile(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "test.mdb")

	err := os.WriteFile(path, []byte{}, 0o600)
	if err != nil {
		t.Fatalf("WriteFile test.mdb: %v", err)
	}

	files, err := discoverInputFiles([]string{path}, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(files) != 1 {
		t.Errorf("expected 1 file, got %d", len(files))
	}
}

func TestDiscoverInputFiles_ignoresNonAccessFiles(t *testing.T) {
	dir := t.TempDir()

	err := os.WriteFile(filepath.Join(dir, "readme.txt"), []byte{}, 0o600)
	if err != nil {
		t.Fatalf("WriteFile readme.txt: %v", err)
	}

	err = os.WriteFile(filepath.Join(dir, "data.xlsx"), []byte{}, 0o600)
	if err != nil {
		t.Fatalf("WriteFile data.xlsx: %v", err)
	}

	// Pass a non-Access file directly — it should be silently ignored
	files, err := discoverInputFiles([]string{filepath.Join(dir, "readme.txt")}, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(files) != 0 {
		t.Errorf("expected 0 files, got %d", len(files))
	}
}

func TestDiscoverInputFiles_recursive(t *testing.T) {
	dir := t.TempDir()
	sub := filepath.Join(dir, "sub")

	err := os.MkdirAll(sub, 0o755)
	if err != nil {
		t.Fatalf("MkdirAll sub: %v", err)
	}

	err = os.WriteFile(filepath.Join(sub, "deep.mdb"), []byte{}, 0o600)
	if err != nil {
		t.Fatalf("WriteFile deep.mdb: %v", err)
	}

	files, err := discoverInputFiles([]string{dir}, true)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(files) != 1 {
		t.Errorf("expected 1 file from recursive walk, got %d", len(files))
	}
}

func TestDiscoverInputFiles_nonRecursiveSkipsDir(t *testing.T) {
	dir := t.TempDir()
	sub := filepath.Join(dir, "sub")

	err := os.MkdirAll(sub, 0o755)
	if err != nil {
		t.Fatalf("MkdirAll sub: %v", err)
	}

	err = os.WriteFile(filepath.Join(sub, "deep.mdb"), []byte{}, 0o600)
	if err != nil {
		t.Fatalf("WriteFile deep.mdb: %v", err)
	}

	files, err := discoverInputFiles([]string{dir}, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(files) != 0 {
		t.Errorf("expected 0 files without recursive flag, got %d", len(files))
	}
}

func TestDiscoverInputFiles_deduplicates(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "test.mdb")

	err := os.WriteFile(path, []byte{}, 0o600)
	if err != nil {
		t.Fatalf("WriteFile test.mdb: %v", err)
	}

	files, err := discoverInputFiles([]string{path, path}, false)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(files) != 1 {
		t.Errorf("expected 1 file after dedup, got %d", len(files))
	}
}

func TestDiscoverInputFiles_missingPath(t *testing.T) {
	_, err := discoverInputFiles([]string{"/nonexistent/path.mdb"}, false)
	if err == nil {
		t.Error("expected error for missing path")
	}
}
