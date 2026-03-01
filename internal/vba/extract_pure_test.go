package vba

import (
	"encoding/binary"
	"log/slog"
	"strings"
	"testing"
)

// buildUncompressedOVBAContainer builds a minimal MS-OVBA compressed container
// with the given payload stored as an uncompressed chunk.
func buildUncompressedOVBAContainer(payload []byte) []byte {
	// Chunk header: uncompressed (bit 15 = 0), magic 0x3 in bits 14:12, size-3 in bits 11:0
	chunkSize := uint16(len(payload) + 2)
	header := uint16(0x3000) | (chunkSize - 3)

	out := make([]byte, 1+2+len(payload))
	out[0] = compressedContainerSig // 0x01
	binary.LittleEndian.PutUint16(out[1:3], header)
	copy(out[3:], payload)

	return out
}

// extractModuleSource – empty data

func TestExtractModuleSource_emptyData(t *testing.T) {
	m := ModuleMapping{ModuleName: "Empty", Data: nil, SourceOffset: 0}

	text, warnings, partial := extractModuleSource(m, slog.New(slog.DiscardHandler))
	if text != "" {
		t.Errorf("expected empty text, got %q", text)
	}

	if len(warnings) == 0 {
		t.Error("expected at least one warning for empty data")
	}

	_ = partial
}

// extractModuleSource – offset beyond data length triggers brute-force scan

func TestExtractModuleSource_offsetBeyondData_noContainer(t *testing.T) {
	// Data that contains no valid compressed container
	data := []byte("garbage data without VBA structure")
	m := ModuleMapping{
		ModuleName:   "Mod",
		Data:         data,
		SourceOffset: 9999, // way beyond data length
	}

	_, warnings, _ := extractModuleSource(m, slog.New(slog.DiscardHandler))

	found := false

	for _, w := range warnings {
		if strings.Contains(w, "exceeds") {
			found = true
		}
	}

	if !found {
		t.Errorf("expected offset-exceeds warning, got warnings: %v", warnings)
	}
}

// extractModuleSource – valid offset with valid compressed data

func TestExtractModuleSource_validOVBAContainer(t *testing.T) {
	vbaText := []byte("Attribute VB_Name = \"TestMod\"\r\nSub Hello()\r\nEnd Sub\r\n")
	container := buildUncompressedOVBAContainer(vbaText)

	m := ModuleMapping{
		ModuleName:   "TestMod",
		Data:         container,
		SourceOffset: 0,
	}

	text, _, _ := extractModuleSource(m, slog.New(slog.DiscardHandler))

	if !strings.Contains(strings.ToLower(text), "sub hello") {
		t.Errorf("expected VBA source in output, got: %q", text)
	}
}

// bruteForceOffsetScan – no compressed container in data

func TestBruteForceOffsetScan_noContainer(t *testing.T) {
	data := []byte("no compressed container here at all")

	_, ok := bruteForceOffsetScan(data)
	if ok {
		t.Error("expected false when no valid compressed container found")
	}
}

// bruteForceOffsetScan – embedded valid container with VBA text

func TestBruteForceOffsetScan_withVBAContainer(t *testing.T) {
	vbaText := []byte("Attribute VB_Name = \"Mod\"\r\nSub Test()\r\nEnd Sub\r\n")
	container := buildUncompressedOVBAContainer(vbaText)

	// Prefix with some garbage bytes before the container
	data := append([]byte("some prefix bytes "), container...)

	text, ok := bruteForceOffsetScan(data)
	if !ok {
		t.Error("expected brute-force to find embedded VBA container")
	}

	if !strings.Contains(strings.ToLower(text), "sub test") {
		t.Errorf("expected VBA source in brute-force result, got: %q", text)
	}
}

// recoverPartialFromRaw – no recoverable fragments

func TestRecoverPartialFromRaw_noFragments(t *testing.T) {
	raw := []byte{0x00, 0x01, 0x02, 0x03, 0xFF}

	text, warnings, partial := recoverPartialFromRaw(raw, nil)
	if text != "" {
		t.Errorf("expected empty text for unrecoverable raw data, got: %q", text)
	}

	if !partial {
		t.Error("expected partial=true when recovery fails")
	}

	hasWarning := false

	for _, w := range warnings {
		if strings.Contains(w, "no recoverable") {
			hasWarning = true
		}
	}

	if !hasWarning {
		t.Errorf("expected 'no recoverable' warning, got: %v", warnings)
	}
}

// ExtractModuleMap – calls ExtractAllModules; test via real fixture

func TestExtractModuleMap_fromStartMDB(t *testing.T) {
	db := testDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	mmap, err := ExtractModuleMap(st, slog.New(slog.DiscardHandler))
	if err != nil {
		t.Fatalf("ExtractModuleMap: %v", err)
	}

	if len(mmap) == 0 {
		t.Fatal("expected at least one module in map")
	}

	// Map keys must match the module names
	for name, module := range mmap {
		if name != module.Name {
			t.Errorf("map key %q != module.Name %q", name, module.Name)
		}
	}
}
