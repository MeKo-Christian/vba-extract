package vba

import (
	"strings"
	"testing"
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

	modules, err := ExtractAllModules(st, false, nil)
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
