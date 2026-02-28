package vba

import (
	"testing"
)

func TestScanOrphanedLvalModules_returnsNoError(t *testing.T) {
	// Start.mdb has a proper MSysAccessStorage so orphaned scan should return
	// empty (all modules are reachable via normal path), but must not error.
	db := testDB(t)

	modules, err := ScanOrphanedLvalModules(db)
	if err != nil {
		t.Fatalf("ScanOrphanedLvalModules: %v", err)
	}

	// Result may be empty (if all modules found via normal path) — that's fine.
	_ = modules
}

// extractModuleNameFromText

func TestExtractModuleNameFromText_findsName(t *testing.T) {
	text := `Attribute VB_Name = "MyModule"
Option Explicit
Sub Foo()
End Sub`

	got := extractModuleNameFromText([]byte(text))
	if got != "MyModule" {
		t.Errorf("expected MyModule, got %q", got)
	}
}

func TestExtractModuleNameFromText_caseInsensitive(t *testing.T) {
	text := `attribute vb_name = "LowerCase"`

	got := extractModuleNameFromText([]byte(text))
	if got != "LowerCase" {
		t.Errorf("expected LowerCase, got %q", got)
	}
}

func TestExtractModuleNameFromText_notFound(t *testing.T) {
	text := "Option Explicit\nSub Foo()\nEnd Sub"

	got := extractModuleNameFromText([]byte(text))
	if got != "" {
		t.Errorf("expected empty string, got %q", got)
	}
}
