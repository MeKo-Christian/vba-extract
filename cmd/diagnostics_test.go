package cmd

import (
	"errors"
	"strings"
	"testing"

	"github.com/MeKo-Christian/accessdump/internal/mdb"
	"github.com/MeKo-Christian/accessdump/internal/vba"
)

func TestLayoutClass(t *testing.T) {
	tests := []struct {
		name      string
		pageSize  int64
		jet       uint32
		wantClass string
	}{
		{name: "jet3 2k", pageSize: mdb.PageSizeJet3, jet: mdb.JetVersion3, wantClass: mdb.LayoutJet32K},
		{name: "legacy 4k header", pageSize: mdb.PageSizeJet4, jet: 259, wantClass: mdb.LayoutLegacy4},
		{name: "jet4", pageSize: mdb.PageSizeJet4, jet: mdb.JetVersion4, wantClass: mdb.LayoutJet44K},
		{name: "unknown", pageSize: 1234, jet: 77, wantClass: mdb.LayoutUnknown},
	}

	for _, tt := range tests {
		got := layoutClass(tt.pageSize, tt.jet)
		if got != tt.wantClass {
			t.Fatalf("%s: layoutClass(%d, %d) = %q, want %q", tt.name, tt.pageSize, tt.jet, got, tt.wantClass)
		}
	}
}

func TestLayoutHint(t *testing.T) {
	if got := layoutHint(mdb.LayoutJet32K); !strings.Contains(got, "Jet 3.x") {
		t.Fatalf("Jet32K hint = %q, want Jet 3.x note", got)
	}

	if got := layoutHint(mdb.LayoutLegacy4); !strings.Contains(got, "Legacy") {
		t.Fatalf("Legacy4 hint = %q, want legacy note", got)
	}

	if got := layoutHint(mdb.LayoutJet44K); got != "" {
		t.Fatalf("Jet44K hint = %q, want empty", got)
	}
}

func TestCommandErrorHint(t *testing.T) {
	cases := []struct {
		err  error
		want string
	}{
		{err: errors.New("mdb: invalid magic bytes: 00112233"), want: "valid Access database header"},
		{err: errors.New("mdb: file too small (42 bytes)"), want: "truncated"},
		{err: errors.New("vba: encrypted project"), want: "encrypted"},
		{err: errors.New("mdb: table layout parsing is not implemented"), want: "not supported"},
	}

	for _, tc := range cases {
		got := commandErrorHint(tc.err)
		if !strings.Contains(strings.ToLower(got), strings.ToLower(tc.want)) {
			t.Fatalf("commandErrorHint(%q) = %q, want substring %q", tc.err.Error(), got, tc.want)
		}
	}
}

func TestExtractionStats(t *testing.T) {
	modules := []vba.ExtractedModule{
		{Name: "A", Partial: false, Warnings: nil},
		{Name: "B", Partial: true, Warnings: []string{"w1", "w2"}},
		{Name: "C", Partial: true, Warnings: []string{"w3"}},
	}

	partials, warnings := extractionStats(modules)
	if partials != 2 {
		t.Fatalf("partials = %d, want 2", partials)
	}

	if warnings != 3 {
		t.Fatalf("warnings = %d, want 3", warnings)
	}
}
