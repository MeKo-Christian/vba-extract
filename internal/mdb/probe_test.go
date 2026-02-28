package mdb

import (
	"os"
	"path/filepath"
	"testing"
)

func TestProbeClassifiesSyntheticJet32K(t *testing.T) {
	path := writeSyntheticDB(t, 2048, 4, JetVersion3)

	got, err := Probe(path)
	if err != nil {
		t.Fatalf("Probe: %v", err)
	}

	if got.LayoutClass != LayoutJet32K {
		t.Fatalf("LayoutClass = %q, want %q", got.LayoutClass, LayoutJet32K)
	}

	if got.PageSize != 2048 {
		t.Fatalf("PageSize = %d, want 2048", got.PageSize)
	}

	if !got.MSysObjectsTDEF {
		t.Fatal("MSysObjectsTDEF = false, want true")
	}
}

func TestProbeClassifiesStartMDBAsJet44K(t *testing.T) {
	path := filepath.Join("..", "..", "testdata", "Start.mdb")
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Skipf("fixture not present: %s", path)
	}

	got, err := Probe(path)
	if err != nil {
		t.Fatalf("Probe: %v", err)
	}

	if got.LayoutClass != LayoutJet44K {
		t.Fatalf("LayoutClass = %q, want %q", got.LayoutClass, LayoutJet44K)
	}

	if !got.MSysObjectsReadable {
		t.Fatalf("MSysObjectsReadable = false, err=%q", got.MSysObjectsError)
	}
}

func TestProbeClassifiesLegacyFixtureAsLegacy4K(t *testing.T) {
	path := filepath.Join("..", "..", "testdata", "jet35", "st990426.mdb")
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Skipf("fixture not present: %s", path)
	}

	got, err := Probe(path)
	if err != nil {
		t.Fatalf("Probe: %v", err)
	}

	if got.LayoutClass != LayoutLegacy4 {
		t.Fatalf("LayoutClass = %q, want %q", got.LayoutClass, LayoutLegacy4)
	}

	if !got.MSysObjectsReadable {
		t.Fatalf("MSysObjectsReadable = false, err=%q", got.MSysObjectsError)
	}
}
