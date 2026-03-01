package mdb

import (
	"encoding/binary"
	"errors"
	"os"
	"testing"
)

func writeSyntheticDB(t *testing.T, pageSize, pageCount int, jetVersion uint32) string {
	t.Helper()

	data := make([]byte, pageSize*pageCount)
	copy(data[0:4], magicBytes[:])
	binary.LittleEndian.PutUint32(data[offsetJetVersion:], jetVersion)

	if pageCount > 2 {
		data[2*pageSize] = PageTypeTDEF
	}

	f, err := os.CreateTemp("", "mdb-pagesize-*.mdb")
	if err != nil {
		t.Fatalf("CreateTemp: %v", err)
	}

	t.Cleanup(func() {
		_ = os.Remove(f.Name())
	})

	_, err = f.Write(data)
	if err != nil {
		_ = f.Close()

		t.Fatalf("Write: %v", err)
	}

	err = f.Close()
	if err != nil {
		t.Fatalf("Close: %v", err)
	}

	return f.Name()
}

func TestOpenJet3Uses2048PageSize(t *testing.T) {
	path := writeSyntheticDB(t, 2048, 4, JetVersion3)

	db, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	t.Cleanup(func() { _ = db.Close() })

	if db.PageSize() != 2048 {
		t.Fatalf("PageSize = %d, want 2048", db.PageSize())
	}

	if !db.IsJet3() {
		t.Fatal("IsJet3() = false, want true")
	}

	if db.IsJet4() {
		t.Fatal("IsJet4() = true, want false")
	}

	if db.PageCount() != 4 {
		t.Fatalf("PageCount = %d, want 4", db.PageCount())
	}

	numRowsOff, rowTableOff := db.DataPageLayoutOffsets()
	if numRowsOff != dataNumRowsJet3 || rowTableOff != dataRowTableJet3 {
		t.Fatalf("DataPageLayoutOffsets = (%#x,%#x), want (%#x,%#x)",
			numRowsOff, rowTableOff, dataNumRowsJet3, dataRowTableJet3)
	}

	page2, err := db.ReadPage(2)
	if err != nil {
		t.Fatalf("ReadPage(2): %v", err)
	}

	if PageType(page2) != PageTypeTDEF {
		t.Fatalf("PageType(page2) = %#x, want %#x", PageType(page2), PageTypeTDEF)
	}

	_, err = db.ReadTableDef(2)
	if errors.Is(err, ErrJet3TableLayoutUnsupported) {
		t.Fatalf("ReadTableDef should use Jet3 parser path, got %v", err)
	}

	_, err = db.ResolveMemo([]byte{1})
	if errors.Is(err, ErrJet3LvalLayoutUnsupported) {
		t.Fatalf("ResolveMemo should use Jet3 LVAL parser path, got %v", err)
	}

	td := &TableDef{db: db}

	_, err = td.parseRow([]byte{0, 0, 0, 0}, nil)
	if errors.Is(err, ErrJet3RowLayoutUnsupported) {
		t.Fatalf("parseRow should use Jet3 row parser path, got %v", err)
	}
}

func TestOpenJet4Uses4096PageSize(t *testing.T) {
	path := writeSyntheticDB(t, 4096, 3, JetVersion4)

	db, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	t.Cleanup(func() { _ = db.Close() })

	if db.PageSize() != 4096 {
		t.Fatalf("PageSize = %d, want 4096", db.PageSize())
	}

	if db.IsJet3() {
		t.Fatal("IsJet3() = true, want false")
	}

	if !db.IsJet4() {
		t.Fatal("IsJet4() = false, want true")
	}

	if db.PageCount() != 3 {
		t.Fatalf("PageCount = %d, want 3", db.PageCount())
	}

	numRowsOff, rowTableOff := db.DataPageLayoutOffsets()
	if numRowsOff != dataNumRows || rowTableOff != dataRowTable {
		t.Fatalf("DataPageLayoutOffsets = (%#x,%#x), want (%#x,%#x)",
			numRowsOff, rowTableOff, dataNumRows, dataRowTable)
	}
}

func TestOpenPrefersPageLayoutOverJetVersionField(t *testing.T) {
	// Header says JetVersion3, but actual layout is 4K pages.
	path := writeSyntheticDB(t, 4096, 4, JetVersion3)

	db, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	t.Cleanup(func() { _ = db.Close() })

	if db.PageSize() != 4096 {
		t.Fatalf("PageSize = %d, want 4096", db.PageSize())
	}
}
