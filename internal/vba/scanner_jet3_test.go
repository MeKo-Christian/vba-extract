package vba

import (
	"encoding/binary"
	"os"
	"path/filepath"
	"testing"

	"github.com/MeKo-Christian/accessdump/internal/mdb"
)

func writeSyntheticJet3ScannerDB(t *testing.T, chainData []byte) string {
	t.Helper()

	const pageSize = 2048
	data := make([]byte, pageSize*6)
	copy(data[0:4], []byte{0x00, 0x01, 0x00, 0x00})
	binary.LittleEndian.PutUint32(data[0x14:], mdb.JetVersion3)

	// Ensure 2K layout detection succeeds:
	// page 1 has a known page type, page 2 is TDEF.
	data[pageSize] = mdb.PageTypeData
	data[2*pageSize] = mdb.PageTypeTDEF

	// Page 3: single data row carrying one LVAL record.
	page3 := data[3*pageSize : 4*pageSize]
	page3[0] = mdb.PageTypeData
	binary.LittleEndian.PutUint16(page3[0x08:], 1) // Jet3 numRows

	record := make([]byte, 4+len(chainData))
	// next pointer = 0 (end of chain)
	copy(record[4:], chainData)

	rowStart := len(page3) - len(record)
	copy(page3[rowStart:], record)
	binary.LittleEndian.PutUint16(page3[0x0A:], uint16(rowStart)) // Jet3 row table start

	f, err := os.CreateTemp("", "jet3-scanner-*.mdb")
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

func pickRecoverableModuleChainData(t *testing.T) ([]byte, string) {
	t.Helper()

	start := filepath.Join("..", "..", "testdata", "Start.mdb")
	if _, err := os.Stat(start); os.IsNotExist(err) {
		t.Skipf("fixture not present: %s", start)
	}

	db, err := mdb.Open(start)
	if err != nil {
		t.Fatalf("Open fixture: %v", err)
	}
	defer db.Close()

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	nodes, err := st.ModuleStreams()
	if err != nil {
		t.Fatalf("ModuleStreams: %v", err)
	}

	for _, node := range nodes {
		if len(node.Data) < 1000 {
			continue
		}

		m := tryExtractModuleFromChain(node.Data)
		if m != nil && m.Name != "" {
			return node.Data, m.Name
		}
	}

	t.Skip("no recoverable module stream found in Start.mdb fixture")
	return nil, ""
}

func TestScanOrphanedLvalModules_Jet3LayoutSynthetic(t *testing.T) {
	chainData, expectedName := pickRecoverableModuleChainData(t)
	path := writeSyntheticJet3ScannerDB(t, chainData)

	db, err := mdb.Open(path)
	if err != nil {
		t.Fatalf("Open synthetic Jet3 DB: %v", err)
	}
	defer db.Close()

	if !db.IsJet3() {
		t.Fatalf("expected synthetic DB to be detected as Jet3 layout, got pageSize=%d", db.PageSize())
	}

	modules, err := ScanOrphanedLvalModules(db)
	if err != nil {
		t.Fatalf("ScanOrphanedLvalModules: %v", err)
	}

	if len(modules) == 0 {
		t.Fatal("expected at least one recovered module from synthetic Jet3 LVAL chain")
	}

	found := false
	for _, m := range modules {
		if m.Name == expectedName {
			found = true
			break
		}
	}

	if !found {
		t.Fatalf("expected recovered module %q not found (got %d modules)", expectedName, len(modules))
	}
}
