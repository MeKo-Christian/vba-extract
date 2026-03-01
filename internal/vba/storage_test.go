package vba

import (
	"os"
	"testing"

	"github.com/MeKo-Christian/accessdump/internal/mdb"
)

const startMDB = "../../testdata/Start.mdb"

func startDB(t *testing.T) *mdb.Database {
	t.Helper()
	_, err := os.Stat(startMDB)
	if os.IsNotExist(err) {
		t.Skip("testdata/Start.mdb not available (proprietary fixture)")
	}
	db, err := mdb.Open(startMDB)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	t.Cleanup(func() { _ = db.Close() })
	return db
}

func TestLoadStorageTree(t *testing.T) {
	db := startDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	if len(st.Nodes) == 0 {
		t.Fatal("storage tree has no nodes")
	}

	if st.Root() == nil {
		t.Fatal("ROOT node not found")
	}
}

func TestVBAFolderAndStreams(t *testing.T) {
	db := startDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	vbaFolder, err := st.VBAFolderNode()
	if err != nil {
		t.Fatalf("VBAFolderNode: %v", err)
	}

	if vbaFolder.Name != "VBA" {
		t.Fatalf("VBA folder name = %q, want %q", vbaFolder.Name, "VBA")
	}

	required, err := st.RequiredStreams()
	if err != nil {
		t.Fatalf("RequiredStreams: %v", err)
	}

	want := []string{"PROJECT", "PROJECTwm", "dir", "_VBA_PROJECT"}
	for _, name := range want {
		if required[name] == nil {
			t.Errorf("required stream %q not found", name)
		}
	}
}

func TestModuleStreamsHaveData(t *testing.T) {
	db := startDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	modules, err := st.ModuleStreams()
	if err != nil {
		t.Fatalf("ModuleStreams: %v", err)
	}

	if len(modules) == 0 {
		t.Fatal("no module streams found")
	}

	nonEmpty := 0

	for _, module := range modules {
		if len(module.Data) > 0 {
			nonEmpty++
		}
	}

	if nonEmpty == 0 {
		t.Fatal("module streams found, but all resolved stream payloads are empty")
	}
}
