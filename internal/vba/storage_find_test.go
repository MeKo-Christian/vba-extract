package vba

import (
	"testing"
)

// Minimal valid PROJECT stream text that ParseProjectStream can parse.
const validProjectStreamText = "ID=proj\r\nModule=Module1\r\nCMG=\"\"\r\n"

func TestFindLikelyProjectNode_findsNodeWithProjectData(t *testing.T) {
	node := &StorageNode{
		ID:   1,
		Name: "PROJECT",
		Type: 2,
		Data: []byte(validProjectStreamText),
	}

	st := &StorageTree{
		Nodes:    []*StorageNode{node},
		ByID:     map[int32]*StorageNode{1: node},
		Children: map[int32][]*StorageNode{},
	}

	got := st.findLikelyProjectNode()
	if got == nil {
		t.Fatal("expected findLikelyProjectNode to return a node")
	}

	if got != node {
		t.Errorf("expected the PROJECT node, got ID=%d", got.ID)
	}
}

func TestFindLikelyProjectNode_emptyTree(t *testing.T) {
	st := &StorageTree{
		Nodes:    nil,
		ByID:     map[int32]*StorageNode{},
		Children: map[int32][]*StorageNode{},
	}

	got := st.findLikelyProjectNode()
	if got != nil {
		t.Errorf("expected nil for empty tree, got %v", got)
	}
}

func TestFindLikelyProjectNode_noMatchingNodes(t *testing.T) {
	node := &StorageNode{
		ID:   1,
		Name: "randomdata",
		Data: []byte("not a project stream"),
	}

	st := &StorageTree{
		Nodes:    []*StorageNode{node},
		ByID:     map[int32]*StorageNode{1: node},
		Children: map[int32][]*StorageNode{},
	}

	got := st.findLikelyProjectNode()
	if got != nil {
		t.Errorf("expected nil for non-project data, got %v", got)
	}
}

// findLikelyDirNode – uses the real sample.mdb storage tree which contains
// an actual dir stream, exercising the decompression-based node detection.

func TestFindLikelyDirNode_withRealFixture(t *testing.T) {
	db := startDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	// The real database has a dir stream; findLikelyDirNode should find it.
	got := st.findLikelyDirNode()
	if got == nil {
		t.Fatal("expected findLikelyDirNode to find dir stream in sample.mdb")
	}
}

func TestFindLikelyDirNode_emptyTree(t *testing.T) {
	st := &StorageTree{
		Nodes:    nil,
		ByID:     map[int32]*StorageNode{},
		Children: map[int32][]*StorageNode{},
	}

	got := st.findLikelyDirNode()
	if got != nil {
		t.Errorf("expected nil for empty tree, got %v", got)
	}
}
