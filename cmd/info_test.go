package cmd

import (
	"bytes"
	"strings"
	"testing"

	"github.com/MeKo-Christian/accessdump/internal/vba"
)

func makeMinimalStorageTree() *vba.StorageTree {
	root := &vba.StorageNode{ID: 0, ParentID: -1, Name: "Root", Type: 1}
	child := &vba.StorageNode{ID: 1, ParentID: 0, Name: "VBA", Type: 1}

	st := &vba.StorageTree{
		Nodes: []*vba.StorageNode{root, child},
		ByID: map[int32]*vba.StorageNode{
			0: root,
			1: child,
		},
		Children: map[int32][]*vba.StorageNode{
			-1: {root},
			0:  {child},
		},
	}

	return st
}

func TestPrintStorageTree_nilRoot(t *testing.T) {
	st := &vba.StorageTree{
		Nodes:    nil,
		ByID:     map[int32]*vba.StorageNode{},
		Children: map[int32][]*vba.StorageNode{},
	}

	var buf bytes.Buffer
	printStorageTree(&buf, st)

	if !strings.Contains(buf.String(), "no root") {
		t.Errorf("expected 'no root' message, got: %q", buf.String())
	}
}

func TestPrintStorageTree_withNodes(t *testing.T) {
	st := makeMinimalStorageTree()

	var buf bytes.Buffer
	printStorageTree(&buf, st)

	out := buf.String()

	if !strings.Contains(out, "storageTree:") {
		t.Errorf("expected storageTree heading, got: %q", out)
	}

	if !strings.Contains(out, "VBA") {
		t.Errorf("expected child node name in output, got: %q", out)
	}
}

func TestPrintForensic_withEmptyTree(t *testing.T) {
	st := &vba.StorageTree{
		Nodes:    nil,
		ByID:     map[int32]*vba.StorageNode{},
		Children: map[int32][]*vba.StorageNode{},
	}

	var buf bytes.Buffer
	printForensic(&buf, st)

	out := buf.String()
	if !strings.Contains(out, "forensic:") {
		t.Errorf("expected forensic: header, got: %q", out)
	}

	if !strings.Contains(out, "hits=0") {
		t.Errorf("expected hits=0 for empty tree, got: %q", out)
	}
}
