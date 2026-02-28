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

func TestPrintStorageTree_cycleDetection(t *testing.T) {
	// Indirect cycle: Root -> A -> B -> A
	root := &vba.StorageNode{ID: 0, ParentID: -1, Name: "Root", Type: 1}
	nodeA := &vba.StorageNode{ID: 1, ParentID: 0, Name: "NodeA", Type: 1}
	nodeB := &vba.StorageNode{ID: 2, ParentID: 1, Name: "NodeB", Type: 1}

	st := &vba.StorageTree{
		Nodes: []*vba.StorageNode{root, nodeA, nodeB},
		ByID: map[int32]*vba.StorageNode{
			0: root,
			1: nodeA,
			2: nodeB,
		},
		Children: map[int32][]*vba.StorageNode{
			-1: {root},
			0:  {nodeA},
			1:  {nodeB},
			2:  {nodeA}, // cycle: B -> A (A is already being visited)
		},
	}

	var buf bytes.Buffer
	// Should not hang or panic
	printStorageTree(&buf, st)

	out := buf.String()
	if !strings.Contains(out, "[cycle]") {
		t.Errorf("expected cycle detection in output, got: %q", out)
	}
}

func TestPrintForensic_manyHitsTruncates(t *testing.T) {
	// Build a tree with 30 nodes that all contain VBA-like project text
	// so ForensicScanStorage generates >25 hits (limit in printForensic)
	nodes := make([]*vba.StorageNode, 30)
	byID := make(map[int32]*vba.StorageNode, 30)
	children := make(map[int32][]*vba.StorageNode)

	root := &vba.StorageNode{ID: 0, ParentID: -1, Name: "Root", Type: 1}
	nodes[0] = root
	byID[0] = root

	// PROJECT text triggers ForensicProjectText hits
	projectData := []byte("ID=proj\r\nModule=Mod\r\nCMG=\"\"\r\n")

	for i := 1; i < 30; i++ {
		n := &vba.StorageNode{
			ID:       int32(i),
			ParentID: 0,
			Name:     "PROJECT",
			Type:     2,
			Data:     projectData,
		}
		nodes[i] = n
		byID[int32(i)] = n
		children[0] = append(children[0], n)
	}

	st := &vba.StorageTree{
		Nodes:    nodes,
		ByID:     byID,
		Children: children,
	}

	var buf bytes.Buffer
	printForensic(&buf, st)

	out := buf.String()
	if !strings.Contains(out, "more hit(s)") {
		t.Errorf("expected truncation message for >25 hits, got: %q", out)
	}
}
