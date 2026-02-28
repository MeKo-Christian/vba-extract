package vba

import (
	"math"
	"testing"
)

// asInt32 – covers all type branches

func TestAsInt32_int32(t *testing.T) {
	v, ok := asInt32(int32(5))
	if !ok || v != 5 {
		t.Errorf("int32: got %d %v", v, ok)
	}
}

func TestAsInt32_int16(t *testing.T) {
	v, ok := asInt32(int16(200))
	if !ok || v != 200 {
		t.Errorf("int16: got %d %v", v, ok)
	}
}

func TestAsInt32_int_inRange(t *testing.T) {
	v, ok := asInt32(int(999))
	if !ok || v != 999 {
		t.Errorf("int: got %d %v", v, ok)
	}
}

func TestAsInt32_int_tooLarge(t *testing.T) {
	_, ok := asInt32(int(math.MaxInt64))
	if ok {
		t.Error("expected false for oversized int")
	}
}

func TestAsInt32_uint16(t *testing.T) {
	v, ok := asInt32(uint16(60000))
	if !ok || v != 60000 {
		t.Errorf("uint16: got %d %v", v, ok)
	}
}

func TestAsInt32_uint32_inRange(t *testing.T) {
	v, ok := asInt32(uint32(100))
	if !ok || v != 100 {
		t.Errorf("uint32: got %d %v", v, ok)
	}
}

func TestAsInt32_uint32_overflow(t *testing.T) {
	_, ok := asInt32(uint32(math.MaxUint32))
	if ok {
		t.Error("expected false for uint32 > MaxInt32")
	}
}

func TestAsInt32_int64_inRange(t *testing.T) {
	v, ok := asInt32(int64(42))
	if !ok || v != 42 {
		t.Errorf("int64: got %d %v", v, ok)
	}
}

func TestAsInt32_int64_overflow(t *testing.T) {
	_, ok := asInt32(int64(math.MaxInt64))
	if ok {
		t.Error("expected false for int64 > MaxInt32")
	}
}

func TestAsInt32_uint64_inRange(t *testing.T) {
	v, ok := asInt32(uint64(7))
	if !ok || v != 7 {
		t.Errorf("uint64: got %d %v", v, ok)
	}
}

func TestAsInt32_uint64_overflow(t *testing.T) {
	_, ok := asInt32(uint64(math.MaxUint64))
	if ok {
		t.Error("expected false for uint64 > MaxInt32")
	}
}

func TestAsInt32_unknownType(t *testing.T) {
	_, ok := asInt32("hello")
	if ok {
		t.Error("expected false for string type")
	}
}

// Root – covers all code paths

func TestRoot_byParentIDZero(t *testing.T) {
	// Standard: root has parentID == 0
	root := &StorageNode{ID: 1, ParentID: 0, Name: "Root"}
	child := &StorageNode{ID: 2, ParentID: 1, Name: "Child"}

	st := &StorageTree{
		Nodes:    []*StorageNode{root, child},
		ByID:     map[int32]*StorageNode{1: root, 2: child},
		Children: map[int32][]*StorageNode{},
	}

	got := st.Root()
	if got != root {
		t.Errorf("expected Root node, got %v", got)
	}
}

func TestRoot_byName(t *testing.T) {
	// No parentID==0; fall back to name "ROOT"
	n1 := &StorageNode{ID: 1, ParentID: 99, Name: "ROOT"}
	n2 := &StorageNode{ID: 2, ParentID: 99, Name: "Child"}

	st := &StorageTree{
		Nodes:    []*StorageNode{n1, n2},
		ByID:     map[int32]*StorageNode{1: n1, 2: n2},
		Children: map[int32][]*StorageNode{},
	}

	got := st.Root()
	if got != n1 {
		t.Errorf("expected ROOT node, got %v", got)
	}
}

func TestRoot_fallbackSmallestID(t *testing.T) {
	// No parentID==0 and no "ROOT" name → pick smallest ID
	n5 := &StorageNode{ID: 5, ParentID: 99, Name: "Alpha"}
	n2 := &StorageNode{ID: 2, ParentID: 99, Name: "Beta"}

	st := &StorageTree{
		Nodes:    []*StorageNode{n5, n2},
		ByID:     map[int32]*StorageNode{5: n5, 2: n2},
		Children: map[int32][]*StorageNode{},
	}

	got := st.Root()
	if got != n2 {
		t.Errorf("expected node with smallest ID (2), got ID=%d", got.ID)
	}
}

func TestRoot_emptyTree(t *testing.T) {
	st := &StorageTree{
		Nodes:    nil,
		ByID:     map[int32]*StorageNode{},
		Children: map[int32][]*StorageNode{},
	}

	if got := st.Root(); got != nil {
		t.Errorf("expected nil for empty tree, got %v", got)
	}
}

// isLikelyModuleStreamName

func TestIsLikelyModuleStreamName_empty(t *testing.T) {
	if isLikelyModuleStreamName("") {
		t.Error("expected false for empty name")
	}
}

func TestIsLikelyModuleStreamName_tooShort(t *testing.T) {
	if isLikelyModuleStreamName("SHORT") {
		t.Error("expected false for name shorter than 12 chars")
	}
}

func TestIsLikelyModuleStreamName_containsVBA(t *testing.T) {
	if isLikelyModuleStreamName("SOMEVBASTREAM") {
		t.Error("expected false for name containing VBA")
	}
}

func TestIsLikelyModuleStreamName_containsPROJECT(t *testing.T) {
	if isLikelyModuleStreamName("MYPROJECTDATA") {
		t.Error("expected false for name containing PROJECT")
	}
}

func TestIsLikelyModuleStreamName_validStream(t *testing.T) {
	// 12+ chars, all uppercase letters, no VBA/PROJECT
	if !isLikelyModuleStreamName("MODULESTREAMX") {
		t.Error("expected true for valid module stream name")
	}
}

func TestIsLikelyModuleStreamName_withUnderscore(t *testing.T) {
	if !isLikelyModuleStreamName("MODULE_STREAM") {
		t.Error("expected true for name with underscore")
	}
}

func TestIsLikelyModuleStreamName_withDigit(t *testing.T) {
	// Digits not allowed
	if isLikelyModuleStreamName("MODULESTREAM1") {
		t.Error("expected false for name with digit")
	}
}

// findByNameGlobal

func TestFindByNameGlobal_findsNode(t *testing.T) {
	target := &StorageNode{ID: 1, Name: "target", Data: []byte{1, 2, 3}}
	other := &StorageNode{ID: 2, Name: "other", Data: []byte{}}

	st := &StorageTree{
		Nodes:    []*StorageNode{target, other},
		ByID:     map[int32]*StorageNode{1: target, 2: other},
		Children: map[int32][]*StorageNode{},
	}

	got := st.findByNameGlobal("TARGET") // case-insensitive
	if got != target {
		t.Errorf("expected target node, got %v", got)
	}
}

func TestFindByNameGlobal_prefersLargerData(t *testing.T) {
	small := &StorageNode{ID: 1, Name: "dir", Data: []byte{1}}
	large := &StorageNode{ID: 2, Name: "dir", Data: []byte{1, 2, 3, 4, 5}}

	st := &StorageTree{
		Nodes:    []*StorageNode{small, large},
		ByID:     map[int32]*StorageNode{1: small, 2: large},
		Children: map[int32][]*StorageNode{},
	}

	got := st.findByNameGlobal("dir")
	if got != large {
		t.Errorf("expected node with larger data, got ID=%d", got.ID)
	}
}

func TestFindByNameGlobal_notFound(t *testing.T) {
	st := &StorageTree{
		Nodes:    []*StorageNode{{ID: 1, Name: "something"}},
		ByID:     map[int32]*StorageNode{},
		Children: map[int32][]*StorageNode{},
	}

	if got := st.findByNameGlobal("missing"); got != nil {
		t.Errorf("expected nil for missing name, got %v", got)
	}
}
