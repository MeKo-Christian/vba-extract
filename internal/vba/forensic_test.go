package vba

import "testing"

func TestForensicScanStorageStartMDB(t *testing.T) {
	db := startDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	report := ForensicScanStorage(st)
	if len(report.Hits) == 0 {
		t.Fatal("forensic scan returned no hits")
	}

	if report.ProjectCandidates == 0 {
		t.Fatal("expected at least one project candidate")
	}

	if report.DirCandidates == 0 {
		t.Fatal("expected at least one dir candidate")
	}
}

func TestSubtreeCycleSafe(t *testing.T) {
	st := &StorageTree{
		ByID:     map[int32]*StorageNode{},
		Children: map[int32][]*StorageNode{},
	}

	n := &StorageNode{ID: 1, ParentID: 1, Name: "MSysAccessStorage_ROOT", Type: 1}
	st.Nodes = []*StorageNode{n}
	st.ByID[1] = n
	st.Children[1] = []*StorageNode{n}

	out := st.subtree(1)
	if len(out) > 1 {
		t.Fatalf("cycle-safe subtree should not explode, got len=%d", len(out))
	}
}

func TestForensicDetectsAccessArtifacts(t *testing.T) {
	st := &StorageTree{}
	st.Nodes = []*StorageNode{
		{
			ID:   10,
			Name: "Blob",
			Type: 2,
			Data: []byte("Form_Order\nCaption=Order\nRecordSource=SELECT ID FROM Orders\nOnClick=Macro"),
		},
	}

	report := ForensicScanStorage(st)
	if report.ArtifactCandidates == 0 {
		t.Fatal("expected artifact candidate to be detected")
	}

	found := false

	for _, hit := range report.Hits {
		if hit.Kind == ForensicAccessArtifact {
			found = true
			break
		}
	}

	if !found {
		t.Fatal("artifact hit kind not present in report")
	}
}
