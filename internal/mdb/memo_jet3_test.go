package mdb

import (
	"encoding/binary"
	"os"
	"testing"
)

func writeSyntheticJet3WithLvalRows(t *testing.T, rows map[int][]byte) string {
	t.Helper()

	path := writeSyntheticDB(t, 2048, 8, JetVersion3)

	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}

	for pageNum, row := range rows {
		pageStart := pageNum * 2048
		page := data[pageStart : pageStart+2048]
		page[0] = PageTypeData
		binary.LittleEndian.PutUint32(page[dataTDefPage:], 2)
		binary.LittleEndian.PutUint16(page[dataNumRowsJet3:], 1)

		rowStart := len(page) - len(row)
		copy(page[rowStart:], row)
		binary.LittleEndian.PutUint16(page[dataRowTableJet3:], uint16(rowStart))
	}

	err = os.WriteFile(path, data, 0o600)
	if err != nil {
		t.Fatalf("WriteFile: %v", err)
	}

	return path
}

func buildLvalRef(bitmask byte, pageNum int, rowID int, memoLen int, inlineData []byte) []byte {
	ref := make([]byte, 12)
	ref[0] = byte(memoLen)
	ref[1] = byte(memoLen >> 8)
	ref[2] = byte(memoLen >> 16)
	ref[3] = bitmask
	binary.LittleEndian.PutUint32(ref[4:], uint32(pageNum<<8|rowID))
	if bitmask == LvalInline && len(inlineData) > 0 {
		ref = append(ref, inlineData...)
	}

	return ref
}

func TestResolveMemoJet3Inline(t *testing.T) {
	path := writeSyntheticDB(t, 2048, 4, JetVersion3)

	db, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	t.Cleanup(func() { _ = db.Close() })

	raw := buildLvalRef(LvalInline, 0, 0, 5, []byte("HELLO"))
	got, err := db.ResolveMemo(raw)
	if err != nil {
		t.Fatalf("ResolveMemo: %v", err)
	}

	if string(got) != "HELLO" {
		t.Fatalf("ResolveMemo inline = %q, want %q", string(got), "HELLO")
	}
}

func TestResolveMemoJet3SinglePage(t *testing.T) {
	path := writeSyntheticJet3WithLvalRows(t, map[int][]byte{
		3: []byte("HELLO"),
	})

	db, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	t.Cleanup(func() { _ = db.Close() })

	raw := buildLvalRef(LvalSingle, 3, 0, 5, nil)
	got, err := db.ResolveMemo(raw)
	if err != nil {
		t.Fatalf("ResolveMemo: %v", err)
	}

	if string(got) != "HELLO" {
		t.Fatalf("ResolveMemo single = %q, want %q", string(got), "HELLO")
	}
}

func TestResolveMemoJet3MultiPageChain(t *testing.T) {
	nextPtr := make([]byte, 4)
	binary.LittleEndian.PutUint32(nextPtr, uint32(4<<8))

	row3 := append(nextPtr, []byte("AB")...)
	row4 := append([]byte{0, 0, 0, 0}, []byte("CD")...)

	path := writeSyntheticJet3WithLvalRows(t, map[int][]byte{
		3: row3,
		4: row4,
	})

	db, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	t.Cleanup(func() { _ = db.Close() })

	raw := buildLvalRef(LvalMultiPage, 3, 0, 4, nil)
	got, err := db.ResolveMemo(raw)
	if err != nil {
		t.Fatalf("ResolveMemo: %v", err)
	}

	if string(got) != "ABCD" {
		t.Fatalf("ResolveMemo chain = %q, want %q", string(got), "ABCD")
	}
}
