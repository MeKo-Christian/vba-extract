package mdb

import (
	"testing"
)

// openJet4DB returns a minimal Jet4 Database for unit testing memo parsing.
// It has no real file pages — page access will error, but the memo
// short-circuit paths (inline, len<12) don't need page I/O.
func openJet4DB(t *testing.T) *Database {
	t.Helper()
	// Use the real sample.mdb so we have an open Jet4 db for inline tests.
	db := testDB(t)

	return db
}

// ResolveMemo – nil/empty input

func TestResolveMemo_nilReturnsNil(t *testing.T) {
	db := openJet4DB(t)

	got, err := db.ResolveMemo(nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if got != nil {
		t.Errorf("expected nil for nil input, got %v", got)
	}
}

func TestResolveMemo_emptySliceReturnsNil(t *testing.T) {
	db := openJet4DB(t)

	got, err := db.ResolveMemo([]byte{})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if got != nil {
		t.Errorf("expected nil for empty slice, got %v", got)
	}
}

// ResolveMemo – short inline data (< 12 bytes)

func TestResolveMemo_shortInlineData(t *testing.T) {
	db := openJet4DB(t)

	payload := []byte{0xAA, 0xBB, 0xCC}

	got, err := db.ResolveMemo(payload)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(got) != len(payload) {
		t.Fatalf("expected %d bytes, got %d", len(payload), len(got))
	}

	for i, b := range payload {
		if got[i] != b {
			t.Errorf("byte[%d]: expected %#x, got %#x", i, b, got[i])
		}
	}
}

// ResolveMemo – LvalInline bitmask (0x80)

func TestResolveMemo_lvalInline(t *testing.T) {
	db := openJet4DB(t)

	// 12-byte header + inline data
	// Header: memoLen=5, bitmask=0x80 (LvalInline), rest=0
	inlineData := []byte("hello")
	raw := make([]byte, 12+len(inlineData))
	raw[0] = byte(len(inlineData)) // memoLen low byte
	raw[3] = LvalInline
	copy(raw[12:], inlineData)

	got, err := db.ResolveMemo(raw)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if string(got) != "hello" {
		t.Errorf("expected 'hello', got %q", string(got))
	}
}

// ResolveMemo – unknown bitmask, valid inline fallback

func TestResolveMemo_unknownBitmask_withInlineData(t *testing.T) {
	db := openJet4DB(t)

	// bitmask = 0x10 (unknown), memoLen=3, followed by 3 bytes at offset 12
	payload := []byte("abc")
	raw := make([]byte, 12+len(payload))
	raw[0] = byte(len(payload)) // memoLen
	raw[3] = 0x10               // unknown bitmask
	copy(raw[12:], payload)

	got, err := db.ResolveMemo(raw)
	if err != nil {
		t.Fatalf("expected inline fallback to succeed: %v", err)
	}

	if string(got) != "abc" {
		t.Errorf("expected 'abc', got %q", string(got))
	}
}

// ResolveMemo – unknown bitmask, no valid inline data → error

func TestResolveMemo_unknownBitmask_noInlineData(t *testing.T) {
	db := openJet4DB(t)

	// bitmask = 0x10 (unknown), memoLen=100, but no data after header → error
	raw := make([]byte, 12) // only 12 bytes, no inline data
	raw[0] = 100            // memoLen = 100 but no data
	raw[3] = 0x10           // unknown bitmask

	_, err := db.ResolveMemo(raw)
	if err == nil {
		t.Error("expected error for unknown bitmask with no valid inline data")
	}
}

// ReadLvalChain – zero pageNum stops immediately

func TestReadLvalChain_zeroPage(t *testing.T) {
	db := openJet4DB(t)

	// pageNum=0 should stop the chain loop immediately and return empty
	got, err := db.ReadLvalChain(0, 0, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(got) != 0 {
		t.Errorf("expected empty result for zero page, got %d bytes", len(got))
	}
}
