package vba

import (
	"encoding/hex"
	"encoding/binary"
	"strings"
	"testing"
)

func buildContainerChunk(compressed bool, payload []byte) []byte {
	chunkSize := len(payload) + 2
	base := uint16(0x3000 | uint16(chunkSize-3))
	if compressed {
		base |= 0x8000
	}

	h := make([]byte, 2)
	binary.LittleEndian.PutUint16(h, base)
	return append(h, payload...)
}

func TestDecompressContainerRejectsInvalidSignature(t *testing.T) {
	_, err := DecompressContainer([]byte{0x00, 0x00, 0x00})
	if err == nil {
		t.Fatal("expected error for invalid signature")
	}
}

func TestDecompressContainerRejectsEmpty(t *testing.T) {
	_, err := DecompressContainer(nil)
	if err == nil {
		t.Fatal("expected error for empty input")
	}
}

func TestDecompressContainerKnownVectors(t *testing.T) {
	vectors := []struct {
		name    string
		inHex   string
		outHex  string
	}{
		{
			name:   "uncompressed-hello",
			inHex:  "01043068656c6c6f",
			outHex: "68656c6c6f",
		},
		{
			name:   "compressed-abc-repeat",
			inHex:  "0105b0084142430320",
			outHex: "414243414243414243",
		},
		{
			name:   "two-uncompressed-chunks",
			inHex:  "010230666f6f0230626172",
			outHex: "666f6f626172",
		},
	}

	for _, vector := range vectors {
		t.Run(vector.name, func(t *testing.T) {
			in, err := hex.DecodeString(vector.inHex)
			if err != nil {
				t.Fatalf("decode input hex: %v", err)
			}
			want, err := hex.DecodeString(vector.outHex)
			if err != nil {
				t.Fatalf("decode output hex: %v", err)
			}

			got, err := DecompressContainer(in)
			if err != nil {
				t.Fatalf("DecompressContainer: %v", err)
			}
			if hex.EncodeToString(got) != vector.outHex {
				t.Fatalf("output mismatch: got=%x want=%x", got, want)
			}
		})
	}
}

func TestDecompressContainerUncompressedChunk(t *testing.T) {
	payload := []byte("hello")
	data := []byte{0x01}
	data = append(data, buildContainerChunk(false, payload)...)

	out, err := DecompressContainer(data)
	if err != nil {
		t.Fatalf("DecompressContainer: %v", err)
	}

	if string(out) != "hello" {
		t.Fatalf("out = %q, want %q", string(out), "hello")
	}
}

func TestDecompressContainerCompressedChunk(t *testing.T) {
	// Tokens: 'A' 'B' 'C' then copy token offset=3 length=6 -> ABCABCABC
	payload := []byte{0x08, 'A', 'B', 'C', 0x03, 0x20}
	data := []byte{0x01}
	data = append(data, buildContainerChunk(true, payload)...)

	out, err := DecompressContainer(data)
	if err != nil {
		t.Fatalf("DecompressContainer: %v", err)
	}

	if string(out) != "ABCABCABC" {
		t.Fatalf("out = %q, want %q", string(out), "ABCABCABC")
	}
}

func TestDecompressContainerMultiChunk(t *testing.T) {
	data := []byte{0x01}
	data = append(data, buildContainerChunk(false, []byte("foo"))...)
	data = append(data, buildContainerChunk(false, []byte("bar"))...)

	out, err := DecompressContainer(data)
	if err != nil {
		t.Fatalf("DecompressContainer: %v", err)
	}

	if string(out) != "foobar" {
		t.Fatalf("out = %q, want %q", string(out), "foobar")
	}
}

func TestDecompressContainerWithFallbackSkipPrefix(t *testing.T) {
	container := []byte{0x01}
	container = append(container, buildContainerChunk(false, []byte("ok"))...)
	input := append([]byte{0xAA, 0xBB, 0xCC}, container...)

	out, strategy, err := DecompressContainerWithFallback(input, false, nil)
	if err != nil {
		t.Fatalf("DecompressContainerWithFallback: %v", err)
	}
	if strategy != StrategySkipPrefix {
		t.Fatalf("strategy = %q, want %q", strategy, StrategySkipPrefix)
	}
	if string(out) != "ok" {
		t.Fatalf("out = %q, want %q", string(out), "ok")
	}
}

func TestDecompressContainerWithFallbackRaw(t *testing.T) {
	input := []byte{0xAA, 0xBB, 0xCC}
	out, strategy, err := DecompressContainerWithFallback(input, false, nil)
	if err != nil {
		t.Fatalf("DecompressContainerWithFallback: %v", err)
	}
	if strategy != StrategyRawPassthru {
		t.Fatalf("strategy = %q, want %q", strategy, StrategyRawPassthru)
	}
	if len(out) != len(input) {
		t.Fatalf("len(out) = %d, want %d", len(out), len(input))
	}
}

func TestDecompressContainerWithFallbackVerboseLog(t *testing.T) {
	container := []byte{0x01}
	container = append(container, buildContainerChunk(false, []byte("v"))...)

	var logs []string
	_, strategy, err := DecompressContainerWithFallback(container, true, func(format string, args ...interface{}) {
		logs = append(logs, format)
	})
	if err != nil {
		t.Fatalf("DecompressContainerWithFallback: %v", err)
	}
	if strategy != StrategyStandard {
		t.Fatalf("strategy = %q, want %q", strategy, StrategyStandard)
	}
	if len(logs) == 0 {
		t.Fatal("expected verbose logs")
	}
}

func TestDirStreamRegressionFromStartMDB(t *testing.T) {
	db := testDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}
	required, err := st.RequiredStreams()
	if err != nil {
		t.Fatalf("RequiredStreams: %v", err)
	}
	dirNode := required["dir"]
	if dirNode == nil || len(dirNode.Data) == 0 {
		t.Skip("dir stream missing in fixture")
	}

	dec, _, err := DecompressContainerWithFallback(dirNode.Data, false, nil)
	if err != nil {
		t.Fatalf("DecompressContainerWithFallback(dir): %v", err)
	}
	if len(dec) == 0 {
		t.Fatal("decompressed dir stream is empty")
	}

	info, err := parseDirRecords(dec)
	if err != nil {
		t.Fatalf("parseDirRecords: %v", err)
	}
	if len(info.Modules) != 15 {
		t.Errorf("modules = %d, want 15", len(info.Modules))
	}

	for _, module := range info.Modules {
		if strings.TrimSpace(module.ModuleName) == "" {
			t.Errorf("module with empty name: %+v", module)
		}
		if strings.TrimSpace(module.StreamName) == "" {
			t.Errorf("module %q has empty StreamName", module.ModuleName)
		}
	}
}

func TestBitCountForDecompressedPos(t *testing.T) {
	cases := []struct {
		pos  int
		want int
	}{
		{0, 4},
		{16, 4},
		{17, 5},
		{32, 5},
		{33, 6},
		{64, 6},
		{65, 7},
		{128, 7},
		{129, 8},
		{256, 8},
		{257, 9},
		{512, 9},
		{513, 10},
		{1024, 10},
		{1025, 11},
		{2048, 11},
		{2049, 12},
		{4096, 12},
	}

	for _, tc := range cases {
		if got := bitCountForDecompressedPos(tc.pos); got != tc.want {
			t.Fatalf("bitCountForDecompressedPos(%d) = %d, want %d", tc.pos, got, tc.want)
		}
	}
}
