# Phase 3 + 4: VBA Project Structure Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Parse the PROJECT and dir streams extracted from MSysAccessStorage to build a complete module→stream mapping, enabling VBA source code extraction in Phase 5.

**Architecture:** Phase 4 (MS-OVBA decompression) is implemented first because the dir stream is compressed. Phase 3 uses Phase 4 to decompress the dir stream, then parses the PROJECT stream for module type classification and the dir stream for the stream name / source offset mapping.

**Tech Stack:** Go stdlib only. No external dependencies. Test against `testdata/Start.mdb` (15 modules, available at project root).

---

## Task 1: MS-OVBA Decompression (Phase 4)

**Files:**
- Create: `internal/vba/decompress.go`
- Create: `internal/vba/decompress_test.go`

---

### Step 1.1: Write the failing tests

Create `internal/vba/decompress_test.go`:

```go
package vba

import (
	"bytes"
	"testing"
)

// minCompressedContainer builds a synthetic MS-OVBA compressed container.
// The input is encoded as all-literal tokens (flag byte 0x00 per 8 bytes).
func buildLiteralContainer(data []byte) []byte {
	var buf bytes.Buffer
	buf.WriteByte(0x01) // signature

	// Write as one compressed chunk (all literals, flag byte = 0x00).
	// One flag byte covers 8 literal tokens; chunk data = flagBytes + literalBytes.
	var chunkData bytes.Buffer
	for i := 0; i < len(data); i += 8 {
		end := i + 8
		if end > len(data) {
			end = len(data)
		}
		chunkData.WriteByte(0x00) // all 8 tokens are literals
		chunkData.Write(data[i:end])
	}

	compressed := chunkData.Bytes()
	// Chunk header: bit 15=1 (compressed), bits 12-14=011, bits 0-11 = len(compressed)-3
	size := uint16(len(compressed) - 3)
	header := uint16(0x8000) | uint16(0x3000) | (size & 0x0FFF)
	buf.WriteByte(byte(header))
	buf.WriteByte(byte(header >> 8))
	buf.Write(compressed)

	return buf.Bytes()
}

func TestDecompress_Signature(t *testing.T) {
	_, err := DecompressContainer([]byte{0x02, 0x00}) // wrong signature
	if err == nil {
		t.Fatal("expected error for invalid signature, got nil")
	}
}

func TestDecompress_Empty(t *testing.T) {
	_, err := DecompressContainer([]byte{0x01}) // only signature, no chunks
	if err != nil {
		t.Fatalf("empty container after signature: %v", err)
	}
}

func TestDecompress_LiteralChunk(t *testing.T) {
	want := []byte("Hello, World!")
	compressed := buildLiteralContainer(want)

	got, err := DecompressContainer(compressed)
	if err != nil {
		t.Fatalf("DecompressContainer: %v", err)
	}
	if !bytes.Equal(got, want) {
		t.Errorf("got %q, want %q", got, want)
	}
}

func TestDecompress_LargeLiteralChunk(t *testing.T) {
	// 200 bytes of incrementing values to exercise chunk boundaries.
	want := make([]byte, 200)
	for i := range want {
		want[i] = byte(i)
	}
	compressed := buildLiteralContainer(want)

	got, err := DecompressContainer(compressed)
	if err != nil {
		t.Fatalf("DecompressContainer: %v", err)
	}
	if !bytes.Equal(got, want) {
		t.Errorf("decompressed %d bytes, want %d; first mismatch at position check", len(got), len(want))
	}
}

func TestDecompress_CopyToken(t *testing.T) {
	// Craft a compressed chunk that uses a copy token for repeated content.
	// Decompressed target: "AAAAAAAAAAAAAAAA" (16 'A' bytes)
	// Strategy: 1 literal 'A', then copy token [offset=1, length=15]
	//
	// At pos=1: bitCount=4, lengthMask=(1<<12)-1=0x0FFF, offsetMask=0xF000
	// length = (token & 0x0FFF) + 3 → we want 15 → token & 0x0FFF = 12
	// offset = (token >> 12) + 1 → we want 1 → token >> 12 = 0
	// token = 0x000C
	want := bytes.Repeat([]byte("A"), 16)

	var chunkData bytes.Buffer
	chunkData.WriteByte(0x02) // flag: bit0=0 (literal), bit1=1 (copy token), rest=0
	chunkData.WriteByte('A')   // literal byte
	// copy token 0x000C little-endian
	chunkData.WriteByte(0x0C)
	chunkData.WriteByte(0x00)

	compressed := chunkData.Bytes()
	size := uint16(len(compressed) - 3)
	header := uint16(0x8000) | uint16(0x3000) | (size & 0x0FFF)

	var container bytes.Buffer
	container.WriteByte(0x01)
	container.WriteByte(byte(header))
	container.WriteByte(byte(header >> 8))
	container.Write(compressed)

	got, err := DecompressContainer(container.Bytes())
	if err != nil {
		t.Fatalf("DecompressContainer: %v", err)
	}
	if !bytes.Equal(got, want) {
		t.Errorf("got %q, want %q", got, want)
	}
}

func TestDecompress_DirStream(t *testing.T) {
	// Regression test: decompress the actual dir stream from Start.mdb.
	// We rely on TestLoadStorageTree passing; skip if MDB is unavailable.
	db := testDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	required, err := st.RequiredStreams()
	if err != nil {
		t.Fatalf("RequiredStreams: %v", err)
	}

	dirNode, ok := required["dir"]
	if !ok || dirNode == nil {
		t.Fatal("dir stream node not found")
	}

	if len(dirNode.Data) == 0 {
		t.Fatal("dir stream data is empty")
	}

	decompressed, err := DecompressContainer(dirNode.Data)
	if err != nil {
		t.Fatalf("DecompressContainer(dir): %v", err)
	}

	if len(decompressed) == 0 {
		t.Fatal("decompressed dir stream is empty")
	}

	t.Logf("dir stream: %d compressed → %d decompressed bytes", len(dirNode.Data), len(decompressed))
}
```

**Step 1.2: Run tests to verify they fail**

```bash
cd /mnt/projekte/Code/vba-extract
go test ./internal/vba/ -run TestDecompress -v
```
Expected: compile error (`DecompressContainer` undefined).

---

### Step 1.3: Implement DecompressContainer

Create `internal/vba/decompress.go`:

```go
package vba

import (
	"encoding/binary"
	"fmt"
)

// DecompressContainer decompresses an MS-OVBA compressed container (§2.4.1).
// The input must start with signature byte 0x01.
func DecompressContainer(data []byte) ([]byte, error) {
	if len(data) == 0 || data[0] != 0x01 {
		if len(data) == 0 {
			return nil, fmt.Errorf("ovba: empty input")
		}
		return nil, fmt.Errorf("ovba: invalid signature byte 0x%02X (want 0x01)", data[0])
	}

	var out []byte
	pos := 1 // skip signature byte

	for pos < len(data) {
		if pos+2 > len(data) {
			return nil, fmt.Errorf("ovba: truncated chunk header at offset %d", pos)
		}
		header := binary.LittleEndian.Uint16(data[pos : pos+2])
		pos += 2

		isCompressed := (header >> 15) != 0
		// bits 12-14 should be 0b011 (signature); we tolerate deviations.
		chunkSize := int(header&0x0FFF) + 3 // compressed data byte count

		if pos+chunkSize > len(data) {
			return nil, fmt.Errorf("ovba: chunk at offset %d claims %d bytes but only %d remain",
				pos-2, chunkSize, len(data)-pos)
		}

		if !isCompressed {
			// Raw chunk: always 4096 bytes follow.
			if chunkSize < 4096 && pos+4096 <= len(data) {
				out = append(out, data[pos:pos+4096]...)
				pos += 4096
			} else {
				out = append(out, data[pos:pos+chunkSize]...)
				pos += chunkSize
			}
			continue
		}

		// Compressed chunk.
		chunk, err := decompressChunk(data[pos : pos+chunkSize])
		if err != nil {
			return nil, fmt.Errorf("ovba: chunk at offset %d: %w", pos-2, err)
		}
		out = append(out, chunk...)
		pos += chunkSize
	}

	return out, nil
}

// decompressChunk decodes one compressed chunk (without its 2-byte header).
func decompressChunk(compressed []byte) ([]byte, error) {
	var out []byte
	src := 0

	for src < len(compressed) {
		flagByte := compressed[src]
		src++

		for bit := 0; bit < 8; bit++ {
			if src >= len(compressed) {
				break
			}

			if (flagByte>>bit)&1 == 0 {
				// Literal byte.
				out = append(out, compressed[src])
				src++
			} else {
				// Copy token.
				if src+1 >= len(compressed) {
					return nil, fmt.Errorf("ovba: truncated copy token at src=%d", src)
				}
				token := binary.LittleEndian.Uint16(compressed[src : src+2])
				src += 2

				bc := bitCountForDecompressedPos(len(out))
				lengthMask := uint16((1 << (16 - bc)) - 1)
				offsetMask := ^lengthMask

				length := int(token&lengthMask) + 3
				offset := int((token&offsetMask)>>(16-bc)) + 1

				start := len(out) - offset
				if start < 0 {
					return nil, fmt.Errorf("ovba: copy token offset %d exceeds decompressed length %d", offset, len(out))
				}

				for i := 0; i < length; i++ {
					out = append(out, out[start+i])
				}
			}
		}
	}

	return out, nil
}

// bitCountForDecompressedPos returns the number of offset bits for the copy
// token based on the current decompressed chunk length (MS-OVBA §2.4.1.3.19).
func bitCountForDecompressedPos(decompressedLength int) int {
	switch {
	case decompressedLength <= 16:
		return 4
	case decompressedLength <= 32:
		return 5
	case decompressedLength <= 64:
		return 6
	case decompressedLength <= 128:
		return 7
	case decompressedLength <= 256:
		return 8
	case decompressedLength <= 512:
		return 9
	case decompressedLength <= 1024:
		return 10
	case decompressedLength <= 2048:
		return 11
	default:
		return 12
	}
}
```

**Step 1.4: Run tests to verify they pass**

```bash
go test ./internal/vba/ -run TestDecompress -v
```
Expected: all `TestDecompress_*` tests PASS.

**Step 1.5: Commit**

```bash
git add internal/vba/decompress.go internal/vba/decompress_test.go
git commit -m "feat(vba): implement MS-OVBA §2.4.1 decompression"
```

---

## Task 2: PROJECT Stream Parser (Phase 3.1)

**Files:**
- Create: `internal/vba/project.go`
- Create: `internal/vba/project_test.go`

The PROJECT stream is a Windows-1252/latin-1 text file (CRLF line endings) stored in the `PROJECT` stream node of MSysAccessStorage. It is NOT compressed.

Format (typical Access database):
```
ID="{GUID}"
Document=Form_Module/&H00000000
Module=Inidatei
Class=MyClass
Name=Start
...
[Host Extender Info]
...
[Workspace]
...
```

Module types:
- `Document=<name>/&H<hexid>` → DocClass (forms/reports)
- `Module=<name>` → standard module (.bas)
- `Class=<name>` → standalone class module (.cls)

---

### Step 2.1: Write the failing tests

Create `internal/vba/project_test.go`:

```go
package vba

import (
	"strings"
	"testing"
)

func TestParseProject_Basic(t *testing.T) {
	raw := "ID=\"{ABCDEF}\"\r\nModule=Modul1\r\nModule=Update\r\nDocument=Form_Start/&H00000000\r\nName=MyProject\r\n"

	info, err := ParseProject([]byte(raw))
	if err != nil {
		t.Fatalf("ParseProject: %v", err)
	}

	if info.Name != "MyProject" {
		t.Errorf("Name = %q, want %q", info.Name, "MyProject")
	}

	if len(info.Modules) != 3 {
		t.Fatalf("got %d modules, want 3", len(info.Modules))
	}

	// Module order must match declaration order.
	if info.Modules[0].Name != "Modul1" || info.Modules[0].Type != ModuleTypeStandard {
		t.Errorf("modules[0] = %+v, want {Name:Modul1, Type:Standard}", info.Modules[0])
	}
	if info.Modules[1].Name != "Update" || info.Modules[1].Type != ModuleTypeStandard {
		t.Errorf("modules[1] = %+v", info.Modules[1])
	}
	if info.Modules[2].Name != "Form_Start" || info.Modules[2].Type != ModuleTypeDocument {
		t.Errorf("modules[2] = %+v, want {Name:Form_Start, Type:Document}", info.Modules[2])
	}
}

func TestParseProject_ClassModules(t *testing.T) {
	raw := "Module=Std\r\nClass=MyClass\r\n"

	info, err := ParseProject([]byte(raw))
	if err != nil {
		t.Fatalf("ParseProject: %v", err)
	}

	if len(info.Modules) != 2 {
		t.Fatalf("got %d modules, want 2", len(info.Modules))
	}
	if info.Modules[1].Type != ModuleTypeClass {
		t.Errorf("modules[1].Type = %v, want ModuleTypeClass", info.Modules[1].Type)
	}
}

func TestParseProject_IgnoresSections(t *testing.T) {
	raw := "Module=Foo\r\n[Host Extender Info]\r\n&H00000001={...}\r\n[Workspace]\r\nFoo=0, 0, 1024, 768\r\n"

	info, err := ParseProject([]byte(raw))
	if err != nil {
		t.Fatalf("ParseProject: %v", err)
	}

	if len(info.Modules) != 1 || info.Modules[0].Name != "Foo" {
		t.Errorf("modules = %+v", info.Modules)
	}
}

func TestParseProject_StartMDB(t *testing.T) {
	db := testDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	required, err := st.RequiredStreams()
	if err != nil {
		t.Fatalf("RequiredStreams: %v", err)
	}

	projectNode := required["PROJECT"]
	if projectNode == nil {
		t.Fatal("PROJECT stream not found")
	}

	info, err := ParseProject(projectNode.Data)
	if err != nil {
		t.Fatalf("ParseProject: %v", err)
	}

	if len(info.Modules) != 15 {
		t.Errorf("got %d modules, want 15", len(info.Modules))
	}

	// Check expected module names are present.
	wantModules := []string{
		"Form_Module", "Form_Übersicht", "Inidatei", "Menueaufbau",
		"Modul1", "SQL", "IPAdresse2", "mod_api_Functions", "Update",
	}
	names := make(map[string]bool, len(info.Modules))
	for _, m := range info.Modules {
		names[m.Name] = true
	}
	for _, want := range wantModules {
		if !names[want] {
			t.Errorf("expected module %q not found in parsed project", want)
		}
	}

	// Check type classification.
	docClasses := 0
	standards := 0
	for _, m := range info.Modules {
		switch m.Type {
		case ModuleTypeDocument:
			docClasses++
		case ModuleTypeStandard:
			standards++
		}
	}
	if docClasses != 6 {
		t.Errorf("got %d Document modules, want 6", docClasses)
	}
	if standards != 9 {
		t.Errorf("got %d standard modules, want 9", standards)
	}

	if !strings.Contains(info.Name, "Start") && info.Name != "" {
		t.Logf("project name = %q", info.Name)
	}

	t.Logf("Parsed PROJECT: name=%q, %d modules (%d DocClass, %d Standard)",
		info.Name, len(info.Modules), docClasses, standards)
}
```

**Step 2.2: Run tests to verify they fail**

```bash
go test ./internal/vba/ -run TestParseProject -v
```
Expected: compile error (`ParseProject` undefined, `ModuleType*` undefined).

---

### Step 2.3: Implement the PROJECT stream parser

Create `internal/vba/project.go`:

```go
package vba

import (
	"bufio"
	"bytes"
	"strings"
)

// ModuleType classifies a VBA module.
type ModuleType int

const (
	ModuleTypeStandard  ModuleType = iota // Module= key in PROJECT stream
	ModuleTypeClass                       // Class= key in PROJECT stream
	ModuleTypeDocument                    // Document= key (forms/reports in Access)
)

// ModuleDecl is one module entry parsed from the PROJECT stream.
type ModuleDecl struct {
	Name   string
	Type   ModuleType
	HostID string // only set for Document type; the &H... part after the slash
}

// ProjectInfo holds the parsed content of a VBA PROJECT stream.
type ProjectInfo struct {
	Name    string
	Modules []ModuleDecl
}

// ParseProject parses the PROJECT stream (latin-1/Windows-1252 text, CRLF).
// Only the module declarations and project name are extracted; all other
// metadata and section content is discarded.
func ParseProject(data []byte) (*ProjectInfo, error) {
	info := &ProjectInfo{}

	scanner := bufio.NewScanner(bytes.NewReader(data))
	inSection := false // true when inside a [...] section that is not the default

	for scanner.Scan() {
		line := strings.TrimRight(scanner.Text(), "\r")

		if strings.HasPrefix(line, "[") {
			// Section headers like [Host Extender Info] stop module parsing.
			inSection = true
			continue
		}

		if inSection {
			continue
		}

		key, value, ok := strings.Cut(line, "=")
		if !ok {
			continue
		}

		switch key {
		case "Name":
			info.Name = value

		case "Module":
			info.Modules = append(info.Modules, ModuleDecl{
				Name: value,
				Type: ModuleTypeStandard,
			})

		case "Class":
			info.Modules = append(info.Modules, ModuleDecl{
				Name: value,
				Type: ModuleTypeClass,
			})

		case "Document":
			// Format: <name>/&H<hexid>
			name, hostID, _ := strings.Cut(value, "/")
			info.Modules = append(info.Modules, ModuleDecl{
				Name:   name,
				Type:   ModuleTypeDocument,
				HostID: hostID,
			})
		}
	}

	return info, scanner.Err()
}
```

**Step 2.4: Run tests to verify they pass**

```bash
go test ./internal/vba/ -run TestParseProject -v
```
Expected: all `TestParseProject_*` tests PASS.

**Step 2.5: Commit**

```bash
git add internal/vba/project.go internal/vba/project_test.go
git commit -m "feat(vba): parse PROJECT stream for module declarations"
```

---

## Task 3: dir Stream Parser (Phase 3.2)

**Files:**
- Create: `internal/vba/dir.go`
- Create: `internal/vba/dir_test.go`

The `dir` stream is MS-OVBA compressed (Task 1). After decompression it contains a sequence of binary records:

```
Record {
    ID:   uint16 LE
    Size: uint32 LE
    Data: [Size]byte
}
```

Relevant record IDs (MS-OVBA §2.3.4.2):

| ID     | Name                      | Content |
|--------|---------------------------|---------|
| 0x000F | PROJECTMODULES            | marks start of module records |
| 0x0013 | PROJECTCOOKIE             | ignore |
| 0x0190 | PROJECTMODULECOUNT        | uint16 module count |
| 0x0019 | MODULENAME                | MBCS module name |
| 0x0047 | MODULENAMEUNICODE         | Unicode name (prefer over 0x0019) |
| 0x001A | MODULESTREAMNAME          | MBCS obfuscated stream name |
| 0x0032 | MODULESTREAMNAME_UNICODE  | Unicode stream name (prefer) |
| 0x0031 | MODULEOFFSET              | uint32 source offset |
| 0x0021 | MODULETYPE_PROCEDURAL     | standard module (size=0) |
| 0x0022 | MODULETYPE_CLASS          | class/document module (size=0) |
| 0x002B | MODULETERMINATOR          | end of module record (size=0) |
| 0x0010 | PROJECTMODULES_TERMINATOR | end of all modules |

All other records: read and skip `Size` bytes.

---

### Step 3.1: Write the failing tests

Create `internal/vba/dir_test.go`:

```go
package vba

import (
	"testing"
)

func TestParseDir_StartMDB(t *testing.T) {
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
	if dirNode == nil {
		t.Fatal("dir stream not found")
	}

	decompressed, err := DecompressContainer(dirNode.Data)
	if err != nil {
		t.Fatalf("DecompressContainer: %v", err)
	}

	dirInfo, err := ParseDir(decompressed)
	if err != nil {
		t.Fatalf("ParseDir: %v", err)
	}

	if len(dirInfo.Modules) != 15 {
		t.Errorf("got %d modules, want 15; modules: %v", len(dirInfo.Modules), moduleNames(dirInfo.Modules))
	}

	// Every module must have a non-empty stream name and a valid offset.
	for _, m := range dirInfo.Modules {
		if m.StreamName == "" {
			t.Errorf("module %q has empty StreamName", m.Name)
		}
		// offset=0 is technically valid, but unusual; warn rather than fail
		t.Logf("  module %-30s stream=%-35s offset=%d type=%d", m.Name, m.StreamName, m.Offset, m.Type)
	}
}

func TestParseDir_ReconcileWithProject(t *testing.T) {
	db := testDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	required, err := st.RequiredStreams()
	if err != nil {
		t.Fatalf("RequiredStreams: %v", err)
	}

	// Parse PROJECT.
	projectInfo, err := ParseProject(required["PROJECT"].Data)
	if err != nil {
		t.Fatalf("ParseProject: %v", err)
	}

	// Parse dir.
	decompressed, err := DecompressContainer(required["dir"].Data)
	if err != nil {
		t.Fatalf("DecompressContainer: %v", err)
	}
	dirInfo, err := ParseDir(decompressed)
	if err != nil {
		t.Fatalf("ParseDir: %v", err)
	}

	// Every name in the PROJECT stream must appear in the dir mapping.
	dirByName := make(map[string]*DirModule, len(dirInfo.Modules))
	for i := range dirInfo.Modules {
		dirByName[dirInfo.Modules[i].Name] = &dirInfo.Modules[i]
	}

	for _, pd := range projectInfo.Modules {
		if _, ok := dirByName[pd.Name]; !ok {
			t.Errorf("PROJECT module %q not found in dir mapping", pd.Name)
		}
	}
}

func moduleNames(mods []DirModule) []string {
	names := make([]string, len(mods))
	for i, m := range mods {
		names[i] = m.Name
	}
	return names
}
```

**Step 3.2: Run tests to verify they fail**

```bash
go test ./internal/vba/ -run TestParseDir -v
```
Expected: compile error (`ParseDir`, `DirInfo`, `DirModule` undefined).

---

### Step 3.3: Implement the dir stream parser

Create `internal/vba/dir.go`:

```go
package vba

import (
	"encoding/binary"
	"fmt"
	"unicode/utf16"
)

// DirModule holds the mapping extracted from one module record in the dir stream.
type DirModule struct {
	Name       string     // display name (from MODULENAME or MODULENAMEUNICODE)
	StreamName string     // obfuscated storage stream name (from MODULESTREAMNAME)
	Offset     uint32     // byte offset of compressed source in the module stream
	Type       ModuleType // ModuleTypeStandard or ModuleTypeClass/Document
}

// DirInfo is the result of parsing a decompressed dir stream.
type DirInfo struct {
	Modules []DirModule
}

const (
	recProjectModules          = 0x000F
	recProjectCookie           = 0x0013
	recProjectModuleCount      = 0x0190
	recModuleName              = 0x0019
	recModuleNameUnicode       = 0x0047
	recModuleStreamName        = 0x001A
	recModuleStreamNameUnicode = 0x0032
	recModuleOffset            = 0x0031
	recModuleTypeProcedural    = 0x0021
	recModuleTypeClass         = 0x0022
	recModuleTerminator        = 0x002B
	recModulesTerminator       = 0x0010
)

// ParseDir parses a decompressed MS-OVBA dir stream (MS-OVBA §2.3.4.2).
// Returns the list of modules with their stream name and source offset.
func ParseDir(data []byte) (*DirInfo, error) {
	r := &dirReader{data: data}
	info := &DirInfo{}

	// Scan forward until we find PROJECTMODULES (0x000F).
	for {
		id, payload, err := r.readRecord()
		if err != nil {
			return nil, fmt.Errorf("dir: scanning for PROJECTMODULES: %w", err)
		}
		if id == recProjectModules {
			_ = payload // PROJECTMODULES has no meaningful payload
			break
		}
		// Unknown records before modules section are silently skipped.
	}

	// Read PROJECTMODULECOUNT (0x0190) — next non-cookie record.
	var moduleCount int
	for {
		id, payload, err := r.readRecord()
		if err != nil {
			return nil, fmt.Errorf("dir: reading module count: %w", err)
		}
		if id == recProjectCookie {
			continue
		}
		if id == recProjectModuleCount {
			if len(payload) < 2 {
				return nil, fmt.Errorf("dir: PROJECTMODULECOUNT too short")
			}
			moduleCount = int(binary.LittleEndian.Uint16(payload))
			break
		}
		return nil, fmt.Errorf("dir: unexpected record 0x%04X before module count", id)
	}

	// Parse each module record.
	for i := 0; i < moduleCount; i++ {
		mod, err := parseModuleRecord(r)
		if err != nil {
			return nil, fmt.Errorf("dir: module %d: %w", i, err)
		}
		info.Modules = append(info.Modules, *mod)
	}

	return info, nil
}

func parseModuleRecord(r *dirReader) (*DirModule, error) {
	mod := &DirModule{Type: ModuleTypeStandard}

	for {
		id, payload, err := r.readRecord()
		if err != nil {
			return nil, err
		}

		switch id {
		case recModuleName:
			// MBCS name; may be overridden by MODULENAMEUNICODE below.
			if mod.Name == "" {
				mod.Name = string(payload)
			}

		case recModuleNameUnicode:
			// UTF-16LE name; preferred over MBCS name.
			if len(payload)%2 == 0 {
				u16 := make([]uint16, len(payload)/2)
				for i := range u16 {
					u16[i] = binary.LittleEndian.Uint16(payload[i*2:])
				}
				mod.Name = string(utf16.Decode(u16))
			}

		case recModuleStreamName:
			// MBCS obfuscated stream name; may be overridden by UNICODE variant.
			if mod.StreamName == "" {
				mod.StreamName = string(payload)
			}

		case recModuleStreamNameUnicode:
			// UTF-16LE stream name; preferred.
			if len(payload)%2 == 0 {
				u16 := make([]uint16, len(payload)/2)
				for i := range u16 {
					u16[i] = binary.LittleEndian.Uint16(payload[i*2:])
				}
				mod.StreamName = string(utf16.Decode(u16))
			}

		case recModuleOffset:
			if len(payload) < 4 {
				return nil, fmt.Errorf("MODULEOFFSET too short (%d bytes)", len(payload))
			}
			mod.Offset = binary.LittleEndian.Uint32(payload)

		case recModuleTypeProcedural:
			mod.Type = ModuleTypeStandard

		case recModuleTypeClass:
			mod.Type = ModuleTypeClass

		case recModuleTerminator:
			return mod, nil

		// All other records (DOCSTRING, HELPTOPIC, COOKIE, READONLY, PRIVATE, …) are skipped.
		}
	}
}

// dirReader reads sequential dir records from a byte slice.
type dirReader struct {
	data []byte
	pos  int
}

func (r *dirReader) readRecord() (id uint16, payload []byte, err error) {
	if r.pos+6 > len(r.data) {
		return 0, nil, fmt.Errorf("dir: unexpected end of data at offset %d (need record header)", r.pos)
	}

	id = binary.LittleEndian.Uint16(r.data[r.pos:])
	size := binary.LittleEndian.Uint32(r.data[r.pos+2:])
	r.pos += 6

	if uint64(r.pos)+uint64(size) > uint64(len(r.data)) {
		return 0, nil, fmt.Errorf("dir: record 0x%04X at offset %d claims %d bytes but only %d remain",
			id, r.pos-6, size, len(r.data)-r.pos)
	}

	payload = r.data[r.pos : r.pos+int(size)]
	r.pos += int(size)

	return id, payload, nil
}
```

**Step 3.4: Run tests to verify they pass**

```bash
go test ./internal/vba/ -run TestParseDir -v
```
Expected: all `TestParseDir_*` tests PASS.

If `TestParseDir_ReconcileWithProject` fails because names in dir use Unicode and PROJECT uses latin-1, check the name comparison. The dir parser prefers MODULENAMEUNICODE (0x0047) which includes Unicode, so names like "Form_Übersicht" should round-trip correctly.

**Step 3.5: Commit**

```bash
git add internal/vba/dir.go internal/vba/dir_test.go
git commit -m "feat(vba): parse dir stream for module→stream mapping"
```

---

## Task 4: Fallback Mapping + Validation (Phase 3.3)

This task adds a reconciliation helper that bridges the PROJECT stream (module type info) and the dir stream (stream name + offset), and provides a fallback when dir parsing fails or names don't align.

**Files:**
- Create: `internal/vba/reconcile.go`
- Create: `internal/vba/reconcile_test.go`

---

### Step 4.1: Write the failing test

Create `internal/vba/reconcile_test.go`:

```go
package vba

import (
	"testing"
)

func TestReconcileModules_StartMDB(t *testing.T) {
	db := testDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	required, err := st.RequiredStreams()
	if err != nil {
		t.Fatalf("RequiredStreams: %v", err)
	}

	modules, err := st.ModuleStreams()
	if err != nil {
		t.Fatalf("ModuleStreams: %v", err)
	}

	projectInfo, err := ParseProject(required["PROJECT"].Data)
	if err != nil {
		t.Fatalf("ParseProject: %v", err)
	}

	decompressed, err := DecompressContainer(required["dir"].Data)
	if err != nil {
		t.Fatalf("DecompressContainer: %v", err)
	}
	dirInfo, err := ParseDir(decompressed)
	if err != nil {
		t.Fatalf("ParseDir: %v", err)
	}

	resolved, warnings := ReconcileModules(projectInfo, dirInfo, modules)

	for _, w := range warnings {
		t.Logf("reconcile warning: %s", w)
	}

	if len(resolved) != 15 {
		t.Errorf("got %d resolved modules, want 15", len(resolved))
	}

	for _, m := range resolved {
		if m.StreamNode == nil {
			t.Errorf("module %q has no stream node", m.Name)
		}
		t.Logf("  %-30s type=%-8s stream=%-35s offset=%d",
			m.Name, moduleTypeName(m.Type), m.StreamName, m.Offset)
	}
}

func moduleTypeName(t ModuleType) string {
	switch t {
	case ModuleTypeStandard:
		return "standard"
	case ModuleTypeClass:
		return "class"
	case ModuleTypeDocument:
		return "document"
	default:
		return "unknown"
	}
}
```

**Step 4.2: Run tests to verify they fail**

```bash
go test ./internal/vba/ -run TestReconcileModules -v
```
Expected: compile error (`ReconcileModules`, `ResolvedModule` undefined).

---

### Step 4.3: Implement reconciliation

Create `internal/vba/reconcile.go`:

```go
package vba

import "fmt"

// ResolvedModule is one fully-reconciled module ready for source extraction.
type ResolvedModule struct {
	Name       string
	Type       ModuleType
	StreamName string       // obfuscated storage name
	Offset     uint32       // byte offset of source in decompressed module stream
	StreamNode *StorageNode // node from MSysAccessStorage (contains raw data)
}

// ReconcileModules merges information from the PROJECT stream, dir stream,
// and the storage tree to produce a slice of ResolvedModules.
//
// It returns a (possibly empty) warnings slice describing any mismatches or
// fallback decisions. Callers should log warnings at verbose level.
func ReconcileModules(
	project *ProjectInfo,
	dir *DirInfo,
	storageModules []*StorageNode,
) ([]ResolvedModule, []string) {
	var warnings []string

	// Build a lookup from module name to DirModule.
	dirByName := make(map[string]*DirModule, len(dir.Modules))
	for i := range dir.Modules {
		dirByName[dir.Modules[i].Name] = &dir.Modules[i]
	}

	// Build a lookup from stream name to StorageNode.
	nodeByStreamName := make(map[string]*StorageNode, len(storageModules))
	for _, n := range storageModules {
		nodeByStreamName[n.Name] = n
	}

	var resolved []ResolvedModule

	for _, pd := range project.Modules {
		dm, ok := dirByName[pd.Name]
		if !ok {
			warnings = append(warnings, fmt.Sprintf("module %q: found in PROJECT but not in dir; skipping", pd.Name))
			continue
		}

		node := nodeByStreamName[dm.StreamName]
		if node == nil {
			// Fallback: try case-insensitive match.
			for name, n := range nodeByStreamName {
				if equalFoldASCII(name, dm.StreamName) {
					node = n
					break
				}
			}
		}
		if node == nil {
			warnings = append(warnings, fmt.Sprintf("module %q: stream %q not found in MSysAccessStorage", pd.Name, dm.StreamName))
		}

		// Use type from PROJECT (more reliable for Document vs Class distinction).
		typ := pd.Type
		if typ == ModuleTypeStandard && dm.Type == ModuleTypeClass {
			// dir knows it's a class even though PROJECT listed it as Module.
			typ = ModuleTypeClass
		}

		resolved = append(resolved, ResolvedModule{
			Name:       pd.Name,
			Type:       typ,
			StreamName: dm.StreamName,
			Offset:     dm.Offset,
			StreamNode: node,
		})
	}

	// Report any dir modules not in PROJECT.
	for _, dm := range dir.Modules {
		found := false
		for _, pd := range project.Modules {
			if pd.Name == dm.Name {
				found = true
				break
			}
		}
		if !found {
			warnings = append(warnings, fmt.Sprintf("dir module %q not listed in PROJECT stream", dm.Name))
		}
	}

	return resolved, warnings
}

// equalFoldASCII is a simple ASCII case-insensitive comparison
// (avoids unicode.ToLower allocations for stream names which are always ASCII).
func equalFoldASCII(a, b string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := 0; i < len(a); i++ {
		ca, cb := a[i], b[i]
		if ca >= 'A' && ca <= 'Z' {
			ca += 32
		}
		if cb >= 'A' && cb <= 'Z' {
			cb += 32
		}
		if ca != cb {
			return false
		}
	}
	return true
}
```

**Step 4.4: Run tests to verify they pass**

```bash
go test ./internal/vba/ -run TestReconcileModules -v
```
Expected: PASS. Also run the full package test to check for regressions:

```bash
go test ./internal/vba/ -v
go test ./internal/mdb/ -v
```
All tests should pass.

**Step 4.5: Commit**

```bash
git add internal/vba/reconcile.go internal/vba/reconcile_test.go
git commit -m "feat(vba): reconcile PROJECT+dir+storage into resolved module list"
```

---

## Final Verification

Run all tests and verify the whole package compiles:

```bash
go build ./...
go test ./... -v
```

Expected outcome:
- `internal/mdb`: all existing tests pass (no regressions)
- `internal/vba`: all new tests pass:
  - `TestDecompress_*` (4 unit tests + 1 regression)
  - `TestParseProject_*` (4 tests including Start.mdb integration)
  - `TestParseDir_*` (2 tests, both using Start.mdb)
  - `TestReconcileModules_StartMDB` (integration test)

---

## Notes for Phase 5 (Next)

After this phase, `ReconcileModules()` provides everything needed for source extraction:

```go
for _, mod := range resolved {
    if mod.StreamNode == nil || len(mod.StreamNode.Data) == 0 {
        // skip / warn
        continue
    }
    // mod.StreamNode.Data = raw module stream bytes
    // mod.Offset = byte offset within DECOMPRESSED module stream where source starts
    source, err := ExtractSource(mod.StreamNode.Data, mod.Offset)
}
```

Phase 5 will implement `ExtractSource` which:
1. Decompresses the full module stream with `DecompressContainer`
2. Seeks to `mod.Offset` in the decompressed bytes
3. Decompresses the nested source container at that offset
4. Decodes the result as Windows-1252 / latin-1 text
