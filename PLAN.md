# vba-extract — VBA Extractor for Access MDB/ACCDB on Linux

## Goal

Build a Go CLI tool (Cobra) that extracts VBA source code from Microsoft Access `.mdb` and `.accdb` files on Linux without requiring Microsoft Access or Wine.

## Background & Research Findings

From hands-on analysis of the IPOffice MDB files at `/mnt/sql-server`:

- **23 unique MDB files** identified across IPOffice/Access2010 and backup folders
- VBA is stored inside `MSysAccessStorage` system table as binary streams
- The storage has a tree structure: `ROOT > VBAProject > VBA > [module streams]`
- The `PROJECT` stream (plain text, latin-1) lists all module names and types
- The `dir` stream contains module-to-stream name mappings and text offsets
- Module stream names are randomly generated (e.g. `BGONIVCSLIGSHEECEWKHKWVUBJRL`)
- VBA streams use **MS-OVBA 2.4.1 compression** (see MS-OVBA spec §2.4.1)
- Access uses a slightly different implementation of this compression than Office docs
- Each module stream contains: compiled p-code + compressed VBA source at a specific offset
- Tools tried and failed: `oletools`, `pcodedmp`, `access_parser` (Python), `mdb-tools` — none can reliably extract VBA source from MDB on Linux

## Architecture

```
vba-extract/
├── cmd/
│   └── root.go              # Cobra CLI (extract, list, info commands)
├── internal/
│   ├── mdb/
│   │   ├── reader.go         # Low-level Jet4 MDB page/table reader
│   │   ├── table.go          # Table definition & row iteration
│   │   ├── column.go         # Column types (text, binary, memo/LV)
│   │   └── memo.go           # Long Value (MEMO/OLE) field reader
│   ├── vba/
│   │   ├── storage.go        # MSysAccessStorage tree parser
│   │   ├── project.go        # PROJECT stream parser (module list)
│   │   ├── dir.go            # dir stream parser (module-stream mapping)
│   │   ├── decompress.go     # MS-OVBA 2.4.1 decompression
│   │   └── extract.go        # Orchestrator: extract all VBA modules
│   └── output/
│       ├── writer.go         # Write .bas/.cls files to disk
│       └── formatter.go      # Optional: syntax highlighting / summary
├── go.mod
├── go.sum
├── main.go
├── PLAN.md
└── testdata/                 # Small test MDB files for development
```

## Specification: MS-OVBA Compression (§2.4.1)

The decompression algorithm is critical and well-documented in the
[MS-OVBA specification](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba):

1. Compressed container starts with signature byte `0x01`
2. Data is split into **chunks** (max 4096 bytes decompressed each)
3. Each chunk has a 2-byte header: bits 0-11 = compressed size, bit 15 = compressed flag
4. Compressed chunks use a **flag byte** followed by 8 tokens:
   - Bit=0: literal byte (copy as-is)
   - Bit=1: copy token (2 bytes encoding offset + length from already-decompressed data)
5. Copy token bit allocation changes dynamically based on decompressed position within chunk

## Specification: Jet4 MDB File Format

Reference: [Jackcess MDB format docs](https://jackcess.sourceforge.io/),
[mdb-tools source](https://github.com/mdbtools/mdbtools),
[libesedb MDB format](https://github.com/libyal/libesedb/blob/main/documentation/Extensible%20Storage%20Engine%20%28ESE%29%20Database%20File%20%28EDB%29%20format.asciidoc)

Key structures:
- Page size: 4096 bytes (Jet4)
- Page 0: Database header (version, key, codepage)
- System tables at known page offsets (MSysObjects at page 2)
- Table definitions contain column metadata + data page pointers
- MEMO/OLE fields use Long Value pages (separate page chain)
- MSysAccessStorage columns: Id, ParentId, Name, Type, Lv (Long Value = binary blob), DateCreate, DateUpdate

---

## Phase 0: Project Bootstrap

Status (27.02.2026): **Completed**

### 0.1 Initialize Go module
- [x] `go mod init github.com/MeKo-Tech/vba-extract`
- [x] Add cobra dependency
- [x] Scaffold `main.go` and `cmd/root.go`
- Result: module initialized, Cobra wired, `go build ./...` passes

### 0.2 Set up basic CLI structure
- [x] `extract` command: extract VBA from one or more MDB files
- [x] `list` command: list VBA modules in an MDB without extracting
- [x] `info` command: show MDB metadata (version, tables, VBA project info)
- [x] Flags: `--output-dir`, `--verbose`, `--format` (flat/tree)
- Result: command stubs in place with placeholder output

### 0.3 Add testdata
- [x] Create or copy a small MDB file with known VBA content for testing
- [x] Use `Start.mdb` (3.6 MB, ~15 VBA modules) as primary test file
- [x] Document expected extraction results
- Result: `testdata/Start.mdb`, `testdata/README.md`, and `testdata/Start.expected.modules.txt` added

---

## Phase 1: MDB File Reader (Jet4 Format)

This is the foundation. We need to read the binary MDB format to access system tables.

### 1.1 Database header parser
- [ ] Implement a database header reader that loads page 0 (4096 bytes)
- [ ] Parse and expose Jet version (Jet3=0x00, Jet4=0x01, ACCDB=0x02)
- [ ] Parse and expose database codepage + encryption flag (if present)
- [ ] Validate magic bytes (`\x00\x01\x00\x00` at offset 0) and return a typed error on mismatch

### 1.2 Page reader
- [ ] Implement a `PageReader` that reads arbitrary pages by index (`pageNum * pageSize`)
- [ ] Enforce/support Jet4 page size (4096 bytes) with clear errors for unsupported sizes
- [ ] Parse and expose page type byte (0x01=data, 0x02=table def, 0x03=index, 0x04=leaf index, 0x05=long value)

### 1.3 Table definition parser
- [ ] Implement MSysObjects parsing (starting at page 2) to locate table definitions
- [ ] Implement table definition parsing (table metadata + column definitions: name, type, offset, length, flags)
- [ ] Implement decoding for required column types:
  - [ ] 0x01: Boolean
  - [ ] 0x03: Short Int
  - [ ] 0x04: Long Int
  - [ ] 0x05: Currency
  - [ ] 0x0A: Text (variable length)
  - [ ] 0x0B: OLE/Binary (Long Value)
  - [ ] 0x0C: Memo (Long Value)

### 1.4 Row reader
- [ ] Implement data page chain traversal for a given table
- [ ] Implement row enumeration using the row offset table stored at the end of data pages
- [ ] Implement variable-length column decoding for rows
- [ ] Implement null bitmap handling (so missing values are distinguishable from empty values)

### 1.5 Long Value (MEMO/OLE) reader
- [ ] Implement Long Value (LV) reading by following LV page chains to reconstruct binary blobs
- [ ] Support multi-page long values (linked pages)
- [ ] Support inline vs overflow long values
- [ ] Validate LV output integrity (length/chain termination), since this is critical for `MSysAccessStorage.Lv`

### 1.6 System table: MSysObjects reader
- [ ] Build a table catalog from MSysObjects (table name → definition page)
- [ ] Expose helpers to resolve system tables by name (e.g. `MSysAccessStorage`)
- [ ] Add minimal sanity checks for system-table presence and friendly errors when missing

---

## Phase 2: MSysAccessStorage Parser

### 2.1 Read MSysAccessStorage table
- [ ] Locate `MSysAccessStorage` via the MSysObjects catalog
- [ ] Implement row decoding for: Id, ParentId, Name, Type, Lv, DateCreate, DateUpdate
- [ ] Build an in-memory node list (and/or map by Id) suitable for tree construction

### 2.2 Navigate VBA storage tree
- [ ] Construct the storage tree (`ROOT > VBAProject > VBA > streams`)
- [ ] Implement lookups to find `VBAProject` and its `VBA` child folder
- [ ] Implement lookup/extraction for required streams: `PROJECT`, `PROJECTwm`, `dir`, `_VBA_PROJECT`
- [ ] Enumerate module streams (all other children of `VBA` folder)

### 2.3 Extract raw stream data
- [ ] Implement raw stream extraction by reading `Lv` (Long Value) for each stream node
- [ ] Ensure multi-page LV assembly is correct (no gaps/duplicates)
- [ ] Add stream-level integrity checks (detect truncation/corruption early)

---

## Phase 3: VBA Project Structure

### 3.1 Parse PROJECT stream
- [ ] Decode `PROJECT` stream as latin-1 text (with explicit charset handling)
- [ ] Parse module declarations and classify module type:
  - [ ] `Module=<name>` → standard module (.bas)
  - [ ] `DocClass=<name>/&H00000000` → class/form module (.cls)
  - [ ] `Class=<name>` → standalone class (.cls)
- [ ] Parse project metadata (Name, HelpFile, Version, etc.) into a struct
- [ ] Parse `[Host Extender Info]` section into a structured representation (or preserve as raw text)

### 3.2 Parse dir stream
- [ ] Decompress `dir` stream using MS-OVBA decompression (Phase 4 dependency)
- [ ] Implement `dir` record parser and support required record types:
  - [ ] `0x0019` MODULENAME: module display name
  - [ ] `0x001A` MODULESTREAMNAME: obfuscated stream name
  - [ ] `0x0032` MODULESTREAMNAME_UNICODE: unicode variant (skip after 0x001A)
  - [ ] `0x0031` MODULEOFFSET: byte offset where source starts in the module stream
  - [ ] `0x0021` MODULETYPE_PROCEDURAL: standard module
  - [ ] `0x0022` MODULETYPE_CLASS/DOCUMENT: class module
  - [ ] `0x002B` MODULE_TERMINATOR: end of module record
- [ ] Build mapping: moduleName → (streamName, sourceOffset, moduleType)

### 3.3 Handle Access-specific dir format
- [ ] Document/encode assumptions about Access-specific `dir` quirks (record ordering, missing records)
- [ ] Implement fallback mapping strategy when `dir` parsing fails (e.g. match by stream order sorted by storage Id)
- [ ] Validate `dir` mapping against `PROJECT` module list and surface mismatches clearly

---

## Phase 4: MS-OVBA Decompression

### 4.1 Implement decompression algorithm
- [ ] Implement `DecompressContainer([]byte) ([]byte, error)` for MS-OVBA §2.4.1
- [ ] Validate container signature byte (must be `0x01`)
- [ ] Implement chunk header parsing (2 bytes LE):
  - [ ] Bits 0-11: compressed chunk size - 3
  - [ ] Bit 12-14: reserved (must be 0b011)
  - [ ] Bit 15: 1=compressed, 0=raw
- [ ] Implement token decoding for compressed chunks:
  - [ ] Read flag byte (8 bits → 8 tokens)
  - [ ] Flag bit=0: copy literal byte
  - [ ] Flag bit=1: decode + execute copy token (2 bytes)
- [ ] Implement copy token decoding (offset/length) based on decompressed position within chunk:
  - [ ] Calculate `bitCount` dynamically (4-12)
  - [ ] Compute `lengthBits = 16 - bitCount`
  - [ ] Compute `offset = (token >> lengthBits) + 1`
  - [ ] Compute `length = (token & ((1 << lengthBits) - 1)) + 3`

### 4.2 Bit count lookup table
- [ ] Implement `bitCountForDecompressedPos(pos int) int` (or lookup table) matching the spec thresholds
```
decompressed <= 16:   bit_count = 4
decompressed <= 32:   bit_count = 5
decompressed <= 64:   bit_count = 6
decompressed <= 128:  bit_count = 7
decompressed <= 256:  bit_count = 8
decompressed <= 512:  bit_count = 9
decompressed <= 1024: bit_count = 10
decompressed <= 2048: bit_count = 11
decompressed <= 4096: bit_count = 12
```

### 4.3 Unit tests for decompression
- [ ] Add unit tests using known compressed/decompressed pairs from the MS-OVBA spec
- [ ] Add unit tests for edge cases: empty input, single chunk, multi-chunk, uncompressed chunks
- [ ] Add a regression test using real `Start.mdb` stream data once `MSysAccessStorage` reading exists

### 4.4 Handle Access-specific compression quirks
- [ ] Add a decompression fallback chain (standard → try skipping leading bytes → treat as raw)
- [ ] Detect and log which fallback path was used (only when `--verbose`)
- [ ] Add tests covering at least one “quirk” scenario once a failing sample is captured

---

## Phase 5: VBA Source Extraction

### 5.1 Extract source from module streams
- [ ] Implement module source extraction pipeline using the `dir` mapping:
  - [ ] Read module stream bytes from `MSysAccessStorage`
  - [ ] Decompress module stream container
  - [ ] Seek to `sourceOffset` (from `dir` MODULEOFFSET)
  - [ ] Decompress the nested source container at `sourceOffset`
  - [ ] Decode the resulting VBA text (latin-1 or project codepage)

### 5.2 Source validation and cleanup
- [ ] Validate extracted text (e.g. must contain `Attribute VB_Name` or other known markers)
- [ ] Strip trailing NUL bytes and other obvious padding
- [ ] Normalize line endings (CRLF → LF)
- [ ] Centralize decoding (latin-1 / Windows-1252 / UTF-8) behind a single helper

### 5.3 Fallback: p-code token scanning
- [ ] Implement “partial recovery” mode for p-code-only modules:
  - [ ] Scan raw binary for readable VBA fragments
  - [ ] Heuristically detect common VBA tokens (`MsgBox`, `DoCmd`, `Sub `, `Function `, ...)
  - [ ] Emit best-effort reconstructed output
  - [ ] Mark output as partial (e.g. header `[PARTIAL - reconstructed from p-code tokens]`)

### 5.4 Fallback: brute-force offset scanning
- [ ] Implement brute-force offset scanning when `dir` mapping is unavailable:
  - [ ] Try candidate offsets within decompressed module stream
  - [ ] Identify likely VBA source starts (`Attribute VB_Name`, `Option Compare`, ...)
  - [ ] Choose the best candidate via a simple scoring heuristic

---

## Phase 6: CLI & Output

### 6.1 `extract` command implementation
- [ ] Wire `extract` command to the extraction pipeline
- [ ] Accept one or more MDB file paths (and optionally globs)
- [ ] Implement `--output-dir` (default: `./vba-output/<db-name>/`)
- [ ] Implement `--flat` output mode (otherwise: one subdirectory per MDB)
- [ ] Write each module as `<moduleName>.bas` or `<moduleName>.cls`
- [ ] Print a summary (modules extracted, LoC, failures)

### 6.2 `list` command implementation
- [ ] Wire `list` command to parse PROJECT/dir and show module names/types/sizes
- [ ] Render output as a readable table by default
- [ ] Add `--json` output mode for machine-readable results

### 6.3 `info` command implementation
- [ ] Wire `info` command to MDB header + catalog parsing
- [ ] Print MDB metadata (version, codepage, table count)
- [ ] Print VBA project summary (name, module count)
- [ ] Optionally print a trimmed storage tree view (ROOT → VBAProject → VBA)

### 6.4 Batch mode
- [ ] Implement `--recursive` directory traversal mode
- [ ] Detect and process `.mdb` and `.accdb` files
- [ ] Add optional duplicate skipping via content hash
- [ ] Print aggregated summary at end (processed/extracted/failed)

### 6.5 Error handling and reporting
- [ ] Implement per-file error reporting (continue on errors by default)
- [ ] Add `--strict` mode to fail-fast / return non-zero when any file fails
- [ ] Use `--verbose` to enable progress + debug logging
- [ ] Implement color-coded terminal output when running on a TTY (auto-disable when output is piped)

---

## Phase 7: Testing & Validation

### 7.1 Unit tests
- [ ] Add unit tests for MDB page reader using synthetic fixtures
- [ ] Add unit tests for column parsing/decoding (all supported types)
- [ ] Add unit tests for MS-OVBA decompression (spec examples + regression samples)
- [ ] Add unit tests for `dir` stream parsing with known module mappings
- [ ] Add unit tests for `PROJECT` parsing with known module lists

### 7.2 Integration tests
- [ ] Add an integration test that runs extraction against `testdata/Start.mdb`
- [ ] Assert module list contains baseline modules from `testdata/Start.expected.modules.txt`
- [ ] Assert at least one extracted file contains `Attribute VB_Name`
- [ ] Add optional integration tests for `AT990426.mdb` and `PPS.mdb` (skipped when fixtures are absent)

### 7.3 Cross-validation
- [ ] Create a cross-validation procedure doc (Windows extraction + comparison steps)
- [ ] Compare extracted module counts and spot-check content/line counts

---

## Phase 8: IPOffice Batch Extraction

### 8.1 Run extraction on all IPOffice MDB files
- [ ] Run extraction over all 23 unique MDB files from `~/Code/MeKo/IPOffice-VBA/mdb-files/`
- [ ] Write outputs to `~/Code/MeKo/IPOffice-VBA/extracted/`
- [ ] Generate a summary report (per-DB module count + failures)

### 8.2 Organize extracted VBA
- [ ] Ensure output is organized as one directory per database
- [ ] Generate a README per database listing modules + brief notes (if derivable)
- [ ] Generate a top-level index file listing all databases and module counts

---

## Key Risks & Mitigations

| Risk | Mitigation |
|------|-----------|
| MS-OVBA decompression may differ in Access vs Office | Implement multiple decompression strategies with fallback chain |
| Long Value (MEMO) pages may be complex to reassemble | Study mdb-tools C source (`libmdb/data.c`) for reference implementation |
| dir stream format may be non-standard in Access | Implement fallback: match by stream order, or brute-force offset scanning |
| Some MDB files may be encrypted | Detect and report; support Jet4 RC4 encryption if feasible |
| P-code-only databases (no source stored) | Implement token scanning fallback; clearly report limitation |
| ACCDB (newer format) differs from MDB | Start with MDB (Jet4), add ACCDB support later |

## References

- [MS-OVBA: Office VBA File Format](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba)
- [mdb-tools source code](https://github.com/mdbtools/mdbtools) — C reference for Jet4 parsing
- [Jackcess](https://jackcess.sourceforge.io/) — Java library for Access, good format documentation
- [access_parser](https://github.com/claroty/access_parser) — Python MDB parser
- [oletools](https://github.com/decalage2/oletools) — MS-OVBA decompression reference in Python
- [Jet4 format research](https://github.com/brianb/mdbtools/blob/master/HACKING) — format internals

## Priority Order

**Start with Phase 4 (decompression)** — this is the core algorithm and can be tested independently.
Then Phase 1 (MDB reader), Phase 2 (storage parser), Phase 3 (project structure), Phase 5 (extraction), Phase 6 (CLI).
Phase 0 (bootstrap) runs in parallel with Phase 4.
