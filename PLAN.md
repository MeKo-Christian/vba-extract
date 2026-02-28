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

### 0.4 Repository & folder scaffolding
- [x] Initialize git repository (`git init`)
- [x] Add `.gitignore`
- [x] Create base directories (`cmd/`, `internal/`, `testdata/`)

---

## Phase 1: MDB File Reader (Jet4 Format)

Status (27.02.2026): **Completed**

This is the foundation. We need to read the binary MDB format to access system tables.

### 1.1 Database header parser
- [x] Implement a database header reader that loads page 0 (4096 bytes)
- [x] Parse and expose Jet version (Jet3=0x00, Jet4=0x01, ACE=0x03/0x04)
- [x] Parse and expose database codepage + sort order
- [x] Validate magic bytes (`\x00\x01\x00\x00` at offset 0) and return a typed error on mismatch
- Result: `internal/mdb/reader.go` — `Open()`, `Header`, `ReadPage()`, `PageType()`

### 1.2 Page reader
- [x] Implement page reader that reads arbitrary pages by index (`pageNum * pageSize`)
- [x] Support Jet4 page size (4096 bytes) with bounds checking
- [x] Parse page type byte (0x01=data, 0x02=TDEF, 0x03=index, 0x04=leaf index, 0x05=usage)
- Result: integrated into `reader.go`

### 1.3 Table definition parser
- [x] Parse TDEF pages (multi-page chain support)
- [x] Parse column definitions: name (UCS-2LE), type, offsets, length, flags
- [x] Support all required column types: Bool, Int, Long, Float, Double, DateTime, Text, Binary, OLE, Memo
- Result: `internal/mdb/table.go` — `ReadTableDef()`, `Column`, `TableDef`
- Note: table type at offset 0x28 (not 0x18 as some docs suggest); col names in Jet4 are UCS-2LE with 2-byte length prefix

### 1.4 Row reader
- [x] Data page discovery via full page scan (checking tdef_pg field)
- [x] Row enumeration from row offset table (2 bytes/entry at offset 0x0E, num_rows at 0x0C)
- [x] Jet4 row format: num_cols at START, [null_mask][num_var_cols][var_table] at END
- [x] Variable-length column decoding with proper type coercion (numeric types stored as var cols)
- [x] Null bitmap handling
- Result: `internal/mdb/column.go` — `ReadRows()`, `parseRow()`
- Note: Jet4 row metadata layout differs from some documentation; var_offset_table stored in reverse order

### 1.5 Long Value (MEMO/OLE) reader
- [x] 12-byte LVAL reference parsing (memo_len, bitmask, lval_dp)
- [x] Inline data (bitmask 0x80), single-page (0x40), multi-page chain (0x00)
- [x] LVAL page record reading via row offset table
- Result: `internal/mdb/memo.go` — `ResolveMemo()`
- Validated: PROJECT stream (1184 bytes) successfully read from MSysAccessStorage

### 1.6 System table: MSysObjects reader
- [x] Build table catalog from MSysObjects (table name → TDEF page via ID)
- [x] `FindTable()` helper to locate any table by name
- [x] `TableNames()` for listing all tables
- Result: `internal/mdb/catalog.go` — `Catalog()`, `FindTable()`, `TableNames()`
- Validated: MSysAccessStorage found at ID=56, all 31 tables enumerated correctly

---

## Phase 2: MSysAccessStorage Parser

Status (27.02.2026): **Completed**

### 2.1 Read MSysAccessStorage table
- [x] Locate `MSysAccessStorage` via the MSysObjects catalog
- [x] Implement row decoding for: Id, ParentId, Name, Type, Lv, DateCreate, DateUpdate
- [x] Build an in-memory node list (and/or map by Id) suitable for tree construction

### 2.2 Navigate VBA storage tree
- [x] Construct the storage tree (`ROOT > VBAProject > VBA > streams`)
- [x] Implement lookups to find `VBAProject` and its `VBA` child folder
- [x] Implement lookup/extraction for required streams: `PROJECT`, `PROJECTwm`, `dir`, `_VBA_PROJECT`
- [x] Enumerate module streams (all other children of `VBA` folder)

### 2.3 Extract raw stream data
- [x] Implement raw stream extraction by reading `Lv` (Long Value) for each stream node
- [x] Ensure multi-page LV assembly is correct (no gaps/duplicates)
- [x] Add stream-level integrity checks (detect truncation/corruption early)
- Result: implemented in `internal/vba/storage.go` with tests in `internal/vba/storage_test.go`

---

## Phase 3: VBA Project Structure

Status (27.02.2026): **In Progress**

### 3.1 Parse PROJECT stream
- [x] Decode `PROJECT` stream as latin-1 text (with explicit charset handling)
- [x] Parse module declarations and classify module type:
  - [x] `Module=<name>` → standard module (.bas)
  - [x] `DocClass=<name>/&H00000000` → class/form module (.cls)
  - [x] `Class=<name>` → standalone class (.cls)
- [x] Parse project metadata (Name, HelpFile, Version, etc.) into a struct
- [x] Parse `[Host Extender Info]` section into a structured representation (or preserve as raw text)

### 3.2 Parse dir stream
- [x] Decompress `dir` stream using MS-OVBA decompression (Phase 4 dependency)
- [x] Implement `dir` record parser and support required record types:
  - [x] `0x0019` MODULENAME: module display name
  - [x] `0x001A` MODULESTREAMNAME: obfuscated stream name
  - [x] `0x0032` MODULESTREAMNAME_UNICODE: unicode variant (skip after 0x001A)
  - [x] `0x0031` MODULEOFFSET: byte offset where source starts in the module stream
  - [x] `0x0021` MODULETYPE_PROCEDURAL: standard module
  - [x] `0x0022` MODULETYPE_CLASS/DOCUMENT: class module
  - [x] `0x002B` MODULE_TERMINATOR: end of module record
- [x] Build mapping: moduleName → (streamName, sourceOffset, moduleType)

### 3.3 Handle Access-specific dir format
- [x] Document/encode assumptions about Access-specific `dir` quirks (record ordering, missing records)
- [x] Implement fallback mapping strategy when `dir` parsing fails (e.g. match by stream order sorted by storage Id)
- [x] Validate `dir` mapping against `PROJECT` module list and surface mismatches clearly
- Result: implemented in `internal/vba/project.go` and `internal/vba/dir.go` with tests in `internal/vba/project_dir_test.go`

---

## Phase 4: MS-OVBA Decompression

Status (27.02.2026): **In Progress**

### 4.1 Implement decompression algorithm
- [x] Implement `DecompressContainer([]byte) ([]byte, error)` for MS-OVBA §2.4.1
- [x] Validate container signature byte (must be `0x01`)
- [x] Implement chunk header parsing (2 bytes LE):
  - [x] Bits 0-11: compressed chunk size - 3
  - [x] Bit 12-14: reserved (must be 0b011)
  - [x] Bit 15: 1=compressed, 0=raw
- [x] Implement token decoding for compressed chunks:
  - [x] Read flag byte (8 bits → 8 tokens)
  - [x] Flag bit=0: copy literal byte
  - [x] Flag bit=1: decode + execute copy token (2 bytes)
- [x] Implement copy token decoding (offset/length) based on decompressed position within chunk:
  - [x] Calculate `bitCount` dynamically (4-12)
  - [x] Compute `lengthBits = 16 - bitCount`
  - [x] Compute `offset = (token >> lengthBits) + 1`
  - [x] Compute `length = (token & ((1 << lengthBits) - 1)) + 3`
- Result: implemented in `internal/vba/decompress.go` with tests in `internal/vba/decompress_test.go`

### 4.2 Bit count lookup table
- [x] Implement `bitCountForDecompressedPos(pos int) int` (or lookup table) matching the spec thresholds
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
- [x] Add unit tests using known compressed/decompressed pairs from the MS-OVBA spec
- [x] Add unit tests for edge cases: empty input, single chunk, multi-chunk, uncompressed chunks
- [x] Add a regression test using real `Start.mdb` stream data once `MSysAccessStorage` reading exists

### 4.4 Handle Access-specific compression quirks
- [x] Add a decompression fallback chain (standard → try skipping leading bytes → treat as raw)
- [x] Detect and log which fallback path was used (only when `--verbose`)
- [x] Add tests covering at least one “quirk” scenario once a failing sample is captured

---

## Phase 5: VBA Source Extraction

Status (27.02.2026): **Completed (core extraction layer)**

### 5.1 Extract source from module streams
- [x] Implement module source extraction pipeline using the `dir` mapping:
  - [x] Read module stream bytes from `MSysAccessStorage`
  - [x] Decompress module stream container
  - [x] Seek to `sourceOffset` (from `dir` MODULEOFFSET)
  - [x] Decompress the nested source container at `sourceOffset`
  - [x] Decode the resulting VBA text (latin-1 or project codepage)

### 5.2 Source validation and cleanup
- [x] Validate extracted text (e.g. must contain `Attribute VB_Name` or other known markers)
- [x] Strip trailing NUL bytes and other obvious padding
- [x] Normalize line endings (CRLF → LF)
- [x] Centralize decoding (latin-1 / Windows-1252 / UTF-8) behind a single helper

### 5.3 Fallback: p-code token scanning
- [x] Implement “partial recovery” mode for p-code-only modules:
  - [x] Scan raw binary for readable VBA fragments
  - [x] Heuristically detect common VBA tokens (`MsgBox`, `DoCmd`, `Sub `, `Function `, ...)
  - [x] Emit best-effort reconstructed output
  - [x] Mark output as partial (e.g. header `[PARTIAL - reconstructed from p-code tokens]`)

### 5.4 Fallback: brute-force offset scanning
- [x] Implement brute-force offset scanning when `dir` mapping is unavailable:
  - [x] Try candidate offsets within decompressed module stream
  - [x] Identify likely VBA source starts (`Attribute VB_Name`, `Option Compare`, ...)
  - [x] Choose the best candidate via a simple scoring heuristic
- Result: implemented in `internal/vba/extract.go` with tests in `internal/vba/extract_test.go`

---

## Phase 6: CLI & Output

Status (27.02.2026): **Completed (core CLI implementation)**

### 6.1 `extract` command implementation
- [x] Wire `extract` command to the extraction pipeline
- [x] Accept one or more MDB file paths (and optionally globs)
- [x] Implement `--output-dir` (default: `./vba-output/<db-name>/`)
- [x] Implement `--flat` output mode (otherwise: one subdirectory per MDB)
- [x] Write each module as `<moduleName>.bas` or `<moduleName>.cls`
- [x] Print a summary (modules extracted, LoC, failures)

### 6.2 `list` command implementation
- [x] Wire `list` command to parse PROJECT/dir and show module names/types/sizes
- [x] Render output as a readable table by default
- [x] Add `--json` output mode for machine-readable results

### 6.3 `info` command implementation
- [x] Wire `info` command to MDB header + catalog parsing
- [x] Print MDB metadata (version, codepage, table count)
- [x] Print VBA project summary (name, module count)
- [x] Optionally print a trimmed storage tree view (ROOT → VBAProject → VBA)

### 6.4 Batch mode
- [x] Implement `--recursive` directory traversal mode
- [x] Detect and process `.mdb` and `.accdb` files
- [x] Add optional duplicate skipping via content hash
- [x] Print aggregated summary at end (processed/extracted/failed)

### 6.5 Error handling and reporting
- [x] Implement per-file error reporting (continue on errors by default)
- [x] Add `--strict` mode to fail-fast / return non-zero when any file fails
- [x] Use `--verbose` to enable progress + debug logging
- [x] Implement color-coded terminal output when running on a TTY (auto-disable when output is piped)
- Result: implemented in `cmd/extract.go`, `cmd/list.go`, `cmd/info.go`, and `cmd/common.go`

---

## Phase 7: Testing & Validation

Status (27.02.2026): **Completed (automated coverage + procedure doc)**

### 7.1 Unit tests
- [x] Add unit tests for MDB page reader using synthetic fixtures
- [x] Add unit tests for column parsing/decoding (all supported types)
- [x] Add unit tests for MS-OVBA decompression (spec examples + regression samples)
- [x] Add unit tests for `dir` stream parsing with known module mappings
- [x] Add unit tests for `PROJECT` parsing with known module lists

### 7.2 Integration tests
- [x] Add an integration test that runs extraction against `testdata/Start.mdb`
- [x] Assert module list contains baseline modules from `testdata/Start.expected.modules.txt`
- [x] Assert at least one extracted file contains `Attribute VB_Name`
- [x] Add optional integration tests for `AT990426.mdb` and `PPS.mdb` (skipped when fixtures are absent)

### 7.3 Cross-validation
- [x] Create a cross-validation procedure doc (Windows extraction + comparison steps)
- [x] Compare extracted module counts and spot-check content/line counts
- Result: coverage in `internal/vba/*_test.go` and procedure in `docs/cross-validation.md`

---

## Phase 8: IPOffice Batch Extraction

Status (27.02.2026): **Completed**

### 8.1 Run extraction on all IPOffice MDB files
- [x] Run extraction over all 23 unique MDB files from `~/Code/MeKo/IPOffice-VBA/mdb-files/`
- [x] Write outputs to `~/Code/MeKo/IPOffice-VBA/extracted/`
- [x] Generate a summary report (per-DB module count + failures)
- Result: processed=23, success=19, failed=4, modules=1895, lines=370872 (rerun 2026-02-28)
- Reports: `docs/reports/2026-02-27-phase8-1-batch-summary.md`, `docs/reports/2026-02-28-phase8-rerun-summary.md`

### 8.2 Organize extracted VBA
- [x] Ensure output is organized as one directory per database
- [x] Generate a README per database listing modules + brief notes (if derivable)
- [x] Generate a top-level index file listing all databases and module counts
- Result: per-database `README.md` files and top-level `INDEX.md` generated in extraction output

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
