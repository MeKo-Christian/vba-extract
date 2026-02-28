# Architecture

## Overview

`accessdump` reads Microsoft Access database files at the raw binary level,
locates their embedded VBA project, and decompresses each module's source code.
It has no runtime dependencies on Windows, ODBC, or any Access library.

## Package structure

```plain
accessdump/
├── main.go                    Entry point
├── cmd/
│   ├── root.go                Cobra CLI root, global flags
│   ├── extract.go             extract command
│   ├── list.go                list command
│   ├── info.go                info command (metadata + forensic scan)
│   └── common.go              Shared I/O helpers
└── internal/
    ├── mdb/
    │   ├── reader.go          Low-level Jet 4 page reader
    │   ├── table.go           Table definition and row iteration
    │   ├── column.go          Column type detection and value reading
    │   ├── memo.go            Long Value (MEMO/OLE) field resolution
    │   └── catalog.go         Database catalog enumeration
    └── vba/
        ├── storage.go         MSysAccessStorage tree parser
        ├── project.go         PROJECT stream parser (module list)
        ├── dir.go             dir stream parser (module-stream mapping)
        ├── decompress.go      MS-OVBA 2.4.1 decompression
        ├── extract.go         Extraction orchestrator
        ├── scanner.go         Forensic LVAL chain scanner (orphaned modules)
        └── forensic.go        Storage tree forensic scoring
```

## Jet 4 / ACE database format

A Jet 4 or ACE database file is a flat array of 4096-byte pages:

- **Page 0** — Database header: magic bytes, version, codepage, encryption key
- **Page 2** — System catalog (`MSysObjects`): table names and root page pointers
- **Data pages** (type `0x01`) — Table rows, indexed by a row-offset table at the start
- **Table-definition pages** (type `0x02`) — Column metadata and data-page pointers
- **Long Value pages** — Chained data pages for MEMO/OLE fields; each record has a 4-byte next-pointer as its first field

The `internal/mdb` package reads these structures. It does not support Jet 3.5 (Access 97, 2048-byte pages) or encrypted databases.

### Long Value (LVAL) chain format

Large field values (MEMO/OLE columns) are stored in LVAL chains:

```plain
12-byte reference in the row:
  bytes 0-2  : total data length (3-byte little-endian)
  byte  3    : storage type  0x00=multi-page  0x40=single-page  0x80=inline
  bytes 4-7  : first page/row pointer  (pageNum << 8) | rowID
  bytes 8-11 : reserved

Each LVAL record:
  bytes 0-3  : next-record pointer  (pageNum << 8) | rowID, or 0 for last
  bytes 4..  : data payload
```

## VBA storage inside Access

Access stores the entire VBA project in the `MSysAccessStorage` system table.
Each row is a storage node with an ID, parent ID, name, type, and an `Lv` MEMO
field containing the binary stream data. The tree structure is:

```plain
ROOT (id=0)
└── VBAProject
    └── VBA
        ├── _VBA_PROJECT   (compiled p-code, not useful for source extraction)
        ├── PROJECT        (plain text, latin-1: module names and types)
        ├── dir            (binary: MS-OVBA compressed module metadata)
        └── <ModuleName>   (binary: MS-OVBA module stream, one per module)
```

The `internal/vba` package parses this tree and extracts source from it.

## MS-OVBA decompression (§2.4.1)

VBA sources are compressed using the MS-OVBA algorithm:

1. Compressed container starts with signature byte `0x01`
2. Data is split into **CompressedChunks** (up to 4096 bytes decompressed each)
3. Each chunk has a 2-byte header: bits 11-0 = compressed payload size − 3, bit 15 = compressed flag, bits 14-12 must be `0x3`
4. Compressed chunks use a **flag byte** followed by 8 tokens:
   - Flag bit = 0: literal byte (copy as-is)
   - Flag bit = 1: 2-byte copy token (back-reference: offset + length)
5. Copy-token bit allocation shifts dynamically as the decompressed window grows

`DecompressContainer` implements the strict algorithm. `DecompressContainerWithFallback` additionally handles streams where a prefix of p-code precedes the container.

## Module stream layout

Each module stream in the `dir` record has a `MODULEOFFSET` field that gives the
byte position where the compressed source begins. Everything before that offset
is compiled p-code (architecture-specific, not source code):

```plain
Module stream bytes:
  [0 .. SourceOffset-1]  compiled p-code (ignored)
  [SourceOffset ..]      MS-OVBA CompressedContainer (VBA source)
```

## Forensic LVAL scanner

Some databases have had their `MSysAccessStorage` rows deleted while the
underlying LVAL page data remains on disk. `ScanOrphanedLvalModules` recovers
these modules with a two-pass algorithm:

**Pass 1** — Read every data page and collect all LVAL next-pointers into a
`targets` set. After this pass, any `(page, row)` tuple **not** in `targets` is
a chain start (not pointed to by any other record).

**Pass 2** — For each chain start whose first content byte is `0x01`
(CompressedContainer signature) or whose row is ≥ 1000 bytes (indicating
p-code-prefixed module data), read up to 256 KB of chain data and scan for a
valid CompressedContainer. If decompression succeeds and the output contains
`Attribute VB_Name`, the module is recovered.

The 256 KB read limit is derived from the largest module observed in production
databases (~154 KB compressed). Duplicate module names are discarded.

## Extraction flow

```plain
loadModules(path)
  ├─ mdb.Open()
  ├─ vba.LoadStorageTree()          -- reads MSysAccessStorage
  │   └─ [on error] ──────────────────────────────────────────┐
  ├─ vba.ExtractAllModules()                                   │
  │   ├─ ParseProjectStream()       -- module list             │
  │   ├─ ParseDirStream()           -- stream offsets          │
  │   └─ extractModuleSource() ×N   -- decompress each module  │
  │                                                            │
  ├─ [on failure or empty] ─────────────────────────────────── ┤
  └─ vba.ScanOrphanedLvalModules()  -- forensic LVAL scan ◄───┘
```
