# vba-extract

Extract VBA source code from Microsoft Access databases (`.mdb` / `.accdb`) on Linux — no Windows, no Wine, no ODBC drivers required.

## What it does

Access databases embed VBA projects as compressed binary streams inside the database file. `vba-extract` reads the raw Jet/ACE page format, locates the VBA project, decompresses each module's source, and writes `.bas` / `.cls` files that can be opened in any editor or version-controlled.

It handles two classes of database:

| Situation                                                      | How it's handled                                             |
| -------------------------------------------------------------- | ------------------------------------------------------------ |
| Standard VBA project (`MSysAccessStorage` intact)              | Reads PROJECT + dir streams, extracts each module            |
| Stripped project (system tables removed but page data on disk) | Two-pass LVAL chain scanner recovers orphaned module streams |
| No VBA at all (pure data database)                             | Reports `modules=0`, no error                                |

## Installation

```sh
go install github.com/MeKo-Tech/vba-extract@latest
```

Or build from source:

```sh
git clone https://github.com/MeKo-Tech/vba-extract
cd vba-extract
go build -o vba-extract .
```

**Requires Go 1.22 or later.**

## Usage

### Extract VBA modules

```sh
# Single file — writes to vba-output/<dbname>/
vba-extract extract MyDatabase.mdb

# Multiple files
vba-extract extract *.mdb

# Recurse into a directory
vba-extract extract --recursive /path/to/databases/

# Write all modules into a single flat directory
vba-extract extract --flat --output-dir ./src MyDatabase.mdb

# Skip duplicate databases (by file hash)
vba-extract extract --dedupe *.mdb

# Stop on first error
vba-extract extract --strict *.mdb

# Show recovery details
vba-extract extract --verbose MyDatabase.mdb
```

### List modules without extracting

```sh
vba-extract list MyDatabase.mdb

# JSON output
vba-extract list --json MyDatabase.mdb
```

Example output:

```
file: MyDatabase.mdb modules: 5
MODULE                         TYPE       STREAM                         BYTES    PARTIAL
Module1                        standard   Module1                        4821     false
Sheet1                         document   Sheet1                         1203     false
ThisWorkbook                   document   ThisWorkbook                    542     false
```

### Database metadata

```sh
# Basic info
vba-extract info MyDatabase.mdb

# Show VBA storage tree
vba-extract info --tree MyDatabase.mdb

# Forensic scan — useful when standard extraction fails
vba-extract info --forensic MyDatabase.mdb
```

## Output

By default, modules are written to `vba-output/<database-name>/`:

```
vba-output/
└── MyDatabase/
    ├── Module1.bas
    ├── UserForm1.cls
    └── Sheet1.cls
```

Use `--output-dir` to specify the root directory, and `--flat` to skip the per-database subdirectory.

## Supported formats

| Format              | Extension | Engine            |
| ------------------- | --------- | ----------------- |
| Access 2000 – 2003  | `.mdb`    | Jet 4.0           |
| Access 2007 – 2019+ | `.accdb`  | ACE               |
| Access 97           | `.mdb`    | Jet 3.5 (partial) |

## How forensic recovery works

When a database's `MSysAccessStorage` system table is missing or its VBA module references have been stripped, `vba-extract` falls back to a two-pass LVAL page scan:

1. **Pass 1** — reads every data page and builds a map of all LVAL chain next-pointers, identifying which records are chain _starts_ vs. _continuations_.
2. **Pass 2** — for each chain start whose content byte pattern suggests a VBA module stream, reads up to 256 KB and scans for a valid MS-OVBA `CompressedContainer`. If decompression succeeds and the output contains `Attribute VB_Name`, the module is recovered.

Duplicate module names (multiple on-disk copies) are deduplicated; only the first found copy is kept.

## Development

```sh
# Run tests
go test ./...

# Run with verbose output on a test file
go run . extract --verbose testdata/Start.mdb
```

Test fixtures live in `testdata/`. The expected extraction results for `Start.mdb` are in `testdata/Start.expected.modules.txt`.

## License

MIT
