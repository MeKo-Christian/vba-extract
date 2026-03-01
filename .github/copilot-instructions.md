# Copilot Instructions for vba-extract

Expert-level guidance for AI coding agents working on the `vba-extract` project.

## Big Picture Architecture

`vba-extract` is a Go-based tool for extracting VBA source code and database schemas from Microsoft Access (`.mdb`/`.accdb`) files by reading raw binary pages (Jet/ACE format).

- **`internal/mdb`**: Low-level database engine. Handles 4096-byte page reading, table/column definitions, and Long Value (LVAL) chain resolution (used for MEMO/OLE fields).
- **`internal/vba`**: VBA project logic. Parses `MSysAccessStorage` to find module streams and implements MS-OVBA decompression ([internal/vba/decompress.go](internal/vba/decompress.go)).
- **`cmd/`**: CLI layer using Cobra. Features commands for `extract`, `schema`, `list`, and `info`.
- **Forensic Recovery**: When standard system tables are missing, the tool uses a two-pass scanner ([internal/vba/scanner.go](internal/vba/scanner.go)) to recover orphaned VBA streams from LVAL pages.

## Project Conventions & Patterns

### 1. Binary Data Handling

- Use `encoding/binary` for reading little-endian data from pages.
- Page size is constant ([internal/mdb/reader.go](internal/mdb/reader.go)): `const PageSize = 4096`.
- Direct byte slicing and `binary.LittleEndian` are preferred over complex struct mapping for raw page data.

### 2. Error Handling

- Use `fmt.Errorf("package: operation: %w", err)` for context-rich error wrapping.
- Library code (`internal/`) should return errors rather than logging or exiting.

### 3. Testing

- **Unit Tests**: Most packages have extensive `_test.go` files. Run with `go test ./...`.
- **Integration Tests**: Rely on `testdata/`. `testdata/sample.mdb` is the primary test fixture.
- **Fixtures**: Use `VBA_FIXTURE_DIR` environment variable to point to a local collection of real `.mdb` files for broader testing.

### 4. Code Style

- Follow standard Go idioms.
- Use `slog` for structured logging (primarily in `cmd/` and for fallback/debug logic in `internal/vba`).

## Common Developer Workflows

- **Run CLI**: `go run . extract --verbose testdata/sample.mdb`
- **Rebuild**: `just build` (if `justfile` is used) or `go build -o vba-extract .`
- **Check Coverage**: `go test -coverprofile=coverage.out ./... && go tool cover -html=coverage.out`

## Key Implementation Details to Note

- **MS-OVBA**: Implementation follows §2.4.1 of the MS-OVBA spec. If source extraction fails or looks like garbage, check `DecompressContainerWithFallback`.
- **LVAL Chains**: Crucial for large VBA modules. A row only contains a 12-byte reference; the actual data is chained across multiple pages.
- **Jet Versions**: The tool focuses on Jet 4.0 and ACE. Jet 3.5 (Access 97) with 2048-byte pages is NOT fully supported.
