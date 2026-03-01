# Jet 3 Support Roadmap

## Objective

Implement reliable Jet 3.x (Access 97 / 2048-byte page layout) support for schema extraction and VBA extraction, without regressing Jet 4/ACE behavior.

## Baseline

- [x] Page-size detection and runtime guards (`IsJet3`, `IsJet4`, explicit `ErrJet3*Unsupported` errors)
- [x] Jet 3 parser entry points exist and are implemented: `readTableDefJet3`, `parseRowJet3`, `resolveMemoJet3`
- [x] Tests for page-size dispatch and guardrails
- [x] Legacy fixture at `testdata/jet35/st990426.mdb` (detected as 4K layout despite non-standard header)

---

## Phase 1: Fixture Curation and Format Characterization

- [x] Probe utility classifies fixtures into `jet3-2k`, `jet4-4k`, `legacy-4k` (`internal/mdb/probe.go`, `probe_test.go`)
- [x] Fixture manifest created (`docs/fixtures.md`) with provenance and expected outcomes
- [ ] **Acquire a true 2048-byte-page Jet 3 `.mdb` file** — all Jet 3 tests currently use synthetic in-memory data
- [ ] Add encrypted Jet 3 file (if available) for negative-path tests

## Phase 2: Jet 3 TDEF Parser (`readTableDefJet3`)

- [x] Jet 3 TDEF offsets and column-entry parsing implemented (`internal/mdb/table.go`)
- [x] Column names, types, fixed/variable offsets, flags, lengths parsed correctly
- [x] `ReadTableDef` dispatch wired — no more `ErrJet3TableLayoutUnsupported` in normal flow
- [x] Synthetic TDEF test: `TestReadTableDefJet3Synthetic`
- [ ] Table-level integration tests against a real Jet 3 fixture:
  - [ ] `MSysObjects` columns
  - [ ] User table with mixed fixed/variable columns
  - [ ] Relationship/query system table discoverability
- [ ] Verify `FindTable`, `Catalog`, and `TableNames` work end-to-end on a real Jet 3 file

## Phase 3: Jet 3 Row Decoder (`parseRowJet3`)

- [x] Jet 3 row layout parsing implemented (`internal/mdb/column.go`)
- [x] Synthetic row test: `TestReadRowsJet3Synthetic` (text, numeric, nullable columns)
- [ ] Row-level integration tests with golden values from a real Jet 3 fixture:
  - [ ] Deleted rows and overflow flags
  - [ ] `ReadRows` returns stable row counts and correct typed values
- [ ] Verify `MSysRelationships` and `MSysQueries` return usable rows from real Jet 3 data

## Phase 4: Jet 3 LVAL/MEMO/OLE (`resolveMemoJet3`)

- [x] Jet 3 memo reference decoding and chain traversal implemented (`internal/mdb/memo.go`)
- [x] Tests: inline memo, single-page LVAL, multi-page chain, truncated/corrupt chains (`memo_jet3_test.go`)
- [ ] Verify `MSysAccessStorage.Lv` resolves on a real Jet 3 fixture (blocked by Phase 1)
- [ ] Confirm no `ErrJet3LvalLayoutUnsupported` for valid Jet 3 files

## Phase 5: VBA Extraction on Jet 3

- [x] `LoadStorageTree`, `RequiredStreams`, module extraction implemented
- [x] `dir` parsing and decompression work unchanged once storage bytes resolve
- [x] Forensic scan infrastructure in place (`internal/vba/forensic.go`, `scanner.go`)
- [ ] End-to-end VBA extraction test on a real Jet 3 file:
  - [ ] Expected module count
  - [ ] Spot-check module names and `Attribute VB_Name` markers
- [ ] Verify forensic scan does not assume 4K page math on Jet 3 page offsets
- [ ] Confirm `go run . extract <jet3file>` succeeds and writes stable modules

## Phase 6: Schema and CLI UX

- [x] `info` command shows `layoutClass` and `pageSize` (`cmd/info.go`, `cmd/diagnostics.go`)
- [ ] Verify `schema` command output parity (tables, FKs, queries) on a Jet 3 fixture
- [ ] Improve user-facing errors:
  - [ ] Encrypted DB
  - [ ] Unsupported/corrupt layout
  - [ ] Partial extraction warnings
- [ ] Ensure error text is actionable and consistent across all Jet 3 error paths

## Phase 7: Quality, Docs, and Release

- [ ] Add regression matrix in CI:
  - [ ] Synthetic parser tests (always run)
  - [ ] Optional fixture-based integration tests (guarded by fixture availability)
- [ ] Update `README.md` support table to reflect Jet 3 status
- [ ] Update `docs/ARCHITECTURE.md` Jet 3 section to match current implementation
- [ ] Add changelog notes and migration/limitations section

> **Blocker for Phases 2–7 integration tests:** a real 2048-byte-page Jet 3 `.mdb` file is needed.
> Parser logic is complete and tested synthetically; real-fixture validation is the remaining gap.

---

# Migration Analysis Roadmap

## Objective

Extend `accessdump` so it produces the artifacts needed to plan and execute migrations from legacy Access applications to modern backend/frontend systems.

The immediate validation case is IPOffice `Start.mdb`, but the roadmap should stay generic enough for similar Access launcher and line-of-business databases.

## Current Status (Baseline)

- [x] `extract` dumps VBA modules/classes/forms
- [x] `schema` dumps tables and some schema metadata
- [x] `list --json` produces JSON output for VBA module inventory
- [ ] Query extraction incomplete — some databases return no SQL for known queries
- [ ] Output not yet structured enough for migration tooling
- [ ] Cross-database dependencies, macro-driven entry points, and form/control metadata not yet surfaced as first-class artifacts

---

## Phase A: Query Recovery and Schema Completeness

- [x] Investigate why `QryUpdate`, `QryErststart`, `Qry_Module_aktiv` in `Start.mdb` are reported without SQL — **root cause found**: these are append/update queries stored in structured Attribute=5/6/7 rows (table/field references), not as SQL text in Attribute=0 Expression; `not-in-table` is correct
- [x] Add a `SQLStatus` reason field to `QueryDef` (`internal/mdb/schema.go`):
  - [x] `found` — SQL read successfully from `MSysQueries`
  - [x] `table-missing` — `MSysQueries` could not be opened
  - [x] `not-in-table` — table opened but no `Attribute=0` row matched the query name
  - [ ] Distinguish parser bug / unexpected row structure (future: `parse-error` status)
- [x] Improve `readQueries` to emit a structured status alongside each query
- [x] Update schema markdown output to surface the status per query (`*SQL not available (table-missing)*`)
- [x] Update DDL output to surface the status in the comment (`SQL not available: not-in-table`)
- [ ] Add fixtures and tests for databases containing:
  - [ ] Standard select queries
  - [ ] Action queries (update/delete/append)
  - [ ] Parameter queries
  - [ ] Nested/system-generated queries
- [x] Ensure missing SQL is clearly labelled with reason in all output formats

## Phase B: Cross-Reference and Dependency Extraction

- [ ] Add VBA analysis pass for cross-database `Run` calls and dynamic launch patterns:
  - [ ] `Run "DatabaseAlias", ...` calls
  - [ ] `DoCmd.OpenForm`, `DoCmd.OpenReport`, `DoCmd.RunMacro`
  - [ ] Shell execution patterns (`Shell`, `CreateObject("WScript.Shell")`)
  - [ ] COM automation of external Access instances
- [ ] Extract a dependency graph artifact:
  - [ ] Local module/function as caller
  - [ ] Target database alias
  - [ ] Target action/form/function
  - [ ] Confidence level for statically vs. dynamically constructed call targets
- [ ] Emit dependency graph in JSON and markdown summary
- [ ] Add tests for common VBA launch patterns and partially dynamic strings

## Phase C: Form and UI Metadata Extraction

- [x] Image extraction from form blobs implemented (`internal/mdb/image.go`)
- [x] Form names parsed from storage dir data (`parseFormDirData`)
- [x] Heuristic UTF-16LE blob scanner extracts UI metadata without full format parsing (`internal/mdb/form_meta.go`):
  - [x] Event bindings: `[Event Procedure]` markers and `=Expr()` expressions
  - [x] Data source / record source (SELECT statements recovered from blob)
  - [ ] Control names and types (partially: names extracted, types not classified)
  - [ ] Startup form / AutoExec references (covered partially via event expressions)
- [x] `FormMeta` struct added to `Schema` (`Forms []FormMeta`)
- [x] `ReadSchema` populates `Forms` via `ScanFormBlobs` (best-effort, silent on missing table)
- [x] Markdown output renders `## Forms` section with RecordSource and event handlers
- [x] Unit tests for `scanBlobStrings`, `classifyBlobStrings`, and integration test against `Start.mdb`
- [ ] Define stable JSON output for form metadata (Phase D dependency)
- [ ] Add fixtures/tests for forms with:
  - [ ] Button click handlers with specific handler names
  - [ ] Timer/startup behavior
  - [ ] Bound controls with table/column RecordSource

## Phase D: Structured Migration Outputs

- [x] `list --json` emits JSON for module inventory
- [ ] Add structured JSON output to `schema` command:
  - [ ] Tables with columns and types
  - [ ] Relationships / foreign keys
  - [ ] Queries with SQL and status
- [ ] Define and document JSON artifact schemas for:
  - [ ] VBA procedures (name, module, parameters, body hash)
  - [ ] Cross-database calls (Phase B output)
  - [ ] UI objects (Phase C output)
  - [ ] Startup entry points (Phase E output)
- [ ] Add optional `--format json` flag to `schema` and `extract` commands (or a new `analyze` command)
- [ ] Document artifact schemas in `docs/`

## Phase E: Startup and Entry-Point Analysis

- [ ] Detect startup paths automatically:
  - [ ] AutoExec macros if present
  - [ ] Startup form from database properties
  - [ ] Timer-triggered forms (OnTimer event bindings)
  - [ ] Common bootstrap procedure names (`Main`, `AutoExec`, `Startup`, …)
- [ ] Add heuristics for launcher-style databases:
  - [ ] Login flow detection (password comparison, user table lookup)
  - [ ] Menu construction patterns
  - [ ] Update/install routines
  - [ ] Environment and version checks
- [ ] Report likely application entry points and bootstrap chains in output
- [ ] Test against `Start.mdb`-style patterns

## Phase F: Security and Modernization Findings

- [ ] Add optional static findings pass for migration relevance:
  - [ ] Plaintext password comparisons
  - [ ] Hardcoded ODBC credentials or connection strings
  - [ ] Shell/batch execution calls
  - [ ] Windows API / registry writes
  - [ ] COM automation of Office/Access
  - [ ] Dynamic SQL string construction (injection risk)
- [ ] Classify findings by category:
  - [ ] `security` — active risk if migrated naively
  - [ ] `migration-blocker` — requires platform-specific replacement
  - [ ] `modernization-hotspot` — good refactor candidate
  - [ ] `possible-dead-code` — never called from detected entry points
- [ ] Emit findings as analysis output only, not build failure
- [ ] Add tests for each finding category with representative VBA snippets

## Phase G: IPOffice / Start.mdb Validation Track

- [ ] Run full analysis pipeline against `Start.mdb` and verify:
  - [ ] Startup flow reconstructed (entry point → login → menu)
  - [ ] Login mechanism identified
  - [ ] Menu definition source identified (table-driven or code-driven)
  - [ ] External database launches detected with aliases and actions
  - [ ] Tables and settings that drive startup and updates listed
- [ ] Compare automated output against `IPOffice-VBA/extracted/Start/ANALYSIS.md` and `MIGRATION.md`
- [ ] File concrete parser/tooling issues for remaining blind spots instead of one-off notes

## Phase H: CLI and Documentation

- [ ] Add migration-analysis command or mode (e.g. `accessdump analyze file.mdb`)
- [ ] Document intended outputs and known limitations in `README.md` and `docs/ARCHITECTURE.md`
- [ ] Provide example workflows in docs for:
  - [ ] Schema extraction
  - [ ] VBA extraction
  - [ ] Migration analysis
  - [ ] JSON export for downstream tools
- [ ] Document confidence/limitations for dynamic VBA pattern parsing

---

## Suggested Delivery Sequence

**Jet 3 track:**

1. Real Jet 3 fixture acquisition
2. Jet 3 TDEF parser + catalog/table integration tests
3. Jet 3 row parser + row/value integration tests
4. Jet 3 LVAL/MEMO parser + chain integration tests
5. Jet 3 VBA end-to-end and forensic fixes
6. CLI/error/doc polish and CI matrix

**Migration analysis track:**

1. Query recovery diagnostics and structured query status
2. Cross-database call extraction and dependency graph output
3. Startup/entry-point heuristics
4. Structured JSON analysis outputs and `analyze` command
5. Form/UI metadata extraction
6. Static security/modernization findings
7. IPOffice `Start.mdb` validation pass and docs/CLI polish
