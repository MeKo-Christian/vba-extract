# Jet 3 Support Roadmap

## Objective

Implement reliable Jet 3.x (Access 97 / 2048-byte page layout) support for schema extraction and VBA extraction, without regressing Jet 4/ACE behavior.

## Current Status (Baseline)

- Page-size detection and runtime guards are in place (`IsJet3`, `IsJet4`, explicit `ErrJet3*Unsupported` errors).
- Jet 3 parser entry points exist but return placeholders:
  - `readTableDefJet3`
  - `parseRowJet3`
  - `resolveMemoJet3`
- Tests exist for page-size dispatch and guardrails.
- A legacy fixture exists at `testdata/jet35/st990426.mdb` (currently detected as 4K layout despite non-standard header value).

## Phase 1: Fixture Curation and Format Characterization

### Tasks

- Build a small fixture set with known behavior:
  - At least 1 true 2048-page Jet 3 file
  - 1 legacy/non-standard file (already present)
  - 1 encrypted Jet 3 file (if available) for negative-path tests
- Add an internal probe utility/test helper that records:
  - Page-size inference
  - `MSysObjects` parse viability
  - TDEF/data-page signatures
- Create a fixture manifest (`docs/fixtures.md`) with provenance and expected outcomes.

### Done When

- CI tests can classify fixtures into `jet3-2k`, `jet4-4k`, `legacy-4k`.
- We can reproduce detection behavior from tests only.

## Phase 2: Jet 3 TDEF Parser (`readTableDefJet3`)

### Tasks

- Implement Jet 3 table-definition offsets and column-entry parsing.
- Parse column names, types, fixed/variable offsets, flags, and lengths.
- Wire `ReadTableDef` dispatch to return real `TableDef` on Jet 3.
- Add table-level tests against real fixture expectations:
  - `MSysObjects` columns
  - A user table with mixed fixed/variable columns
  - Relationship/query system table discoverability

### Done When

- `FindTable`, `Catalog`, and `TableNames` work on true Jet 3 fixture.
- No `ErrJet3TableLayoutUnsupported` in normal Jet 3 flow.

## Phase 3: Jet 3 Row Decoder (`parseRowJet3` + data-page handling)

### Tasks

- Implement Jet 3 row layout parsing (header, null map, variable offsets, fixed area).
- Validate and adapt row-offset table flags/masks if Jet 3 differs from Jet 4 assumptions.
- Keep `ReadRows` shared where possible; branch only layout-specific logic.
- Add row-level tests with golden values from known tables:
  - Text, numeric, datetime, nullable columns
  - Deleted rows and overflow flags behavior

### Done When

- `ReadRows` returns stable row counts and correct typed values for Jet 3 fixture tables.
- Schema extraction uses actual row content from `MSysRelationships` and `MSysQueries`.

## Phase 4: Jet 3 LVAL/MEMO/OLE (`resolveMemoJet3`)

### Tasks

- Implement Jet 3 memo reference decoding and chain traversal.
- Validate pointer format, row addressing, and continuation semantics.
- Add stress tests for:
  - Inline memo
  - Single-page LVAL
  - Multi-page LVAL chain
  - Truncated/corrupt chains (error-path behavior)

### Done When

- `MSysAccessStorage.Lv` resolves on true Jet 3 fixture (or fails with explicit corruption errors).
- No `ErrJet3LvalLayoutUnsupported` for valid Jet 3 files.

## Phase 5: VBA Extraction on Jet 3

### Tasks

- Verify `LoadStorageTree`, `RequiredStreams`, and module extraction with Jet 3 data.
- Ensure `dir` parsing and decompression work unchanged once storage bytes resolve.
- Add end-to-end `extract` tests:
  - Expected module count
  - Spot-check module names/content markers (`Attribute VB_Name`)
- Validate forensic fallback path against Jet 3 page-size assumptions.

### Done When

- `go run . extract <jet3file>` succeeds and writes stable modules.
- Forensic scan does not assume 4K page math on Jet 3.

## Phase 6: Schema and CLI UX

### Tasks

- Verify `schema` command output parity (tables, FKs, queries) on Jet 3 fixtures.
- Improve user-facing errors:
  - encrypted DB
  - unsupported/corrupt layout
  - partial extraction warnings
- Add `info` output hints for detected layout (`pageSize`, `layoutClass`).

### Done When

- `schema` and `info` produce useful output for Jet 3 with no placeholder errors.
- Error text is actionable and consistent.

## Phase 7: Quality, Docs, and Release

### Tasks

- Add regression matrix in CI:
  - synthetic parser tests
  - optional fixture-based integration tests (guarded by fixture availability)
- Resolve docs inconsistency:
  - `README.md` support table
  - `docs/ARCHITECTURE.md` Jet 3 section
- Add changelog notes and migration/limitations section.

### Done When

- Documentation matches reality.
- CI protects both Jet 3 and Jet 4/ACE paths.

## Risks and Mitigations

- Non-standard header versions (e.g. `259`): rely on page-layout probes first, version second.
- Fixture confidentiality/size: keep fixtures in ignored local `testdata/`, add synthetic tests to keep CI deterministic.
- Parser divergence: isolate Jet 3 logic in dedicated functions with minimal shared assumptions.

## Suggested Delivery Sequence (PRs)

1. Fixture + probe + docs manifest.
2. Jet 3 TDEF parser + catalog/table tests.
3. Jet 3 row parser + row/value tests.
4. Jet 3 LVAL/MEMO parser + chain tests.
5. Jet 3 VBA end-to-end and forensic fixes.
6. CLI/error/doc polish and CI matrix.
