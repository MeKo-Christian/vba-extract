# Fixture Manifest

This document tracks local MDB fixtures used for parser and layout regression tests.

## Classification Labels
- `jet3-2k`: 2048-byte page layout (true Jet 3.x storage).
- `jet4-4k`: 4096-byte Jet 4/ACE layout.
- `legacy-4k`: 4096-byte layout with legacy/non-standard header version values (for example `jetVersion=259`).
- `unknown`: unable to classify by current probe.

## Active Fixtures
| Label | Path | Source | Expected |
|---|---|---|---|
| `jet3-2k` | synthetic temp fixture (tests) | generated in `internal/mdb/page_size_test.go` | `PageSize=2048`, `MSysObjects` page signature present |
| `jet4-4k` | `testdata/Start.mdb` | repository test fixture | `LayoutClass=jet4-4k`, `MSysObjectsReadable=true` |
| `legacy-4k` | `testdata/jet35/st990426.mdb` | copied from `/mnt/sql-server/20251217/st990426.mdb` | `LayoutClass=legacy-4k`, `MSysObjectsReadable=true` |

## Discovery Scan (2026-02-28)
Probe scan over `/mnt/sql-server` (`1261` `.mdb` files):
- `jet4-4k`: `1125`
- `legacy-4k`: `136`
- `jet3-2k`: `0`
- `unknown`: `0`

Smallest known `legacy-4k` candidate:
- `/mnt/sql-server/GS_Transfer/IP/IPOffice/Werkzeug/IPOffice/Pflege.mdb` (`382976` bytes)

No encrypted/unknown fixtures were identified by the current heuristic scan.

## Notes
- `testdata/` is gitignored in this repo, so large/private fixtures remain local by design.
- CI-safe coverage is provided by synthetic fixtures plus optional local fixtures when present.

## Regression Commands
- Synthetic Jet3 coverage (always CI-safe):
  - `go test ./internal/mdb -run 'TestProbeClassifiesSyntheticJet32K|TestOpenJet3Uses2048PageSize|TestReadTableDefJet3Synthetic|TestReadRowsJet3Synthetic|TestResolveMemoJet3'`
  - `go test ./internal/vba -run 'TestScanOrphanedLvalModules_Jet3LayoutSynthetic'`
- Optional local fixture coverage:
  - `go test ./internal/vba -run TestExtractAllModulesFromLegacyFixture`
  - `go test ./cmd -run TestLoadSchema_legacyFixture`
