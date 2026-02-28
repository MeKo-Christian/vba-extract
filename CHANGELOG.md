# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added
- Jet3 parser coverage for table definition, row decoding, and MEMO/LVAL traversal.
- Jet3 VBA extraction regression tests:
  - synthetic 2K forensic scanner path
  - legacy fixture end-to-end extraction checks
- CLI diagnostics helpers for layout classification and actionable error hints.

### Changed
- `info` command now reports `pageSize`, `layoutClass`, and `layoutHint`.
- `extract` and `schema` command error output now includes troubleshooting hints and probe context when available.
- `extract` command now reports partial-recovery and warning counts per file.
- CI unit workflow now includes explicit Jet3 synthetic regression tests and optional legacy-fixture tests when available.

### Documentation
- Updated support matrix and limitations in `README.md` to reflect current Jet3 status.
- Updated `docs/ARCHITECTURE.md` to describe 2K/4K layout detection and current parser scope.
- Expanded `docs/fixtures.md` with regression command matrix.
