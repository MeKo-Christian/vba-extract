# Cross-Validation Procedure (Windows Access vs Linux vba-extract)

This procedure validates extraction quality by comparing Linux output from `vba-extract` with VBA source exported from Microsoft Access on Windows.

## 1) Prepare fixture

- Pick one MDB/ACCDB file (start with `Start.mdb`).
- Ensure the exact same file bytes are used on both systems.

## 2) Linux extraction

Run:

```bash
go run . extract /path/to/Start.mdb --output-dir ./tmp/linux-out
```

Collect:
- module file count
- per-module file names
- total line count

## 3) Windows Access export

On Windows with Access:

1. Open database in Access.
2. Open VBA editor (`ALT+F11`).
3. Export each module/class/form code module to files (`.bas` / `.cls`).
4. Save to `windows-out/<db-name>/`.

## 4) Normalize before comparison

For both outputs:
- normalize line endings to LF
- trim trailing whitespace
- keep original module filenames

## 5) Compare results

Check:
- module count matches
- expected module names exist on both sides
- per-module line counts are close (or equal)
- spot-check content for key modules (`Inidatei`, `SQL`, `mod_api_Functions`, one `Form_*`)

Suggested commands on Linux:

```bash
find tmp/linux-out -type f | sort
find windows-out -type f | sort
wc -l tmp/linux-out/**/*.bas tmp/linux-out/**/*.cls 2>/dev/null
wc -l windows-out/**/*.bas windows-out/**/*.cls 2>/dev/null
```

## 6) Record mismatches

For each mismatch, record:
- module name
- missing/extra status
- line-count delta
- content delta type (header/attributes/body/encoding)

## 7) Acceptance criteria

A fixture is considered validated when:
- all expected modules are present
- no critical module is missing
- body content is semantically equivalent for spot-checked modules
- any differences are documented and explained (encoding, Access metadata lines, etc.)
