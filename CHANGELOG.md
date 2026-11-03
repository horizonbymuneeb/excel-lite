# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- docs suite (`docs/`) covering reading, writing, CLI, API reference, recipes, performance,
  compatibility, migration, troubleshooting, architecture and browser usage
- contributor tooling: ESLint, Prettier, coverage, docs link checker, browser-bundle checker
- CI: Node 22/24 matrix, lint, typecheck, coverage, fixture drift, pack check, browser bundle

## [0.1.0] - 2026-11-03

First release. Read and write `.xlsx` in Node 22+ and in browsers.

### Added

- `readSheet(source, { sheet, header })` — async row iteration over a path, `Buffer`,
  `Uint8Array`, `ArrayBuffer`, `Blob` or async byte stream
- `readWorkbook(source)` and `listSheets(source)` — sheet names, per-sheet iterators, row
  counts without materialising cells
- value support: shared strings, inline strings, rich-text runs, numbers, percentages,
  booleans, error values (`#DIV/0!`), ISO date cells, custom number formats and 1904
  date-system workbooks (returned as UTC `Date`)
- sparse rows read as `null`, blank rows skipped, header mode with duplicate/blank header
  handling
- `toXlsx(sheets)`, `writeSheet(dest, rows, options)` and `writeWorkbook(dest, sheets)` —
  multiple sheets, shared strings, number formats, column widths, freeze panes, auto filter,
  bold/fill/alignment
- sheet-name sanitisation (31 characters, forbidden characters, duplicates)
- CLI: `excel-lite info | to-json | to-csv` with NDJSON streaming, `--pretty`, `--out`,
  `--delimiter`, `--sheet`, `--header`
- dual ESM/CJS build with TypeScript declarations; no Node builtins in the browser graph
- 157 tests across reader, writer, API, CLI and edge cases, cross-validated against
  ExcelJS, SheetJS and openpyxl; 93% statement coverage
- benchmark harness (`npm run bench`) reporting write/read time, peak RSS and file size

### Fixed before first release

- `dateToSerial` wrote 1900-03-01 as serial 60 (the phantom 1900-02-29 slot), so the cell
  rendered one day early; serials at and after the 1900 leap-bug boundary now match Excel
- numeric cells now honour a date-like column format instead of silently falling back to
  General
- header mode produces unique keys even when a duplicate header collides with a generated
  `_2` key, and extra columns can no longer overwrite an identically named header key
- the CLI no longer prints the `excel-lite:` prefix twice

### Notes

- formulas are read as their cached values; formatting is written but not read back
- peak read memory is roughly 8× the file size because the sheet XML is decompressed into
  memory first; a truly streaming parser is planned for 0.2
- `.xls` (legacy binary), charts, pivot tables, conditional formatting, data validation and
  images are not supported
