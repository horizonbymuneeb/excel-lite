# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.1] - 2026-11-04

### Changed

- `package.json` now carries `repository`, `homepage` and `bugs`, so the npm page links back to
  the source, issue tracker and docs

## [0.1.2] - 2026-11-04

### Added

- **live documentation site**: <https://horizonbymuneeb.github.io/excel-lite/> — the pages in
  `docs/` rendered as a static site and deployed by `.github/workflows/docs.yml`, so the docs
  and the code ship together
- `site/verify.py`: headless Chromium pass over every page at 1440×980 and 390×844 (horizontal
  overflow, console errors), 667 internal links and anchors, and text-contrast measurements

### Changed

- `homepage` in `package.json` points at the documentation site instead of the README anchor

## [0.1.3] - 2026-11-04

### Added

- **self-hosted dynamic badges**: tests, coverage, browser-bundle size and docs page count are
  published as shields.io endpoint JSON under `/badges/` on the documentation site, so a badge
  changes when the repository numbers change instead of going stale (refreshed by every docs
  deploy through `scripts/metrics.mjs`)

### Changed

- the README badge row now links each number to where it comes from (CI, docs site, bundle notes)
- the docs workflow also rebuilds when `src/` or the metrics script changes, since both move
  the badge numbers

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
