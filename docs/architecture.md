# Architecture

`excel-lite` is about 1,600 lines of TypeScript in eight files. This page explains what each file
does, how a read and a write actually flow, why the parser is regex-driven, and where to make a
change.

## On this page

- [Module map](#module-map)
- [How a read flows](#how-a-read-flows)
- [How a write flows](#how-a-write-flows)
- [Why the parser is regex-driven](#why-the-parser-is-regex-driven)
- [Tests and fixtures](#tests-and-fixtures)
- [Where to change what](#where-to-change-what)
- [Build and packaging](#build-and-packaging)

## Module map

| File | Lines (0.1.0) | Job |
|---|---|---|
| `src/index.ts` | ~14 | public entry point: re-exports the functions and types, nothing else |
| `src/types.ts` | ~63 | the public type surface: `CellValue`, `Source`, `ReadOptions`, `SheetInfo`, `ColumnSpec`, `RowInput`, `SheetInput`, `WriteOptions`, `SheetRow` |
| `src/read.ts` | ~473 | workbook opening (zip parts, relationships, 1904 flag), sheet resolution, shared strings, style→date detection, row iteration, header keys, the public readers and date helpers |
| `src/write.ts` | ~536 | OOXML generation: sheet XML, shared-string table, style table, column widths, freeze/filter, package parts, zip assembly, the public writers |
| `src/xml.ts` | ~258 | the XML layer: a tiny DOM for metadata parts, a regex row scanner for worksheets, entity escaping/decoding, row counting |
| `src/zip.ts` | ~44 | fflate wrappers: inflate only requested entries, 512 MiB per-entry guard, UTF-8 encode/decode, zip at level 6 |
| `src/fs.ts` | ~46 | Node file access through `process.getBuiltinModule('fs/promises')` so no `node:fs` import enters the bundle |
| `src/cli.ts` | ~167 | argument parsing, `info`/`to-json`/`to-csv`, CSV quoting, NDJSON vs `--pretty`, EPIPE handling |

There are no other source files, and `fflate` is the only runtime dependency. Tests, fixtures,
scripts and docs live outside `src/` (see [tests and fixtures](#tests-and-fixtures)).

## How a read flows

```
readSheet(source, options)
  │
  ├─ toBytes(source)                     src/read.ts
  │    path → fs.readFileBytes (src/fs.ts) · Uint8Array/ArrayBuffer/Blob · async iterable → one buffer
  │
  ├─ WorkbookReader.open(bytes)          src/read.ts
  │    ├─ unzipEntries([xl/workbook.xml, xl/_rels/workbook.xml.rels])   src/zip.ts (filter: only those entries)
  │    ├─ parseXml(workbook.xml)  → <workbook>, <workbookPr date1904>, <sheets><sheet r:id name>
  │    ├─ parseXml(rels)          → r:id → resolved part path (relative targets normalised)
  │    └─ sheet refs: name + part, falling back to xl/worksheets/sheetN.xml
  │
  ├─ resolveSheet(options.sheet)         0-based index, exact name, then case-insensitive name
  │
  └─ rows(index, header)                 generator
       ├─ loadSharedStrings()            inflate xl/sharedStrings.xml once → parseXml → <si> → string[]
       ├─ loadStyles()                   inflate xl/styles.xml → numFmts + cellXfs → boolean[] "is date style"
       └─ getSheetXml(index)             inflate ONE sheet part (~ only requested) → decodeUtf8 → string
            │
            ├─ scanRows(xml)             src/xml.ts: regex over <row>/<c> → RawRow{ rowNumber, cells[{ref,type,style,value}] }
            ├─ toArray(cells)            positional: column from cell ref (columnIndexFromRef) or cursor
            │    └─ cellValue(cell)      type switch: s / inlineStr / str / b / e / d / numeric+dateStyle
            └─ header mode               makeKeys(first non-empty row) → object per row
```

Notable design points:

- **Only the requested sheet is inflated.** `unzipEntries` passes a `filter` to `fflate.unzipSync`,
  so `xl/theme1.xml`, images and the other sheets stay compressed. That is what keeps a multi-sheet
  workbook cheap to read one sheet from.
- **Row parsing is lazy, the XML is not.** The strings and bytes for the sheet are held for the
  whole iteration; `scanRows` yields rows as regex matches are found. This is the memory model in
  [performance](performance.md#the-memory-model).
- **Positions come from cell refs when present.** `A1`, `C1` gaps become `null`; a file that omits
  `r` attributes falls back to sequential positions. Trailing empty cells are trimmed.
- **Styles are consulted only for dates.** `loadStyles` builds a boolean per `cellXfs` entry from
  built-in date format ids and the custom-format heuristic. It is re-parsed on each `rows()` call
  (small, but not cached).
- **Metadata is parsed as a DOM; worksheets are not.** `xl/workbook.xml`, the rels, styles and
  shared strings are small, so `parseXml` is fine there. A worksheet can be tens of megabytes, so it
  never becomes a DOM.

## How a write flows

```
toXlsx(sheets)                            src/write.ts
  │
  ├─ for each sheet:
  │    ├─ collect rows                   object rows → values by column key; array rows as-is
  │    ├─ infer columns                  from the first object row when `columns` is absent
  │    ├─ header row (when enabled)     bold style, header text, counted into widths
  │    ├─ body cells                     cellXml(value, ref, styleId, column, styles, shared)
  │    │     string → SharedStrings.add(text)          (dedupe, 32 767-char cap)
  │    │     number → <v>                              (non-finite → shared string)
  │    │     boolean → t="b"
  │    │     Date    → dateToSerial + a date-formatted xf
  │    │     null/'' → no cell
  │    ├─ widths                        explicit width, else clamp(8, maxLen+2, 60) when autoWidth
  │    ├─ <cols>, <sheetViews><pane> (freeze), <sheetData>, <autoFilter> (header-row range)
  │    └─ sheet XML string
  │
  ├─ package parts                       [Content_Types].xml, _rels/.rels, xl/workbook.xml,
  │                                      xl/_rels/workbook.xml.rels, xl/styles.xml, xl/sharedStrings.xml
  │
  └─ zipFiles(files)                     fflate.zipSync({ level: 6 }) → Uint8Array
```

Style and string tables are deduplicated by key:

- `StyleBook` keeps one `xf` per distinct `(numFmtId, bold, fill, align)` and one `<fill>` per
  colour; custom number formats are interned from `numFmtId` 164 upwards.
- `SharedStrings` deduplicates text and tracks `count` (references) and `uniqueCount` separately,
  because Excel expects those attributes to differ.
- Sheet names go through `uniqueSheetName`: forbidden characters replaced, 31-character cap,
  case-insensitive duplicate suffixing.

`writeSheet(dest, rows, options)` is `writeWorkbook(dest, [{ ...options, rows }])`; both build the
bytes with `toXlsx` and then write them through `src/fs.ts`. The whole workbook is in memory — there
is no streaming writer ([performance](performance.md#the-write-path)).

## Why the parser is regex-driven

`src/xml.ts` is not a general XML parser, on purpose:

- **Input is machine-generated.** OOXML worksheet parts are written by Excel, LibreOffice, openpyxl,
  ExcelJS or SheetJS. They use a small, predictable subset: elements, attributes with quoted values,
  the five predefined entities and numeric character references.
- **A DOM costs more than the file.** A 400k-row sheet with `<row>`, `<c>`, `<v>` elements would
  create millions of node objects; the row scanner matches and discards instead, so only the current
  row is materialised.
- **Metadata is small.** For workbook/rels/styles/sharedStrings a small DOM (`parseXml`) is simpler
  than chasing SAX callbacks, and those parts are kilobytes, not megabytes.
- **Security.** `<!DOCTYPE>` and processing instructions are skipped, so external entities and
  "billion laughs" expansion have no path in. Only named entities and numeric references are
  decoded. See [SECURITY.md](../SECURITY.md).
- **Namespace prefixes are stripped** by `localName`, so `x:sheet` and `sheet` are the same.

The trade-offs, which are real: it is not schema-validating, it requires attributes to be quoted
(the format always quotes them), it does not resolve arbitrary entities, and it can be confused by
pathological XML inside an attribute value. For the inputs this package accepts (files produced by
spreadsheet tools), that is a good trade. If you need a conforming XML parser, this is not it.

`escapeXml` on the write side does the mirror-image minimal job: escape `& < > " '`, drop characters
XML 1.0 forbids (control characters other than tab/LF/CR, U+FFFE, U+FFFF), leave everything else —
including emoji and other non-BMP characters — alone.

## Tests and fixtures

| Location | What it holds |
|---|---|
| `test/fixtures.test.ts` | reads openpyxl-generated fixtures: every value type, blank rows, both sheets, row counts, custom formats, bytes/stream sources |
| `test/roundtrip.test.ts` | write→read of every supported type; dates; sparse rows; sheet names; multi-sheet; reads ExcelJS- and SheetJS-written files; asserts ExcelJS and SheetJS can read our output; error messages |
| `test/cli.test.ts` | spawns `dist/cli.js`: info, NDJSON, `--pretty`, CSV quoting, `--help`, missing file exit code |
| `fixtures/*.xlsx` | workbooks produced by **openpyxl 3.1.5**, never by our own writer; regenerated by `scripts/make-fixtures.py` |
| `scripts/make-fixtures.py` | the fixture generator (Python + openpyxl) |
| `scripts/verify-with-openpyxl.py` | optional third-implementation check of a file we wrote |
| `scripts/bench.mjs` | the benchmark behind [performance](performance.md) |

Tests import `excel-lite` by package name (Node's self-reference via `exports`) and run against the
**built `dist`**, not `src` — so packaging bugs fail the test suite. `npm test` builds first. CI runs
Node 22 and 24, typecheck, lint, tests, coverage, a browser-bundle check (no Node builtins, ≤120 kB,
writes and reads in a browser-like bundle), a docs link check, and a fixture-drift check that
regenerates the openpyxl fixtures and fails if they differ from the committed ones.

Adding a fixture: put the generator code in `scripts/make-fixtures.py` (not a hand-edited file), run
it, commit the `.xlsx`, and add a test. A fixture produced by a different implementation is the most
valuable kind of contribution ([CONTRIBUTING.md](../CONTRIBUTING.md)).

## Where to change what

| You want to… | Change | Watch out for |
|---|---|---|
| read a new cell type or attribute | `cellValue` / `toArray` in `src/read.ts`; `scanRows` in `src/xml.ts` if it is not already captured | add a fixture from another implementation, not a round trip |
| change how dates are detected | `BUILTIN_DATE_FORMATS` and `isDateFormat` in `src/read.ts` | false positives turn numbers into `Date`s |
| change header-key rules | `makeKeys` / `uniqueKey` in `src/read.ts` | dedupe against every used key; `column_N` naming |
| add a write option (style, merge, comment…) | `ColumnSpec`/`SheetInput` in `src/types.ts`, `buildSheetXml`/`cellXml` in `src/write.ts`, styles in `StyleBook` | both writers must stay in sync; update `docs/api.md` |
| change column widths or freeze/filter XML | `buildSheetXml`, `parseFreeze` in `src/write.ts` | Excel is picky about `cols`/`pane` attribute order |
| add a CLI flag or command | `parseArgs` and `main` in `src/cli.ts` | exit codes: 0 success (incl. EPIPE), 1 anything else |
| add an export | `src/index.ts` (and `src/types.ts` for types) | keep the API flat; every export needs a page in `docs/api.md` |
| improve large-file memory | `WorkbookReader.getSheetXml`/`rows` in `src/read.ts` (chunked inflate + incremental scanner) | this is the 0.2 streaming-reader work |
| change an error message | throw site (`src/read.ts`, `src/write.ts`, `src/zip.ts`, `src/fs.ts`) | messages start with `excel-lite:`; docs quote them verbatim |
| take `node:fs` out of a bundler's graph | already done in `src/fs.ts`; keep it that way | `scripts/check-browser-bundle.mjs` fails CI if a builtin reappears |
| change what gets published | `package.json` `files` | `docs/` ships, `examples/`, `test/` and `fixtures/` do not |

## Build and packaging

- **Build**: `tsup` (`tsup.config.ts`) emits ESM (`dist/index.js`, `dist/cli.js`), CJS
  (`dist/index.cjs`, `dist/cli.cjs`), `.d.ts`/`.d.cts` and source maps; `target: node22`, no code
  splitting. `dist/` is gitignored and rebuilt by `npm test`.
- **Entry points**: `exports` in `package.json` maps `import`/`require`/`types` for the package root;
  `bin` maps `excel-lite` → `dist/cli.js`.
- **Published files**: `dist`, `docs`, `README.md`, `CHANGELOG.md`, `LICENSE`.
- **Package size sanity**: `npm run check:bundle` limits the browser bundle to 120 kB and asserts it
  contains no Node builtins; measured at ~58 kB unminified ([browser](browser.md#bundle-size)).

## See also

- [Reading](reading.md) and [Writing](writing.md) — the behaviours shown in the flows.
- [Performance](performance.md) — why the read flow costs what it costs.
- [Browser](browser.md) — the `src/fs.ts` trick in context.
- [CONTRIBUTING.md](../CONTRIBUTING.md) — setup, rules and PR checklist.
