# excel-lite

[![npm](https://img.shields.io/npm/v/excel-lite.svg)](https://www.npmjs.com/package/excel-lite)
[![ci](https://github.com/horizonbymuneeb/excel-lite/actions/workflows/ci.yml/badge.svg)](https://github.com/horizonbymuneeb/excel-lite/actions/workflows/ci.yml)
[![docs](https://img.shields.io/endpoint?url=https://horizonbymuneeb.github.io/excel-lite/badges/docs.json)](https://horizonbymuneeb.github.io/excel-lite/)
[![tests](https://img.shields.io/endpoint?url=https://horizonbymuneeb.github.io/excel-lite/badges/tests.json)](https://github.com/horizonbymuneeb/excel-lite/actions/workflows/ci.yml)
[![coverage](https://img.shields.io/endpoint?url=https://horizonbymuneeb.github.io/excel-lite/badges/coverage.json)](https://horizonbymuneeb.github.io/excel-lite/performance.html)
[![browser bundle](https://img.shields.io/endpoint?url=https://horizonbymuneeb.github.io/excel-lite/badges/bundle.json)](https://horizonbymuneeb.github.io/excel-lite/browser.html)
[![license](https://img.shields.io/npm/l/excel-lite.svg)](LICENSE)
[![node](https://img.shields.io/node/v/excel-lite.svg)](https://nodejs.org)

Read and write `.xlsx` in Node and the browser. No native builds, one 30 kB dependency
(`fflate`), and an API small enough to keep in your head.

**[Docs site](https://horizonbymuneeb.github.io/excel-lite/)** ·
**[Docs source](docs/README.md)** ·
[Reading](docs/reading.md) ·
[Writing](docs/writing.md) ·
[API](docs/api.md) ·
[CLI](docs/cli.md) ·
[Recipes](docs/recipes.md) ·
[Troubleshooting](docs/troubleshooting.md)

```ts
import { readSheet, writeSheet } from 'excel-lite';

for await (const row of readSheet('report.xlsx', { header: true })) {
  console.log(row.email, row.amount); // dates are Dates, numbers are numbers
}

await writeSheet('report.xlsx', rows, {
  columns: [
    { key: 'email', width: 32 },
    { key: 'amount', format: '#,##0.00' },
    { key: 'joined', format: 'yyyy-mm-dd' },
  ],
  freeze: 'A2',
});
```

## Why this exists

Two Excel packages own this space and both left npm behind:

| | downloads/week | last npm release | state |
|---|---|---|---|
| `xlsx` (SheetJS) | ~15.5M | **0.18.5, March 2022** | fixes ship on their own CDN, not npm |
| `exceljs` | ~26.2M | **v4.4.0, 37 months ago** | 800+ open issues, no release since |
| `excel-lite` | you are here | — | ESM + CJS, typed, tested against both |

If you run `npm i xlsx` today you get the 2022 build. `excel-lite` is not trying to be a
spreadsheet engine — it does the two things applications actually do (import rows,
export rows) and does them on a modern runtime.

## Install

```bash
npm install excel-lite
```

- **Node 22+** for anything that takes a file path — the library reads builtins through
  `process.getBuiltinModule`, so there is no `node:fs` import in the bundle.
- **Browsers, workers, Bun, Deno** for everything else: `toXlsx()` returns a `Uint8Array`
  and `readSheet()` takes a `Blob`/`ArrayBuffer`/`Uint8Array`. The esbuild browser bundle is
  ~57 kB raw, 31 kB minified and 13 kB gzipped, and contains zero Node builtins.

## Reading

```ts
import { readSheet, readWorkbook, listSheets } from 'excel-lite';

// stream rows as arrays
for await (const row of readSheet('data.xlsx')) console.log(row);

// first row as keys
for await (const row of readSheet('data.xlsx', { header: true })) {
  console.log(row.name);
}

// pick a sheet, then iterate more than once
const workbook = await readWorkbook('data.xlsx');
console.log(workbook.sheetNames); // ['Data', 'Notes']
for await (const row of workbook.sheet('Notes', { header: true })) console.log(row);

// sheet names + row counts, without parsing cells
console.log(await listSheets('data.xlsx')); // [{ name: 'Data', index: 0, rowCount: 50001 }]
```

`readSheet` accepts a file path, `Buffer`/`Uint8Array`, `ArrayBuffer`, `Blob`, or an async
iterable of byte chunks (fetch bodies, Node streams).

What the reader understands:

- shared strings, inline strings and rich-text runs
- numbers, percentages (as numbers), booleans and error values (`'#DIV/0!'`)
- dates and times, including custom `numFmt` patterns and 1904 date-system workbooks —
  returned as `Date` objects in **UTC**, so a server in Karachi and a server in Lisbon
  return the same instant
- sparse rows (`A` and `C` present, `B` missing → `null` in the middle), blank rows skipped
- sheet lookup by name or index, `header: true` with unique keys even when headers repeat
- Excel's phantom 1900-02-29 (serial 60) has no representable date; it reads as 1900-03-01
  and is written back as serial 61

## Writing

```ts
import { toXlsx, writeSheet, writeWorkbook } from 'excel-lite';

await writeWorkbook('book.xlsx', [
  {
    name: 'Report',
    rows: [{ email: 'a@b.co', amount: 1234.5, joined: new Date('2024-05-17') }],
    columns: [{ key: 'email' }, { key: 'amount', format: '#,##0.00' }, { key: 'joined', format: 'yyyy-mm-dd' }],
    freeze: 'A2',      // freeze the header row
    autoFilter: true,  // filter dropdowns (default when a header is written)
  },
  { name: 'Raw', rows: [[1, 2, 3], [4, 5, 6]] },
]);

// browser / worker / any environment without fs
const bytes: Uint8Array = toXlsx([{ rows }]);
```

Column options: `key`, `header`, `width` (Excel units, auto-computed by default),
`format` (any Excel number format), `bold`, `fill` (`RRGGBB` / `AARRGGBB`), `align`.

Cells accept `string | number | boolean | Date | null | undefined`. Invalid XML characters
are stripped — Excel refuses to open a workbook that contains them. Sheet names are
sanitised (31 char limit, `[]:*?/\` removed, duplicates renamed).

## CLI

```bash
npx excel-lite info report.xlsx
npx excel-lite to-json report.xlsx --header > rows.ndjson
npx excel-lite to-csv report.xlsx --out report.csv
```

`to-json` is NDJSON by default (one row per line — pipe it into anything), or a single
array with `--pretty`.

## Benchmarks

50,000 rows × 6 columns, Node 26, median of 3 runs. Each run gets its own process, RSS is
`ru_maxrss` minus a bare-Node baseline, and **every reader reads the same SheetJS-written
file** so nobody gets a file their own writer shaped.

| | excel-lite | ExcelJS | SheetJS |
|---|---|---|---|
| write 50,000 rows | **1.10 s** | 2.73 s | 1.62 s |
| read 50,000 rows | **450 ms** | 2.19 s | 1.94 s |
| peak RSS (read) | **62 MB** | 340 MB | 182 MB |
| file size written | **1.54 MB** | 1.51 MB | 11.54 MB |

Read is ~4.9× faster than ExcelJS and ~4.3× faster than SheetJS; peak memory is ~5.5× and
~2.9× lower. Write is ~2.5× faster than ExcelJS. RSS moves ±15% between runs, times ±5%.

Reproduce it: `npm run bench` (add a row count: `npm run bench -- 200000`).

The RSS number is the honest weak spot: rows are parsed lazily, but the sheet XML is
decompressed into memory before the first row arrives, so peak memory scales with the amount
of cell data — 62 MB for the 50k-row benchmark above, ~218 MB for a 100k-row/3.5 MB file
(measured by `examples/streaming-large-file.mjs`). Fine for the files applications usually
handle; a truly streaming parser that keeps memory flat is the headline item for
[0.2](ROADMAP.md).

## What it deliberately does not do

- **Formulas**: read as their cached value. SheetJS keeps the formula text; `excel-lite`
  returns `7` where the cell says `=SUM(A1:A3)`.
- **Styles**: writes bold, fill, alignment, widths, freeze panes and number formats.
  Column `bold`/`fill`/`align` apply to data cells; the header row is always bold.
  Does not read formatting back out.
- **Legacy `.xls`**: only the OOXML container (`.xlsx` / `.xlsm` read; macros are ignored).
- **Charts, pivot tables, conditional formatting, data validation, images**: not modelled.
- **Cell-level random access**: rows in, rows out. If you need `sheet.getCell('B7')`,
  that is ExcelJS's job.

## Verified against other implementations

`excel-lite` is not self-certifying:

- the suite reads fixtures generated by **openpyxl** (`fixtures/`, 9 files, regenerated by
  `scripts/make-fixtures.py`) and a hand-written OOXML fixture no library would produce
- the suite reads workbooks written by **ExcelJS** and **SheetJS**, and asserts those two
  read back every value, date and number format our writer produces
- **openpyxl** read-back is checked by the opt-in `scripts/verify-with-openpyxl.py`; it is
  not part of `npm test` because it needs Python

```bash
npm test            # 157 tests
npm run typecheck
npm run lint
npm run coverage    # 93% statements against src
npm run bench
npm run check:links
npm run check:bundle
```

## API

```ts
readSheet(source, { sheet?, header? }): AsyncGenerator<CellValue[] | Record<string, CellValue>>
readWorkbook(source): Promise<{ sheetNames, sheet(name?, { header? }), sheets() }>
listSheets(source): Promise<{ name, index, rowCount }[]>
toXlsx(sheets: SheetInput[]): Uint8Array
writeSheet(dest, rows, options): Promise<Uint8Array>
writeWorkbook(dest, sheets): Promise<Uint8Array>
```

## Docs

The same pages are published as a site: **<https://horizonbymuneeb.github.io/excel-lite/>** —
built from this repository by `.github/workflows/docs.yml`, so the docs and the code can never
drift. Source of truth stays in [`docs/`](docs/README.md).

| Page | What it covers |
|---|---|
| [Getting started](docs/getting-started.md) | install, first read, first write, first CLI command |
| [Reading](docs/reading.md) | sources, sheet selection, header rules, value mapping, dates, errors |
| [Writing](docs/writing.md) | every option, number formats, sheet names, Excel limits |
| [API reference](docs/api.md) | every export, signature, parameter and thrown error |
| [CLI](docs/cli.md) | commands, flags, exit codes, piping |
| [Recipes](docs/recipes.md) | 12 end-to-end recipes (Express, Postgres `COPY`, S3, browser download, zod…) |
| [Performance](docs/performance.md) | benchmark method, memory model, when to use something else |
| [Compatibility](docs/compatibility.md) | what is read, written, ignored, and what is untested |
| [Migration](docs/migration.md) | moving from SheetJS or ExcelJS |
| [Troubleshooting](docs/troubleshooting.md) | 17 concrete failures and their fixes |
| [Architecture](docs/architecture.md) | how the reader/writer work, where to change what |
| [Browser](docs/browser.md) | bundlers, no-Node-builtins mechanism, React example |

## Contributing

Contributions are welcome — the highest-value thing you can send is a workbook produced by a
different implementation that this reader gets wrong. See [CONTRIBUTING.md](CONTRIBUTING.md)
for setup, test policy and the PR checklist; [ROADMAP.md](ROADMAP.md) says what is already
planned; [SECURITY.md](SECURITY.md) explains how to report a vulnerability privately.

## License

MIT © [Muneeb Anjum](https://github.com/horizonbymuneeb)
