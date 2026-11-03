# excel-lite

Read and write `.xlsx` in Node and the browser. No native builds, one 30 kB dependency
(`fflate`), and an API small enough to keep in your head.

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
  and `readSheet()` takes a `Blob`/`ArrayBuffer`/`Uint8Array`. A browser-target bundle with
  esbuild is ~57 kB and contains zero Node builtins.

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
- sheet lookup by name or index, `header: true` with duplicate/blank header handling

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
| write 50,000 rows | **1.15 s** | 2.84 s | 1.69 s |
| read 50,000 rows | **457 ms** | 2.17 s | 2.04 s |
| peak RSS (read) | **80 MB** | 388 MB | 221 MB |
| file size written | **1.54 MB** | 1.51 MB | 11.54 MB |

Read is ~4.7× faster than ExcelJS and ~4.5× faster than SheetJS; peak memory is ~4.8×
lower than ExcelJS. RSS moves ±15% between runs, times ±5%.

Reproduce it: `npm run bench` (add a row count: `npm run bench -- 200000`).

The RSS number is the honest weak spot: rows are parsed lazily, but the sheet XML is
decompressed into memory first, so peak RSS lands around 8× the file size. That is fine
for the 1–50 MB files applications usually handle; a truly streaming parser is v0.2.

## What it deliberately does not do

- **Formulas**: read as their cached value. SheetJS keeps the formula text; `excel-lite`
  returns `7` where the cell says `=SUM(A1:A3)`.
- **Styles**: writes bold, fill, alignment, widths, freeze panes and number formats.
  Does not read formatting back out.
- **Legacy `.xls`**: only the OOXML container (`.xlsx` / `.xlsm` read; macros are ignored).
- **Charts, pivot tables, conditional formatting, data validation, images**: not modelled.
- **Cell-level random access**: rows in, rows out. If you need `sheet.getCell('B7')`,
  that is ExcelJS's job.

## Verified against other implementations

`excel-lite` is not self-certifying:

- tests read fixtures generated by **openpyxl** (`fixtures/`, regenerated by
  `scripts/make-fixtures.py`)
- tests read workbooks written by **ExcelJS** and **SheetJS**
- tests write workbooks and assert **ExcelJS**, **SheetJS** and **openpyxl** all read the
  same values, dates and number formats back

```bash
npm test        # 25 tests
npm run bench
npm run typecheck
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

## License

MIT
