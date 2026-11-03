# Migrating from SheetJS or ExcelJS

`excel-lite` covers the two jobs most applications actually do — import rows, export rows — and
deliberately does not cover workbook editing. This page maps the common calls from `xlsx`
(SheetJS) and `exceljs` to their equivalents, and lists what will not port 1:1.

## On this page

- [From SheetJS (`xlsx`)](#from-sheetjs-the-xlsx-package)
- [From ExcelJS](#from-exceljs)
- [Behaviour differences that bite](#behaviour-differences-that-bite)
- [What will not port 1:1](#what-will-not-port-one-to-one)
- [Migration checklist](#migration-checklist)

## From SheetJS (the xlsx package)

### Reading a file

```js
// before — SheetJS
import XLSX from 'xlsx';

const workbook = XLSX.readFile('report.xlsx', { cellDates: true });
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

// after — excel-lite
import { readSheet } from 'excel-lite';

const rows = [];
for await (const row of readSheet('report.xlsx', { header: true })) rows.push(row);
```

`sheet_to_json` with default options skips blank rows and gives objects keyed by the first row —
the same shape as `{ header: true }`. If you used `header: 1` you get arrays, the same as default
`readSheet()`.

### Reading from bytes

```js
// before
const workbook = XLSX.read(buffer, { type: 'buffer' });

// after — the Buffer is already a Uint8Array
const workbook = await readWorkbook(buffer);
```

### Writing a file

```js
// before
const sheet = XLSX.utils.json_to_sheet(rows);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, sheet, 'Report');
XLSX.writeFile(workbook, 'report.xlsx');

// after
import { writeSheet } from 'excel-lite';

await writeSheet('report.xlsx', rows, {
  name: 'Report',
  columns: [{ key: 'email', width: 32 }, { key: 'amount', format: '#,##0.00' }],
});
```

Array-of-arrays input works too: `writeSheet(path, [[1, 2], [3, 4]])` — with no `columns`, no
header row is written.

### Sheet and value access

| SheetJS | excel-lite |
|---|---|
| `workbook.SheetNames` | `workbook.sheetNames` (from `readWorkbook`) |
| `workbook.Sheets['Data']` | `workbook.sheet('Data')` — an async iterable, not a sheet object |
| `XLSX.utils.sheet_to_json(ws)` | `for await (const row of readSheet(file))` |
| `ws['B7'].v` | not supported — rows only |
| `ws['B7'].f` (formula text) | not supported — the cached value is returned |
| `ws['B7'].w` (formatted text) | not supported — the reader returns typed values |
| `ws['!cols']` (widths) | not read |

## From ExcelJS

### Reading a file

```js
// before — ExcelJS
import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile('report.xlsx');
const sheet = workbook.getWorksheet('Data');
sheet.eachRow((row, rowNumber) => {
  console.log(rowNumber, row.getCell(1).value, row.getCell(2).value);
});

// after — excel-lite
import { readWorkbook } from 'excel-lite';

const workbook = await readWorkbook('report.xlsx');
for await (const row of workbook.sheet('Data')) {
  console.log(row[0], row[1]);
}
```

ExcelJS gives you a cell object with `.value`, `.numFmt`, `.font`, `.fill`, `.address` and so on.
`excel-lite` gives you the value only — that is the trade for the memory and speed difference in
[performance](performance.md).

### Writing a file

```js
// before — ExcelJS
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Report');
sheet.columns = [
  { header: 'Email', key: 'email', width: 32 },
  { header: 'Amount', key: 'amount', width: 14, style: { numFmt: '#,##0.00' } },
];
for (const row of rows) sheet.addRow(row);
sheet.views = [{ state: 'frozen', ySplit: 1 }];
await workbook.xlsx.writeFile('report.xlsx');

// after — excel-lite
import { writeSheet } from 'excel-lite';

await writeSheet('report.xlsx', rows, {
  name: 'Report',
  columns: [
    { key: 'email', header: 'Email', width: 32 },
    { key: 'amount', header: 'Amount', format: '#,##0.00' },
  ],
  freeze: 'A2',
});
```

`autoFilter`, computed widths and the bold header row come for free with the `excel-lite` version;
borders, vertical alignment, merges and per-cell styles do not exist.

### Style and structure access

| ExcelJS | excel-lite |
|---|---|
| `worksheet.getCell('B7').value` | not supported — iterate rows |
| `cell.numFmt` / `.font` / `.fill` / `.alignment` | not read back (written only, per column) |
| `worksheet.mergeCells(...)` | not supported |
| `worksheet.getImages()` / `addImage` | not supported |
| `worksheet.eachRow()` | `for await (const row of workbook.sheet(...))` |
| `worksheet.columns = [...]` | `columns: [...]` in `SheetInput` |
| `worksheet.views` (freeze) | `freeze: 'A2'` |
| `worksheet.autoFilter` | `autoFilter: true` (header-row range only) |
| `cell.comment`, `cell.note`, hyperlinks | not supported |
| `ExcelJS.stream.xlsx.WorkbookReader/Writer` | not supported — the reader is lazy per row but not streamed from disk |

## Behaviour differences that bite

Beyond the API shape, these are the semantic differences to check when porting:

1. **Dates.**
   - SheetJS without `cellDates: true` returns serial numbers. With it, dates are anchored to
     **local** midnight — on a `+05:00` machine, serial `45429` came back as
     `2024-05-16T18:59:48Z`, i.e. the previous day in UTC.
   - ExcelJS returns `Date` objects (in our tests, the same UTC instant).
   - `excel-lite` always returns `Date` at **UTC midnight** for date cells, regardless of the
     machine's timezone. If your code formatted SheetJS dates with local getters, it now gets UTC
     getters' answers — that is usually the fix, not a workaround.
2. **Percentages.** SheetJS formats a percentage as text in some modes; `excel-lite` returns the
   underlying number (`12.5%` → `0.125`) and never the display string.
3. **Error cells.** ExcelJS exposes `{ error: '#N/A' }`; `excel-lite` returns the `string` `'#N/A'`.
4. **Blank cells.** SheetJS `sheet_to_json` omits keys unless you pass `defval`; `excel-lite`
   `header: true` always includes every key, with `null` for missing cells.
5. **Missing sheets.** `workbook.Sheets['Nope']` is `undefined` in SheetJS; `readSheet(file, { sheet: 'Nope' })`
   throws with the list of available names.
6. **Formula text.** SheetJS keeps `cell.f`; `excel-lite` returns the cached value and drops the
   formula. Read-only formula inspection needs SheetJS.
7. **Style read-back.** ExcelJS reads formatting; `excel-lite` does not. If a port reads styles it
   must keep ExcelJS for that step ([recipes](recipes.md#11-exceljs-interop) shows the hybrid).
8. **Merges.** ExcelJS exposes merged ranges; `excel-lite` returns the top-left value and `null`
   for the covered cells, with no merge metadata.
9. **Formats.** SheetJS reads `.xls`, `.ods`, `.csv`, HTML and more; `excel-lite` reads `.xlsx`
   and `.xlsm` only. Convert legacy inputs first.
10. **Memory.** ExcelJS builds a full workbook model; `excel-lite` parses rows lazily. The
    benchmark in [performance](performance.md) shows ~4× less peak RSS for the read path, and
    SheetJS-written files are ~7.5× larger than either ExcelJS's or `excel-lite`'s output.

## What will not port one to one

These have no `excel-lite` equivalent in 0.1 — plan for a second tool or a workflow change:

- cell-level formatting (borders, fonts other than bold, vertical alignment, text wrap);
- reading formatting back out of a file;
- merged cells (read or write);
- formulas (writing them, or reading their text);
- comments, notes, hyperlinks;
- images, charts, pivot tables, sparklines, shapes;
- conditional formatting and data validation;
- named ranges, defined names, print settings, page setup, protection;
- row heights, hidden rows/columns, tab colours, sheet visibility;
- `.xls` input, `.ods`, `.csv`, HTML tables;
- random access by cell (`sheet.getCell('B7')`) and in-place workbook edits;
- SheetJS's `cell.f` / `cell.w` / `cell.z` metadata, its `sheet_to_json` options (`defval`,
  `raw`, `dateNF`, `blankrows`, `header: 'A'`), and dense-mode worksheets;
- ExcelJS's rich cell values (`{ formula }`, `{ richText }`, `{ hyperlink }`, `{ error }`),
  `eachSheet`, `getSheetValues`, streaming reader/writer, and style objects;
- macros: `excel-lite` never writes a macro-enabled package.

## Migration checklist

1. Replace read calls with `readSheet()` / `readWorkbook()`; keep a row loop with `try/catch`,
   because reader errors are thrown during iteration.
2. Replace write calls with `writeSheet()` / `writeWorkbook()`; move column metadata into
   `columns` and freeze state into `freeze`.
3. Audit date handling: convert local-time formatting to `getUTC*`/`toISOString`, and write dates
   with `Date.UTC(...)` or ISO date strings.
4. Replace formula/format reads with the cached value, or keep ExcelJS for those steps and hand
   values between the two ([interop](recipes.md#11-exceljs-interop)).
5. Replace `.xls` inputs with a one-time conversion — `excel-lite` refuses them by design.
6. Run the new code against your real files, not just fixtures: dates, unicode, sparse rows and
   percentages are where ports break.
7. Cap upload sizes: peak memory is several times the file size ([performance](performance.md)).

## See also

- [Compatibility](compatibility.md) — what is written, read and ignored.
- [Reading](reading.md) / [Writing](writing.md) — the behaviour behind the mapping.
- [Recipes](recipes.md#11-exceljs-interop) — using excel-lite and ExcelJS together.
