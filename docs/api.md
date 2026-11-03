# API reference

Every name exported from `excel-lite` in 0.1.0, with signatures, parameters, return values,
thrown errors and a runnable example. Behaviour details live in [reading](reading.md) and
[writing](writing.md); this page is the shape of the API.

## On this page

- [Types](#types)
- [readSheet](#readsheet)
- [readWorkbook](#readworkbook) and [Workbook](#workbook)
- [listSheets](#listsheets)
- [WorkbookReader](#workbookreader)
- [toXlsx](#toxlsx)
- [writeSheet](#writesheet)
- [writeWorkbook](#writeworkbook)
- [serialToDate](#serialtodate) and [dateToSerial](#datetoserial)
- [columnLetter](#columnletter) and [columnIndexFromRef](#columnindexfromref)
- [Known type gaps](#known-type-gaps)

## Types

```ts
type CellValue = string | number | boolean | Date | null;
type RowObject = Record<string, CellValue>;
type SheetRow = CellValue[] | RowObject;

type Source = string | Uint8Array | ArrayBuffer | AsyncIterable<Uint8Array>;

interface ReadOptions {
  sheet?: string | number; // sheet name or 0-based index (numeric names fall back by name);
                           // default: the first sheet
  header?: boolean;        // first non-empty row becomes object keys
}

interface SheetInfo {
  name: string;
  index: number;    // 0-based position in the workbook
  rowCount: number; // physical <row> elements, header row included
}

interface ColumnSpec {
  key: string;                                        // property on object rows
  header?: string;                                    // header text; default: key
  width?: number;                                     // Excel character units
  format?: string;                                    // Excel number format code
  bold?: boolean;                                     // data cells only
  fill?: string;                                      // RRGGBB or AARRGGBB, '#' allowed
  align?: 'left' | 'center' | 'right';                // data cells only
}

type RowInput = CellValue[] | Record<string, CellValue | undefined>;

interface SheetInput {
  name?: string;
  rows: Iterable<RowInput>;
  columns?: ColumnSpec[];
  header?: boolean;
  freeze?: string | { row?: number; col?: number };
  autoFilter?: boolean;
  autoWidth?: boolean;
}

interface WriteOptions extends Omit<SheetInput, 'rows' | 'name'> {
  name?: string;
}
```

All types are exported as **types only** (`export type`), so import them with
`import type { CellValue } from 'excel-lite'` or `import { type CellValue } from 'excel-lite'`.

## readSheet

```ts
function readSheet(source: Source, options: ReadOptions & { header: true }): AsyncGenerator<RowObject>;
function readSheet(source: Source, options?: ReadOptions & { header?: false }): AsyncGenerator<CellValue[]>;
function readSheet(source: Source, options?: ReadOptions): AsyncGenerator<SheetRow>; // implementation signature
```

Reads one sheet as an async generator of rows. The workbook is opened (unzip + metadata) on the
**first** `next()`, so errors appear while iterating.

| Parameter | Type | Default | Meaning |
|---|---|---|---|
| `source` | `Source` | required | file path (Node 22.3+), `Uint8Array`/`Buffer`, `ArrayBuffer`, `Blob`, or async iterable of byte chunks |
| `options.sheet` | `string \| number` | `0` | sheet name (exact, then case-insensitive) or 0-based index; a number outside the index range is matched against a sheet name of that number |
| `options.header` | `boolean` | `false` | first non-empty row becomes the object keys |

**Returns** an `AsyncGenerator`: `RowObject` per row when `header: true`, `CellValue[]` otherwise.

**Throws** while iterating: `Error` for a non-zip file, a missing `xl/workbook.xml`, a malformed
workbook, a missing sheet part, an unknown sheet name, or the 512 MiB decompression limit;
`RangeError` for an out-of-range sheet index; `TypeError` for an unsupported source; the
platform's `ENOENT` for a missing path. Exact messages are in [reading](reading.md#errors-thrown).

```js
import { readSheet } from 'excel-lite';

for await (const row of readSheet('report.xlsx', { header: true })) {
  console.log(row.email, row.amount);
}

for await (const row of readSheet('report.xlsx', { sheet: 'Notes' })) {
  console.log(row[0], row[1]);
}
```

## readWorkbook

```ts
function readWorkbook(source: Source): Promise<Workbook>;
```

Opens the workbook once and returns handles for its sheets. Use it when you read several sheets
or iterate the same sheet more than once: `readSheet()` re-opens the file on every call.

| Parameter | Type | Default | Meaning |
|---|---|---|---|
| `source` | `Source` | required | same sources as `readSheet` |

**Returns** `Promise<Workbook>`. **Throws** the same open-time errors as `readSheet`, but as a
rejected promise (this is a plain async function, not a generator).

```js
import { readWorkbook } from 'excel-lite';

const workbook = await readWorkbook('book.xlsx');
console.log(workbook.sheetNames); // ['Data', 'Notes']

for await (const row of workbook.sheet('Data', { header: true })) console.log(row);
for await (const row of workbook.sheet(1)) console.log(row); // by index
```

### Workbook

The interface returned by `readWorkbook`. It is declared in `src/read.ts` but **not
re-exported** from the package entry point in 0.1.0 — see [known type gaps](#known-type-gaps).

```ts
interface Workbook {
  sheetNames: string[];
  sheet<T extends boolean = false>(
    sheet?: string | number,
    options?: { header?: T },
  ): AsyncGenerator<T extends true ? RowObject : CellValue[]>;
  sheets<T extends boolean = false>(
    options?: { header?: T },
  ): AsyncGenerator<T extends true ? RowObject : CellValue[]>[]; // note: an array
}
```

| Member | Returns | Notes |
|---|---|---|
| `sheetNames` | `string[]` | names in workbook order |
| `sheet(nameOrIndex?, { header? })` | `AsyncGenerator` | a **new** generator per call, so call it again to re-iterate; default: first sheet |
| `sheets({ header? })` | `AsyncGenerator[]` | one generator per sheet, in order; loop over the array and iterate each |

`sheet()` and `sheets()` use the same name/index resolution and header rules as `readSheet`,
including `RangeError`/`Error` for a bad sheet.

```js
const workbook = await readWorkbook(bytes);
for (const stream of workbook.sheets({ header: true })) {
  for await (const row of stream) console.log(row);
}
```

## listSheets

```ts
function listSheets(source: Source): Promise<SheetInfo[]>;
```

Sheet names, 0-based indexes and row counts, without materialising cell values. Row counts come
from counting `<row>` elements, so the sheet XML is still inflated — this is not free metadata.

| Parameter | Type | Default | Meaning |
|---|---|---|---|
| `source` | `Source` | required | same sources as `readSheet` |

**Returns** `Promise<SheetInfo[]>`. **Throws** the open-time errors and a missing sheet part
error (one sheet at a time, in order).

```js
import { listSheets } from 'excel-lite';

console.log(await listSheets('book.xlsx'));
// [ { name: 'Data', index: 0, rowCount: 50001 }, { name: 'Notes', index: 1, rowCount: 3 } ]
```

## WorkbookReader

```ts
class WorkbookReader {
  static open(source: Source): Promise<WorkbookReader>;
  readonly sheetNames: string[];
  resolveSheet(sheet: string | number | undefined): number;
  rowCount(index: number): number;
  rows(index: number, header: boolean): Generator<CellValue[] | RowObject>;
}
```

A lower-level handle on an opened workbook. `readWorkbook()` is a friendlier wrapper around it;
use `WorkbookReader` when you want the sync `rows()` generator or `rowCount()` directly.

| Member | Signature | Notes |
|---|---|---|
| `WorkbookReader.open` | `(source) => Promise<WorkbookReader>` | unzips the workbook and its relationships; throws the open-time errors |
| `sheetNames` | `string[]` | workbook order |
| `resolveSheet` | `(sheet) => number` | resolves name/index to a 0-based index; throws `RangeError` or `Error` |
| `rowCount` | `(index) => number` | counts `<row>` elements; throws if the sheet part is missing |
| `rows` | `(index, header) => Generator` | a **synchronous** generator of rows; throws synchronously when iterated |

```js
import { WorkbookReader } from 'excel-lite';

const reader = await WorkbookReader.open('book.xlsx');
console.log(reader.sheetNames);

const index = reader.resolveSheet('Data');
console.log('rows:', reader.rowCount(index));
for (const row of reader.rows(index, true)) console.log(row); // sync iteration
```

The constructor is private; always use `WorkbookReader.open`. The instance keeps the whole
workbook buffer in memory, so it is cheaper to reuse than to call `readSheet()` per sheet.

## toXlsx

```ts
function toXlsx(sheets: SheetInput[]): Uint8Array;
```

Builds an `.xlsx` workbook in memory. Works in Node, browsers, workers, Bun and Deno — no file
system, no Node builtins.

| Parameter | Type | Default | Meaning |
|---|---|---|---|
| `sheets` | `SheetInput[]` | required, at least one | one entry per worksheet |

**Returns** `Uint8Array` — a complete zip container (`PK…`).

**Throws** `Error` for an empty array, an invalid `freeze` string, or an invalid `fill` colour.
Exact messages: [writing](writing.md#errors-thrown).

```js
import { toXlsx } from 'excel-lite';

const bytes = toXlsx([
  { name: 'Report', rows: [{ id: 1, amount: 1234.5 }], columns: [{ key: 'id' }, { key: 'amount', format: '#,##0.00' }] },
]);

// e.g. in a request handler
return new Response(bytes, {
  headers: {
    'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'content-disposition': 'attachment; filename="report.xlsx"',
  },
});
```

## writeSheet

```ts
function writeSheet(destination: string, rows: Iterable<RowInput>, options?: WriteOptions): Promise<Uint8Array>;
```

One-sheet convenience wrapper: `writeWorkbook(destination, [{ ...options, rows }])`.

| Parameter | Type | Default | Meaning |
|---|---|---|---|
| `destination` | `string` | required | file path (Node 22.3+); overwritten if it exists |
| `rows` | `Iterable<RowInput>` | required | array rows or object rows, see [writing](writing.md#rows-arrays-and-objects) |
| `options` | `WriteOptions` | `{}` | `name`, `columns`, `header`, `freeze`, `autoFilter`, `autoWidth` |

**Returns** `Promise<Uint8Array>` — the bytes written. **Throws** the `toXlsx` errors, the
filesystem error for the path, and the Node-version error in non-Node runtimes.

```js
import { writeSheet } from 'excel-lite';

await writeSheet('report.xlsx', rows, {
  name: 'Report',
  columns: [{ key: 'email', width: 32 }, { key: 'amount', format: '#,##0.00' }],
  freeze: 'A2',
});
```

## writeWorkbook

```ts
function writeWorkbook(destination: string, sheets: SheetInput[]): Promise<Uint8Array>;
```

Writes several sheets to a file.

| Parameter | Type | Default | Meaning |
|---|---|---|---|
| `destination` | `string` | required | file path (Node 22.3+); overwritten if it exists |
| `sheets` | `SheetInput[]` | required, at least one | sheets in tab order |

**Returns** `Promise<Uint8Array>`. **Throws** the `toXlsx` errors plus filesystem errors.

```js
import { writeWorkbook } from 'excel-lite';

await writeWorkbook('book.xlsx', [
  { name: 'Summary', rows: summaryRows, columns: [{ key: 'total', format: '#,##0.00' }] },
  { name: 'Raw', rows: [[1, 2, 3]] },
]);
```

## serialToDate

```ts
function serialToDate(serial: number, date1904?: boolean): Date;
```

Converts an Excel serial number to a `Date` in **UTC**.

| Parameter | Type | Default | Meaning |
|---|---|---|---|
| `serial` | `number` | required | days since the Excel epoch; fractions are times of day |
| `date1904` | `boolean` | `false` | use the 1904 date system |

**Returns** `Date`. Serials below 61 are shifted by one day to account for Excel's phantom
1900-02-29; serial `0` is `1899-12-30T00:00:00.000Z`. Never throws (an out-of-range number
would produce an invalid `Date`, not an exception).

```js
import { serialToDate } from 'excel-lite';

serialToDate(45429).toISOString();            // '2024-05-17T00:00:00.000Z'
serialToDate(45429.5).toISOString();          // '2024-05-17T12:00:00.000Z'
serialToDate(45429, true).toISOString();      // '2028-05-18T00:00:00.000Z' (1904 system)
```

## dateToSerial

```ts
function dateToSerial(date: Date): number;
```

Reverse of `serialToDate`: a JS `Date` becomes an Excel serial number (1900 date system).
Values before the epoch produce negative serials; an invalid `Date` returns `NaN`.

| Parameter | Type | Meaning |
|---|---|---|
| `date` | `Date` | any instant; only `.getTime()` is used |

**Returns** `number`.

```js
import { dateToSerial } from 'excel-lite';

dateToSerial(new Date(Date.UTC(2024, 4, 17))); // 45429
dateToSerial(new Date(Date.UTC(1900, 0, 1)));  // 1 (serial 1 is 1900-01-01)
dateToSerial(new Date('nope'));                // NaN
```

## columnLetter

```ts
function columnLetter(index: number): string;
```

0-based column index → letters. `0 → A`, `25 → Z`, `26 → AA`, `16383 → XFD`.
Negative input returns `''`; there is no upper bound, so indexes past `XFD` produce letters Excel
does not accept.

```js
import { columnLetter } from 'excel-lite';

columnLetter(0);     // 'A'
columnLetter(26);    // 'AA'
columnLetter(16383); // 'XFD'
```

## columnIndexFromRef

```ts
function columnIndexFromRef(ref: string): number;
```

Cell reference → 0-based column index. Reads leading ASCII letters and stops at the first
non-letter, so `'C12'` and `'C'` both give `2`. Lowercase letters are accepted (`'a1'` → `0`).
Returns `-1` when the reference does not start with a letter (`'$A$1'`, `'12'`, `''`).

| Parameter | Type | Meaning |
|---|---|---|
| `ref` | `string` | a cell reference or column reference |

**Returns** `number`.

```js
import { columnIndexFromRef } from 'excel-lite';

columnIndexFromRef('A');          // 0
columnIndexFromRef('C12');        // 2
columnIndexFromRef('XFD1048576'); // 16383
columnIndexFromRef('12');         // -1
```

## Known type gaps

Two gaps in 0.1.0's types, both runtime-supported:

1. **`Source` does not include `Blob`.** `readSheet(blob)` works in browsers and Node 18+, but
   TypeScript rejects it. Wrap it (`new Uint8Array(await blob.arrayBuffer())`) or cast.
2. **`Workbook` is not exported.** The interface exists in the declarations but the package entry
   point does not re-export it, so `import type { Workbook } from 'excel-lite'` fails. Infer it:

```ts
import { readWorkbook } from 'excel-lite';

type Workbook = Awaited<ReturnType<typeof readWorkbook>>;
```

The internal `SheetMatrix` type is also exported from `src/types.ts` but not from the entry point;
it is unused by the public API.

## See also

- [Reading](reading.md) — behaviour behind `readSheet`, `readWorkbook`, `listSheets`.
- [Writing](writing.md) — behaviour behind `toXlsx`, `writeSheet`, `writeWorkbook`.
- [Migration](migration.md) — mapping from SheetJS and ExcelJS APIs.
