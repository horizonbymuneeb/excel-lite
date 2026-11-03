# Reading workbooks

`readSheet()` is the entry point: give it a workbook and it yields rows. This page documents
every reader behaviour — sources, sheet selection, header mode, the value mapping, dates,
sparse rows, row counts and the exact errors you can hit.

## On this page

- [Sources](#sources)
- [Iterating](#iterating)
- [Selecting a sheet](#selecting-a-sheet)
- [Header mode](#header-mode)
- [Value mapping](#value-mapping)
- [Dates, times and timezones](#dates-times-and-timezones)
- [Sparse rows and blank rows](#sparse-rows-and-blank-rows)
- [Row counts and listSheets](#row-counts-and-listsheets)
- [Errors thrown](#errors-thrown)
- [What the reader ignores](#what-the-reader-ignores)
- [Limits](#limits)

## Sources

`readSheet(source)` and every other reader accept these:

| Source | Works in | Notes |
|---|---|---|
| `string` — a file path | Node 22.3+ | read with `fs/promises`; a missing file rejects with the usual `ENOENT` error |
| `Uint8Array` | everywhere | includes Node `Buffer` (a `Buffer` is a `Uint8Array`) |
| `ArrayBuffer` | everywhere | wrapped without copying |
| `Blob` | browsers, workers, Node 18+ | `await source.arrayBuffer()` under the hood |
| `AsyncIterable<Uint8Array>` | everywhere | fetch bodies, Node `Readable` streams, S3 `Body` |

```js
import { createReadStream } from 'node:fs';
import { readSheet } from 'excel-lite';

// path
for await (const row of readSheet('data.xlsx')) { /* … */ }

// bytes
const response = await fetch('https://example.com/report.xlsx');
for await (const row of readSheet(new Uint8Array(await response.arrayBuffer()))) { /* … */ }

// stream: chunks are collected, then the zip is parsed
for await (const row of readSheet(createReadStream('data.xlsx'))) { /* … */ }
```

Two honest details:

- An async iterable is **collected into one buffer first**, then unzipped. It saves you from
  holding the file and the response in two places, but it is not a streaming parser.
- In 0.1.0 the exported `Source` type is `string | Uint8Array | ArrayBuffer | AsyncIterable<Uint8Array>`
  — `Blob` works at runtime but is missing from the type. For TypeScript, use
  `new Uint8Array(await blob.arrayBuffer())`, or cast: `readSheet(blob as unknown as Source)`.

Anything else throws a `TypeError`:

```
excel-lite: unsupported source — pass a file path, Uint8Array, ArrayBuffer, Blob or async iterable of bytes
```

## Iterating

`readSheet()` returns an async generator. Opening (unzip + workbook metadata) happens on the
first `next()`, so **errors surface while iterating, not at the call**:

```js
// does NOT throw here — the generator is lazy
const rows = readSheet('missing.xlsx');

try {
  for await (const row of rows) { /* … */ }
} catch (error) {
  console.error(error.message); // ENOENT: no such file or directory, open 'missing.xlsx'
}
```

Because it is lazy you can stop early — the reader stops parsing rows:

```js
for await (const row of readSheet('huge.xlsx')) {
  if (row[0] === 'STOP') break;
}
```

The sheet XML is still decompressed before the first row; see [performance](performance.md).
Each `readSheet()` call re-opens and re-parses the workbook. To read several sheets from one
open, use `readWorkbook()`:

```js
import { readWorkbook } from 'excel-lite';

const workbook = await readWorkbook('book.xlsx');
console.log(workbook.sheetNames); // ['Data', 'Notes']

for await (const row of workbook.sheet('Data', { header: true })) console.log(row);
for await (const row of workbook.sheet('Notes')) console.log(row);

// every sheet as a separate async generator, in workbook order
for (const stream of workbook.sheets({ header: true })) {
  for await (const row of stream) console.log(row);
}
```

`workbook.sheet()` returns a fresh generator on every call, so you can iterate the same sheet
twice. `workbook.sheets()` returns an **array of generators** (not one generator), which is why
the loop above takes a stream at a time.

`readWorkbook()` and `listSheets()` are plain async functions, so they reject immediately on a
bad file instead of on first iteration.

## Selecting a sheet

```js
readSheet(file);                          // the first sheet
readSheet(file, { sheet: 1 });            // 0-based index: the second sheet
readSheet(file, { sheet: 'Notes' });      // by name, exact match first
readSheet(file, { sheet: 'notes' });      // then case-insensitive match
readSheet(file, { sheet: 99 });           // RangeError
```

Rules, in order:

1. `sheet` omitted → index `0`.
2. A number is a **0-based index** into the workbook's sheet order. If the number is outside
   the workbook's range, it is matched against a sheet whose *name* is that number — so
   `{ sheet: 2 }` opens a sheet literally named `"2"` when there is no index `2`. If nothing
   matches, it throws `RangeError`.
3. A string is matched exactly first, then case-insensitively. Matching is exact apart from
   case — no trimming, no fuzzy matching.
4. No match throws with the available names listed.

Hidden sheets are read like any other: the reader does not look at the sheet `state` attribute.

## Header mode

`header: true` turns the **first non-empty row** into keys. Every later row is an object with
one property per key.

```js
for await (const row of readSheet('people.xlsx', { header: true })) console.log(row);
// { name: 'Ali', qty: 3, joined: 2023-07-04T00:00:00.000Z }
```

Key generation rules, in order:

| Situation in the header row | Resulting key | Example |
|---|---|---|
| text | the text, trimmed | `' Email '` → `email` |
| number / boolean / date | `String(value).trim()` | `42` → `'42'`; a `Date` becomes `String(date)` in your **local** timezone |
| empty cell or missing cell (a gap means a hole → `null`) | `column_<1-based column number>` | `A`, `C` present, `B` empty → `column_2` |
| repeated text | `<key>_2`, `<key>_3`, … | `name`, `name`, `name` → `name`, `name_2`, `name_3` |
| text equal to a key already used | the next free suffix | header `name`, blank, `name`, `name_2` → `name`, `column_2`, `name_2`, `name_2_2` |

Data-row rules:

- Every header key is present on every object. A cell that does not exist is `null`, never
  `undefined` and never omitted.
- Cells beyond the header row's width become `column_<1-based position>`, or the next free
  `column_N` if that name is already a key: a header of `['a', 'b']` and a row of `[1, 2, 3]`
  gives `{ a: 1, b: 2, column_3: 3 }`; a header of `['a', 'column_3']` gives
  `{ a: 1, column_3: 2, column_3_2: 3 }`.
- Keys are deduplicated against **every key already used**, including generated ones, so a later
  column never overwrites an earlier one. A header that literally says `name_2` beside two
  columns called `name` becomes `name_2_2`; a header containing `column_5` pushes a data cell to
  `column_5_2`.

Header keys should be text. A numeric or date header cell is stringified, and `String(date)`
includes the machine's local timezone offset, so the same file produces different keys on
different servers.

If the sheet has only one row, header mode yields zero rows. Blank rows never become the header
(first **non-empty** row wins). Reading with `header: true` on a file without a header row
silently eats your first data row — that is what the option asks for.

## Value mapping

The reader maps OOXML cell types and number formats to five JavaScript types: `string`,
`number`, `boolean`, `Date`, `null`.

| In the file | `t` attribute | You get | Notes |
|---|---|---|---|
| shared string | `s` | `string` | the shared-string table is loaded once, lazily; `<r>` rich-text runs are concatenated, phonetic `<rPh>` runs are ignored; an out-of-range index reads as `null` |
| inline string | `inlineStr` | `string` | `<t>` runs concatenated |
| formula string result | `str` | `string` | the cached text result |
| rich text (shared) | `s` | `string` | `Hello` + ` World` → `'Hello World'` |
| number | `n` or absent | `number` | also percentages: `12.5%` stored as `0.125` reads as `0.125` |
| number in a date-formatted cell | `n` or absent | `Date` | see [dates](#dates-times-and-timezones) |
| boolean | `b` | `boolean` | `1`/`true` → `true`, `0`/`false` → `false` |
| error value | `e` | `string` | `'#DIV/0!'`, `'#N/A'`, `'#VALUE!'` — the error text, not an exception |
| ISO date cell | `d` | `Date` | parsed with `new Date(value)`; an unparseable value is returned as its raw `string` |
| empty cell | any | `null` | also an empty shared-string reference |

There is no other type. `undefined` is never returned, numbers are never strings, and dates are
never numbers (unless the cell's format does not look like a date — see below).

Edge cases the code makes:

- A numeric cell whose text is not finite (`<v>NaN</v>`) is returned as the raw `string`.
- A shared-string cell with no index (`<c t="s"><v></v></c>`) is `null`.
- Trailing empty cells are trimmed, so `[1, null, null]` reads as `[1]`.
- Reading does not coerce: a number stored as text (cell type shared string) stays a `string`.

## Dates, times and timezones

Excel has no date type. A date is a number of days since an epoch, and the cell's number
format decides whether to show it as a date. The reader uses the same rule: **a numeric cell
whose style has a date/time format becomes a `Date`; everything else stays a number.**

`Date` objects are built in **UTC**, so the result does not depend on the machine's timezone:

| Stored serial | Cell style | You get |
|---|---|---|
| `45429` | `yyyy-mm-dd` | `2024-05-17T00:00:00.000Z` |
| `45429.5` | `yyyy-mm-dd hh:mm:ss` | `2024-05-17T12:00:00.000Z` |
| `0.270833…` | `h:mm` (a time-only cell) | `1899-12-30T06:30:00.000Z` — Excel's "zero date" |
| `45429` | `General` or no style | `45429` — a plain number |
| `45429` | any format | `2028-05-18T00:00:00.000Z` in a 1904-date-system workbook |

The timezone consequence, spelled out:

- `toISOString()` and `getUTCFullYear()`/`getUTCMonth()`/`getUTCDate()` give the date that is
  in the file.
- `getFullYear()`/`getMonth()`/`getDate()`/`toLocaleDateString()` use the **server's** local
  timezone. In a timezone behind UTC (e.g. `America/New_York`), UTC midnight is the previous
  evening locally, so `row.joined.getDate()` can return the day before. This is the classic
  "dates are off by one day" bug, and it is a presentation bug, not a parsing bug.
- The same asymmetry applies when writing: `new Date(2024, 4, 17)` is local midnight, which is
  an instant that may fall on the previous UTC day. Use `new Date(Date.UTC(2024, 4, 17))` or
  `new Date('2024-05-17')` (ISO date strings are parsed as UTC) — see [writing](writing.md#dates).

Which formats count as dates:

- built-in format ids `14`–`22`, `27`–`36`, `45`–`47`, `50`–`58` (Excel's date, time, duration
  and locale date variants);
- custom formats (`numFmtId >= 164`) whose code contains `y`, `m`, `d`, `h` or `s` after
  removing quoted literals (`"days"`), escaped characters (`\d`) and bracketed colour/locale
  sections (`[Red]`, `[$-409]`). Bracketed time parts (`[h]`, `[m]`, `[s]`) are kept.

So a number with a date-like custom format reads as a `Date` even if it is not a date, and a
real date stored with `General` reads as a number. This is a format-based heuristic, which is
the same thing Excel does; it is not guesswork about the value.

Excel's 1900 leap-year bug and the 1904 date system are handled: serials `1`–`60` are shifted by
one day, a workbook with `<workbookPr date1904="1"/>` is read with the 1904 epoch, and
`serialToDate(serial, true)` exposes the 1904 conversion for your own use.

Time-only values come back as a `Date` on **1899-12-30**. If all you want is the wall-clock
time, read `getUTCHours()`, `getUTCMinutes()`, `getUTCSeconds()` from that `Date`:

```js
const time = new Date(Date.UTC(1899, 11, 30, 6, 30)); // what a 06:30 cell reads as
console.log(time.getUTCHours(), time.getUTCMinutes()); // 6 30
```

## Sparse rows and blank rows

Rows are positional. When a file has cells in `A` and `C` but nothing in `B`, the row reads as
`['a', null, 'c']` — the hole is filled with `null`, never `undefined`, never compacted.

```js
import { readSheet, writeSheet } from 'excel-lite';

await writeSheet('sparse.xlsx', [['a', null, 'c'], [null, null], ['d']], { header: false });

for await (const row of readSheet('sparse.xlsx')) console.log(row);
// [ 'a', null, 'c' ]
// [ 'd' ]
```

Why only two rows: a row that contains no cells at all is **skipped** (the writer does not emit
it either, which is why the second input row is gone). "Blank" means no cells after parsing —
a row whose only cells are empty strings also disappears, because empty strings are written as
no cell.

Trailing empty cells are trimmed; interior ones are `null`. With `header: true` the same rules
apply, except that missing object properties are `null` rather than absent.

## Row counts and listSheets

```js
import { listSheets } from 'excel-lite';

console.log(await listSheets('book.xlsx'));
// [ { name: 'Data', index: 0, rowCount: 50001 }, { name: 'Notes', index: 1, rowCount: 3 } ]
```

- `rowCount` counts the `<row>` elements in the sheet part, **including the header row**.
  A sheet with a header and 3 data rows reports `4`.
- It does not parse cell values, but it does inflate the sheet XML to count, so on a large
  workbook `listSheets()` costs time and memory proportional to the largest sheet. It is not
  free metadata.
- A row element that exists but contains no cells still counts, so `rowCount` can be larger
  than the number of rows you get by iterating.
- `index` is the 0-based sheet position and is what `{ sheet: index }` expects.

## Errors thrown

Every message starts with `excel-lite:` (the CLI adds a second prefix, see [CLI](cli.md#exit-codes-and-errors)).

| When | Error | Message |
|---|---|---|
| source is not a zip | `Error` | `excel-lite: could not read the archive — the file is not a zip (.xlsx/.xlsm) container` |
| no `xl/workbook.xml` in the archive | `Error` | `excel-lite: xl/workbook.xml is missing — this is not an .xlsx file (legacy .xls binaries are not supported)` |
| workbook root element is not `<workbook>` | `Error` | `excel-lite: workbook.xml is malformed` |
| workbook lists no sheets | `Error` | `excel-lite: the workbook has no sheets` |
| sheet part missing from the archive | `Error` | `excel-lite: sheet "<name>" is missing from the archive (<part>)` |
| `sheet` index out of range | `RangeError` | `excel-lite: sheet index <n> is out of range (workbook has <n> sheet(s))` |
| `sheet` name not found | `Error` | `excel-lite: sheet "<name>" not found — available: <names>` |
| unsupported source value | `TypeError` | `excel-lite: unsupported source — pass a file path, Uint8Array, ArrayBuffer, Blob or async iterable of bytes` |
| path used without Node 22.3+ | `Error` | `excel-lite: reading or writing by file path needs Node 22+. In browsers and older runtimes pass bytes instead: readSheet(uint8Array) / toXlsx(sheets).` |
| a zip entry inflates past 512 MiB | `Error` | `excel-lite: <part> inflates to <n> MiB, above the 512 MiB safety limit — refusing to decompress (see SECURITY.md)` |
| file does not exist | `Error` | the platform's `ENOENT: no such file or directory, open '<path>'` |

See [troubleshooting](troubleshooting.md) for what to do about each.

## What the reader ignores

- **Formulas.** You get the cached value (`7`), never the formula text (`=SUM(A1:A3)`). A
  formula cell with no cached value reads as `null`. Formula text is on the [roadmap](../ROADMAP.md).
- **Formatting.** Bold, fill, alignment, widths, freeze state and merged ranges are not
  returned. Number formats are read only to decide "number or `Date`".
- **Merged cells.** Only the top-left cell holds the value; the other cells in the merge read
  as `null`/missing. Merges are not expanded, and there is no merge metadata.
- **Everything that does not affect cell values**: comments, hyperlinks, images, charts, pivot
  tables, conditional formatting, data validation, named ranges, macros (`vbaProject.bin`),
  print settings, hidden rows/columns, sheet tab colours.

The [compatibility matrix](compatibility.md) lists this as a table.

## Limits

- **512 MiB per zip entry.** Before inflating a needed part, the reader checks its declared
  uncompressed size and refuses anything larger, so a zip bomb cannot exhaust memory. It is a
  fixed limit in 0.1; see [`SECURITY.md`](../SECURITY.md).
- **Memory scales with the decompressed sheet XML**, not with the rows you consume. The model
  and measured numbers are in [performance](performance.md).
- **No random access.** There is no `getCell('B7')`; rows in, rows out. That is ExcelJS's model.
- **OOXML only.** `.xlsx` and `.xlsm` (macros ignored). Legacy `.xls` binaries are rejected with
  the message above.

## See also

- [Writing](writing.md) — the other half of the round trip.
- [API reference](api.md#readsheet) — exact signatures.
- [Troubleshooting](troubleshooting.md) — dates off by a day, memory, "not a zip".
