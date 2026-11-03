# Writing workbooks

Three functions write: `toXlsx()` builds bytes in memory, `writeSheet()` writes one sheet to a
path, `writeWorkbook()` writes several sheets to a path. All three share the same options.

## On this page

- [Minimal example](#minimal-example)
- [SheetInput](#sheetinput)
- [Rows: arrays and objects](#rows-arrays-and-objects)
- [Columns](#columns)
- [Header row](#header-row)
- [Dates](#dates)
- [Other cell values](#other-cell-values)
- [Cell styles](#cell-styles)
- [Freeze panes and auto filter](#freeze-panes-and-auto-filter)
- [Column widths](#column-widths)
- [Sheet names](#sheet-names)
- [Multiple sheets](#multiple-sheets)
- [toXlsx vs writeSheet vs writeWorkbook](#toxlsx-vs-writesheet-vs-writeworkbook)
- [Limits and what is not enforced](#limits-and-what-is-not-enforced)
- [Errors thrown](#errors-thrown)

## Minimal example

```js
import { writeSheet } from 'excel-lite';

await writeSheet('out.xlsx', [
  { name: 'Ada', score: 91.5 },
  { name: 'Linus', score: 77 },
], {
  columns: [
    { key: 'name', header: 'Name', width: 24 },
    { key: 'score', header: 'Score', format: '0.00' },
  ],
});
```

That writes a header row (because `columns` is given), two data rows, a bold header, computed
column widths and the number format. The function resolves with the `Uint8Array` it wrote.

## SheetInput

One object per worksheet, passed to `toXlsx(sheets)` or `writeWorkbook(path, sheets)`:

| Field | Type | Default | Meaning |
|---|---|---|---|
| `name` | `string` | `Sheet1`, `Sheet2`, … | tab name; sanitised (see [sheet names](#sheet-names)) |
| `rows` | `Iterable<RowInput>` | required | the data rows — array or object rows, see below |
| `columns` | `ColumnSpec[]` | inferred for object rows | column order, headers, widths and formats |
| `header` | `boolean` | `true` when `columns` is non-empty | write row 1 as the header row |
| `freeze` | `string \| { row?, col? }` | none | frozen pane, see [freeze panes](#freeze-panes-and-auto-filter) |
| `autoFilter` | `boolean` | `true` when a header is written | filter dropdowns on the header row |
| `autoWidth` | `boolean` | `true` | compute widths for columns without an explicit `width` |

`writeSheet(dest, rows, options)` takes the same fields minus `rows`, plus `name`:

```ts
writeSheet('report.xlsx', rows, { name: 'Report', columns, freeze: 'A2' })
```

## Rows: arrays and objects

Both forms work, and you can mix them in one sheet:

```js
// array rows: positional, no keys
await writeSheet('a.xlsx', [[1, 'one'], [2, 'two']], {
  columns: [{ key: 'id' }, { key: 'label' }],
});

// object rows: values picked by column key
await writeSheet('b.xlsx', [{ id: 1, label: 'one' }], {
  columns: [{ key: 'id' }, { key: 'label' }],
});
```

Object-row rules:

- With `columns`, each row is written as `columns.map(column => row[column.key])`. Keys not in
  `columns` are **dropped**, and column order is the `columns` order — not the object's.
- Without `columns`, the columns are inferred from the **first** object row's own keys
  (`Object.keys`, insertion order). Later rows that carry extra keys are still written only
  through the inferred columns, so extra keys disappear. Give explicit `columns` when rows are
  not uniform.
- `undefined` and `null` values produce no cell, exactly like a missing key.

Array-row rules:

- Values are positional. If a row is longer than `columns` (or there is no `columns`), the extra
  values are written in the following cells with no style.
- Short rows simply end early.
- `columns` still drives the header row, formats, widths and freeze/filter ranges.

## Columns

```js
const columns = [
  { key: 'email', header: 'Email', width: 32, bold: true, fill: '1F2937', align: 'left' },
  { key: 'amount', header: 'Amount', format: '#,##0.00' },
  { key: 'joined', header: 'Joined', format: 'yyyy-mm-dd' },
];
```

| Field | Type | Default | Meaning |
|---|---|---|---|
| `key` | `string` | required | property name on object rows; also the default header text |
| `header` | `string` | `key` | header text written to row 1 |
| `width` | `number` | computed | column width in Excel character units |
| `format` | `string` | `General` | any Excel number format code, e.g. `#,##0.00`, `0.0%`, `yyyy-mm-dd` |
| `bold` | `boolean` | `false` | bold **data** cells in this column |
| `fill` | `string` | none | solid fill colour, `RRGGBB` or `AARRGGBB`, leading `#` allowed |
| `align` | `'left' \| 'center' \| 'right'` | none | horizontal alignment of **data** cells |

Styles are per column, not per cell: every data cell in the column gets the same `bold`, `fill`,
`align` and `format`. There is no cell-level styling in 0.1 — that is on the
[roadmap](../ROADMAP.md).

## Header row

The header row is written when `header` is true **and** there is at least one column:
`header ?? columns.length > 0`. Consequences:

- `columns: [...]` and no `header` → a header is written.
- `rows` only, as objects → columns are inferred, so a header is written.
- `rows: [[1, 2]]` with no `columns` → **no header row**; `header: true` alone does nothing
  because there is nothing to write as a header.
- `header: false` with columns → data starts in row 1, no header, and no auto filter.

Header cells are always **bold**. A `ColumnSpec`'s `bold`, `fill`, `align` and `format` apply to
its data cells, not to the header cell. Header text is the `header` field, else `key`.

## Dates

A `Date` cell is written as an Excel serial number plus a number format, so it stays a real
date in Excel rather than text:

```js
await writeSheet('dates.xlsx', [{ joined: new Date(Date.UTC(2024, 4, 17)) }], {
  columns: [{ key: 'joined', format: 'yyyy-mm-dd' }],
});
```

Rules:

- The format is `column.format` when it looks like a date/time format. Otherwise the writer
  picks `yyyy-mm-dd hh:mm:ss` if the value has a non-zero UTC time part, and `yyyy-mm-dd` if
  it does not.
- The conversion uses **UTC**: `Date.UTC(2024, 4, 17)` becomes serial `45429` (the date Excel
  shows as 2024-05-17). A `Date` built from local parts —
  `new Date(2024, 4, 17)` in a `+05:00` timezone — is `2024-05-16T19:00:00Z`, so it is written as
  2024-05-16 19:00. Use `Date.UTC(...)` or an ISO date string (`new Date('2024-05-17')`) when
  you mean a calendar date.
- An invalid `Date` (`new Date('nope')`) is omitted, so the cell reads back as `null`.
- Reading returns `Date` objects in UTC, so a write → read round trip preserves the instant,
  including the time part down to the second.
- A **number** in a column with a date-looking `format` is written with that format too, because
  the format belongs to the column. Excel then displays the serial as a date (serial `45429` shows
  2024-05-17) and the reader returns a `Date`. If you want a number displayed as a number, do not
  give its column a date format.

Excel's serial quirks are mirrored by `dateToSerial()`: serials below 61 are shifted by one day
because of the phantom 1900-02-29, so 1900-01-01 writes as `1`. 1904-date-system workbooks are
read correctly but the writer always produces a 1900-date-system file.

## Other cell values

| Value | Written as | Notes |
|---|---|---|
| `string` | shared-string table entry | XML-escaped; characters forbidden by XML 1.0 (most control characters, U+FFFE, U+FFFF) are stripped; truncated to 32,767 characters |
| `number` | numeric cell | `NaN`, `Infinity`, `-Infinity` are written as **text**; write `null` instead if you mean "empty" |
| `boolean` | boolean cell (`t="b"`) | reads back as `true`/`false` |
| `Date` | numeric cell + date format | see [dates](#dates) |
| `null` / `undefined` | no cell | reads back as `null` |
| `''` (empty string) | no cell | also reads back as `null`; there is no way to write an empty-string cell |

String truncation and character stripping happen silently at the boundary — Excel refuses to
open a workbook whose XML contains characters it does not allow, so dropping them is the safe
behaviour. If you need the exact bytes back, validate before writing.

## Cell styles

The writer emits a minimal style sheet: two fonts (regular, bold), solid fills for the colours
you use, and one `xf` per distinct combination of (number format, bold, fill, align). Shared
style records mean 50,000 cells with the same format cost one style.

What you can write: number formats, bold, solid fill, horizontal alignment, column widths,
freeze panes, auto filter. That is the whole list — no borders, no vertical alignment, no text
wrap, no merged cells, no comments, no images, no conditional formatting. See
[compatibility](compatibility.md).

## Freeze panes and auto filter

```js
await writeSheet('a.xlsx', rows, {
  columns,
  freeze: 'A2',      // freeze row 1: the header stays visible
  autoFilter: true,  // filter dropdowns on the header row (default when a header is written)
});
```

`freeze` accepts:

| Value | Effect |
|---|---|
| `'A2'` | freeze row 1 (1 row, 0 columns) |
| `'B3'` | freeze the first column and the first two rows |
| `'B2'` | freeze the first column and row 1 |
| `'A1'` or `{ row: 0, col: 0 }` | no freeze |
| `{ row: 2 }` | freeze the first two rows |
| `{ col: 1 }` | freeze the first column |
| `{ row: 1, col: 1 }` | same as `'B2'` |

Only a single frozen pane (top/left) is supported; there are no split panes, and no zoom or
selection state.

`autoFilter` writes an `<autoFilter>` element whose range is **the header row only**
(`A1:D1` for four columns). Excel shows the dropdowns, and ExcelJS reads the ref back as
`A1:D1`. The writer does not extend the range over the data rows; if you need the filter range
to cover the data for a downstream tool, edit the ref after writing or open the file in Excel
and re-apply the filter.

## Column widths

- An explicit `width` wins.
- Otherwise, with `autoWidth` (the default), the width is
  `min(60, max(8, longestVisibleValue + 2))`, measured over the header text and all data cells
  in that column. `Date` cells count as 19 characters, booleans as 5, `null` as 0.
- `autoWidth: false` omits the width for columns without an explicit `width` (the column gets
  Excel's default width).
- Widths are Excel character units, not pixels.

## Sheet names

Names are sanitised rather than rejected, so a bad `name` never fails a write:

| Input | Written name |
|---|---|
| `'a/b*c?'` | `'a b c'` — `\ / ? * [ ] :` become spaces, whitespace collapses, trimmed |
| `' a b '` | `'a b'` |
| `''` or `'***'` | `'Sheet'` |
| `'x'.repeat(40)` | the first 31 characters — Excel's hard limit |
| two sheets both named `'Data'` | `'Data'`, `'Data (2)'` — duplicates are compared case-insensitively |
| `'DATA'` when `'data'` already exists | `'DATA (2)'` |

Numbers in the name are not special-cased; the name is used as text.

## Multiple sheets

```js
import { writeWorkbook } from 'excel-lite';

await writeWorkbook('book.xlsx', [
  {
    name: 'Report',
    rows: orders,
    columns: [{ key: 'id' }, { key: 'amount', format: '#,##0.00' }],
    freeze: 'A2',
  },
  { name: 'Raw', rows: [[1, 2, 3], [4, 5, 6]] },
]);
```

Sheets appear in array order and are named as given (or `Sheet1`, `Sheet2`, …). The first sheet
is the active tab. There is no per-sheet tab colour, visibility or protection.

## toXlsx vs writeSheet vs writeWorkbook

| Function | Signature | Input | Writes a file | Use it in |
|---|---|---|---|---|
| `toXlsx` | `(sheets: SheetInput[]) => Uint8Array` | sheets array | no | browsers, workers, HTTP responses, tests |
| `writeSheet` | `(dest: string, rows, options?) => Promise<Uint8Array>` | one sheet | yes | Node scripts, one-sheet exports |
| `writeWorkbook` | `(dest: string, sheets: SheetInput[]) => Promise<Uint8Array>` | sheets array | yes | Node scripts, multi-sheet exports |

`writeSheet` is exactly `writeWorkbook(dest, [{ ...options, rows }])`. The file is overwritten
if it exists; parent directories are not created (an `ENOENT` from the filesystem propagates).
Both file writers return the bytes as well, so you can write and upload without re-reading.

Writing is **in memory**: all rows are collected, then the sheet XML, shared strings, styles and
zip are built. There is no streaming writer, so peak memory grows with the number of rows you
write — see [performance](performance.md).

## Limits and what is not enforced

Excel's own limits are:

| Limit | Excel's value | Enforced by excel-lite? |
|---|---|---|
| rows per sheet | 1,048,576 | **no** |
| columns per sheet | 16,384 (A … XFD) | **no** |
| characters in a cell | 32,767 | **yes** — silently truncated |
| characters in a sheet name | 31 | **yes** — truncated |
| sheet name characters | no `\ / ? * [ ] :` | **yes** — replaced with spaces |

Because row and column limits are not checked, a sheet with 20,000 columns or 2,000,000 rows is
written as-is: the file is a valid zip but Excel will report a problem and may repair or refuse
it. Keep exports within Excel's limits yourself. There is no error for exceeding them.

## Errors thrown

| When | Error | Message |
|---|---|---|
| `toXlsx([])` or no sheets | `Error` | `excel-lite: toXlsx needs at least one sheet` |
| `freeze` string is not an A1-style cell | `Error` | `excel-lite: freeze must look like "A2", got "<value>"` |
| `fill` is not 6 or 8 hex digits | `Error` | `excel-lite: invalid colour "<value>" — use RRGGBB or AARRGGBB` |
| destination directory does not exist | `Error` | the platform's `ENOENT` error |
| destination path without Node 22.3+ | `Error` | `excel-lite: reading or writing by file path needs Node 22+…` |

Two readable edge cases: `freeze: 'A0'` and `freeze: 'A1'` are accepted but produce no pane,
because both splits are `<= 0`. `freeze: 'Sheet1!A2'` throws — it is not a cell reference.

## See also

- [Reading](reading.md) — what comes back out, and the UTC date rule.
- [API reference](api.md#toxlsx) — every signature and type.
- [Recipes](recipes.md) — styled reports, CSV bridges, browser downloads.
- [Compatibility](compatibility.md) — what Excel/LibreOffice will actually see.
