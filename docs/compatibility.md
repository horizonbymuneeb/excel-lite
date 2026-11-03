# Compatibility

What `excel-lite` writes, what it reads, what it silently ignores, and which other
implementations have actually been tested against it. If a feature is missing here, it is
missing from the code — nothing on this page is aspirational.

## On this page

- [File formats](#file-formats)
- [What is written](#what-is-written)
- [What is read](#what-is-read)
- [What is ignored](#what-is-ignored)
- [Styles: written vs read](#styles-written-vs-read)
- [Tested against other implementations](#tested-against-other-implementations)
- [Not tested](#not-tested)
- [Round-trip guarantees](#round-trip-guarantees)

## File formats

| Format | Read | Write | Notes |
|---|---|---|---|
| `.xlsx` (OOXML spreadsheet) | yes | yes | the format this package is for |
| `.xlsm` (macro-enabled) | yes | no | read as data; `vbaProject.bin` and other macro parts are ignored |
| `.xls` (legacy binary, BIFF) | no | no | rejected: "this is not an .xlsx file (legacy .xls binaries are not supported)" |
| `.csv`, `.ods`, `.numbers` | no | no | out of scope; CSV is a five-line conversion on either side ([recipes](recipes.md)) |

Writer output is always a plain `.xlsx` package with `[Content_Types].xml` declaring a
spreadsheet (not macro-enabled). Writing to a path named `*.xlsm` would produce a file with a
macro-enabled extension and no macros — untested, and not recommended.

## What is written

| Feature | Written? | Details |
|---|---|---|
| Text cells | yes | shared-string table; XML-escaped; XML-1.0-illegal characters stripped; truncated at 32,767 |
| Numbers | yes | finite numbers as numeric cells; `NaN`/`±Infinity` are written as **text** |
| Booleans | yes | `t="b"` |
| Dates | yes | serial + number format (`yyyy-mm-dd` or `yyyy-mm-dd hh:mm:ss`, or your `format`) |
| Empty cells (`null`, `undefined`, `''`) | yes | written as no cell at all |
| Error values (`#DIV/0!`) | as text | writing `'#DIV/0!'` produces a text cell, not an Excel error cell |
| Number formats | yes | any format code; custom codes are stored as `numFmtId` 164+ |
| Column widths | yes | explicit, or computed (`min(60, max(8, len+2))`) |
| Frozen panes | yes | one top/left pane only |
| Auto filter | yes | element written with the **header-row range only** (e.g. `A1:D1`) |
| Bold | yes | header row always; data cells per column |
| Solid fill | yes | `RRGGBB` / `AARRGGBB` |
| Horizontal alignment | yes | `left` / `center` / `right` |
| Multiple sheets | yes | order preserved, first sheet active |
| Sheet-name sanitisation | yes | 31 characters, forbidden characters replaced, duplicates renamed |
| Borders, vertical alignment, wrap text | **no** | would need per-cell styles |
| Merged cells | **no** | [roadmap](../ROADMAP.md) item |
| Row heights, hidden rows/columns, tab colours | **no** | not modelled |
| Formulas | **no** | values are written; there is no formula API |
| Comments, hyperlinks, images, charts, pivot tables | **no** | not modelled |
| Conditional formatting, data validation | **no** | not modelled |
| Named ranges, print settings, protection | **no** | not modelled |
| Macros | **no** | writer never produces a macro-enabled package |

## What is read

| Read | Result |
|---|---|
| Cell values | `string`, `number`, `boolean`, `Date` or `null` — the mapping table is in [reading](reading.md#value-mapping) |
| Shared strings, inline strings, rich-text runs | concatenated text |
| Percentages | the underlying fraction as a number (`12.5%` → `0.125`) |
| Error values | the error text as a `string` (`'#DIV/0!'`) |
| ISO date cells (`t="d"`) | `Date` |
| Number formats | used **only** to decide whether a number is a date; not returned |
| 1904 date-system workbooks | dates converted correctly |
| Sheet names, order, count | yes |
| Hidden sheets | read like visible ones (the `state` attribute is ignored) |
| Sparse rows | `null` in the holes |
| Blank rows | skipped |
| Formulas | the **cached value only** — never the formula text |
| Merged cells | the top-left cell's value only; the covered cells read as absent/`null` |
| Everything else | ignored |

There is no API to read a cell's format, style, comment, hyperlink, merge range, row height or
sheet visibility.

## What is ignored

These parts may exist in the file; the reader does not look at them and the writer does not
produce them:

comments and notes, hyperlinks, images and drawings, charts, pivot tables, conditional
formatting, data validation, defined names / named ranges, macros (`vbaProject.bin`), external
links, workbook and document properties, print areas and page setup, protection, tab colours,
sheet visibility, row/column grouping and outlines, autofilter state (only the element is
written, no filter values), cell comments, sparklines, themes beyond what Excel's defaults
provide.

If one of these matters to your workflow, `excel-lite` is the wrong tool for that step — but it
composes well: read values with `excel-lite`, then open the same file with ExcelJS or openpyxl
for the parts that need a full workbook model ([recipes](recipes.md#11-exceljs-interop)).

## Styles: written vs read

This asymmetry is the thing people trip over, so it is worth saying plainly:

- **Written**: bold, solid fill, horizontal alignment, column widths, freeze panes, auto filter,
  number formats.
- **Not read back**: everything above. A write → read round trip returns the **values** only.
  The reader looks at `xl/styles.xml` solely to answer "is this cell's number format a date?".

That means "read a styled template, fill it, write it back" is **not** supported in 0.1. The
[roadmap](../ROADMAP.md) has formatting read-back for 0.3.

## Tested against other implementations

`excel-lite` is not self-certifying. The automated test suite cross-checks against these, with
the versions pinned in `package.json`/`package-lock.json`:

| Implementation | Version used | Direction | Where |
|---|---|---|---|
| **ExcelJS** | 4.4.0 | reads ExcelJS-written files; ExcelJS reads ours (values, dates, number formats) | `test/roundtrip.test.ts` |
| **SheetJS (`xlsx`)** | 0.18.5 | reads SheetJS-written files; SheetJS reads ours (raw values) | `test/roundtrip.test.ts` |
| **openpyxl** | 3.1.5 | reads openpyxl-generated fixtures (values, dates, formats, headers, sparse cells, 1904 dates, unicode) | `fixtures/`, `test/fixtures.test.ts`, `scripts/make-fixtures.py` |

For the other direction — openpyxl reading what `excel-lite` writes — there is an opt-in script
rather than an automated test (it needs Python and openpyxl installed):

```bash
node -e "import('excel-lite').then(({ writeSheet }) => writeSheet('/tmp/sample.xlsx', [{ name: 'Muneeb', amount: 1234.5, when: new Date(Date.UTC(2024, 4, 17)) }], { columns: [{ key: 'name' }, { key: 'amount', format: '#,##0.00' }, { key: 'when', format: 'yyyy-mm-dd' }] }))"
python3 scripts/verify-with-openpyxl.py /tmp/sample.xlsx
```

CI regenerates the openpyxl fixtures with `python3 scripts/make-fixtures.py` and fails if the
committed fixtures differ, so the fixture files in the repo really are openpyxl output and not
edited by hand.

## Not tested

Be clear about the boundary — these were **not** verified by this project:

| Implementation | Status |
|---|---|
| Microsoft Excel (desktop, Windows/macOS) | **untested** — no automated check opens our files in Excel |
| Microsoft Excel Online | **untested** |
| LibreOffice Calc | **untested** |
| Google Sheets import | **untested** |
| Apple Numbers | **untested** |
| WPS Office | **untested** |

The output is standard OOXML and every tested reader accepts it, but "Excel opens it without a
repair prompt" is a claim nobody has automated here. If you hit a file Excel complains about,
that is a bug worth reporting with the file ([SUPPORT.md](../SUPPORT.md)) — and the most likely
causes are listed in [troubleshooting](troubleshooting.md#excel-or-libreoffice-complains-about-the-file).

## Round-trip guarantees

What survives `read → write → read` exactly:

- strings, including unicode, emoji, quotes, newlines and commas (XML-illegal characters do not
  survive the write, by design);
- numbers (including negative and very small/large finite values);
- booleans;
- dates, as UTC instants, to the second;
- `null` stays `null`;
- sheet names (after sanitisation) and sheet order;
- column order and header text.

What does not survive:

- number formatting, bold/fill/alignment, widths, freeze state, filter state — not read back;
- formulas (cached value only, and the formula is not written back);
- merged cells, comments, images, charts and every other ignored part;
- strings longer than 32,767 characters (truncated);
- XML-illegal control characters (stripped).

## See also

- [Reading](reading.md) — the read side in detail.
- [Writing](writing.md) — the write side in detail.
- [Migration](migration.md) — if you expected SheetJS/ExcelJS features here.
- [Troubleshooting](troubleshooting.md) — when a file will not open.
