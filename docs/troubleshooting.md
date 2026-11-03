# Troubleshooting

Concrete failures with the exact message you see and the fix. If your problem is not here, the
[reading](reading.md#errors-thrown) and [writing](writing.md#errors-thrown) pages list every error
the code can throw, and [SUPPORT.md](../SUPPORT.md) says where to ask.

## On this page

- [`could not read the archive` / not a zip](#could-not-read-the-archive-not-a-zip)
- [`xl/workbook.xml is missing`](#xlworkbookxml-is-missing)
- [`workbook.xml is malformed` or `the workbook has no sheets`](#workbookxml-is-malformed-or-the-workbook-has-no-sheets)
- [Sheet not found](#sheet-not-found)
- [Dates shifted by a day](#dates-shifted-by-a-day)
- [Times come back as 1899 dates](#times-come-back-as-1899-dates)
- [Empty cells: null vs undefined](#empty-cells-null-vs-undefined)
- [Memory blowup on a large file](#memory-blowup-on-a-large-file)
- [Excel or LibreOffice complains about the file](#excel-or-libreoffice-complains-about-the-file)
- [Bundler errors about node:fs](#bundler-errors-about-nodefs)
- [CJS/ESM interop](#cjs-and-esm-interop)
- [EPIPE, ECONNRESET or SIGPIPE](#epipe-econnreset-or-sigpipe)
- [`excel-lite: not found` (npm bin)](#excel-lite-not-found-npm-bin)
- [Node version is too old for paths](#node-version-is-too-old-for-paths)
- [TypeScript rejects a Blob or a Workbook type](#typescript-rejects-a-blob-or-a-workbook-type)
- [Header keys are wrong](#header-keys-are-wrong)
- [`inflates to … above the 512 MiB safety limit`](#an-entry-inflates-above-the-512-mib-safety-limit)

## Could not read the archive / not a zip

```
excel-lite: could not read the archive — the file is not a zip (.xlsx/.xlsm) container
```

The bytes are not a zip container. In practice:

- an HTML error page saved as `.xlsx` (a failed download — check the response status);
- a CSV or TSV renamed to `.xlsx`;
- a legacy `.xls` binary (BIFF), which is not a zip either;
- a truncated download.

Fix: open the file with `unzip -l file.xlsx` (or `file file.xlsx`). If it is CSV, convert it to
real `.xlsx` ([recipes](recipes.md#4-csv-to-xlsx)). If it is `.xls`, re-export it as `.xlsx` first;
`excel-lite` does not read BIFF ([compatibility](compatibility.md#file-formats)).

## xl/workbook.xml is missing

```
excel-lite: xl/workbook.xml is missing — this is not an .xlsx file (legacy .xls binaries are not supported)
```

The zip is valid but is not a spreadsheet: a `.docx`, `.pptx`, `.zip` or another OOXML package.
Check the extension and the source of the upload. For a `.xlsx` that genuinely lacks the part, the
file is corrupt — regenerate it from the producer.

## workbook.xml is malformed or the workbook has no sheets

```
excel-lite: workbook.xml is malformed
excel-lite: the workbook has no sheets
```

The zip has `xl/workbook.xml` but its content is not usable: wrong root element, truncated XML, or
an empty `<sheets>` element. This is a corrupt or non-standard producer, not a supported input
variant. Re-save the file in Excel/LibreOffice/openpyxl and try again.

## Sheet not found

```
excel-lite: sheet "Nope" not found — available: Data, Notes
```

or, for an index:

```
RangeError: excel-lite: sheet index 5 is out of range (workbook has 2 sheet(s))
```

Fix: use one of the names in the message. Remember:

- `index` is **0-based** — the second sheet is `{ sheet: 1 }`;
- names match exactly first, then case-insensitively, but are not trimmed or fuzzy-matched;
- hidden sheets are included, so a name can be "invisible" in Excel but valid here;
- `listSheets(file)` or `readWorkbook(file).sheetNames` prints the real names.

## Dates shifted by a day

Symptom: `row.joined` prints as `2024-05-16` for a cell that says `2024-05-17`, or
`toLocaleDateString()` shows the previous day.

Cause: dates are **UTC** `Date` objects. `getDate()`, `getMonth()`, `toLocaleDateString()` and
`Intl.DateTimeFormat` without a timezone use the machine's local timezone; in a timezone behind
UTC, UTC midnight is the previous evening.

```js
row.joined.toISOString();               // '2024-05-17T00:00:00.000Z'  ← correct date
row.joined.getUTCDate();                // 17
row.joined.getDate();                   // 16 in America/New_York
row.joined.toLocaleDateString('en-CA', { timeZone: 'UTC' }); // '2024-05-17'
```

Fix: format with UTC getters, `toISOString().slice(0, 10)`, or pass `timeZone: 'UTC'` to
`Intl`. Do not "correct" the `Date` — it is right.

The same trap exists on the write side: `new Date(2024, 4, 17)` is **local** midnight, which is
the previous UTC day in a positive-offset timezone, so it is written as the previous date. Use
`new Date(Date.UTC(2024, 4, 17))` or `new Date('2024-05-17')`.

## Times come back as 1899 dates

Symptom: a time-only cell (`06:30`) reads as `1899-12-30T06:30:00.000Z`.

Cause: Excel has no time type. A time is a fraction of a day, stored as a serial number below `1`,
and the reader converts serials to `Date`s from the 1899-12-30 epoch. That is the format's
behaviour, not a parsing error.

```js
const time = row.clock; // Date at 1899-12-30T06:30:00.000Z
const hhmm = `${String(time.getUTCHours()).padStart(2, '0')}:${String(time.getUTCMinutes()).padStart(2, '0')}`;
```

Fix: read the UTC time parts (as above) and ignore the date. Do not add a base date to the `Date`
object; build a string or a new `Date` with the date you actually want.

## Empty cells: null vs undefined

Symptom: `row.email` is `null`, or a property "does not exist" when you expect `undefined`.

The reader never returns `undefined`:

- with `header: false`, missing cells are `null` and trailing `null`s are trimmed, so the row is
  shorter;
- with `header: true`, **every** header key exists on every object, with `null` for a missing cell;
- cells beyond the header row's width appear as `column_<n>`.

Fix: test with `row.email == null` (covers both), or `row.email ?? ''`. Do not use
`Object.hasOwn(row, 'email')` to detect an empty cell — it will be present with `null`. Excel
cannot distinguish an empty cell from a cell that never existed, so neither can the reader.

Also note: `''` (empty string) is written as *no cell* and therefore reads back as `null`; there is
no empty-string cell in this writer.

## Memory blowup on a large file

Symptom: reading a 200 MB `.xlsx` kills the process, or a 50 MB upload takes gigabytes of RSS.

Cause: the sheet XML is decompressed into memory as a whole (and then decoded to a UTF-16 string),
so peak RSS is several times the compressed file size — see
[performance](performance.md#the-memory-model). The 512 MiB per-entry guard stops zip bombs, not
legitimate large sheets.

Fix, in order of preference:

1. Cap upload size at the edge and reject oversized files early.
2. Pre-convert huge exports to CSV at the source and read the CSV.
3. Split the workbook into smaller files per sheet or per date range.
4. If you must parse it in-process, raise the process heap (`node --max-old-space-size=4096`) and
   accept the RSS — or wait for the streaming reader on the [roadmap](../ROADMAP.md).

Also: `listSheets()` inflates every sheet to count rows; `readWorkbook().sheetNames` does not.

## Excel or LibreOffice complains about the file

Typical messages: "We found a problem with some content…", "file is corrupt", a repair prompt, or
the file opens empty.

Causes and fixes, most likely first:

1. **Too many rows or columns.** The writer does not enforce Excel's limits (1,048,576 rows,
   16,384 columns). A 20,000-column sheet is a valid zip and an invalid Excel sheet — trim the
   data. See [writing](writing.md#limits-and-what-is-not-enforced).
2. **A non-finite number.** `NaN`/`Infinity` are written as text, which is legal but surprising;
   prefer `null`.
3. **A Date in a text column, or a number in a date-formatted column.** Both are legal, but the
   display may surprise you: a number in a `yyyy-mm-dd` column is shown as that date, because the
   format is applied to the column. Remove the format if the value is meant to be a plain number.
4. **Two sheets with names Excel dislikes.** Sanitisation handles `\ / ? * [ ] :`, 31 characters
   and duplicates, but check the resulting `listSheets()` names anyway.
5. **A truncated or partially written file** — write to a temp path and rename, and check the
   write completed before handing the file out.
6. **A file written by something else** — confirm with `unzip -l file.xlsx` that the zip has
   `[Content_Types].xml`, `xl/workbook.xml` and `xl/worksheets/sheet1.xml`.

Files produced by the tested paths open in ExcelJS, SheetJS and openpyxl; desktop Excel itself is
[not tested](compatibility.md#not-tested), so a reproducible Excel complaint is worth an issue
with the file attached.

## Bundler errors about node:fs

Symptom: Vite/webpack/esbuild fails with `Could not resolve "node:fs"` (or `fs`, `path`) when
bundling anything that imports `excel-lite`.

Cause: `excel-lite` has **no** `node:fs` import — it reaches `fs/promises` through
`process.getBuiltinModule()` at runtime (`src/fs.ts`), so the browser bundle contains zero Node
builtins. The error usually comes from:

- a different dependency in the same import graph (an S3/SQL client, `dotenv`, a CLI helper);
- an old copy of `excel-lite` (pre-0.1 builds had a static import);
- a bundler configured with `target: 'node'` for browser code.

Fix: find the actual importer with a bundle analysis (`vite build --debug`, `webpack --stats
detailed`), or exclude it from the client bundle. Verify `excel-lite` alone bundles clean:

```bash
npx esbuild --bundle --platform=browser --format=esm node_modules/excel-lite/dist/index.js --outfile=/tmp/check.js
grep -c 'node:' /tmp/check.js   # 0
```

See [browser](browser.md) for bundler configuration.

## CJS and ESM interop

Symptom: `ERR_REQUIRE_ESM`, `Cannot find module`, or `require` returns `{ default: … }` in a
CommonJS project.

The package ships both builds and an `exports` map:

```js
// ESM
import { readSheet } from 'excel-lite';
// CommonJS
const { readSheet } = require('excel-lite'); // dist/index.cjs — named exports, no .default
```

Fix:

- in a `"type": "module"` project, use `import`; a `.cjs` file must use `require`;
- from TypeScript, use `"module": "node16"`/`"nodenext"` or `"bundler"` module resolution so the
  `exports` map is honoured;
- do not write `require('excel-lite').default` — the CJS build has named exports;
- in a CJS file that needs the ESM build (rare), `const { readSheet } = await import('excel-lite')`.

## EPIPE, ECONNRESET or SIGPIPE

Symptom: your Node script that pipes `excel-lite` rows somewhere crashes with
`Error: write EPIPE` when the consumer exits (`| head`, a closed HTTP response, `less … q`).

The CLI handles this itself (exit 0, no stack trace — see [CLI](cli.md#exit-codes-and-errors)).
Your own script must do the same:

```js
process.stdout.on('error', (error) => {
  if (error.code === 'EPIPE') process.exit(0); // reader went away: success, not failure
  throw error;
});
```

For HTTP responses, check `res.writableEnded`/`destroyed` before writing, or listen for
`'close'` on the response and stop the loop. The same applies to `ECONNRESET` on a socket:
the client disconnected; stop writing and clean up.

## excel-lite not found (npm bin)

Symptom: `excel-lite: command not found` after `npm install excel-lite`.

Cause: dependency binaries are not on the global `PATH`; only `npx` or an npm script sees them.

Fix:

```bash
npx excel-lite info report.xlsx            # resolves the local install
npm exec -- excel-lite info report.xlsx    # same
npx --yes excel-lite info report.xlsx      # install-if-missing (prefix --yes in CI)
```

or add a script to `package.json` (`"cli": "excel-lite"`, run with `npm run cli -- info file.xlsx`).
For a global command: `npm i -g excel-lite`. If you use `nvm`/`volta`, reinstall after switching
Node versions so the bin shim points at the right runtime. The binary needs Node 22.3+ (next
entry).

## Node version is too old for paths

```
excel-lite: reading or writing by file path needs Node 22+. In browsers and older runtimes pass bytes instead: readSheet(uint8Array) / toXlsx(sheets).
```

Path access uses `process.getBuiltinModule`, which exists from Node **22.3.0** (`process.getBuiltinModule`
also landed in 20.16.0, but this package's `engines` floor is 22). On Node 22.0–22.2, in Bun, Deno,
browsers or edge runtimes, use bytes:

```js
const bytes = new Uint8Array(await file.arrayBuffer()); // or fs.readFile
for await (const row of readSheet(bytes)) { /* … */ }
const out = toXlsx([{ rows }]);                          // keep or upload the Uint8Array
```

## TypeScript rejects a Blob or a Workbook type

Two known gaps in 0.1.0's types (both work at runtime):

```ts
// Blob works, but Source does not include it yet
import type { Source } from 'excel-lite';
readSheet(blob as unknown as Source);
// or avoid the cast:
readSheet(new Uint8Array(await blob.arrayBuffer()));

// Workbook is not re-exported from the entry point
import { readWorkbook } from 'excel-lite';
type Workbook = Awaited<ReturnType<typeof readWorkbook>>;
```

Both are listed in [API reference](api.md#known-type-gaps) and should disappear in 0.2.

## Header keys are wrong

Three distinct cases:

1. **Wrong keys** — `header: true` uses the first *non-empty* row. If the file has a title row
   (`'Q3 Report'`) above the real header, that title becomes the only key. Read `header: false` and
   skip rows yourself, or fix the file upstream.
2. **`column_2` keys** — the header cell is blank or absent. A gap in the header row becomes
   `column_<1-based position>`.
3. **`name_2` keys** — duplicate header text; each repeat gets `_2`, `_3`. The suffix is chosen
   against every key already used, so a header literally titled `name_2` becomes `name_2_2` rather
   than overwriting anything. Rename headers if the suffixes make the keys awkward.

Also avoid non-text headers: a date or number header is stringified, and `String(Date)` embeds the
local timezone, so keys differ between machines.

## An entry inflates above the 512 MiB safety limit

```
excel-lite: xl/worksheets/sheet1.xml inflates to 812 MiB, above the 512 MiB safety limit — refusing to decompress (see SECURITY.md)
```

The zip entry is larger than the fixed 512 MiB guard, which exists so a zip bomb cannot exhaust
memory. It is a hard limit in 0.1; a configurable `maxEntryBytes` option is on the
[roadmap](../ROADMAP.md).

Fix: split the sheet, or convert it to CSV before importing. Do not raise the limit by patching the
package unless you control the input.

## See also

- [Reading](reading.md) — behaviour and error list.
- [Writing](writing.md) — options and error list.
- [Performance](performance.md) — memory numbers behind the large-file advice.
- [Compatibility](compatibility.md) — what is tested and what is not.
