# Getting started

`excel-lite` turns a `.xlsx` file into rows and rows back into a `.xlsx` file. This page gets
you from zero to a first read, a first write and a first CLI command.

## On this page

- [Requirements](#requirements)
- [Install](#install)
- [First read](#first-read)
- [First write](#first-write)
- [First CLI command](#first-cli-command)
- [Verifying the install](#verifying-the-install)
- [Common first-run problems](#common-first-run-problems)
- [Next steps](#next-steps)

## Requirements

| Runtime | What works | Notes |
|---|---|---|
| Node 22.3+ | everything, including file paths | path access goes through `process.getBuiltinModule`, added in Node 22.3 |
| Node 22.0–22.2 | everything that takes bytes | path calls throw a clear error telling you to pass bytes |
| Browsers, web workers | bytes only | `toXlsx()` returns a `Uint8Array`; `readSheet()` takes bytes or a `Blob` |
| Bun, Deno, edge runtimes | bytes only | the package has no `node:fs` import in its module graph |

`package.json` declares `"engines": { "node": ">=22" }`. CI runs Node 22 and 24. If you need
to support older Node, pass bytes instead of paths — see [browser](browser.md).

## Install

```bash
npm install excel-lite
```

```js
// ESM
import { readSheet, writeSheet, toXlsx } from 'excel-lite';

// CommonJS
const { readSheet, writeSheet, toXlsx } = require('excel-lite');
```

TypeScript declarations ship in the package (`dist/index.d.ts`), so no `@types` package is
needed. No native build, no postinstall step, one runtime dependency (`fflate`).

## First read

You need a file. Either use one you already have, or create one in the next section. Then:

```js
import { readSheet } from 'excel-lite';

const rows = [];
for await (const row of readSheet('report.xlsx')) {
  rows.push(row); // each row is a CellValue[]: string | number | boolean | Date | null
}

console.log(rows[0]); // the header row, if the file has one
```

`readSheet()` returns an async generator, so rows are parsed as you iterate and you can stop
early. With `header: true` the first non-empty row becomes the object keys:

```js
for await (const row of readSheet('report.xlsx', { header: true })) {
  console.log(row.email, row.amount, row.joined);
}
```

Prefer objects? This reads every row into memory first — fine for a report, not for a
million-row import:

```js
const all = [];
for await (const row of readSheet('report.xlsx', { header: true })) all.push(row);
```

## First write

```js
import { writeSheet } from 'excel-lite';

const rows = [
  { email: 'ada@example.com', amount: 1234.5, joined: new Date(Date.UTC(2024, 4, 17)) },
  { email: 'linus@example.com', amount: 99, joined: new Date(Date.UTC(2024, 5, 1)) },
];

await writeSheet('report.xlsx', rows, {
  columns: [
    { key: 'email', header: 'Email', width: 32 },
    { key: 'amount', header: 'Amount', format: '#,##0.00' },
    { key: 'joined', header: 'Joined', format: 'yyyy-mm-dd' },
  ],
  freeze: 'A2', // freeze the header row
});
```

`writeSheet()` writes the file and returns the same bytes, so the same call works in a browser
where you ignore the path argument — see `toXlsx()` in [writing](writing.md#toxlsx-vs-writesheet-vs-writeworkbook).

Both functions return promises; always `await` them.

## First CLI command

No install needed to look at a file:

```bash
npx excel-lite info report.xlsx
```

```
report.xlsx
   0  Sheet1  3 rows
```

```bash
npx excel-lite to-json report.xlsx --header          # NDJSON, one row per line
npx excel-lite to-csv  report.xlsx --out report.csv  # CSV with proper quoting
```

`to-json` streams one JSON object per line (pipe it into `jq`, a database loader, or another
process); `--pretty` emits a single JSON array instead. All flags are in [CLI](cli.md).

## Verifying the install

Run this script — it writes a workbook in memory, reads the bytes back, and prints one row.
It needs no file system, so it also proves the package loaded correctly in your runtime:

```js
import assert from 'node:assert/strict';
import { readSheet, toXlsx } from 'excel-lite';

const bytes = toXlsx([{ rows: [{ ok: true, when: new Date(Date.UTC(2024, 0, 1)) }] }]);
for await (const row of readSheet(bytes, { header: true })) {
  console.log(row); // { ok: true, when: 2024-01-01T00:00:00.000Z }
  assert.equal(row.ok, true);
}
console.log('excel-lite works');
```

Expected output:

```
{ ok: true, when: 2024-01-01T00:00:00.000Z }
excel-lite works
```

Also useful:

```bash
npx excel-lite --help     # usage and flags, exit code 0
npm ls excel-lite         # the installed version
```

## Common first-run problems

| Symptom | Cause | Fix |
|---|---|---|
| `Cannot find module 'excel-lite'` | not installed, or wrong Node context | `npm install excel-lite`; check you are in the project that has the dependency |
| `ENOENT: no such file or directory` | path is relative to the process working directory | resolve the path (`path.join(import.meta.dirname, ...)`) or run from the right directory |
| `excel-lite: could not read the archive …` | the file is not a `.xlsx`/`.xlsm` zip — often a `.csv` renamed, or a legacy `.xls` | export a real `.xlsx`; `.xls` is not supported ([compatibility](compatibility.md)) |
| `excel-lite: reading or writing by file path needs Node 22+` | Node older than 22.3, or a browser/edge runtime | pass bytes instead of a path |
| `SyntaxError: Cannot use import statement outside a module` | ESM syntax in a CommonJS project | use `require('excel-lite')` or rename to `.mjs` |
| dates print one day early | you formatted a UTC `Date` with local-time getters | use `getUTC*` / `toISOString()` ([reading](reading.md#dates-times-and-timezones)) |

More failures, with exact messages: [troubleshooting](troubleshooting.md).

## Next steps

- [Reading](reading.md) — every value type, header rules, sparse rows, sheet selection.
- [Writing](writing.md) — columns, number formats, styles, freeze panes, filters.
- [Recipes](recipes.md) — uploads, CSV bridges, Postgres COPY, S3/R2, zod, ExcelJS interop.
- [`examples/`](../examples/README.md) — runnable scripts for each of the above.
