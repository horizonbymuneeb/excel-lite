# excel-lite documentation

`excel-lite` reads and writes `.xlsx` workbooks. It is small on purpose: rows in, rows out,
one runtime dependency (`fflate`), no native builds, dual ESM/CJS, TypeScript types included.
It does not try to be a spreadsheet engine — see [what it deliberately does not do](../README.md#what-it-deliberately-does-not-do).

- **Node 22+** for anything that takes a file path.
- **Browsers, workers, Bun, Deno** for everything else: `toXlsx()` returns a `Uint8Array` and
  `readSheet()` accepts bytes or a `Blob`.
- **Dates are UTC `Date` objects**, numbers are numbers, booleans are booleans, empty cells are `null`.

## Quick start

```bash
npm install excel-lite
```

```js
import { readSheet, writeSheet } from 'excel-lite';

for await (const row of readSheet('report.xlsx', { header: true })) {
  console.log(row.email, row.amount); // dates are Dates, numbers are numbers
}

await writeSheet('report.xlsx', rows, {
  columns: [{ key: 'email', width: 32 }, { key: 'amount', format: '#,##0.00' }],
  freeze: 'A2',
});
```

```bash
npx excel-lite info report.xlsx
npx excel-lite to-json report.xlsx --header > rows.ndjson
```

## The manual

| Page | What it covers | Read it when |
|---|---|---|
| [Getting started](getting-started.md) | install, first read, first write, first CLI run, verifying the install | you have never used the package |
| [Reading](reading.md) | every reader behaviour: sources, sheets, header mode, value mapping, dates, sparse rows, row counts, errors | you are importing data |
| [Writing](writing.md) | every writer option: rows, columns, formats, dates, widths, freeze, filters, sheet names, limits | you are exporting data |
| [API reference](api.md) | every export, signature, parameter, return value, thrown error, example | you know what you want and need the exact shape |
| [CLI](cli.md) | `info`, `to-json`, `to-csv`, NDJSON vs `--pretty`, `--out`, `--delimiter`, exit codes, piping | you want rows out of a file without writing code |
| [Recipes](recipes.md) | uploads, fetch, CSV bridges, Postgres COPY, S3/R2, zod validation, ExcelJS and openpyxl interop | you want working code for a real task |
| [Performance](performance.md) | benchmark method, measured numbers, the memory model, when to use something else | files are large or memory is tight |
| [Compatibility](compatibility.md) | what is written, read, ignored, and which readers are actually tested | you need to know what survives a round trip |
| [Migration](migration.md) | moving from `xlsx` (SheetJS) or ExcelJS, with before/after code | you already have code to port |
| [Troubleshooting](troubleshooting.md) | 17 concrete failures and the exact fix | something threw or looks wrong |
| [Architecture](architecture.md) | module map, read flow, write flow, why the parser is regex-driven, where to change what | you are contributing |
| [Browser](browser.md) | bundlers, why there is no `node:fs`, Vite/React upload and download, bundle size | you are shipping to the browser |

Working code lives in [`examples/`](../examples/README.md) — every script there runs with
`node examples/<name>.mjs` after `npm install && npm run build` in the repository.

## Conventions used in these docs

- Snippets are ESM. CommonJS works too: `const { readSheet } = require('excel-lite')`.
- Every error message starts with `excel-lite:` — exact strings are listed in
  [troubleshooting](troubleshooting.md) and the [API reference](api.md).
- "Cell" means one value from one row; "sheet" means a worksheet tab; "workbook" means the `.xlsx` file.
- Dates are always `Date` objects at UTC midnight for date-only cells. The timezone
  consequences are spelled out in [reading](reading.md#dates-times-and-timezones).
- A feature is only documented here if the source implements it. If the code does not
  support something, it is named in [compatibility](compatibility.md) as unsupported.

## Project documents

- [`README.md`](../README.md) — the one-page pitch and benchmark table
- [`CHANGELOG.md`](../CHANGELOG.md) — what changed between releases
- [`ROADMAP.md`](../ROADMAP.md) — what is planned for 0.2 and later
- [`CONTRIBUTING.md`](../CONTRIBUTING.md) — development setup and the rules for tests and docs
- [`SECURITY.md`](../SECURITY.md) — threat model, untrusted files, the 512 MiB decompression limit
- [`SUPPORT.md`](../SUPPORT.md) — where to ask for help
