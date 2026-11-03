# Examples

Runnable Node ESM scripts for `excel-lite`. Each one is a single file with a header comment saying
what it shows and how to run it, and each one asserts its own result — a non-zero exit code means
something is wrong.

## Run them

From the repository root:

```bash
npm install          # once
npm run build        # once — the examples import the built package
node examples/basic-read.mjs
```

They write to a temp directory and delete it afterwards (the styled report keeps its output only
when you pass a path). `exceljs-interop.mjs` needs `exceljs`, which is already a devDependency of
this repo; in your own project `npm i exceljs` if you want that example.

| Script | What it shows | Run |
|---|---|---|
| [`basic-read.mjs`](basic-read.mjs) | read a workbook as arrays and as objects, `listSheets`, missing cells as `null`, UTC dates | `node examples/basic-read.mjs` |
| [`basic-write.mjs`](basic-write.mjs) | `writeSheet` with columns/formats/freeze, `writeWorkbook` with several sheets, `toXlsx` bytes | `node examples/basic-write.mjs` |
| [`streaming-large-file.mjs`](streaming-large-file.mjs) | 100k-row write and read, early exit, read-only peak RSS in a child process, throughput | `node examples/streaming-large-file.mjs [rows]` |
| [`styled-report.mjs`](styled-report.mjs) | bold/filled headers, alignment, widths, freeze, auto filter, number and date formats | `node examples/styled-report.mjs [out.xlsx]` |
| [`csv-bridge.mjs`](csv-bridge.mjs) | CSV → `.xlsx` and `.xlsx` → CSV/NDJSON with a small dependency-free CSV parser | `node examples/csv-bridge.mjs` |
| [`cli-pipeline.mjs`](cli-pipeline.mjs) | driving the CLI from Node: `info`, NDJSON, `--pretty`, `--delimiter`, `--out`, EPIPE | `node examples/cli-pipeline.mjs` |
| [`browser-usage.mjs`](browser-usage.mjs) | bytes only: `Uint8Array`, `ArrayBuffer`, `Blob`, chunk iterable — no file paths | `node examples/browser-usage.mjs` |
| [`exceljs-interop.mjs`](exceljs-interop.mjs) | excel-lite → ExcelJS (values and styles) and ExcelJS → excel-lite (values) | `node examples/exceljs-interop.mjs` |

## What each one prints

`basic-read.mjs` (excerpt):

```
--- header: true
Ali 3 19.99 2023-07-04T00:00:00.000Z true
Sara null 0.5 null false
...
ok: 3 rows, missing cells are null, dates are UTC
```

`streaming-large-file.mjs` (numbers vary by machine):

```
wrote 100,000 rows in 3000 ms — 3.5 MB on disk
sheet info: [ { name: 'Sheet1', index: 0, rowCount: 100001 } ]
read 100,000 rows in 1522 ms (qty total 24950000)
first 5 ids in 619 ms: 1, 2, 3, 4, 5
read-only peak RSS: 229.0 MB (100,001 rows) for a 3.5 MB file
```

The last line is the important one: memory tracks the decompressed sheet XML, not the number of
rows you consumed. See [performance](../docs/performance.md#the-memory-model).

`cli-pipeline.mjs` ends with:

```
first line through `| head -1`: {"id":1,"note":"row 1"}
exit code — cli: 0

ok: info, NDJSON, --pretty, --delimiter, --out and | head -1 (EPIPE) all behave
```

## Shell equivalents

`cli-pipeline.mjs` mirrors these commands; here they are on their own:

```bash
node dist/cli.js info report.xlsx
node dist/cli.js to-json report.xlsx --header | jq -c '.email'
node dist/cli.js to-json big.xlsx | head -100          # EPIPE handled: exit 0
node dist/cli.js to-csv report.xlsx --header --out report.csv
node dist/cli.js to-json report.xlsx --sheet Notes --pretty
```

After a global install (`npm i -g excel-lite`) or in a consumer project, `npx excel-lite` replaces
`node dist/cli.js` ([CLI](../docs/cli.md)).

## Notes

- Examples import `excel-lite` by package name. Inside this repository Node resolves that through
  the package's own `exports` map, which is why `npm run build` is needed first (the map points at
  `dist/`).
- `examples/` is not part of the published npm package (`package.json` `files` lists `dist`,
  `docs`, `README.md`, `CHANGELOG.md`, `LICENSE`). They are repository documentation you can run.
- Each script cleans up after itself; nothing is written into the repository.
- Every example uses the public API only. If one of them breaks, that is a bug — please report it.

## See also

- [Getting started](../docs/getting-started.md) — the smallest possible read and write.
- [Recipes](../docs/recipes.md) — the same patterns with Express, Next.js, Postgres, S3 and zod.
- [API reference](../docs/api.md) — signatures used by these scripts.
