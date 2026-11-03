# Performance and memory

Two things decide whether `excel-lite` fits a job: how fast it moves rows, and how much memory a
file costs. This page gives the benchmark method, the published numbers, a reproduction on a
second machine, and the memory model — including where the "≈8× the file size" rule of thumb
comes from and when it understates the cost.

## On this page

- [What the benchmark does](#what-the-benchmark-does)
- [Published numbers](#published-numbers)
- [Reproduction](#reproduction)
- [The memory model](#the-memory-model)
- [Measured memory scenarios](#measured-memory-scenarios)
- [The write path](#the-write-path)
- [When to use something else](#when-to-use-something-else)
- [Practical guidance](#practical-guidance)

## What the benchmark does

`npm run bench` (`scripts/bench.mjs`) compares `excel-lite`, ExcelJS and SheetJS on the same
workload:

- **Data**: 50,000 rows × 6 columns (`id`, `name`, `qty`, `price` with `#,##0.00`, `date` with
  `yyyy-mm-dd`, `note`). `npm run bench -- 200000` raises the row count.
- **Isolation**: every measurement runs in its **own child process**, so peak RSS belongs to that
  operation and not to whatever ran before it in the same process.
- **Median of 3** runs per library and direction.
- **Same input file for every reader**: each library writes its own file, then *every* library
  reads the **SheetJS-written** file. No library gets to read a file its own writer shaped.
- **Memory**: `process.resourceUsage().maxRSS` of the child, minus the `ru_maxrss` of a bare
  `node -e` process (the baseline for module loading and the runtime).
- **File size**: `stat` of the file each writer produced.

Two honest caveats about the harness:

1. The child process builds the 50,000 source row objects **before** it measures anything,
   including in read mode. That array (tens of MB) is therefore inside the reported read RSS for
   every library. It is the same handicap for all three, but it is not part of reading a file.
2. RSS is a peak, not a steady state, and it moves between runs (README says ±15%; times ±5%).

## Published numbers

From the project README, measured on its author's machine (Node 26):

| | excel-lite | ExcelJS | SheetJS |
|---|---|---|---|
| write 50,000 rows | **1.10 s** | 2.73 s | 1.62 s |
| read 50,000 rows | **450 ms** | 2.19 s | 1.94 s |
| peak RSS (read) | **62 MB** | 340 MB | 182 MB |
| file size written | **1.54 MB** | 1.51 MB | 11.54 MB |

Read is about 4.9× faster than ExcelJS and 4.3× faster than SheetJS on that file; peak memory is
about 5.5× and 2.9× lower. Write is about 2.5× faster than ExcelJS. The file-size column is worth noticing: the readers all read the
11.54 MB SheetJS file, so a SheetJS export is roughly 7.5× larger than the same data written by
ExcelJS or `excel-lite`.

## Reproduction

Run `npm run bench` yourself. On this shared-CPU machine, across runs:

| Metric | Range observed | Notes |
|---|---|---|
| read 50k rows | 450–609 ms | ExcelJS 2.17–2.42 s, SheetJS 1.94–2.38 s |
| write 50k rows | 1.10–1.41 s | ExcelJS 2.73–3.39 s, SheetJS 1.62–1.76 s |
| peak RSS, read | 62–95 MB | ExcelJS 340–388 MB, SheetJS 182–221 MB |

Reads are consistently 4–5× faster than both alternatives and use 3–5× less peak memory; write is
2–3× faster than ExcelJS. Treat the numbers as order-of-magnitude evidence, not a contract.

One caveat about the harness itself: an earlier version of `scripts/bench.mjs` built the 50,000
source rows for read measurements too, which added ~30 MB of harness objects to the read peak. The
current script only builds rows for the write measurement; the numbers above are from the fixed
version.

## The memory model

`readSheet()` parses rows lazily, but it is **not** a streaming parser. In `src/read.ts`:

1. The source is turned into one `Uint8Array` (`toBytes`). An async iterable of chunks is
   concatenated first, so the whole file is in memory.
2. `WorkbookReader.open()` unzips only `xl/workbook.xml` and its relationships, but keeps the
   compressed buffer for later.
3. For the sheet you iterate, `getSheetXml()` inflates that one part (`fflate.unzipSync` with a
   filter), then `decodeUtf8()` turns the bytes into a **JavaScript string**. Strings in V8 are
   UTF-16 (2 bytes per character for the ASCII-heavy XML in practice), so the sheet is now held
   twice: as inflated bytes and as a string.
4. `scanRows()` matches `<row>`/`<c>` elements with regexes over that string and yields rows as
   they are matched, so row objects do not accumulate — but the string and the byte buffer stay
   alive for the whole iteration.
5. `xl/sharedStrings.xml` is loaded lazily by `parseXml()` into a small DOM, mapped to an array of
   strings, and the DOM is discarded. Peak therefore includes a DOM node (with `attrs` and
   `children` objects) per `<si>` — the cost grows with the number of **distinct** strings, not
   with the number of cells.
6. `xl/styles.xml` is parsed to build the "is this style a date?" table. It is small, but it is
   re-parsed on every `rows()` call, not cached across calls.

So the model is roughly:

```
peak RSS ≈ file bytes
         + inflated sheet XML bytes
         + UTF-16 copy of the inflated sheet XML (≈ 2× those bytes)
         + unique shared strings (array + parsed DOM transiently)
         + the rows you keep in your own code
```

For a typical compressed workbook (2.5–7× XML compression) that works out to roughly **5–10× the
compressed file size** for benchmark-shaped data. For a very small file with a large sheet the
multiple is much higher, because the compression ratio is high — the measured table below has a
1.54 MB file peaking at 42 MB (27×).

## Measured memory scenarios

Read-only, one child process per measurement, Node 26.7.0, baseline subtracted. The "unique
strings" count is what drives the shared-string cost; the rows are the same 6-column shape.

| Scenario | File | Inflated sheet XML | Unique strings | Peak RSS | Read time |
|---|---|---|---|---|---|
| 25k rows, `note` unique per row | 0.87 MB | 5.2 MB | 25,000 | 76 MB | ~650 ms |
| 50k rows, `note` unique per row | 1.73 MB | 10.4 MB | 50,000 | 120 MB | ~1.0 s |
| 100k rows, `note` unique per row | 3.47 MB | 20.9 MB | 100,000 | 186 MB | ~1.8 s |
| 200k rows, `note` unique per row | 6.99 MB | 42.8 MB | 200,000 | 311 MB | ~3.4 s |
| 50k rows, repeated values (bench shape) | 1.54 MB | 10.3 MB | 5 | 42 MB | ~0.55 s |
| 50k rows, SheetJS-written (stored, not compressed) | 11.54 MB | 11.5 MB | inline | 64 MB | ~0.5 s |

The 50k/unique-string row and the 50k/repeated row are the same data shape: the 3× difference in
peak (120 MB vs 42 MB) is the shared-string table — 50,000 unique strings become 50,000 DOM
elements, then 50,000 array entries.

The 11.54 MB SheetJS file peaks at only 64 MB because it is *stored*, not compressed: the reader
never has to hold both a small compressed buffer and a large inflated one. Meanwhile the 1.54 MB
`excel-lite` file inflates to 10.3 MB, so its 42 MB peak is **27× its file size** — which is why
the "8×" rule of thumb should be read as "8× for benchmark-shaped files".

Peak RSS is also not proportional to what you consume: breaking out of the loop after 5 rows
stops row parsing, but the sheet string and its inflated bytes were already paid for before the
first row was yielded. Iterating to the end adds row garbage but not the dominant cost.

## The write path

`toXlsx()` is fully in memory and has three phases:

1. **Collect.** Every row is materialised into an array of value arrays. A generator as input
   does not help — the writer still builds the full body first.
2. **Build.** One XML string per sheet (plus the shared-string table, styles and package parts).
   All sheet strings coexist until the zip is produced.
3. **Zip.** `fflate.zipSync` at compression level 6 produces the output buffer while the sheet
   strings are still alive.

So peak write memory ≈ your rows + all sheet XML + the compressed output. Measured: 50,000 rows
with 50,000 unique strings peaks at ~162 MB and takes ~2.2 s (unique strings are deduplicated
through a `Map`, which is also why the bench's 5-distinct-strings workload writes in ~1.2 s).

There is no streaming writer in 0.1. For a fixed-memory export, either cap the export size or
write CSV — `excel-lite` can read CSV-sized volumes fine, it just cannot stream them out.

## When to use something else

| Requirement | What to use | Why |
|---|---|---|
| Read a 200 MB+ sheet in constant memory | wait for the streaming reader on the [roadmap](../ROADMAP.md), or pre-convert to CSV | the sheet XML is inflated whole |
| Random access (`sheet.getCell('B7')`), editing a workbook, reading formatting back | ExcelJS | `excel-lite` is rows-in/rows-out and does not read styles |
| Formula text as well as cached values | SheetJS | `excel-lite` returns the cached value only |
| Legacy `.xls` | SheetJS or a one-off conversion | OOXML only ([compatibility](compatibility.md)) |
| Charts, images, pivot tables, merged cells | not supported by any of this reader's API — use a full workbook library | those parts are ignored |
| CSV in / CSV out only | no spreadsheet library | `excel-lite` is for the `.xlsx` side |

## Practical guidance

- **Cap uploads.** Peak memory is several times the file size, and a 512 MiB cap on a single
  decompressed part is a guard, not a budget. Treat untrusted `.xlsx` uploads as a
  denial-of-service surface and limit size at the edge ([SECURITY.md](../SECURITY.md)).
- **Reuse one open workbook.** `readWorkbook()` parses the archive once; calling `readSheet()` per
  sheet re-inflates the workbook metadata each time.
- **`listSheets()` is not free.** It inflates each sheet to count rows; for a big file prefer
  `workbook.sheetNames` when you do not need counts.
- **Stream the output, not the input.** NDJSON rows to stdout (`process.stdout.write`) keeps your
  own memory flat; the CLI's NDJSON mode does this. `--pretty` buffers.
- **Break early** when you only need validation or the first matches — row parsing stops, and the
  file/XML cost is the fixed part you already paid.
- **Measure, do not guess.** `npm run bench` and `node examples/streaming-large-file.mjs` print
  real numbers for your machine.

## See also

- [Reading](reading.md#limits) — the 512 MiB guard and what is ignored.
- [Architecture](architecture.md#how-a-read-flows) — the same flow as code.
- [ROADMAP](../ROADMAP.md) — streaming reader, `maxEntryBytes` option.
