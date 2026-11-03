# Command line

The package ships one binary, `excel-lite`, for turning `.xlsx` files into text without writing
a script. It is the same reader as the library: paths only, no stdin input.

## On this page

- [Usage](#usage)
- [Commands](#commands)
- [Flags](#flags)
- [Output formats](#output-formats)
- [Examples](#examples)
- [Piping](#piping)
- [Exit codes and errors](#exit-codes-and-errors)
- [Notes and limits](#notes-and-limits)

## Usage

```
excel-lite — read .xlsx without the ceremony

Usage:
  excel-lite info <file.xlsx>
  excel-lite to-json <file.xlsx> [options]
  excel-lite to-csv  <file.xlsx> [options]

Options:
  --sheet <name|index>   sheet to read (default: first)
  --header               use the first row as keys
  --out <file>           write to a file instead of stdout
  --pretty               to-json: emit one JSON array instead of NDJSON
  --delimiter <char>     to-csv: field separator (default: ,)
  -h, --help             show this help

Examples:
  excel-lite info report.xlsx
  excel-lite to-json report.xlsx --sheet Data --header > rows.ndjson
  excel-lite to-csv report.xlsx --out report.csv
```

Run it with `npx excel-lite …` (no install), from a project dependency
(`node_modules/.bin/excel-lite`), or globally (`npm i -g excel-lite`). The binary is
`dist/cli.js` and declares `#!/usr/bin/env node`.

## Commands

| Command | Output | Use |
|---|---|---|
| `info <file>` | the file name, then one line per sheet: index, name, row count | check what is in a workbook before reading it |
| `to-json <file>` | NDJSON (one JSON value per line) by default; one JSON array with `--pretty` | feed a loader, `jq`, or another process |
| `to-csv <file>` | RFC-4180-style CSV with proper quoting; a header line only when `--header` is given | spreadsheets, `COPY`, `awk` |

`info` output example:

```
report.xlsx
   0  Data   4 rows
   1  Notes  3 rows
```

The index is 0-based (the same value `--sheet <index>` expects), right-aligned to two characters;
the name column is padded to the longest name (minimum width 4).

## Flags

| Flag | Applies to | Meaning |
|---|---|---|
| `--sheet <name\|index>` | all commands | sheet to read; a value of only digits is a 0-based index, but if no such index exists it is matched against a sheet with that numeric name; anything else is a name (exact, then case-insensitive) |
| `--header` | `to-json`, `to-csv` | use the first non-empty row as keys; `to-json` emits objects, `to-csv` emits a header line |
| `--out <file>` | `to-json`, `to-csv` | write to a file instead of stdout; stdout stays empty |
| `--pretty` | `to-json` | collect all rows and print one indented JSON array instead of NDJSON; ignored (buffered is still buffered) for `to-csv` |
| `--delimiter <char>` | `to-csv` | field separator, default `,`; pass a tab as `$'\t'` in bash/zsh |
| `-h`, `--help` | — | print usage and exit 0 |

Flags and arguments may appear in any order. Unknown options fail; a missing file argument for a
command fails; extra positional arguments fail.

## Output formats

### NDJSON (`to-json`, default)

One JSON value per line, written as rows arrive. With `--header` each line is a JSON object;
without it each line is a JSON array.

```bash
excel-lite to-json report.xlsx --header
```

```json
{"name":"Ali","qty":3,"price":19.99,"when":"2023-07-04T00:00:00.000Z","active":true,"note":"plain text"}
{"name":"Ünïcode ✓","qty":0,"price":0.5,"when":"2022-11-30T13:45:30.000Z","active":false,"note":"comma, \"quote\""}
```

Dates become ISO-8601 strings (`JSON.stringify` on a `Date`), `null` stays `null`, and numbers
and booleans keep their JSON types.

### Single JSON array (`to-json --pretty`)

```bash
excel-lite to-json report.xlsx --header --pretty
```

```json
[
  {
    "name": "Ali",
    "qty": 3
  }
]
```

This buffers every row in memory before printing, so use NDJSON for large sheets.

### CSV (`to-csv`)

Fields are quoted only when they contain the delimiter, a double quote, `\n` or `\r`; quotes are
doubled. Dates are ISO strings, `null`/missing cells become empty fields, booleans are
`true`/`false`.

```bash
excel-lite to-csv report.xlsx --header --delimiter ';'
```

```
name;qty;price;when;active;note
Ali;3;19.99;2023-07-04T00:00:00.000Z;true;plain text
Ünïcode ✓;0;0.5;2022-11-30T13:45:30.000Z;false;"comma, ""quote"""
```

Without `--header`, `to-csv` writes arrays with no header line. `--pretty` has no effect on CSV.

## Examples

```bash
# what is in this file?
npx excel-lite info report.xlsx

# NDJSON to stdout, then count rows
npx excel-lite to-json report.xlsx --header | wc -l

# one column, via jq
npx excel-lite to-json report.xlsx --header | jq -r '.email'

# a second sheet by name, or by index
npx excel-lite to-json report.xlsx --sheet Notes
npx excel-lite to-json report.xlsx --sheet 1

# write files (stdout stays empty)
npx excel-lite to-json report.xlsx --header --out rows.ndjson
npx excel-lite to-csv  report.xlsx --header --out report.csv

# tab-separated, for `cut`/`awk`
npx excel-lite to-csv report.xlsx --delimiter $'\t' | cut -f1

# pretty JSON array for a small sheet
npx excel-lite to-json report.xlsx --sheet Notes --pretty
```

## Piping

The CLI writes to stdout and does not read stdin, so the pipeline shape is always
`excel-lite <file> | something`:

```bash
# load straight into Postgres (server-side COPY)
npx excel-lite to-csv report.xlsx --header \
  | psql "$DATABASE_URL" -c "COPY report FROM STDIN WITH (FORMAT csv, HEADER true)"

# import line-by-line from Node
npx excel-lite to-json report.xlsx --header | node import.mjs

# stop early — the CLI handles the broken pipe and exits 0
npx excel-lite to-json huge.xlsx | head -100

# upload without touching disk
npx excel-lite to-csv report.xlsx --header | gzip > report.csv.gz
```

**EPIPE.** When the reader on the other end exits first (`head`, `less`, a crashed consumer), the
CLI exits **0** silently instead of printing an `EPIPE` stack trace. That makes
`excel-lite … | head` safe inside scripts and under `set -o pipefail`. The CLI itself has no
streaming write concern: it writes rows as it parses them, so `head` also stops the parsing work.

## Exit codes and errors

| Code | Meaning |
|---|---|
| `0` | success, including `--help` and a closed pipe (EPIPE) |
| `1` | anything that failed: unknown command/flag, missing file argument, unreadable file, bad sheet, write error |

Errors go to **stderr** and are prefixed with `excel-lite: ` — but only once: the CLI leaves the
prefix alone when the message already starts with it, so a library error reads:

```
$ excel-lite to-json report.xlsx --sheet Nope
excel-lite: sheet "Nope" not found — available: Data, Notes
```

Exact messages for the common failures:

| Command line | Exit | stderr |
|---|---|---|
| `info` (no file) | 1 | `excel-lite: "info" needs a .xlsx file` |
| `--bogus` | 1 | `excel-lite: unknown option --bogus` |
| `frobnicate file.xlsx` | 1 | `excel-lite: unknown command "frobnicate" (try --help)` |
| `info /nope/missing.xlsx` | 1 | `excel-lite: ENOENT: no such file or directory, open '/nope/missing.xlsx'` |
| `to-json file.xlsx --out` | 1 | `excel-lite: --out needs a value` |
| `to-json file.xlsx --sheet Nope` | 1 | `excel-lite: sheet "Nope" not found — available: …` |

The CLI does not add a stack trace; the message is the whole error.

## Notes and limits

- **Paths only.** There is no `-` for stdin and no URL input: the first argument after the command
  must be a `.xlsx`/`.xlsm` file path. Convert bytes to text with the library if that is what you
  have.
- **Numeric sheet names** are reachable with `--sheet 2` only when index 2 does not exist — the
  CLI tries the index first, then the name. A workbook where both exist cannot distinguish them
  from the CLI; pass `--sheet "2"` (quoted, still digits, so still index-first) or use the
  library with `{ sheet: '2' }`.
- **`--header` skips the first row**, so on a file without a header row you lose a data row. That
  is the same rule as the library.
- **`--pretty` buffers.** NDJSON is the streaming format; a 500k-row sheet with `--pretty` is held
  in memory before anything is printed.
- **Row counts are physical `<row>` counts**, header included — same as
  [`listSheets()`](reading.md#row-counts-and-listsheets).
- The CLI is statically imported as Node code in the published package (`dist/cli.js`); it is not
  usable from a browser.

## See also

- [Reading](reading.md) — what each value becomes.
- [Recipes](recipes.md) — CSV/NDJSON pipelines in code.
- [Troubleshooting](troubleshooting.md#epipe-econnreset-or-sigpipe) — broken pipes in your own scripts.
