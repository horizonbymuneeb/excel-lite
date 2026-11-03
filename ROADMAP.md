# Roadmap

This is a small library with an intentionally flat API. Items below are ordered by what
users actually hit, not by what is fun to build. Discussion happens in the issue tracker.

## 0.2 — memory and formats

- **True streaming reader.** Today the sheet XML is decompressed into memory before rows are
  parsed, so peak memory grows with the amount of cell data (62 MB for the 50k-row benchmark,
  ~218 MB for 100k rows — `docs/performance.md`). A chunked
  inflate → incremental row scanner keeps memory independent of file size.
- **`maxEntryBytes` as a public option.** The zip-bomb guard is currently a fixed 512 MiB.
- **Merge cells on write.** Read already surfaces the top-left value; writing `merges` is a
  small addition that report generators need.
- **Hyperlinks and comments on read** (as row metadata, not styling).

## 0.3 — formatting read-back

- Read back what we already write: number formats, bold/fill/alignment, column widths, freeze
  state. Needed for "read a template, fill it, write it back" workflows.
- Cell-level write options so a value can carry its own style.

## Considering

- Formula text on read (cached value stays the default) — `docs/migration.md` explains why
  this is not a silent default.
- Defined names / named ranges.
- `.xls` legacy support via an optional adapter package, so core stays small.
- A `@excel-lite/react` helper for download buttons (probably better as a docs recipe).

## Non-goals

- Being a spreadsheet engine: no formula evaluation, no recalculation, no charts, no pivot
  tables, no images, no VBA.
- Random access by cell (`sheet.getCell('B7')`) — that is ExcelJS's model. Iterate rows.
- Supporting every OOXML part in existence. If a part does not affect row values, it is
  ignored.

## How to influence the roadmap

Open an issue with the workflow you are trying to build, not just the feature name. If a
workaround exists (and many do — see `docs/recipes.md`), it goes in the docs instead of the
API. PRs that arrive with tests, a fixture from a real-world file, and docs updates move to
the front of the queue.
