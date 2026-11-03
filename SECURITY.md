# Security policy

## Supported versions

| Version | Supported |
| ------- | --------- |
| 0.1.x   | ✅        |

## Reporting a vulnerability

Please do **not** open a public issue for a security problem.

- Preferred: [GitHub Security Advisories](https://github.com/horizonbymuneeb/excel-lite/security/advisories/new)
- Or email: horizonbymuneeb@users.noreply.github.com

Include the version, a minimal reproducer (a script or a small `.xlsx`), the impact you
expect, and whether the file came from an untrusted source. You will get an acknowledgement
within a few days; a fix and coordinated disclosure usually follow within 14 days.

## Threat model

excel-lite parses untrusted input: workbooks arrive from uploads, email attachments and
third-party APIs. The design choices that matter:

- **No macro execution.** `.xlsm` files are read as data; `vbaProject.bin` is never touched.
- **No DTD / external entity processing.** The XML reader decodes the five predefined
  entities and numeric character references only. `<!DOCTYPE>` is skipped, so XXE and
  entity-expansion ("billion laughs") attacks have no path in.
- **Zip entries are inflated on demand.** Only the parts needed for the requested sheet are
  decompressed, and a single entry larger than 512 MiB is rejected before inflation, which
  bounds the damage a zip bomb can cause. This is a hard limit for now — open an issue if
  you legitimately need larger sheets.
- **No network `eval`, no dynamic code execution, no shell.** The library never executes
  cell content. Formula text is not evaluated; cached values are returned.
- **No filesystem access unless you pass a path.** In browsers the bundle contains no Node
  builtins at all (`docs/browser.md` explains the mechanism).
- **Written files are XML-escaped** and characters forbidden by XML 1.0 are stripped, so a
  cell value cannot break out of the document structure.

Known non-issues: a workbook can still exhaust memory on a very large sheet — peak memory grows
with decompressed cell data, not with the compressed file size (62 MB for a 50k-row benchmark,
~218 MB for 100k rows; see `docs/performance.md`). Treat the size of files you accept from
untrusted users as a denial-of-service budget, and cap uploads at your edge.

## Dependency policy

Runtime dependencies: `fflate` only. Dependabot watches devDependencies and the lockfile.
