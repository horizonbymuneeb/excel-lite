# Contributing to excel-lite

Thanks for taking the time. This project is small on purpose, which makes it easy to
contribute to — but the parts that ship (reader, writer, CLI) are used on real financial and
operational data, so "it works on my one file" is not the bar. Tests and honest docs are.

## Ways to contribute

- **Report a bug** with a minimal `.xlsx` fixture or a script that generates one, plus what
  you expected and what you got. Issues that can be reproduced are fixed first.
- **Add a fixture** produced by a _different_ implementation (Excel, LibreOffice, Google
  Sheets, openpyxl, ExcelJS, SheetJS, a customer export). Cross-implementation fixtures are
  the most valuable thing this project can receive.
- **Improve the docs.** If the docs and the code disagree, that is a bug — send a PR with
  whichever one is wrong.
- **Performance** work needs numbers from `npm run bench`, not guesses.

## Development setup

```bash
git clone <your fork>
cd excel-lite
npm install --include=dev   # this box's npm config omits devDependencies by default
npm test                    # builds dist, then runs vitest
npm run typecheck
npm run lint
npm run coverage
npm run bench               # 50k rows vs ExcelJS and SheetJS, ~1 minute
```

Node 22+ is required for anything that takes a file path (`process.getBuiltinModule`);
everything else runs in any modern runtime.

Optional, only if you touch third-party fixtures:

```bash
python3 -m venv /tmp/venv-openpyxl
/tmp/venv-openpyxl/bin/pip install openpyxl
/tmp/venv-openpyxl/bin/python scripts/make-fixtures.py
```

## Project layout

| Path           | What lives there                                                             |
| -------------- | ---------------------------------------------------------------------------- |
| `src/read.ts`  | workbook parsing: zip parts, rels, styles, shared strings, row iteration     |
| `src/write.ts` | OOXML generation: sheet XML, shared strings, styles, package parts           |
| `src/xml.ts`   | the XML reader: small DOM for metadata, row scanner for worksheets           |
| `src/zip.ts`   | fflate wrappers (`unzipEntries` only inflates what is asked for)             |
| `src/fs.ts`    | Node fs access without a `node:fs` import (browser-safe)                     |
| `src/cli.ts`   | `info`, `to-json`, `to-csv`                                                  |
| `test/`        | vitest suites, run against the built `dist` (so packaging bugs fail tests)   |
| `fixtures/`    | workbooks produced by openpyxl — never by our own writer                     |
| `docs/`        | the manual: reading, writing, API, CLI, recipes, performance, compatibility… |
| `examples/`    | runnable scripts (`node examples/basic-read.mjs`)                            |
| `scripts/`     | benchmark, fixture generator, verification helpers                           |

Architecture details, including where to change what: [`docs/architecture.md`](docs/architecture.md).

## Rules of the road

1. **No new runtime dependencies** without discussing it in an issue first. `fflate` is the
   one dependency and it earns its place; the package must stay tiny and auditable.
2. **Tests are the specification.** A behaviour change needs a test that fails before and
   passes after. Fixtures from other implementations beat hand-built ones.
3. **Do not test against our own writer only.** A reader/writer pair can agree on a
   wrong format forever. Every format change should be cross-checked with ExcelJS, SheetJS
   or openpyxl.
4. **Docs are part of the change.** User-visible behaviour means an update to `README.md`
   and the relevant page in `docs/`. The docs link checker fails CI on broken links.
5. **Formatting is automated.** `npm run format` before pushing; `npm run lint` must pass.
6. **Keep the public API small.** New exported names are a maintenance cost; prefer an
   option on an existing function.

## Commit and PR conventions

- Commits: [Conventional Commits](https://www.conventionalcommits.org/) —
  `feat:`, `fix:`, `docs:`, `test:`, `perf:`, `chore:`, `refactor:`.
- One logical change per PR. Split refactors from behaviour changes.
- PR description: what changed, why, how it was verified (commands + output), and what is
  deliberately out of scope.

PR checklist (the template asks the same):

- [ ] `npm test` passes
- [ ] `npm run typecheck` and `npm run lint` pass
- [ ] new behaviour has a test, and the test would fail without the change
- [ ] files written by the change were read back with ExcelJS/SheetJS/openpyxl
- [ ] `README.md` / `docs/` updated for user-visible changes
- [ ] `CHANGELOG.md` has an entry under "Unreleased"

## Release process (maintainers)

1. Update `CHANGELOG.md`; bump `version` in `package.json`.
2. `npm test && npm run lint && npm run typecheck && npm pack --dry-run`.
3. Commit, tag `vX.Y.Z`, push.
4. The release workflow publishes with provenance (`npm publish --provenance`); or publish
   locally with `npm publish --access public` if the workflow is unavailable.
5. Verify from the registry: `npm i excel-lite@X.Y.Z` in an empty directory and read a file.

## Question or unsure?

Open a discussion or an issue — asking before writing code is cheaper for everyone than a
large PR that has to be reshaped.
