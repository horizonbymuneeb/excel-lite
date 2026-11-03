## What and why

<!-- One paragraph. What behaviour changes, and what problem it solves. Link the issue. -->

Fixes #

## How it was verified

<!-- Commands you ran and what they printed. "npm test passes" is the minimum; cross-implementation
     evidence (ExcelJS / SheetJS / openpyxl reading the output) is what we actually need for
     format changes. -->

```
npm test
npm run typecheck
npm run lint
```

- [ ] new behaviour has a test that fails without the change
- [ ] files produced by this change were read back by ExcelJS, SheetJS or openpyxl
- [ ] `README.md` and the relevant `docs/` page are updated (user-visible changes)
- [ ] `CHANGELOG.md` has an entry under "Unreleased"
- [ ] no new runtime dependency (or it was agreed in an issue first)

## Deliberately out of scope

<!-- What this PR does not try to fix, so reviewers do not expect it. -->
