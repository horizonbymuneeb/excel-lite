// CommonJS smoke test: `require('excel-lite')` must resolve to the CJS build.
// Run by test/api.test.ts with the Node binary; not collected by vitest itself.
const { toXlsx, readSheet, columnLetter, columnIndexFromRef } = require('excel-lite');

(async () => {
  const rows = [];
  for await (const row of readSheet(toXlsx([{ rows: [['cjs', 1]] }]))) rows.push(row);
  process.stdout.write(
    JSON.stringify({
      rows,
      letter: columnLetter(26),
      index: columnIndexFromRef('XFD1'),
      exports: Object.keys(require('excel-lite')).sort(),
    }),
  );
})().catch((error) => {
  process.stderr.write(`${error && error.stack ? error.stack : error}\n`);
  process.exit(1);
});
