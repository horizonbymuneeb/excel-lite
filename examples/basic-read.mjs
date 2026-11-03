/**
 * basic-read — read a workbook as row arrays and as objects, inspect its sheets.
 *
 * Run from the package root (after `npm install && npm run build`):
 *   node examples/basic-read.mjs
 *
 * Creates a temporary workbook so the example is self-contained, then reads it
 * back. Replace the path with your own file to read real data.
 */
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { listSheets, readSheet, writeSheet } from 'excel-lite';

const dir = mkdtempSync(join(tmpdir(), 'excel-lite-basic-read-'));
const file = join(dir, 'people.xlsx');

try {
  // --- 1. make a file to read (skip this if you already have one) ----------
  await writeSheet(
    file,
    [
      { name: 'Ali', qty: 3, price: 19.99, joined: new Date(Date.UTC(2023, 6, 4)), active: true },
      // a row with a missing value: qty and joined are absent
      { name: 'Sara', price: 0.5, active: false },
      { name: 'Nina', qty: 1299, price: 1234567.891, joined: new Date(Date.UTC(2024, 1, 29)), active: true },
    ],
    {
      columns: [
        { key: 'name' },
        { key: 'qty' },
        { key: 'price', format: '#,##0.00' },
        { key: 'joined', format: 'yyyy-mm-dd' },
        { key: 'active' },
      ],
    },
  );

  // --- 2. objects, keyed by the first row --------------------------------
  console.log('--- header: true');
  for await (const row of readSheet(file, { header: true })) {
    console.log(row.name, row.qty, row.price, row.joined?.toISOString() ?? null, row.active);
  }

  // --- 3. plain arrays, header row included ------------------------------
  console.log('\n--- header: false');
  for await (const row of readSheet(file)) console.log(row);

  // --- 4. a specific sheet (name or 0-based index) -----------------------
  console.log('\n--- sheets in the workbook');
  console.log(await listSheets(file));

  // --- 5. dates are UTC Dates; missing cells are null --------------------
  const rows = [];
  for await (const row of readSheet(file, { header: true })) rows.push(row);
  assert.equal(rows.length, 3);
  assert.equal(rows[1].qty, null, 'a cell that was never written reads as null');
  assert.ok(rows[0].joined instanceof Date);
  assert.equal(rows[0].joined.toISOString(), '2023-07-04T00:00:00.000Z');
  console.log('\nok: 3 rows, missing cells are null, dates are UTC');
} finally {
  rmSync(dir, { recursive: true, force: true });
}
