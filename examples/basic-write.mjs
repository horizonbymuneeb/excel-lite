/**
 * basic-write — write one sheet and a multi-sheet workbook, with columns and formats.
 *
 * Run from the package root (after `npm install && npm run build`):
 *   node examples/basic-write.mjs
 *
 * Both helpers return the bytes they wrote, so the same code path works in a
 * browser where you would keep the Uint8Array instead of passing a path.
 */
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { listSheets, readSheet, toXlsx, writeSheet, writeWorkbook } from 'excel-lite';

const dir = mkdtempSync(join(tmpdir(), 'excel-lite-basic-write-'));
const single = join(dir, 'report.xlsx');
const multi = join(dir, 'book.xlsx');

try {
  const people = [
    { email: 'a@example.com', amount: 1234.5, joined: new Date(Date.UTC(2024, 4, 17)) },
    { email: 'b@example.com', amount: -20, joined: new Date(Date.UTC(2025, 0, 2)) },
  ];

  // --- 1. one sheet: headers, widths, number formats, frozen header row ---
  const bytes = await writeSheet(single, people, {
    name: 'People',
    columns: [
      { key: 'email', header: 'Email', width: 32 },
      { key: 'amount', header: 'Amount', format: '#,##0.00' },
      { key: 'joined', header: 'Joined', format: 'yyyy-mm-dd' },
    ],
    freeze: 'A2', // freeze row 1
  });
  console.log(`wrote ${single} (${bytes.length} bytes, file size ${statSync(single).size})`);

  // --- 2. several sheets in one workbook ---------------------------------
  await writeWorkbook(multi, [
    {
      name: 'Summary',
      rows: people,
      columns: [
        { key: 'email', bold: true },
        { key: 'amount', format: '#,##0.00' },
        { key: 'joined', format: 'yyyy-mm-dd' },
      ],
      autoFilter: true,
    },
    {
      name: 'Raw',
      rows: [
        [1, 2, 3],
        [4, 5, 6],
      ],
    }, // array rows, no header
    { name: 'Notes', rows: [{ note: 'written by excel-lite' }] },
  ]);
  console.log('sheets:', await listSheets(multi));

  // --- 3. in-memory only: no path, no filesystem ------------------------
  const inMemory = toXlsx([{ rows: [['hello', 42]] }]);
  assert.ok(inMemory instanceof Uint8Array);
  assert.equal(String.fromCharCode(inMemory[0], inMemory[1]), 'PK', 'xlsx is a zip container');
  console.log(`toXlsx returned ${inMemory.length} bytes (no path involved)`);

  // --- 4. read back what we wrote ---------------------------------------
  for await (const row of readSheet(single, { header: true })) console.log(row);
  const headers = [];
  for await (const row of readSheet(multi, { sheet: 0 })) {
    headers.push(row);
    break;
  }
  console.log('Summary header row:', headers[0]);
  assert.deepEqual(headers[0], ['email', 'amount', 'joined']);
  console.log('\nok: single sheet, multi-sheet workbook and in-memory bytes all written');
} finally {
  rmSync(dir, { recursive: true, force: true });
}
