/**
 * streaming-large-file — row-by-row iteration over a large sheet, with timings and memory.
 *
 * Run from the package root (after `npm install && npm run build`):
 *   node examples/streaming-large-file.mjs [rows]     # default 100000
 *
 * Shows three things that matter for big exports:
 *   1. rows are parsed lazily — you can stop after the first few,
 *   2. the whole sheet XML is still decompressed into memory first, so peak RSS
 *      grows with the sheet, not with the number of rows you consumed,
 *   3. through-put on this machine, measured rather than claimed.
 */
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtempSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { listSheets, readSheet, writeSheet } from 'excel-lite';

const ROWS = Number(process.argv[2] ?? 100_000);
const dir = mkdtempSync(join(tmpdir(), 'excel-lite-streaming-'));
const file = join(dir, `${ROWS}.xlsx`);

const names = ['Muneeb', 'Sara', 'Ali', 'Nina', 'Ünïcode ✓', 'Chen', 'Amara', 'Diego'];
const rows = Array.from({ length: ROWS }, (_, index) => ({
  id: index + 1,
  name: names[index % names.length],
  qty: (index * 7) % 500,
  price: Math.round(((index * 1337) % 100_000) / 100) + 0.99,
  when: new Date(Date.UTC(2020, 0, 1) + (index % 1500) * 86_400_000),
  note: index % 97 === 0 ? null : `note ${index}`,
}));

const mb = (bytes) => `${(bytes / 1024 / 1024).toFixed(1)} MB`;

try {
  // --- write ------------------------------------------------------------
  let started = performance.now();
  await writeSheet(file, rows, {
    columns: [
      { key: 'id', header: 'ID' },
      { key: 'name' },
      { key: 'qty' },
      { key: 'price', format: '#,##0.00' },
      { key: 'when', header: 'date', format: 'yyyy-mm-dd' },
      { key: 'note' },
    ],
    freeze: 'A2',
  });
  const writeMs = performance.now() - started;
  const size = statSync(file).size;
  console.log(`wrote ${ROWS.toLocaleString()} rows in ${writeMs.toFixed(0)} ms — ${mb(size)} on disk`);

  // --- row count without reading cells ----------------------------------
  console.log('sheet info:', await listSheets(file));

  // --- full read --------------------------------------------------------
  started = performance.now();
  let count = 0;
  let totalQty = 0;
  for await (const row of readSheet(file, { header: true })) {
    count++;
    totalQty += row.qty ?? 0;
  }
  const readMs = performance.now() - started;
  console.log(`read ${count.toLocaleString()} rows in ${readMs.toFixed(0)} ms (qty total ${totalQty})`);
  assert.equal(count, ROWS);

  // --- early exit: stop after 5 rows ------------------------------------
  started = performance.now();
  const first = [];
  for await (const row of readSheet(file, { header: true })) {
    first.push(row.ID);
    if (first.length === 5) break; // the generator stops, and so does the work
  }
  console.log(`first 5 ids in ${(performance.now() - started).toFixed(0)} ms: ${first.join(', ')}`);
  assert.deepEqual(first, [1, 2, 3, 4, 5]);

  // --- memory -----------------------------------------------------------
  // The number below mixes writing and reading in this process (the row array
  // this script built is still alive), so measure a clean read in its own
  // process — that is the number that matters for a server job.
  const peak = process.resourceUsage().maxRSS * 1024;
  const readOnly = await new Promise((resolve, reject) => {
    const code = `
      import('excel-lite').then(async ({ readSheet }) => {
        let rows = 0;
        for await (const _row of readSheet(process.argv[1])) rows++;
        process.stdout.write(JSON.stringify({ rows, rss: process.resourceUsage().maxRSS * 1024 }));
      });
    `;
    execFile(process.execPath, ['-e', code, file], { maxBuffer: 1 << 20 }, (error, stdout) => {
      if (error) reject(error);
      else resolve(JSON.parse(stdout));
    });
  });
  console.log(
    `read-only peak RSS: ${mb(readOnly.rss)} (${readOnly.rows.toLocaleString()} rows) for a ${mb(size)} file`,
  );
  console.log(`this process (write + read): ${mb(peak)}`);
  console.log('memory tracks the inflated sheet XML, not the row count you consumed — see docs/performance.md');
  console.log('\nok: large file written, streamed, and stopped early');
} finally {
  rmSync(dir, { recursive: true, force: true });
}
