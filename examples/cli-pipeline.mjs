/**
 * cli-pipeline — drive the `excel-lite` CLI from Node and pipe it into other processes.
 *
 * Run from the package root (after `npm install && npm run build`):
 *   node examples/cli-pipeline.mjs
 *
 * The CLI lives at dist/cli.js in this repo; after `npm i -g excel-lite` (or
 * `npx excel-lite`) it is the `excel-lite` command. The shell equivalents:
 *
 *   excel-lite info report.xlsx
 *   excel-lite to-json report.xlsx --header | jq -c '.email'
 *   excel-lite to-json big.xlsx | head -100        # EPIPE is handled: exit 0
 *   excel-lite to-csv report.xlsx --out report.csv
 *   excel-lite to-json report.xlsx --sheet Notes --pretty
 */
import assert from 'node:assert/strict';
import { execFile, spawn } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';
import { writeSheet } from 'excel-lite';

const run = promisify(execFile);
const cli = new URL('../dist/cli.js', import.meta.url).pathname;
const dir = mkdtempSync(join(tmpdir(), 'excel-lite-cli-'));
const file = join(dir, 'input.xlsx');
const big = join(dir, 'big.xlsx');

try {
  await writeSheet(
    file,
    [
      { email: 'a@example.com', amount: 10.5, when: new Date(Date.UTC(2026, 0, 2)) },
      { email: 'b@example.com', amount: 99, when: new Date(Date.UTC(2026, 0, 3)) },
    ],
    { columns: [{ key: 'email' }, { key: 'amount', format: '#,##0.00' }, { key: 'when', format: 'yyyy-mm-dd' }] },
  );

  // --- information about the workbook -----------------------------------
  const info = await run(process.execPath, [cli, 'info', file]);
  process.stdout.write(info.stdout);

  // --- NDJSON: one JSON object per line, parse it as it arrives ----------
  const { stdout: ndjson } = await run(process.execPath, [cli, 'to-json', file, '--header']);
  const rows = ndjson
    .trim()
    .split('\n')
    .map((line) => JSON.parse(line));
  console.log('NDJSON rows:', rows.length, rows[0]);
  assert.equal(rows[1].amount, 99);

  // --- --pretty: one JSON array (the CLI buffers the whole sheet) --------
  const { stdout: json } = await run(process.execPath, [cli, 'to-json', file, '--pretty', '--header']);
  assert.ok(Array.isArray(JSON.parse(json)));

  // --- CSV with a custom delimiter --------------------------------------
  const { stdout: tsv } = await run(process.execPath, [cli, 'to-csv', file, '--delimiter', '\t']);
  assert.match(tsv, /\t/);

  // --- write to a file instead of stdout --------------------------------
  const outFile = join(dir, 'out.csv');
  await run(process.execPath, [cli, 'to-csv', file, '--out', outFile, '--header']);
  console.log('--out file starts with:', (await run('head', ['-n', '1', outFile])).stdout.trim());

  // --- piping into another process, and stopping early ------------------
  // 20k rows: more than one pipe buffer, so the CLI really does get EPIPE
  // when `head` exits. It treats EPIPE as success instead of crashing.
  await writeSheet(
    big,
    Array.from({ length: 20_000 }, (_, i) => ({ id: i + 1, note: `row ${i + 1}` })),
    {
      columns: [{ key: 'id' }, { key: 'note' }],
    },
  );

  // `child.stdout.destroy()` is what `| head -1` does to the pipe: the reader
  // goes away, the next write fails with EPIPE.
  const child = spawn(process.execPath, [cli, 'to-json', big, '--header'], {
    stdio: ['ignore', 'pipe', 'inherit'],
  });
  let firstLine = '';
  child.stdout.on('data', (chunk) => {
    firstLine += chunk;
    if (firstLine.includes('\n')) child.stdout.destroy(); // "head -1" hangs up
  });
  const cliCode = await new Promise((resolve) => child.on('close', resolve));
  console.log('first line through `| head -1`:', firstLine.split('\n')[0]);
  console.log(`exit code — cli: ${cliCode}`);
  assert.equal(cliCode, 0, 'EPIPE must not fail the CLI');

  console.log('\nok: info, NDJSON, --pretty, --delimiter, --out and | head -1 (EPIPE) all behave');
} finally {
  rmSync(dir, { recursive: true, force: true });
}
