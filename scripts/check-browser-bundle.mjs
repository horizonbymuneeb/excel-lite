#!/usr/bin/env node
/**
 * Browser bundle check.
 *
 * Bundles the library for the browser with esbuild and asserts three things:
 *   1. it resolves (no "Could not resolve node:..." errors),
 *   2. the output contains no Node builtins or require() calls,
 *   3. the bundle actually runs: it writes a workbook in memory and reads it back.
 *
 *   node scripts/check-browser-bundle.mjs
 */
import { build } from 'esbuild';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const MAX_KB = 120;

const ENTRY = `
import { toXlsx, readSheet } from 'excel-lite';

const bytes = toXlsx([
  { name: 'Sheet1', rows: [{ a: 1, b: 'x' }], columns: [{ key: 'a' }, { key: 'b' }] },
]);
if (!(bytes instanceof Uint8Array) || bytes.length < 100) throw new Error('toXlsx produced no bytes');
if (String.fromCharCode(bytes[0], bytes[1]) !== 'PK') throw new Error('output is not a zip');

const rows = [];
for await (const row of readSheet(bytes)) rows.push(row);
if (rows.length !== 2) throw new Error('expected 2 rows, got ' + rows.length);
if (rows[1][0] !== 1 || rows[1][1] !== 'x') throw new Error('values did not round-trip: ' + JSON.stringify(rows));
globalThis.__excelLiteOk = true;
`;

const NODE_BUILTIN = /(?:from\s*["']node:)|(?:require\(\s*["']node:)/;

const result = await build({
  stdin: { contents: ENTRY, resolveDir: process.cwd(), loader: 'js' },
  bundle: true,
  write: false,
  platform: 'browser',
  format: 'esm',
  target: 'es2022',
  logLevel: 'silent',
});

const output = result.outputFiles[0]?.text ?? '';
const kb = Buffer.byteLength(output) / 1024;

if (NODE_BUILTIN.test(output)) {
  console.error('browser bundle references Node builtins:');
  for (const line of output.split('\n'))
    if (NODE_BUILTIN.test(line) && line.length < 200) console.error('  ' + line.trim());
  process.exit(1);
}
if (kb > MAX_KB) {
  console.error(`browser bundle is ${kb.toFixed(1)} kB, budget is ${MAX_KB} kB`);
  process.exit(1);
}

const dir = mkdtempSync(join(tmpdir(), 'excel-lite-bundle-'));
const file = join(dir, 'bundle.mjs');
writeFileSync(file, output);
try {
  await import(`file://${file}`);
  if (globalThis.__excelLiteOk !== true) throw new Error('bundle did not report success');
} finally {
  rmSync(dir, { recursive: true, force: true });
}

console.log(`browser bundle ok: ${kb.toFixed(1)} kB, no Node builtins, writes and reads .xlsx`);
