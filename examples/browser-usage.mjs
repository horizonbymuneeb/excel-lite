/**
 * browser-usage — the browser/worker path: bytes in, bytes out, no file paths.
 *
 * Run from the package root (after `npm install && npm run build`):
 *   node examples/browser-usage.mjs
 *
 * Nothing in this file imports `node:fs` or anything else Node-only, so the same
 * code runs unchanged in a browser, a web worker, Bun or Deno. `toXlsx()` returns
 * a Uint8Array; `readSheet()` accepts Uint8Array, ArrayBuffer or Blob.
 *
 * In a browser you would hand the bytes to the user with a Blob URL — see
 * docs/browser.md and docs/recipes.md ("browser download").
 */
import assert from 'node:assert/strict';
import { readSheet, toXlsx } from 'excel-lite';

// --- write in memory -------------------------------------------------------
const bytes = toXlsx([
  {
    name: 'Orders',
    rows: [
      { id: 1, customer: 'Ada', total: 19.99, placed: new Date(Date.UTC(2026, 0, 2)) },
      { id: 2, customer: 'Linus', total: 42, placed: new Date(Date.UTC(2026, 0, 3)) },
    ],
    columns: [
      { key: 'id' },
      { key: 'customer' },
      { key: 'total', format: '#,##0.00' },
      { key: 'placed', format: 'yyyy-mm-dd' },
    ],
    freeze: 'A2',
  },
]);

assert.ok(bytes instanceof Uint8Array);
console.log(`toXlsx -> Uint8Array (${bytes.length} bytes, starts with "${String.fromCharCode(...bytes.slice(0, 2))}")`);

// --- the four ways to get bytes into readSheet ----------------------------
async function collect(source) {
  const rows = [];
  for await (const row of readSheet(source, { header: true })) rows.push(row);
  return rows;
}

const fromBytes = await collect(bytes);
const fromArrayBuffer = await collect(bytes.slice().buffer);
const fromBlob = await collect(
  new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
);

// An async iterable of chunks is what a fetch body gives you in the browser.
async function* chunks() {
  for (let offset = 0; offset < bytes.length; offset += 2048) yield bytes.subarray(offset, offset + 2048);
}
const fromChunks = await collect(chunks());

console.log('rows via Uint8Array   :', fromBytes.length);
console.log('rows via ArrayBuffer  :', fromArrayBuffer.length);
console.log('rows via Blob         :', fromBlob.length);
console.log('rows via async chunks :', fromChunks.length);

assert.deepEqual(fromBytes, fromBlob);
assert.deepEqual(fromBytes, fromChunks);
assert.equal(fromBytes[1].customer, 'Linus');
assert.ok(fromBytes[0].placed instanceof Date);

// --- turning the bytes into a download (browser only, no Node required) ---
// Kept as a function so this example runs in Node without a DOM.
function downloadInBrowser(name, data) {
  const url = URL.createObjectURL(
    new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
  );
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}
console.log(`\ndownloadInBrowser(${'report.xlsx'}) is ready for the browser (not called here)`);
void downloadInBrowser;

console.log('ok: wrote and read .xlsx using only bytes — no file system, no Node builtins');
