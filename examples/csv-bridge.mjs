/**
 * csv-bridge — CSV text to .xlsx, and .xlsx back to CSV, with no extra dependencies.
 *
 * Run from the package root (after `npm install && npm run build`):
 *   node examples/csv-bridge.mjs
 *
 * The CSV parser below is deliberately small (quotes, doubled quotes, newlines
 * inside fields). Use a real CSV library if your data has other dialects — the
 * point here is that the xlsx side needs no help.
 */
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { readSheet, writeSheet } from 'excel-lite';

/** Split CSV text into string rows. Handles "quoted, fields" and "" escapes. */
function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (quoted) {
      if (char === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (char === '"') {
        quoted = false;
      } else {
        field += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ',') {
      row.push(field);
      field = '';
    } else if (char === '\n' || char === '\r') {
      if (char === '\r' && text[i + 1] === '\n') i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += char;
    }
  }
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows.filter((fields) => fields.some((value) => value.trim() !== ''));
}

/** Quote one CSV field only when it needs it. */
function csvField(value) {
  const text = value instanceof Date ? value.toISOString() : value === null || value === undefined ? '' : String(value);
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

const dir = mkdtempSync(join(tmpdir(), 'excel-lite-csv-'));
const file = join(dir, 'from-csv.xlsx');

const csv = [
  'sku,description,price,in_stock',
  'A-1,Plain widget,19.99,true',
  'A-2,"Widget, large",24.50,false',
  'B-7,"He said ""hello""",5,true',
].join('\n');

try {
  // --- CSV -> xlsx ------------------------------------------------------
  const [head, ...data] = parseCsv(csv);
  const rows = data.map((fields) =>
    Object.fromEntries(
      head.map((key, index) => {
        const raw = fields[index] ?? '';
        // Coerce numbers and booleans; keep everything else as text.
        if (raw === 'true' || raw === 'false') return [key, raw === 'true'];
        return [key, raw !== '' && Number.isFinite(Number(raw)) ? Number(raw) : raw];
      }),
    ),
  );

  await writeSheet(file, rows, {
    columns: [
      { key: 'sku', width: 10 },
      { key: 'description', width: 24 },
      { key: 'price', format: '#,##0.00' },
      { key: 'in_stock', header: 'In stock' },
    ],
    freeze: 'A2',
  });
  console.log(`CSV -> ${file}: ${rows.length} data rows`);

  // --- xlsx -> CSV ------------------------------------------------------
  let out = '';
  let headerKeys = null;
  for await (const row of readSheet(file, { header: true })) {
    if (!headerKeys) {
      headerKeys = Object.keys(row);
      out += `${headerKeys.map(csvField).join(',')}\n`;
    }
    out += `${headerKeys.map((key) => csvField(row[key])).join(',')}\n`;
  }
  process.stdout.write(out);
  assert.match(out, /"Widget, large"/);
  assert.match(out, /"He said ""hello"""/);

  // --- xlsx -> NDJSON (the other useful bridge) -------------------------
  let ndjson = '';
  for await (const row of readSheet(file, { header: true })) ndjson += `${JSON.stringify(row)}\n`;
  const parsed = ndjson
    .trim()
    .split('\n')
    .map((line) => JSON.parse(line));
  assert.equal(parsed.length, 3);
  assert.equal(parsed[1].price, 24.5);
  console.log(`\nok: round-tripped ${parsed.length} rows CSV -> xlsx -> CSV/NDJSON`);
} finally {
  rmSync(dir, { recursive: true, force: true });
}
