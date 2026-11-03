/**
 * styled-report — a formatted report: bold/filled headers, aligned cells, widths,
 * frozen header row, auto filter, number and date formats.
 *
 * Run from the package root (after `npm install && npm run build`):
 *   node examples/styled-report.mjs [out.xlsx]
 *
 * Pass a path to keep the file (open it in Excel/LibreOffice to look at it),
 * or run without arguments to use a temp file that is deleted afterwards.
 */
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { listSheets, readSheet, writeWorkbook } from 'excel-lite';

const keep = process.argv[2] ? resolve(process.argv[2]) : null;
const dir = keep ? null : mkdtempSync(join(tmpdir(), 'excel-lite-styled-'));
const file = keep ?? join(dir, 'report.xlsx');

const rows = [
  { region: 'North', units: 1240, revenue: 184320.5, margin: 0.18, updated: new Date(Date.UTC(2026, 0, 31)) },
  { region: 'South', units: 980, revenue: 129450, margin: 0.12, updated: new Date(Date.UTC(2026, 0, 31)) },
  { region: 'West', units: 1510, revenue: 241900.75, margin: 0.21, updated: new Date(Date.UTC(2026, 0, 31)) },
];

const header = (key, title, extra = {}) => ({
  key,
  header: title,
  bold: true,
  fill: '1F2937',
  align: 'left',
  ...extra,
});

try {
  await writeWorkbook(file, [
    {
      name: 'Revenue',
      rows,
      header: true,
      freeze: 'A2', // header row stays visible
      autoFilter: true, // filter dropdowns on the header row
      columns: [
        header('region', 'Region', { width: 14 }),
        { ...header('units', 'Units', { align: 'right' }), format: '#,##0' },
        { ...header('revenue', 'Revenue', { align: 'right' }), format: '#,##0.00' },
        { ...header('margin', 'Margin', { align: 'right' }), format: '0.0%' },
        { ...header('updated', 'Updated', { align: 'right' }), format: 'yyyy-mm-dd' },
      ],
    },
    {
      name: 'Notes',
      rows: [
        { note: 'The header row is bold automatically; bold/fill/align on a column style its data cells.' },
        { note: 'Number formats are Excel format codes, e.g. #,##0.00 or 0.0%.' },
        { note: 'The auto filter covers the header row and Excel shows dropdowns for each column.' },
      ],
      columns: [{ key: 'note', width: 80 }],
      autoFilter: false,
    },
  ]);

  console.log(`wrote ${file} (${statSync(file).size} bytes)`);
  console.log('sheets:', await listSheets(file));

  // Read the values back — formatting is write-only, the reader returns values.
  for await (const row of readSheet(file, { header: true })) console.log(row);
  const read = [];
  for await (const row of readSheet(file, { header: true })) read.push(row);
  assert.equal(read.length, 3);
  assert.equal(read[0].Revenue, 184320.5);
  assert.equal(read[0].Updated.toISOString(), '2026-01-31T00:00:00.000Z');
  console.log('\nok: styled report written; values read back unchanged');
} finally {
  if (dir) rmSync(dir, { recursive: true, force: true });
}
