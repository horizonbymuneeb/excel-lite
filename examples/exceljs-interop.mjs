/**
 * exceljs-interop — move between excel-lite and ExcelJS in both directions.
 *
 * Run from the package root (after `npm install && npm run build`):
 *   node examples/exceljs-interop.mjs
 *
 * ExcelJS is a devDependency here; in your own project install it yourself
 * (`npm i exceljs`) only if you need it. excel-lite writes formatting but does
 * not read it back, so the common pattern is: write rows with excel-lite, then
 * open the result in ExcelJS if you need to inspect or extend formatting.
 */
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { readSheet, writeSheet } from 'excel-lite';

const dir = mkdtempSync(join(tmpdir(), 'excel-lite-exceljs-'));
const ours = join(dir, 'from-excel-lite.xlsx');
const theirs = join(dir, 'from-exceljs.xlsx');

try {
  // --- 1. excel-lite writes, ExcelJS reads ------------------------------
  await writeSheet(
    ours,
    [
      { name: 'Ali', amount: 1234.56, when: new Date(Date.UTC(2023, 6, 4)), done: true },
      { name: 'Nina', amount: -2, when: new Date(Date.UTC(2022, 10, 30)), done: false },
    ],
    {
      columns: [
        { key: 'name', header: 'Name', width: 24, bold: true, fill: '1F2937', align: 'left' },
        { key: 'amount', header: 'Amount', format: '#,##0.00' },
        { key: 'when', header: 'When', format: 'yyyy-mm-dd' },
        { key: 'done', header: 'Done' },
      ],
      freeze: 'A2',
      autoFilter: true,
    },
  );

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(ours);
  const sheet = workbook.getWorksheet('Sheet1');

  console.log('ExcelJS reads it back:');
  console.log('  A1 font bold:', sheet.getCell('A1').font?.bold);
  console.log('  A1 fill:', JSON.stringify(sheet.getCell('A1').fill));
  console.log('  B2 numFmt:', sheet.getCell('B2').numFmt);
  console.log('  C2 value:', sheet.getCell('C2').value?.toISOString?.() ?? sheet.getCell('C2').value);
  console.log('  freeze:', JSON.stringify(sheet.views));
  console.log('  autoFilter ref:', sheet.autoFilter);
  console.log('  column A width:', sheet.getColumn(1).width);

  assert.equal(sheet.getCell('A1').font?.bold, true);
  assert.equal(sheet.getCell('B2').numFmt, '#,##0.00');
  assert.equal(sheet.getCell('A2').value, 'Ali');
  assert.equal(sheet.getCell('D3').value, false);
  assert.equal(sheet.getCell('C2').value.toISOString(), '2023-07-04T00:00:00.000Z');
  assert.equal(sheet.views[0].state, 'frozen');
  assert.equal(sheet.autoFilter, 'A1:D1');

  // --- 2. ExcelJS writes, excel-lite reads ------------------------------
  const theirsBook = new ExcelJS.Workbook();
  const source = theirsBook.addWorksheet('Data');
  source.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Amount', key: 'amount', width: 12 },
    { header: 'When', key: 'when', width: 14, style: { numFmt: 'yyyy-mm-dd' } },
    { header: 'Done', key: 'done' },
  ];
  source.addRow({ name: 'Sara', amount: 500, when: new Date(Date.UTC(2025, 0, 2)), done: true });
  source.addRow({ name: 'Chen', amount: 0.25, when: new Date(Date.UTC(2024, 1, 29)), done: false });
  await theirsBook.xlsx.writeFile(theirs);

  console.log('\nexcel-lite reads the ExcelJS file:');
  for await (const row of readSheet(theirs, { header: true })) {
    console.log(' ', row);
  }
  const rows = [];
  for await (const row of readSheet(theirs, { header: true })) rows.push(row);
  assert.equal(rows.length, 2);
  assert.equal(rows[0].Amount, 500); // ExcelJS headers are the object keys
  assert.equal(rows[1].When.toISOString(), '2024-02-29T00:00:00.000Z');
  assert.equal(rows[1].Done, false);

  console.log('\nok: excel-lite -> ExcelJS (values + styles) and ExcelJS -> excel-lite (values)');
} finally {
  rmSync(dir, { recursive: true, force: true });
}
