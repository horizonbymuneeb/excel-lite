import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { listSheets, readSheet, readWorkbook, toXlsx, writeSheet } from 'excel-lite';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const tmp = mkdtempSync(join(tmpdir(), 'excel-lite-'));
const file = (name: string) => join(tmp, name);

const sampleRows = [
  { name: 'Muneeb', age: 27, joined: new Date(Date.UTC(2024, 4, 17)), active: true, note: 'plain' },
  { name: 'Sara', age: 31.5, joined: new Date(Date.UTC(2025, 0, 2, 9, 30)), active: false, note: null },
  { name: 'Ünïcode ✓', age: 0, joined: null, active: true, note: 'a, "quoted"\nnewline' },
];

const columns = [
  { key: 'name', width: 24 },
  { key: 'age', format: '#,##0.00' },
  { key: 'joined', format: 'yyyy-mm-dd' },
  { key: 'active' },
  { key: 'note' },
];

async function collect<T>(iterable: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const item of iterable) out.push(item);
  return out;
}

describe('roundtrip', () => {
  it('writes and reads back every supported value type', async () => {
    await writeSheet(file('types.xlsx'), sampleRows, { name: 'People', columns });
    const rows = await collect(readSheet(file('types.xlsx'), { header: true }));

    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual({
      name: 'Muneeb',
      age: 27,
      joined: new Date(Date.UTC(2024, 4, 17)),
      active: true,
      note: 'plain',
    });
    expect(rows[1]?.joined).toEqual(new Date(Date.UTC(2025, 0, 2, 9, 30)));
    expect(rows[2]?.note).toBe('a, "quoted"\nnewline');
    expect(rows[2]?.joined).toBeNull();
  });

  it('keeps column order and header names', async () => {
    await writeSheet(file('order.xlsx'), [{ b: 1, a: 2, c: 3 }], {
      columns: [{ key: 'a' }, { key: 'b' }, { key: 'c', header: 'C label' }],
    });
    const rows = await collect(readSheet(file('order.xlsx'), { header: true }));
    expect(rows).toEqual([{ a: 2, b: 1, 'C label': 3 }]);
  });

  it('writes multiple sheets and reads each by name or index', async () => {
    const bytes = toXlsx([
      {
        name: 'First',
        rows: [
          ['a', 1],
          ['b', 2],
        ],
      },
      { name: 'Second', rows: [['c', 3]] },
    ]);
    writeFileSync(file('multi.xlsx'), bytes);

    const workbook = await readWorkbook(file('multi.xlsx'));
    expect(workbook.sheetNames).toEqual(['First', 'Second']);
    expect(await collect(workbook.sheet('Second'))).toEqual([['c', 3]]);
    expect(await collect(readSheet(file('multi.xlsx'), { sheet: 0, header: false }))).toEqual([
      ['a', 1],
      ['b', 2],
    ]);
    expect((await listSheets(file('multi.xlsx'))).map((sheet) => sheet.rowCount)).toEqual([2, 1]);
  });

  it('preserves dates with and without a time part', async () => {
    const midnight = new Date(Date.UTC(2020, 0, 1));
    const withTime = new Date(Date.UTC(2020, 0, 1, 13, 45, 30));
    await writeSheet(file('dates.xlsx'), [[midnight, withTime]], { header: false });
    const rows = await collect(readSheet(file('dates.xlsx')));
    expect(rows[0]?.[0]).toEqual(midnight);
    expect(rows[0]?.[1]).toEqual(withTime);
  });

  it('survives strings that Excel would normally choke on', async () => {
    const nasty = ' <tag> & "quotes" \'apostrophe\' \u0001 control \u00e9 \u{1f600} ';
    await writeSheet(file('nasty.xlsx'), [[nasty]]);
    const rows = await collect(readSheet(file('nasty.xlsx')));
    expect(rows[0]?.[0]).toContain('<tag> & "quotes"');
    expect(rows[0]?.[0]).toContain('😀');
    expect(rows[0]?.[0]).not.toContain('\u0001');
  });

  it('skips blank rows but keeps interior nulls', async () => {
    await writeSheet(file('sparse.xlsx'), [['a', null, 'c'], [null, null], ['d']], { header: false });
    const rows = await collect(readSheet(file('sparse.xlsx')));
    expect(rows).toEqual([['a', null, 'c'], ['d']]);
  });
});

describe('cross validation', () => {
  it('reads a file written by ExcelJS', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Data');
    sheet.columns = [
      { header: 'Name', key: 'name', width: 20 },
      { header: 'Amount', key: 'amount', width: 12 },
      { header: 'When', key: 'when', width: 14, style: { numFmt: 'yyyy-mm-dd' } },
      { header: 'Done', key: 'done' },
    ];
    sheet.addRow({ name: 'Ali', amount: 1234.56, when: new Date(Date.UTC(2023, 6, 4)), done: true });
    sheet.addRow({ name: 'Nina', amount: -2, when: new Date(Date.UTC(2022, 10, 30)), done: false });
    await workbook.xlsx.writeFile(file('exceljs.xlsx'));

    const rows = await collect(readSheet(file('exceljs.xlsx'), { header: true }));
    expect(rows[0]).toEqual({
      Name: 'Ali',
      Amount: 1234.56,
      When: new Date(Date.UTC(2023, 6, 4)),
      Done: true,
    });
  });

  it('reads a file written by SheetJS', async () => {
    const sheet = XLSX.utils.aoa_to_sheet([
      ['id', 'label', 'ratio'],
      [1, 'alpha', 0.25],
      [2, 'beta', 0.5],
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
    XLSX.writeFile(workbook, file('sheetjs.xlsx'));

    const rows = await collect(readSheet(file('sheetjs.xlsx'), { header: true }));
    expect(rows).toEqual([
      { id: 1, label: 'alpha', ratio: 0.25 },
      { id: 2, label: 'beta', ratio: 0.5 },
    ]);
  });

  it('writes files that ExcelJS and SheetJS can read', async () => {
    await writeSheet(file('ours.xlsx'), sampleRows, { columns });

    const exceljs = new ExcelJS.Workbook();
    await exceljs.xlsx.readFile(file('ours.xlsx'));
    const sheet = exceljs.getWorksheet('Sheet1')!;
    expect(sheet.getCell('A2').value).toBe('Muneeb');
    expect(sheet.getCell('B2').value).toBe(27);
    expect(sheet.getCell('C2').value).toEqual(new Date(Date.UTC(2024, 4, 17)));
    expect(sheet.getCell('D2').value).toBe(true);
    expect(sheet.getCell('B3').numFmt).toBe('#,##0.00');

    const workbook = XLSX.readFile(file('ours.xlsx'));
    const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets.Sheet1!, { raw: true });
    expect(json[0]).toMatchObject({ name: 'Muneeb', age: 27, active: true });
  });

  it('is readable from a Buffer, Uint8Array, ArrayBuffer and an async byte stream', async () => {
    const bytes = toXlsx([{ rows: [['x', 1]] }]);
    expect(await collect(readSheet(bytes))).toEqual([['x', 1]]);
    expect(await collect(readSheet(bytes.buffer as ArrayBuffer))).toEqual([['x', 1]]);

    async function* chunks() {
      for (let offset = 0; offset < bytes.length; offset += 512) {
        yield bytes.subarray(offset, offset + 512);
      }
    }
    expect(await collect(readSheet(chunks()))).toEqual([['x', 1]]);
  });
});

describe('errors', () => {
  it('rejects non-xlsx input with a readable message', async () => {
    const path = file('bad.xlsx');
    writeFileSync(path, 'not a zip at all');
    await expect(collect(readSheet(path))).rejects.toThrow(/not a zip|workbook\.xml/);
  });

  it('reports unknown sheet names with the available ones', async () => {
    await writeSheet(file('named.xlsx'), [['a']], { name: 'Only' });
    await expect(collect(readSheet(file('named.xlsx'), { sheet: 'Nope' }))).rejects.toThrow(/only/i);
  });

  it('sanitises sheet names Excel would reject', async () => {
    const bytes = toXlsx([
      { name: 'a/b*c?', rows: [[1]] },
      { name: 'a/b*c?', rows: [[2]] },
      { name: 'x'.repeat(40), rows: [[3]] },
    ]);
    writeFileSync(file('names.xlsx'), bytes);
    const names = (await listSheets(file('names.xlsx'))).map((sheet) => sheet.name);
    expect(names[0]).not.toMatch(/[\\/?*[\]:]/);
    expect(names[0]).not.toBe(names[1]);
    expect(names[2]!.length).toBeLessThanOrEqual(31);
  });

  it('refuses an empty workbook', () => {
    expect(() => toXlsx([])).toThrow(/at least one sheet/i);
  });
});

describe('cross round-trip: ours -> third party -> ours', () => {
  const dateOnly = new Date(Date.UTC(2024, 4, 17));
  const dateTime = new Date(Date.UTC(2024, 4, 17, 13, 45, 30));
  const values: (string | number | boolean | Date | null)[][] = [
    ['text', 'a & b <c>', 'Ünïcode ✓ 😀'],
    [42, -3.25, 10_000_000_000],
    [true, false, null],
    [dateOnly, dateTime, null],
  ];
  // Rows stop at their last non-null cell, so the trailing nulls are not rows of nulls.
  const expected: (string | number | boolean | Date | null)[][] = [
    ['text', 'a & b <c>', 'Ünïcode ✓ 😀'],
    [42, -3.25, 10_000_000_000],
    [true, false],
    [dateOnly, dateTime],
  ];

  it('survives a rewrite by ExcelJS', async () => {
    const ours = file('cross-ours.xlsx');
    writeFileSync(ours, toXlsx([{ name: 'Values', rows: values }]));

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(ours);
    const viaExcelJs = file('cross-exceljs.xlsx');
    await workbook.xlsx.writeFile(viaExcelJs);

    expect(await collect(readSheet(viaExcelJs))).toEqual(expected);
  });

  it('survives a rewrite by SheetJS', async () => {
    const ours = file('cross-ours-2.xlsx');
    const bytes = toXlsx([{ name: 'Values', rows: values }]);
    writeFileSync(ours, bytes);

    const workbook = XLSX.read(bytes, { type: 'buffer', cellNF: true, cellStyles: true });
    const out = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer', cellNF: true, cellStyles: true });
    const viaSheetJs = file('cross-sheetjs.xlsx');
    writeFileSync(viaSheetJs, out);

    expect(await collect(readSheet(viaSheetJs))).toEqual(expected);
  });

  it('keeps every sheet through an ExcelJS rewrite', async () => {
    const ours = file('cross-multi.xlsx');
    writeFileSync(
      ours,
      toXlsx([
        { name: 'First', rows: [['a', 1]] },
        { name: 'Second', rows: [['b', 2]] },
      ]),
    );
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(ours);
    const rewritten = file('cross-multi-exceljs.xlsx');
    await workbook.xlsx.writeFile(rewritten);

    const reopened = await readWorkbook(rewritten);
    expect(reopened.sheetNames).toEqual(['First', 'Second']);
    expect(await collect(reopened.sheet('First'))).toEqual([['a', 1]]);
    expect(await collect(reopened.sheet('Second'))).toEqual([['b', 2]]);
  });
});
