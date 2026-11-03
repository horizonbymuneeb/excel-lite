import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { listSheets, readSheet, toXlsx, writeSheet, writeWorkbook } from 'excel-lite';
import type { CellValue, RowInput } from 'excel-lite';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const tmp = mkdtempSync(join(tmpdir(), 'excel-lite-writer-'));
const temp = (name: string) => join(tmp, name);

async function collect<T>(iterable: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const item of iterable) out.push(item);
  return out;
}

/** Load our bytes with ExcelJS so the assertions are made by an independent reader. */
async function load(bytes: Uint8Array | string): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  if (typeof bytes === 'string') await workbook.xlsx.readFile(bytes);
  else await workbook.xlsx.load(Buffer.from(bytes));
  return workbook;
}

async function sheetOf(bytes: Uint8Array | string, index = 1): Promise<ExcelJS.Worksheet> {
  return (await load(bytes)).getWorksheet(index)!;
}

interface ReadOptionsLite {
  header?: boolean;
  sheet?: string | number;
}

async function readAll(path: string, options?: ReadOptionsLite): Promise<CellValue[][]> {
  return (await collect(readSheet(path, options as never))) as CellValue[][];
}

describe('sheet names', () => {
  it('writes multiple sheets in the given order', async () => {
    const path = temp('multi.xlsx');
    await writeWorkbook(path, [
      { name: 'First', rows: [['a', 1]] },
      { name: 'Second', rows: [['b', 2]] },
      { name: 'Third', rows: [[3]] },
    ]);
    expect((await listSheets(path)).map((sheet) => sheet.name)).toEqual(['First', 'Second', 'Third']);
    expect(await readAll(path, { sheet: 'Second' })).toEqual([['b', 2]]);
  });

  it('renames duplicate sheet names case-insensitively', async () => {
    const names = (
      await listSheets(
        toXlsx([
          { name: 'Data', rows: [[1]] },
          { name: 'data', rows: [[2]] },
          { name: 'DATA', rows: [[3]] },
        ]),
      )
    ).map((sheet) => sheet.name);
    expect(names).toEqual(['Data', 'data (2)', 'DATA (3)']);
  });

  it('strips characters Excel rejects and truncates to 31 characters', async () => {
    const names = (
      await listSheets(
        toXlsx([
          { name: 'a/b*c?', rows: [[1]] },
          { name: 'a/b*c?', rows: [[2]] },
          { name: 'x'.repeat(40), rows: [[3]] },
        ]),
      )
    ).map((sheet) => sheet.name);
    expect(names).toEqual(['a b c', 'a b c (2)', 'x'.repeat(31)]);
  });

  it('falls back to Sheet / Sheet2 for empty and missing names', async () => {
    const names = (await listSheets(toXlsx([{ name: '', rows: [[1]] }, { rows: [[2]] }]))).map((sheet) => sheet.name);
    // An empty name becomes "Sheet"; a sheet with no name gets the positional default.
    expect(names).toEqual(['Sheet', 'Sheet2']);
  });

  it('keeps unicode sheet names intact', async () => {
    const workbook = await load(toXlsx([{ name: 'Résumé 2024', rows: [[1]] }]));
    expect(workbook.worksheets[0]?.name).toBe('Résumé 2024');
  });
});

describe('rows and columns', () => {
  it('writes column headers by default when columns are given', async () => {
    const sheet = await sheetOf(
      toXlsx([{ rows: [{ a: 1, b: 2 }], columns: [{ key: 'a', header: 'A label' }, { key: 'b' }] }]),
    );
    expect(sheet.getCell('A1').value).toBe('A label');
    expect(sheet.getCell('B1').value).toBe('b');
    expect(sheet.getRow(2).values).toEqual([undefined, 1, 2]);
  });

  it('writes no header when header is false', async () => {
    const bytes = toXlsx([{ rows: [[10, 'x']], columns: [{ key: 'a' }, { key: 'b' }], header: false }]);
    const sheet = await sheetOf(bytes);
    expect(sheet.getCell('A1').value).toBe(10);
    expect(sheet.getCell('B1').value).toBe('x');
    const path = temp('noheader.xlsx');
    writeFileSync(path, bytes);
    expect(await readAll(path)).toEqual([[10, 'x']]);
  });

  it('writes array rows in the column order given', async () => {
    const path = temp('arrays.xlsx');
    const bytes = await writeSheet(
      path,
      [
        [1, 'one'],
        [2, 'two'],
      ],
      {
        columns: [{ key: 'id' }, { key: 'label' }],
      },
    );
    expect(await readAll(path, { header: true })).toEqual([
      { id: 1, label: 'one' },
      { id: 2, label: 'two' },
    ] as never);
    expect(await (await load(bytes)).getWorksheet(1)!.getRow(1).values).toEqual([undefined, 'id', 'label']);
  });

  it('writes mixed object and array rows against the same columns', async () => {
    const rows: RowInput[] = [[1, 'one'], { id: 2, label: 'two' }];
    const path = temp('mixed.xlsx');
    await writeSheet(path, rows, { columns: [{ key: 'id' }, { key: 'label' }] });
    expect(await readAll(path, { header: false })).toEqual([
      ['id', 'label'],
      [1, 'one'],
      [2, 'two'],
    ]);
  });

  it('derives columns from the first object row when none are given', async () => {
    const path = temp('derived.xlsx');
    await writeSheet(path, [
      { b: 1, a: 2 },
      { a: 3, b: 4 },
    ]);
    expect(await readAll(path, { header: true })).toEqual([
      { b: 1, a: 2 },
      { b: 4, a: 3 },
    ] as never);
  });

  it('keeps cells past the declared columns for array rows', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [[1, 2, 3, 4]], columns: [{ key: 'a' }, { key: 'b' }] }]));
    expect(sheet.getRow(2).values).toEqual([undefined, 1, 2, 3, 4]);
    expect(sheet.autoFilter).toBe('A1:D1');
  });

  it('skips null and undefined but keeps interior holes', async () => {
    const path = temp('nulls.xlsx');
    await writeSheet(path, [['a', null, 'c'], [undefined, undefined], ['d']], { header: false });
    expect(await readAll(path)).toEqual([['a', null, 'c'], ['d']]);
    const sheet = await sheetOf(path);
    expect(sheet.getCell('B2').value).toBeNull();
  });

  it('does not write empty strings', async () => {
    const path = temp('empty-strings.xlsx');
    await writeSheet(path, [['a', '', 'c']], { header: false });
    // An empty string has no cell in the XML, so it reads back as a missing value.
    expect(await readAll(path)).toEqual([['a', null, 'c']]);
    expect((await sheetOf(path)).getCell('B1').value).toBeNull();
  });
});

describe('number formats', () => {
  it('applies custom number formats to data cells', async () => {
    const path = temp('numfmt.xlsx');
    await writeSheet(path, [{ amount: 1234.5, ratio: 0.125 }], {
      header: false,
      columns: [
        { key: 'amount', format: '#,##0.00' },
        { key: 'ratio', format: '0.00%' },
      ],
    });
    const sheet = await sheetOf(path);
    expect(sheet.getCell('A1').numFmt).toBe('#,##0.00');
    expect(sheet.getCell('B1').numFmt).toBe('0.00%');
    expect(sheet.getCell('A1').value).toBe(1234.5);
  });

  it('keeps quoted literals in a format and does not read them back as dates', async () => {
    const path = temp('quoted-format.xlsx');
    await writeSheet(path, [[42]], {
      header: false,
      columns: [{ key: 'kg', format: '"kg"#,##0' }],
    });
    expect((await sheetOf(path)).getCell('A1').numFmt).toBe('"kg"#,##0');
    expect(await readAll(path)).toEqual([[42]]);
  });

  it('uses a date format for Dates without a time part', async () => {
    const path = temp('dates-only.xlsx');
    const when = new Date(Date.UTC(2024, 4, 17));
    await writeSheet(path, [[when]], { header: false });
    const sheet = await sheetOf(path);
    expect(sheet.getCell('A1').numFmt).toBe('yyyy-mm-dd');
    expect(sheet.getCell('A1').value).toEqual(when);
    expect(await readAll(path)).toEqual([[when]]);
  });

  it('uses a date-time format for Dates with a time part', async () => {
    const path = temp('dates-time.xlsx');
    const when = new Date(Date.UTC(2024, 0, 1, 13, 45, 30));
    await writeSheet(path, [[when]], { header: false });
    const sheet = await sheetOf(path);
    expect(sheet.getCell('A1').numFmt).toBe('yyyy-mm-dd hh:mm:ss');
    expect(await readAll(path)).toEqual([[when]]);
  });

  it('honours an explicit date-like column format', async () => {
    const path = temp('dates-custom.xlsx');
    const when = new Date(Date.UTC(2024, 4, 17));
    await writeSheet(path, [{ when }], {
      header: false,
      columns: [{ key: 'when', format: 'dd/mm/yyyy' }],
    });
    expect((await sheetOf(path)).getCell('A1').numFmt).toBe('dd/mm/yyyy');
    expect(await readAll(path)).toEqual([[when]]);
  });

  it('does not let a non-date column format override a Date cell', async () => {
    const when = new Date(Date.UTC(2024, 4, 17));
    const sheet = await sheetOf(
      toXlsx([{ rows: [[when]], columns: [{ key: 'when', format: '#,##0.00' }], header: false }]),
    );
    expect(sheet.getCell('A1').numFmt).toBe('yyyy-mm-dd');
  });
});

describe('cell styling', () => {
  it('applies bold, solid fill and horizontal alignment', async () => {
    const sheet = await sheetOf(
      toXlsx([
        {
          header: false,
          rows: [['bold', 'red', 'centered']],
          columns: [
            { key: 'a', bold: true },
            { key: 'b', fill: '#ff0000' },
            { key: 'c', align: 'center' },
          ],
        },
      ]),
    );
    expect(sheet.getCell('A1').font?.bold).toBe(true);
    expect(sheet.getCell('B1').fill).toMatchObject({
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF0000' },
    });
    expect(sheet.getCell('C1').alignment?.horizontal).toBe('center');
    // Untouched cells keep the default (no) fill.
    expect(sheet.getCell('A1').fill).toMatchObject({ pattern: 'none' });
  });

  it('accepts AARRGGBB fill colours and a leading hash', async () => {
    const sheet = await sheetOf(
      toXlsx([
        {
          header: false,
          rows: [[1, 2]],
          columns: [
            { key: 'a', fill: '80FF0000' },
            { key: 'b', fill: '#00FF00' },
          ],
        },
      ]),
    );
    expect(sheet.getCell('A1').fill).toMatchObject({ fgColor: { argb: '80FF0000' } });
    expect(sheet.getCell('B1').fill).toMatchObject({ fgColor: { argb: 'FF00FF00' } });
  });

  it('writes the header row in bold', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [{ a: 1 }], columns: [{ key: 'a' }] }]));
    expect(sheet.getCell('A1').font?.bold).toBe(true);
  });

  it('rejects a malformed fill colour', () => {
    expect(() => toXlsx([{ rows: [[1]], columns: [{ key: 'a', fill: 'red' }] }])).toThrow(/invalid colour/);
  });
});

describe('layout', () => {
  it('writes explicit column widths', async () => {
    const sheet = await sheetOf(
      toXlsx([
        {
          rows: [[1, 2]],
          columns: [
            { key: 'a', width: 30 },
            { key: 'b', width: 12 },
          ],
        },
      ]),
    );
    expect(sheet.getColumn(1).width).toBe(30);
    expect(sheet.getColumn(2).width).toBe(12);
  });

  it('computes widths from the content when they are not given', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [['hello world', 1]] }]));
    expect(sheet.getColumn(1).width).toBe(13); // 11 characters + 2
    expect(sheet.getColumn(2).width).toBe(8); // minimum width
  });

  it('skips width computation with autoWidth: false', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [['hello world', 1]], autoWidth: false }]));
    expect(sheet.getColumn(1).width).toBeUndefined();
  });

  it('freezes the top row from "A2"', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [[1]], freeze: 'A2' }]));
    expect(sheet.views[0]).toMatchObject({ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2' });
  });

  it('freezes both axes from "B2"', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [[1, 2]], freeze: 'B2' }]));
    expect(sheet.views[0]).toMatchObject({ state: 'frozen', xSplit: 1, ySplit: 1, topLeftCell: 'B2' });
  });

  it('accepts the object freeze form', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [[1, 2]], freeze: { row: 1, col: 2 } }]));
    expect(sheet.views[0]).toMatchObject({ state: 'frozen', xSplit: 2, ySplit: 1, topLeftCell: 'C2' });
  });

  it('writes no pane for "A1" and rejects malformed references', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [[1]], freeze: 'A1' }]));
    expect(sheet.views[0]?.state).toBe('normal');
    expect(() => toXlsx([{ rows: [[1]], freeze: 'A' }])).toThrow(/freeze must look like "A2"/);
  });

  it('adds an auto-filter when a header is written', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [{ a: 1, b: 2 }], columns: [{ key: 'a' }, { key: 'b' }] }]));
    expect(sheet.autoFilter).toBe('A1:B1');
  });

  it('omits the auto-filter when autoFilter is false', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [{ a: 1 }], columns: [{ key: 'a' }], autoFilter: false }]));
    expect(sheet.autoFilter).toBeUndefined();
  });

  it('writes no auto-filter without a header row', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [[1, 2]], header: false }]));
    expect(sheet.autoFilter).toBeUndefined();
  });
});

describe('value handling', () => {
  it('writes booleans as boolean cells', async () => {
    const path = temp('bools.xlsx');
    await writeSheet(path, [[true, false]], { header: false });
    const sheet = await sheetOf(path);
    expect(sheet.getCell('A1').value).toBe(true);
    expect(sheet.getCell('B1').value).toBe(false);
    expect(await readAll(path)).toEqual([[true, false]]);
  });

  it('truncates strings longer than 32,767 characters', async () => {
    const path = temp('long.xlsx');
    const long = 'x'.repeat(40_000);
    await writeSheet(path, [[long]], { header: false });
    const sheet = await sheetOf(path);
    expect(String(sheet.getCell('A1').value)).toHaveLength(32_767);
    expect((await readAll(path))[0]?.[0]).toHaveLength(32_767);
  });

  it('strips XML control characters but keeps tab, newline and carriage return', async () => {
    const path = temp('controls.xlsx');
    await writeSheet(path, [['a\u0000b\u0001c\u0008d\te\nf\rg']], { header: false });
    expect(await readAll(path)).toEqual([['abcd\te\nf\rg']]);
  });

  it('preserves unicode and emoji', async () => {
    const path = temp('unicode.xlsx');
    const text = '漢字 😀 ünïcode ✓';
    await writeSheet(path, [[text]], { header: false });
    expect(await readAll(path)).toEqual([[text]]);
    expect((await sheetOf(path)).getCell('A1').value).toBe(text);
  });

  it('writes a row of 200 columns', async () => {
    const path = temp('wide.xlsx');
    const wide = Array.from({ length: 200 }, (_, index) => index);
    await writeSheet(path, [wide], { header: false });
    const rows = await readAll(path);
    expect(rows[0]).toHaveLength(200);
    expect(rows[0]?.[199]).toBe(199);
  });

  it('writes an empty sheet when the row set is empty', async () => {
    const path = temp('empty.xlsx');
    await writeWorkbook(path, [{ name: 'Empty', rows: [] }]);
    expect(await listSheets(path)).toEqual([{ name: 'Empty', index: 0, rowCount: 0 }]);
    expect(await readAll(path)).toEqual([]);
    expect((await sheetOf(path)).actualRowCount).toBe(0);
  });

  it('does not coerce strings that look like numbers, booleans or dates', async () => {
    const path = temp('stringy.xlsx');
    await writeSheet(path, [['42', 'true', '2024-05-17', '007']], { header: false });
    expect(await readAll(path)).toEqual([['42', 'true', '2024-05-17', '007']]);

    const sheet = await sheetOf(path);
    expect(typeof sheet.getCell('A1').value).toBe('string');
    expect(typeof sheet.getCell('B1').value).toBe('string');
    expect(typeof sheet.getCell('D1').value).toBe('string');
  });

  it('writes finite numbers as numeric cells, not text', async () => {
    const sheet = await sheetOf(toXlsx([{ rows: [[42, -3.25]], header: false }]));
    expect(typeof sheet.getCell('A1').value).toBe('number');
    expect(typeof sheet.getCell('B1').value).toBe('number');
  });

  it('round-trips very large and very small floats', async () => {
    const path = temp('floats.xlsx');
    const values = [1e308, 5e-324, -1e308, 1 / 3, Number.MAX_SAFE_INTEGER];
    await writeSheet(path, [values], { header: false });
    const [row] = await readAll(path);
    expect(row).toEqual(values);
  });

  it('returns the bytes it wrote', async () => {
    const path = temp('bytes.xlsx');
    const bytes = await writeSheet(path, [['a', 1]], { name: 'Bytes' });
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(Buffer.from(bytes).equals(readFileSync(path))).toBe(true);

    const multi = await writeWorkbook(temp('bytes-multi.xlsx'), [{ rows: [[1]] }, { rows: [[2]] }]);
    expect(multi).toBeInstanceOf(Uint8Array);
  });
});

process.on('exit', () => rmSync(tmp, { recursive: true, force: true }));
