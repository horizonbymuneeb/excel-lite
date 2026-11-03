import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { listSheets, readSheet, readWorkbook, toXlsx } from 'excel-lite';
import type { CellValue, RowObject } from 'excel-lite';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const tmp = mkdtempSync(join(tmpdir(), 'excel-lite-reader-'));
const fixture = (name: string) => join(process.cwd(), 'fixtures', name);
const temp = (name: string) => join(tmp, name);

async function collect<T>(iterable: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const item of iterable) out.push(item);
  return out;
}

async function writeExcelJs(name: string, build: (workbook: ExcelJS.Workbook) => void): Promise<string> {
  const workbook = new ExcelJS.Workbook();
  build(workbook);
  const path = temp(name);
  await workbook.xlsx.writeFile(path);
  return path;
}

describe('cell types', () => {
  it('reads shared strings, including repeated string-table indexes', async () => {
    const path = await writeExcelJs('shared.xlsx', (workbook) => {
      const sheet = workbook.addWorksheet('S');
      sheet.getCell('A1').value = 'shared';
      sheet.getCell('A2').value = 'shared';
      sheet.getCell('B1').value = 'other';
    });
    expect(await collect(readSheet(path))).toEqual([
      ['shared', 'other'],
      ['shared'],
    ]);
  });

  it('concatenates every run of a rich-text shared string', async () => {
    const path = await writeExcelJs('rich.xlsx', (workbook) => {
      const sheet = workbook.addWorksheet('S');
      sheet.getCell('A1').value = {
        richText: [{ text: 'Bold ' }, { text: 'plain ' }, { font: { italic: true }, text: 'italic' }],
      };
    });
    expect(await collect(readSheet(path))).toEqual([['Bold plain italic']]);
  });

  it('reads inline strings, joining inline rich-text runs', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-shapes.xlsx')))) as CellValue[][];
    expect(rows[0]?.[0]).toBe('Bold plain italic');
    expect(rows[0]?.[1]).toBe('plain');

    const handmade = (await collect(readSheet(fixture('handmade-cells.xlsx')))) as CellValue[][];
    expect(handmade[0]?.[2]).toBe('inline');
    expect(handmade[4]).toEqual(['ab']);
  });

  it('returns cached formula results, not formula text', async () => {
    const path = await writeExcelJs('formula.xlsx', (workbook) => {
      const sheet = workbook.addWorksheet('S');
      sheet.getCell('A1').value = 'a';
      sheet.getCell('B1').value = { formula: 'CONCATENATE(A1,"b")', result: 'ab' };
      sheet.getCell('C1').value = { formula: '1+1', result: 2 };
    });
    expect(await collect(readSheet(path))).toEqual([['a', 'ab', 2]]);

    const handmade = (await collect(readSheet(fixture('handmade-cells.xlsx')))) as CellValue[][];
    expect(handmade[0]?.[3]).toBe('cached');
  });

  it('returns error cells as their string code', async () => {
    const shapes = (await collect(readSheet(fixture('openpyxl-shapes.xlsx')))) as CellValue[][];
    expect(shapes[0]?.[2]).toBe('#DIV/0!');

    const handmade = (await collect(readSheet(fixture('handmade-cells.xlsx')))) as CellValue[][];
    expect(handmade[0]?.[4]).toBe('#N/A');
  });

  it('parses t="d" ISO date cells, with and without a UTC designator', async () => {
    const rows = (await collect(readSheet(fixture('handmade-cells.xlsx')))) as CellValue[][];
    const expected = new Date(Date.UTC(2024, 4, 17, 12, 34, 56));
    expect(rows[0]?.[5]).toEqual(expected);
    // 2024-05-17T18:04:56+05:30 is the same instant, expressed with an offset.
    expect(rows[0]?.[6]).toEqual(expected);
  });

  it('reads booleans written as 1/0 and as "true"/"false"', async () => {
    const rows = (await collect(readSheet(fixture('handmade-cells.xlsx')))) as CellValue[][];
    expect(rows[0]?.[7]).toBe(true);
    expect(rows[0]?.[8]).toBe(false);
    expect(rows[0]?.[9]).toBe(true);

    const shapes = (await collect(readSheet(fixture('openpyxl-shapes.xlsx')))) as CellValue[][];
    expect(shapes[0]?.[3]).toBe(true);
  });

  it('decodes XML entities in shared strings and inline strings', async () => {
    const rows = (await collect(readSheet(fixture('handmade-cells.xlsx')))) as CellValue[][];
    expect(rows[0]?.[10]).toBe('a & b < c > d');
    const shapes = (await collect(readSheet(fixture('openpyxl-shapes.xlsx')))) as CellValue[][];
    expect(shapes[1]).toEqual(['x', null, 'z']);
  });

  it('reads numbers stored as text as strings and missing string-table indexes as null', async () => {
    const rows = (await collect(readSheet(fixture('handmade-cells.xlsx')))) as CellValue[][];
    // L1 points past the end of the shared-string table, M1 has an empty <v>.
    expect(rows[0]?.[11]).toBeNull();
    expect(rows[0]?.[12]).toBeNull();
    expect(rows[0]?.[13]).toBe(42);
  });
});

describe('dates', () => {
  it('reads a 1904 date-system workbook as UTC dates', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-1904.xlsx'), { header: true }))) as RowObject[];
    expect(rows[0]).toEqual({
      label: 'aligned',
      date: new Date(Date.UTC(2024, 0, 2)),
      datetime: new Date(Date.UTC(2024, 0, 2, 3, 4, 5)),
      time: new Date(Date.UTC(1904, 0, 1, 6, 30)),
    });
    // Serial 0 in the 1904 system is the epoch day itself.
    expect(rows[1]).toEqual({
      label: 'epoch',
      date: new Date(Date.UTC(1904, 0, 1)),
      datetime: new Date(Date.UTC(1904, 0, 1)),
      time: new Date(Date.UTC(1904, 0, 1)),
    });
  });

  it('applies the 1900 leap-year correction across the boundary serials', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-serial-dates.xlsx'), { header: true }))) as RowObject[];
    expect(rows.map((row) => [row.serial, row.as_date])).toEqual([
      [0, new Date(Date.UTC(1899, 11, 30))],
      [1, new Date(Date.UTC(1900, 0, 1))],
      [2, new Date(Date.UTC(1900, 0, 2))],
      [59, new Date(Date.UTC(1900, 1, 28))],
      // Excel's phantom 1900-02-29 (serial 60) and serial 61 both land on 1900-03-01.
      [60, new Date(Date.UTC(1900, 2, 1))],
      [61, new Date(Date.UTC(1900, 2, 1))],
      [36525, new Date(Date.UTC(1999, 11, 31))],
    ]);
  });

  it('anchors time-only cells at the epoch of the workbook date system', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-nondates.xlsx'), { header: true }))) as RowObject[];
    expect(rows[0]?.time).toEqual(new Date(Date.UTC(1899, 11, 30, 6, 30)));
    expect(rows[0]?.elapsed).toEqual(new Date(Date.UTC(1899, 11, 30, 18, 45)));
    expect(rows[1]?.time).toEqual(new Date(Date.UTC(1899, 11, 30, 0, 0)));
    expect(rows[1]?.elapsed).toEqual(new Date(Date.UTC(1899, 11, 30, 1, 30)));
  });

  it('returns dates regardless of the process time zone', async () => {
    // The reader builds Dates from UTC arithmetic, so `Date#getTime()` is stable.
    const rows = (await collect(readSheet(fixture('openpyxl-mixed.xlsx'), { header: true }))) as RowObject[];
    expect((rows[0]?.when as Date).getTime()).toBe(Date.UTC(2023, 6, 4));
    expect((rows[1]?.when as Date).getTime()).toBe(Date.UTC(2022, 10, 30, 13, 45, 30));
  });
});

describe('number formats', () => {
  it('keeps currency, percent, text, general and quoted formats numeric', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-nondates.xlsx'), { header: true }))) as RowObject[];
    expect(rows[0]).toMatchObject({
      money: 1234.5,
      pct: 0.125,
      text: 12345,
      general: 42,
      quoted: 42,
      red: 42,
      literal: 42,
    });
    for (const key of ['money', 'pct', 'text', 'general', 'quoted', 'red', 'literal']) {
      expect(typeof rows[0]?.[key]).toBe('number');
    }
    expect(rows[1]).toMatchObject({ money: -20, pct: 1, text: 7, general: -3.25, quoted: 99, red: -1, literal: 3 });
  });

  it('does not turn numeric formats with quoted date letters into dates', async () => {
    // '"days"0' and '"kg"#,##0' contain y/m/d only inside quoted literals.
    const rows = (await collect(readSheet(fixture('openpyxl-nondates.xlsx'), { header: true }))) as RowObject[];
    expect(rows[0]?.literal).toBe(42);
    expect(rows[1]?.literal).toBe(3);
  });

  it('turns [h]:mm:ss elapsed-time formats into Dates', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-nondates.xlsx'), { header: true }))) as RowObject[];
    expect(rows[0]?.elapsed).toBeInstanceOf(Date);
    expect(rows[1]?.elapsed).toBeInstanceOf(Date);
  });
});

describe('row shapes', () => {
  it('fills sparse rows with null instead of undefined', async () => {
    const shapes = (await collect(readSheet(fixture('openpyxl-shapes.xlsx')))) as CellValue[][];
    expect(shapes[1]).toEqual(['x', null, 'z']);
    expect(shapes[1]?.includes(undefined as unknown as CellValue)).toBe(false);

    const handmade = (await collect(readSheet(fixture('handmade-cells.xlsx')))) as CellValue[][];
    expect(handmade[1]).toEqual([1, null, 3]);
  });

  it('trims trailing empty cells', async () => {
    const rows = (await collect(readSheet(fixture('handmade-cells.xlsx')))) as CellValue[][];
    expect(rows[2]).toEqual(['shared']); // A3 plus empty B3/C3
    expect(rows[3]).toEqual([5]); // A5 plus empty B5/C5
    expect(rows[5]).toEqual(['shared']); // A7 plus empty B7
  });

  it('skips rows that contain no values at all', async () => {
    const rows = await collect(readSheet(fixture('handmade-cells.xlsx')));
    // Seven <row> elements, but row 4 is completely empty and must not be yielded.
    expect(rows).toHaveLength(6);
    expect(rows[3]).toEqual([5]);
  });

  it('reads a row of 200 columns without losing cells', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-wide.xlsx'), { header: true }))) as RowObject[];
    expect(rows).toHaveLength(2);
    expect(Object.keys(rows[0]!)).toHaveLength(200);
    expect(rows[0]?.c0).toBe(0);
    expect(rows[0]?.c199).toBe(199);
    expect(rows[1]?.c199).toBe(288.5);
  });

  it('exposes the top-left value of a merged range only', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-shapes.xlsx')))) as CellValue[][];
    expect(rows[2]).toEqual(['merged block']);
  });
});

describe('sheets', () => {
  it('looks a sheet up by zero-based index', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-mixed.xlsx'), { sheet: 1 }))) as CellValue[][];
    expect(rows[0]).toEqual(['key', 'value']);
    expect(rows[2]?.[0]).toBe('updated');
  });

  it('looks a sheet up by exact name', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-mixed.xlsx'), { sheet: 'Notes' }))) as CellValue[][];
    expect(rows).toHaveLength(3);
  });

  it('falls back to a case-insensitive name lookup', async () => {
    for (const name of ['notes', 'NOTES', 'nOtEs']) {
      const rows = (await collect(readSheet(fixture('openpyxl-mixed.xlsx'), { sheet: name }))) as CellValue[][];
      expect(rows[0]).toEqual(['key', 'value']);
    }
  });

  it('lists the available sheets on an unknown name', async () => {
    await expect(collect(readSheet(fixture('openpyxl-mixed.xlsx'), { sheet: 'Nope' }))).rejects.toThrow(
      /sheet "Nope" not found — available: Data, Notes/,
    );
  });

  it('rejects an out-of-range sheet index', async () => {
    await expect(collect(readSheet(fixture('openpyxl-mixed.xlsx'), { sheet: 2 }))).rejects.toThrow(
      /sheet index 2 is out of range \(workbook has 2 sheet\(s\)\)/,
    );
    await expect(collect(readSheet(fixture('openpyxl-mixed.xlsx'), { sheet: -1 }))).rejects.toThrow(RangeError);
  });

  it('stacks every sheet through readWorkbook().sheets()', async () => {
    const workbook = await readWorkbook(fixture('openpyxl-mixed.xlsx'));
    expect(workbook.sheetNames).toEqual(['Data', 'Notes']);
    const stacked: number[] = [];
    for (const generator of workbook.sheets()) {
      const rows = await collect(generator);
      stacked.push(rows.length);
    }
    expect(stacked).toEqual([4, 3]);
  });

  it('lets the same workbook sheet be iterated more than once', async () => {
    const workbook = await readWorkbook(fixture('openpyxl-mixed.xlsx'));
    const first = await collect(workbook.sheet('Notes', { header: true }));
    const second = await collect(workbook.sheet('Notes', { header: true }));
    expect(first).toHaveLength(2);
    expect(second).toEqual(first);
    expect(first[0]).toEqual({ key: 'owner', value: 'finance' });
  });

  it('counts physical rows for listSheets but skips empty ones while iterating', async () => {
    expect(await listSheets(fixture('handmade-cells.xlsx'))).toEqual([
      // Raw has 7 <row> elements (one empty); Rows has 3 (one empty, one numbering gap).
      { name: 'Raw', index: 0, rowCount: 7 },
      { name: 'Rows', index: 1, rowCount: 3 },
    ]);
    const raw = await collect(readSheet(fixture('handmade-cells.xlsx'), { sheet: 0 }));
    const rows = await collect(readSheet(fixture('handmade-cells.xlsx'), { sheet: 'Rows' }));
    expect(raw).toHaveLength(6);
    expect(rows).toEqual([[1], [4]]);
  });

  it('does not count <rowBreaks> as rows', async () => {
    // The Raw sheet has a rowBreaks element; rowCount must stay 7, not 8.
    const sheets = await listSheets(fixture('handmade-cells.xlsx'));
    expect(sheets[0]?.rowCount).toBe(7);
  });

  it('reports the sheet names of a workbook', async () => {
    const workbook = await readWorkbook(fixture('openpyxl-1904.xlsx'));
    expect(workbook.sheetNames).toEqual(['Mac']);
  });
});

describe('header mode', () => {
  it('handles duplicate, blank and numeric headers plus extra columns', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-headers.xlsx'), { header: true }))) as RowObject[];
    expect(rows[0]).toEqual({
      name: 'a',
      name_2: 'b',
      '123': 'c',
      column_4: 'd',
      name_3: 'e',
      // Row 2 is one column longer than the header row.
      column_6: 'f',
    });
    expect(rows[1]).toEqual({ name: 'g', name_2: null, '123': null, column_4: null, name_3: null });
  });

  it('maps short rows onto the header with nulls', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-headers.xlsx'), { header: true }))) as RowObject[];
    expect(rows[1]?.name).toBe('g');
    expect(rows[1]?.name_3).toBeNull();
  });

  it('yields no objects for a sheet with only a header row', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-wide.xlsx'), { header: true }))) as RowObject[];
    expect(rows).toHaveLength(2);
  });
});

describe('input sources', () => {
  it('accepts a path, Buffer, Uint8Array, ArrayBuffer, Blob and an async byte stream', async () => {
    const path = fixture('openpyxl-mixed.xlsx');
    const expected = await collect(readSheet(path));

    const buffer = readFileSync(path);
    expect(await collect(readSheet(buffer))).toEqual(expected);

    const view = new Uint8Array(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength));
    expect(await collect(readSheet(view))).toEqual(expected);
    expect(await collect(readSheet(view.buffer as ArrayBuffer))).toEqual(expected);
    expect(await collect(readSheet(new Blob([buffer])))).toEqual(expected);

    async function* chunks() {
      for (let offset = 0; offset < buffer.length; offset += 700) {
        yield new Uint8Array(buffer.subarray(offset, offset + 700));
      }
    }
    expect(await collect(readSheet(chunks()))).toEqual(expected);
  });

  it('rejects source types it cannot turn into bytes', async () => {
    await expect(collect(readSheet(42 as never))).rejects.toThrow(TypeError);
    await expect(collect(readSheet(42 as never))).rejects.toThrow(/unsupported source/);
    await expect(collect(readSheet({} as never))).rejects.toThrow(/unsupported source/);
  });

  it('reports a missing file with the Node error', async () => {
    await expect(collect(readSheet(join(tmp, 'does-not-exist.xlsx')))).rejects.toThrow(/ENOENT/);
  });
});

process.on('exit', () => rmSync(tmp, { recursive: true, force: true }));
