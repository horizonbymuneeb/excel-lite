import { describe, expect, it } from 'vitest';
import { execFile } from 'node:child_process';
import { createRequire } from 'node:module';
import { promisify } from 'node:util';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  WorkbookReader,
  columnIndexFromRef,
  columnLetter,
  dateToSerial,
  listSheets,
  readSheet,
  readWorkbook,
  serialToDate,
  toXlsx,
  writeSheet,
  writeWorkbook,
} from 'excel-lite';

const run = promisify(execFile);
const fixture = (name: string) => join(process.cwd(), 'fixtures', name);

const DOCUMENTED_EXPORTS = [
  'WorkbookReader',
  'columnIndexFromRef',
  'columnLetter',
  'dateToSerial',
  'listSheets',
  'readSheet',
  'readWorkbook',
  'serialToDate',
  'toXlsx',
  'writeSheet',
  'writeWorkbook',
];

const DOCUMENTED_TYPES = [
  'CellValue',
  'ColumnSpec',
  'ReadOptions',
  'RowInput',
  'RowObject',
  'SheetInfo',
  'SheetInput',
  'SheetRow',
  'Source',
  'WriteOptions',
];

describe('package surface', () => {
  it('exposes every documented runtime export as a function', () => {
    const actual = {
      WorkbookReader,
      columnIndexFromRef,
      columnLetter,
      dateToSerial,
      listSheets,
      readSheet,
      readWorkbook,
      serialToDate,
      toXlsx,
      writeSheet,
      writeWorkbook,
    };
    expect(Object.keys(actual).sort()).toEqual(DOCUMENTED_EXPORTS);
    for (const [name, value] of Object.entries(actual)) {
      expect(typeof value, `${name} should be a function`).toBe('function');
    }
  });

  it('resolves the same API through CJS require()', () => {
    const require = createRequire(import.meta.url);
    const cjs = require('excel-lite') as Record<string, unknown>;
    expect(Object.keys(cjs).sort()).toEqual(DOCUMENTED_EXPORTS);
    for (const name of DOCUMENTED_EXPORTS) expect(typeof cjs[name]).toBe('function');
  });

  it('works end to end from a real .cjs file', async () => {
    const { stdout } = await run(process.execPath, [join(process.cwd(), 'test', 'cjs-smoke.cjs')]);
    const result = JSON.parse(stdout) as {
      rows: unknown[];
      letter: string;
      index: number;
      exports: string[];
    };
    expect(result.rows).toEqual([['cjs', 1]]);
    expect(result.letter).toBe('AA');
    expect(result.index).toBe(16383);
    expect(result.exports).toEqual(DOCUMENTED_EXPORTS);
  });

  it('ships the documented types in the built declaration file', () => {
    const declarations = readFileSync(join(process.cwd(), 'dist', 'index.d.ts'), 'utf8');
    for (const name of DOCUMENTED_EXPORTS) {
      expect(declarations, `${name} missing from dist/index.d.ts`).toContain(name);
    }
    for (const name of DOCUMENTED_TYPES) {
      expect(declarations, `${name} missing from dist/index.d.ts`).toContain(`type ${name}`);
    }
    // readSheet keeps the overload that turns `header: true` into objects.
    expect(declarations).toContain('AsyncGenerator<RowObject>');
    expect(declarations).toContain('AsyncGenerator<CellValue[]>');
  });
});

describe('columnLetter', () => {
  it('maps zero-based indexes to Excel letters', () => {
    expect(columnLetter(0)).toBe('A');
    expect(columnLetter(1)).toBe('B');
    expect(columnLetter(25)).toBe('Z');
    expect(columnLetter(26)).toBe('AA');
    expect(columnLetter(27)).toBe('AB');
    expect(columnLetter(51)).toBe('AZ');
    expect(columnLetter(52)).toBe('BA');
    expect(columnLetter(701)).toBe('ZZ');
    expect(columnLetter(702)).toBe('AAA');
    expect(columnLetter(16383)).toBe('XFD');
  });
});

describe('columnIndexFromRef', () => {
  it('parses cell references, including lowercase ones', () => {
    expect(columnIndexFromRef('A1')).toBe(0);
    expect(columnIndexFromRef('B2')).toBe(1);
    expect(columnIndexFromRef('Z9')).toBe(25);
    expect(columnIndexFromRef('AA1')).toBe(26);
    expect(columnIndexFromRef('AB12')).toBe(27);
    expect(columnIndexFromRef('XFD1')).toBe(16383);
    expect(columnIndexFromRef('c12')).toBe(2);
    expect(columnIndexFromRef('ABC123')).toBe(730);
  });

  it('returns -1 when the reference has no column part', () => {
    expect(columnIndexFromRef('')).toBe(-1);
    expect(columnIndexFromRef('1')).toBe(-1);
    expect(columnIndexFromRef('7A')).toBe(-1);
  });

  it('is the inverse of columnLetter', () => {
    for (const index of [0, 1, 25, 26, 51, 52, 701, 702, 16383]) {
      expect(columnIndexFromRef(`${columnLetter(index)}1`)).toBe(index);
    }
  });
});

describe('serialToDate', () => {
  it('maps the 1900 boundary serials', () => {
    expect(serialToDate(1)).toEqual(new Date(Date.UTC(1900, 0, 1)));
    expect(serialToDate(59)).toEqual(new Date(Date.UTC(1900, 1, 28)));
    // Excel's phantom 1900-02-29 means serials 60 and 61 both describe 1900-03-01.
    expect(serialToDate(60)).toEqual(new Date(Date.UTC(1900, 2, 1)));
    expect(serialToDate(61)).toEqual(new Date(Date.UTC(1900, 2, 1)));
    expect(serialToDate(36525)).toEqual(new Date(Date.UTC(1999, 11, 31)));
    expect(serialToDate(45293)).toEqual(new Date(Date.UTC(2024, 0, 2)));
  });

  it('keeps the fractional part as UTC time of day', () => {
    expect(serialToDate(45293.5)).toEqual(new Date(Date.UTC(2024, 0, 2, 12)));
    expect(serialToDate(0.2708333333333333)).toEqual(new Date(Date.UTC(1899, 11, 30, 6, 30)));
  });

  it('uses the 1904 epoch when asked', () => {
    expect(serialToDate(0, true)).toEqual(new Date(Date.UTC(1904, 0, 1)));
    expect(serialToDate(1, true)).toEqual(new Date(Date.UTC(1904, 0, 2)));
    expect(serialToDate(43831, true)).toEqual(new Date(Date.UTC(2024, 0, 2)));
  });
});

describe('dateToSerial', () => {
  it('maps 1900-era dates to the serials Excel uses', () => {
    expect(dateToSerial(new Date(Date.UTC(1900, 0, 1)))).toBe(1);
    expect(dateToSerial(new Date(Date.UTC(1900, 1, 28)))).toBe(59);
    expect(dateToSerial(new Date(Date.UTC(2024, 0, 2)))).toBe(45293);
  });

  it('round-trips the phantom-day window without losing a day', () => {
    for (const day of [1, 2, 59, 60, 61]) {
      const date = serialToDate(day);
      expect(serialToDate(dateToSerial(date)).getTime()).toBe(date.getTime());
    }
  });

  // Serial 60 is Excel's phantom 1900-02-29 — a date JavaScript cannot represent, so it
  // reads as 1900-03-01 and writes back as 61. Every other serial must round-trip exactly.
  it('round-trips serials 1..40000 except the phantom 1900-02-29 slot', () => {
    const serials = new Set<number>([1, 2, 3, 4, 58, 59, 61, 62, 63, 36525, 40000]);
    for (let serial = 62; serial <= 40000; serial += 137) serials.add(serial);
    const failures: Array<[number, number]> = [];
    for (const serial of [...serials].sort((a, b) => a - b)) {
      const back = dateToSerial(serialToDate(serial));
      if (back !== serial) failures.push([serial, back]);
    }
    expect(failures).toEqual([]);
  });

  it('round-trips whole-day dates over a long range', () => {
    const failures: number[] = [];
    for (let serial = 62; serial <= 40000; serial += 997) {
      const date = serialToDate(serial);
      const back = serialToDate(dateToSerial(date));
      if (back.getTime() !== date.getTime()) failures.push(serial);
    }
    expect(failures).toEqual([]);
  });

  it('round-trips dates with a time part to the millisecond', () => {
    const dates = [
      new Date(Date.UTC(2024, 0, 2, 3, 4, 5)),
      new Date(Date.UTC(1999, 11, 31, 23, 59, 59)),
      new Date(Date.UTC(2035, 6, 4, 12, 0, 0)),
    ];
    for (const date of dates) {
      expect(Math.abs(serialToDate(dateToSerial(date)).getTime() - date.getTime())).toBeLessThan(1);
    }
  });
});

describe('builders', () => {
  it('toXlsx rejects an empty or non-array sheet list', () => {
    expect(() => toXlsx([])).toThrow(/at least one sheet/);
    expect(() => toXlsx('nope' as never)).toThrow(/at least one sheet/);
  });

  it('WorkbookReader exposes sheet names, resolution and row counts', async () => {
    const reader = await WorkbookReader.open(fixture('openpyxl-mixed.xlsx'));
    expect(reader.sheetNames).toEqual(['Data', 'Notes']);
    expect(reader.resolveSheet(undefined)).toBe(0);
    expect(reader.resolveSheet(1)).toBe(1);
    expect(reader.resolveSheet('notes')).toBe(1);
    expect(reader.rowCount(0)).toBe(4);
    expect(reader.rowCount(1)).toBe(3);
    const rows: string[] = [];
    for (const row of reader.rows(1, false)) rows.push(String((row as string[])[0]));
    expect(rows).toEqual(['key', 'owner', 'updated']);
  });

  it('returns the shapes documented for readWorkbook and listSheets', async () => {
    const workbook = await readWorkbook(fixture('openpyxl-mixed.xlsx'));
    expect(Object.keys(workbook).sort()).toEqual(['sheet', 'sheetNames', 'sheets']);
    expect(workbook.sheetNames).toEqual(['Data', 'Notes']);
    expect(await listSheets(fixture('openpyxl-mixed.xlsx'))).toEqual([
      { name: 'Data', index: 0, rowCount: 4 },
      { name: 'Notes', index: 1, rowCount: 3 },
    ]);
  });

  it('writeSheet and writeWorkbook are available with the documented arity', () => {
    expect(writeSheet.length).toBeLessThanOrEqual(3);
    expect(writeWorkbook.length).toBe(2);
  });
});
