import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { listSheets, readSheet, readWorkbook } from 'excel-lite';
import type { CellValue, RowObject } from 'excel-lite';

const fixture = (name: string) => join(process.cwd(), 'fixtures', name);

async function collect<T>(iterable: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const item of iterable) out.push(item);
  return out;
}

describe('files written by openpyxl', () => {
  it('reads every value type, skipping the blank row', async () => {
    const rows = await collect(readSheet(fixture('openpyxl-mixed.xlsx'), { header: true }));
    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual({
      name: 'Ali',
      qty: 3,
      price: 19.99,
      when: new Date(Date.UTC(2023, 6, 4)),
      active: true,
      note: 'plain text',
    });
    expect(rows[1]).toEqual({
      name: 'Ünïcode ✓',
      qty: 0,
      price: 0.5,
      when: new Date(Date.UTC(2022, 10, 30, 13, 45, 30)),
      active: false,
      note: 'comma, "quote"',
    });
    expect(rows[2]?.note).toBe('😀 emoji');
    expect(rows[2]?.when).toEqual(new Date(Date.UTC(2024, 1, 29)));
  });

  it('exposes both sheets with the right names', async () => {
    const workbook = await readWorkbook(fixture('openpyxl-mixed.xlsx'));
    expect(workbook.sheetNames).toEqual(['Data', 'Notes']);
    const notes = await collect(workbook.sheet('Notes'));
    expect(notes).toEqual([
      ['key', 'value'],
      ['owner', 'finance'],
      ['updated', new Date(Date.UTC(2024, 0, 15))],
    ]);
  });

  it('counts rows without parsing cells', async () => {
    expect(await listSheets(fixture('openpyxl-mixed.xlsx'))).toEqual([
      { name: 'Data', index: 0, rowCount: 4 },
      { name: 'Notes', index: 1, rowCount: 3 },
    ]);
  });

  it('turns custom date/time formats into Date but leaves money and percent numeric', async () => {
    const rows = (await collect(readSheet(fixture('openpyxl-formats.xlsx'), { header: true }))) as RowObject[];
    expect(rows[0]).toEqual({
      label: 'a',
      money: 1234.5,
      pct: 0.125,
      date: new Date(Date.UTC(2020, 0, 1)),
      datetime: new Date(Date.UTC(2020, 0, 1, 6, 30)),
      sparse: 'x',
    });
    // 1999-12-31 is serial 36525, well past the 1900 leap-year bug window.
    expect(rows[1]?.date).toEqual(new Date(Date.UTC(1999, 11, 31)));
    expect(rows[1]?.sparse).toBeNull();
    expect(rows[2]?.datetime).toBeNull();
  });

  it('accepts the file as bytes and as a stream', async () => {
    const bytes = new Uint8Array(readFileSync(fixture('openpyxl-mixed.xlsx')));
    const rows = (await collect(readSheet(bytes))) as CellValue[][];
    expect(rows[0]?.slice(0, 2)).toEqual(['name', 'qty']);

    async function* stream() {
      for (let offset = 0; offset < bytes.length; offset += 1024) yield bytes.subarray(offset, offset + 1024);
    }
    const streamed = (await collect(readSheet(stream(), { header: true }))) as RowObject[];
    expect(streamed).toHaveLength(3);
  });
});
