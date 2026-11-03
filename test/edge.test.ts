import { describe, expect, it } from 'vitest';
import { zipSync } from 'fflate';
import { listSheets, readSheet, toXlsx, writeSheet } from 'excel-lite';
import type { CellValue, RowInput } from 'excel-lite';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const tmp = mkdtempSync(join(tmpdir(), 'excel-lite-edge-'));
const temp = (name: string) => join(tmp, name);
const encode = (text: string) => new TextEncoder().encode(text);

const WORKBOOK_HEADER =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

const WORKBOOK_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
  '</Relationships>';

const ONE_SHEET = `${WORKBOOK_HEADER}<sheets><sheet name="Only" sheetId="1" r:id="rId1"/></sheets></workbook>`;
const SHEET_WITH_ONE_CELL =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>';

function archive(files: Record<string, string>): Uint8Array {
  const encoded: Record<string, Uint8Array> = {};
  for (const [name, body] of Object.entries(files)) encoded[name] = encode(body);
  return zipSync(encoded);
}

async function collect<T>(iterable: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const item of iterable) out.push(item);
  return out;
}

describe('malformed workbooks', () => {
  it('rejects a zip whose xl/workbook.xml is missing', async () => {
    const bytes = archive({ 'xl/styles.xml': '<styleSheet/>' });
    await expect(collect(readSheet(bytes))).rejects.toThrow(/xl\/workbook\.xml is missing/);
  });

  it('rejects a workbook whose workbook.xml is not XML', async () => {
    const bytes = archive({ 'xl/workbook.xml': 'this is not xml' });
    await expect(collect(readSheet(bytes))).rejects.toThrow(/workbook\.xml is malformed/);
  });

  it('rejects a workbook with no sheets', async () => {
    const bytes = archive({ 'xl/workbook.xml': `${WORKBOOK_HEADER}<sheets/></workbook>` });
    await expect(collect(readSheet(bytes))).rejects.toThrow(/the workbook has no sheets/);
  });

  it('reports a sheet part that the archive does not contain', async () => {
    const bytes = archive({
      'xl/workbook.xml': ONE_SHEET,
      'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
    });
    await expect(collect(readSheet(bytes))).rejects.toThrow(/sheet "Only" is missing from the archive/);
    await expect(listSheets(bytes)).rejects.toThrow(/sheet "Only" is missing from the archive/);
  });

  it('falls back to xl/worksheets/sheet1.xml when the workbook has no relationships', async () => {
    const bytes = archive({
      'xl/workbook.xml': `${WORKBOOK_HEADER}<sheets><sheet name="Fallback" sheetId="1"/></sheets></workbook>`,
      'xl/worksheets/sheet1.xml': SHEET_WITH_ONE_CELL,
    });
    expect(await collect(readSheet(bytes))).toEqual([[7]]);
    expect((await listSheets(bytes))[0]?.name).toBe('Fallback');
  });

  it('rejects bytes that are not a zip archive', async () => {
    await expect(collect(readSheet(new Uint8Array([1, 2, 3, 4])))).rejects.toThrow(/not a zip/);
  });

  it('rejects a legacy binary file that was renamed to .xlsx', async () => {
    const path = temp('legacy.xlsx');
    writeFileSync(path, Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));
    await expect(collect(readSheet(path))).rejects.toThrow(/not a zip/);
  });
});

describe('writer corner cases', () => {
  it('stores non-finite numbers as text instead of emitting invalid XML', async () => {
    const path = temp('non-finite.xlsx');
    await writeSheet(path, [[Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]], { header: false });
    expect(await collect(readSheet(path))).toEqual([['NaN', 'Infinity', '-Infinity']]);
  });

  it('skips invalid Date objects', async () => {
    const path = temp('invalid-date.xlsx');
    await writeSheet(path, [[new Date(Number.NaN)]], { header: false });
    expect(await collect(readSheet(path))).toEqual([]);

    const withValue = temp('invalid-date-2.xlsx');
    await writeSheet(withValue, [['kept', new Date(Number.NaN)]], { header: false });
    expect(await collect(readSheet(withValue))).toEqual([['kept']]);
  });

  it('accepts any iterable of rows, not just arrays', async () => {
    function* rows(): Generator<RowInput> {
      yield ['a', 1];
      yield ['b', 2];
    }
    const path = temp('generator.xlsx');
    await writeSheet(path, rows(), { header: false });
    expect(await collect(readSheet(path))).toEqual([
      ['a', 1],
      ['b', 2],
    ]);
  });

  it('treats a whitespace-only sheet name like an empty one', async () => {
    const sheets = await listSheets(toXlsx([{ name: '   ', rows: [[1]] }]));
    expect(sheets[0]?.name).toBe('Sheet');
  });

  it('writes an empty sheet for an empty object row', async () => {
    const bytes = toXlsx([{ rows: [{}] }]);
    expect(await collect(readSheet(bytes))).toEqual([]);
  });
});

describe('header detection', () => {
  it('uses the first non-empty row as the header', async () => {
    const path = temp('first-non-empty.xlsx');
    await writeSheet(
      path,
      [
        [null, null],
        ['name', 'age'],
        ['Ali', 27],
      ],
      { header: false },
    );
    const rows = (await collect(readSheet(path, { header: true }))) as Array<Record<string, CellValue>>;
    expect(rows).toEqual([{ name: 'Ali', age: 27 }]);
  });

  it('skips fully blank leading rows when header mode is off', async () => {
    const path = temp('blank-leading.xlsx');
    await writeSheet(
      path,
      [
        [null, null],
        [1, 2],
      ],
      { header: false },
    );
    expect(await collect(readSheet(path))).toEqual([[1, 2]]);
  });
});

process.on('exit', () => rmSync(tmp, { recursive: true, force: true }));
