import { decodeUtf8, unzipEntries } from './zip.js';
import {
  countRows,
  findChild,
  findChildren,
  localName,
  parseXml,
  scanRows,
  type RawCell,
  type XmlNode,
} from './xml.js';
import type { CellValue, ReadOptions, RowObject, SheetInfo, SheetRow, Source } from './types.js';

const WORKBOOK_PART = 'xl/workbook.xml';
const WORKBOOK_RELS_PART = 'xl/_rels/workbook.xml.rels';
const SHARED_STRINGS_PART = 'xl/sharedStrings.xml';
const STYLES_PART = 'xl/styles.xml';

const DAY_MS = 86_400_000;
const EXCEL_EPOCH_1900 = Date.UTC(1899, 11, 30);
const EXCEL_EPOCH_1904 = Date.UTC(1904, 0, 1);

/**
 * Number formats Excel ships in the file format itself. 14–22, 45–47 and the
 * locale variants in 27–36 / 50–58 render as dates or times.
 */
const BUILTIN_DATE_FORMATS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52,
  53, 54, 55, 56, 57, 58,
]);

// ---------------------------------------------------------------------------
// byte sources
// ---------------------------------------------------------------------------

export async function toBytes(source: Source): Promise<Uint8Array> {
  if (typeof source === 'string') {
    const { readFile } = await import('node:fs/promises');
    const buffer = await readFile(source);
    return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
  }
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  if (typeof Blob !== 'undefined' && source instanceof Blob) {
    return new Uint8Array(await source.arrayBuffer());
  }
  const iterable = source as AsyncIterable<Uint8Array>;
  if (typeof iterable[Symbol.asyncIterator] === 'function') {
    const chunks: Uint8Array[] = [];
    let total = 0;
    for await (const chunk of iterable) {
      chunks.push(chunk);
      total += chunk.length;
    }
    const out = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      out.set(chunk, offset);
      offset += chunk.length;
    }
    return out;
  }
  throw new TypeError(
    'excel-lite: unsupported source — pass a file path, Uint8Array, ArrayBuffer, Blob or async iterable of bytes',
  );
}

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

function normalizePath(path: string): string {
  const parts: string[] = [];
  for (const segment of path.split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') parts.pop();
    else parts.push(segment);
  }
  return parts.join('/');
}

function resolveTarget(basePart: string, target: string): string {
  if (target.startsWith('/')) return normalizePath(target);
  const dir = basePart.includes('/') ? basePart.slice(0, basePart.lastIndexOf('/')) : '';
  return normalizePath(`${dir}/${target}`);
}

/** `C12` -> 2 (zero based). Returns -1 when the reference has no column part. */
export function columnIndexFromRef(ref: string): number {
  let index = 0;
  let sawLetter = false;
  for (let i = 0; i < ref.length; i++) {
    const code = ref.charCodeAt(i);
    if (code >= 65 && code <= 90) {
      index = index * 26 + (code - 64);
      sawLetter = true;
    } else if (code >= 97 && code <= 122) {
      index = index * 26 + (code - 96);
      sawLetter = true;
    } else break;
  }
  return sawLetter ? index - 1 : -1;
}

/**
 * Excel counts days from 1899-12-30, but claims 1900-02-29 existed (serial 60).
 * Serials below 61 are therefore shifted by one day. Dates come back in UTC so
 * the result does not depend on the machine's time zone.
 */
export function serialToDate(serial: number, date1904 = false): Date {
  if (date1904) return new Date(EXCEL_EPOCH_1904 + serial * DAY_MS);
  const adjusted = serial >= 1 && serial < 61 ? serial + 1 : serial;
  return new Date(EXCEL_EPOCH_1900 + adjusted * DAY_MS);
}

function isDateFormat(numFmtId: number, customFormats: Map<number, string>): boolean {
  if (BUILTIN_DATE_FORMATS.has(numFmtId)) return true;
  if (numFmtId < 164) return false;
  const code = customFormats.get(numFmtId);
  if (!code) return false;
  const stripped = code
    .replace(/"[^"]*"/g, '') // quoted literals
    .replace(/\\./g, '') // escaped characters
    .replace(/\[(?![hms])[^\]]*\]/gi, '') // [Red], [$-409] but keep [h] [m] [s]
    .replace(/\bGeneral\b/gi, '')
    .replace(/\bAM\/PM\b/gi, '')
    .replace(/\bA\/P\b/gi, '');
  return /[ymdhs]/i.test(stripped);
}

/** Text of a shared-string item, ignoring phonetic (<rPh>) runs. */
function sharedStringText(si: XmlNode): string {
  let out = '';
  for (const child of si.children) {
    const name = localName(child.name);
    if (name === 't') out += child.text;
    else if (name === 'r') {
      const t = findChild(child, 't');
      if (t) out += t.text;
    }
  }
  return out;
}

function makeKeys(row: CellValue[]): string[] {
  const keys: string[] = [];
  const seen = new Map<string, number>();
  for (let i = 0; i < row.length; i++) {
    const value = row[i];
    let key = value === null || value === undefined ? '' : String(value).trim();
    if (key === '') key = `column_${i + 1}`;
    const count = seen.get(key) ?? 0;
    seen.set(key, count + 1);
    keys.push(count === 0 ? key : `${key}_${count + 1}`);
  }
  return keys;
}

interface SheetRef {
  name: string;
  part: string;
}

// ---------------------------------------------------------------------------
// reader
// ---------------------------------------------------------------------------

export class WorkbookReader {
  readonly sheetNames: string[];
  private readonly bytes: Uint8Array;
  private readonly refs: SheetRef[];
  private readonly date1904: boolean;
  private sharedStrings: string[] | null = null;
  private dateStyles: boolean[] = [];

  private constructor(bytes: Uint8Array, refs: SheetRef[], date1904: boolean) {
    this.bytes = bytes;
    this.refs = refs;
    this.sheetNames = refs.map((ref) => ref.name);
    this.date1904 = date1904;
  }

  static async open(source: Source): Promise<WorkbookReader> {
    const bytes = await toBytes(source);
    let parts: Record<string, Uint8Array>;
    try {
      parts = unzipEntries(bytes, [WORKBOOK_PART, WORKBOOK_RELS_PART]);
    } catch {
      throw new Error('excel-lite: could not read the archive — the file is not a zip (.xlsx/.xlsm) container');
    }
    const workbookBytes = parts[WORKBOOK_PART];
    if (!workbookBytes) {
      throw new Error(
        'excel-lite: xl/workbook.xml is missing — this is not an .xlsx file (legacy .xls binaries are not supported)',
      );
    }

    const workbookDoc = parseXml(decodeUtf8(workbookBytes));
    const workbook = findChild(workbookDoc, 'workbook');
    if (!workbook) throw new Error('excel-lite: workbook.xml is malformed');

    const properties = findChild(workbook, 'workbookPr');
    const date1904 = properties?.attrs.date1904 === '1' || properties?.attrs.date1904 === 'true';

    const relTargets = new Map<string, string>();
    const relsBytes = parts[WORKBOOK_RELS_PART];
    if (relsBytes) {
      const relsDoc = parseXml(decodeUtf8(relsBytes));
      for (const relationship of findChildren(relsDoc, 'Relationship')) {
        const id = relationship.attrs.Id;
        const target = relationship.attrs.Target;
        if (id && target) relTargets.set(id, resolveTarget(WORKBOOK_PART, target));
      }
    }

    const sheetNodes = findChildren(findChild(workbook, 'sheets'), 'sheet');
    if (sheetNodes.length === 0) throw new Error('excel-lite: the workbook has no sheets');

    const refs: SheetRef[] = sheetNodes.map((sheet, index) => {
      const name = sheet.attrs.name ?? `Sheet${index + 1}`;
      const relationshipId = sheet.attrs.id; // r:id, prefix stripped by the parser
      const part =
        (relationshipId ? relTargets.get(relationshipId) : undefined) ??
        `xl/worksheets/sheet${index + 1}.xml`;
      return { name, part };
    });

    return new WorkbookReader(bytes, refs, date1904);
  }

  private getSheetXml(index: number): string {
    const ref = this.refs[index];
    const parts = unzipEntries(this.bytes, [ref.part]);
    const bytes = parts[ref.part];
    if (!bytes) throw new Error(`excel-lite: sheet "${ref.name}" is missing from the archive (${ref.part})`);
    return decodeUtf8(bytes);
  }

  private loadSharedStrings(): string[] {
    if (this.sharedStrings) return this.sharedStrings;
    const parts = unzipEntries(this.bytes, [SHARED_STRINGS_PART]);
    const bytes = parts[SHARED_STRINGS_PART];
    if (!bytes) {
      this.sharedStrings = [];
      return this.sharedStrings;
    }
    const doc = parseXml(decodeUtf8(bytes));
    const strings = findChildren(findChild(doc, 'sst'), 'si').map(sharedStringText);
    this.sharedStrings = strings;
    return strings;
  }

  private loadStyles(): void {
    const parts = unzipEntries(this.bytes, [STYLES_PART]);
    const bytes = parts[STYLES_PART];
    if (!bytes) {
      this.dateStyles = [];
      return;
    }
    const doc = parseXml(decodeUtf8(bytes));
    const stylesheet = findChild(doc, 'styleSheet');
    const customFormats = new Map<number, string>();
    for (const format of findChildren(findChild(stylesheet, 'numFmts'), 'numFmt')) {
      customFormats.set(Number(format.attrs.numFmtId ?? 0), format.attrs.formatCode ?? '');
    }
    const xfs = findChildren(findChild(stylesheet, 'cellXfs'), 'xf');
    this.dateStyles = xfs.map((xf) => isDateFormat(Number(xf.attrs.numFmtId ?? 0), customFormats));
  }

  resolveSheet(sheet: string | number | undefined): number {
    if (sheet === undefined || sheet === null) return 0;
    if (typeof sheet === 'number') {
      if (sheet < 0 || sheet >= this.refs.length) {
        throw new RangeError(
          `excel-lite: sheet index ${sheet} is out of range (workbook has ${this.refs.length} sheet(s))`,
        );
      }
      return sheet;
    }
    const exact = this.sheetNames.indexOf(sheet);
    if (exact !== -1) return exact;
    const lower = sheet.toLowerCase();
    const relaxed = this.sheetNames.findIndex((name) => name.toLowerCase() === lower);
    if (relaxed !== -1) return relaxed;
    throw new Error(`excel-lite: sheet "${sheet}" not found — available: ${this.sheetNames.join(', ')}`);
  }

  private cellValue(cell: RawCell, sharedStrings: string[]): CellValue {
    switch (cell.type) {
      case 's': {
        if (cell.value === '') return null;
        const index = Number(cell.value);
        const text = sharedStrings[index];
        return text === undefined ? null : text;
      }
      case 'inlineStr':
      case 'str':
        return cell.value;
      case 'b':
        return cell.value === '1' || cell.value.toLowerCase() === 'true';
      case 'e':
        return cell.value; // #DIV/0!, #N/A, …
      case 'd': {
        const date = new Date(cell.value);
        return Number.isNaN(date.getTime()) ? cell.value : date;
      }
      default: {
        if (cell.value === '') return null;
        const numeric = Number(cell.value);
        if (!Number.isFinite(numeric)) return cell.value;
        return this.dateStyles[cell.style] ? serialToDate(numeric, this.date1904) : numeric;
      }
    }
  }

  private toArray(cells: RawCell[], sharedStrings: string[]): CellValue[] {
    const values: CellValue[] = [];
    let cursor = 0;
    for (const cell of cells) {
      let index = cursor;
      if (cell.ref) {
        const parsed = columnIndexFromRef(cell.ref);
        if (parsed >= 0) index = parsed;
      }
      values[index] = this.cellValue(cell, sharedStrings);
      cursor = index + 1;
    }
    let end = values.length;
    while (end > 0 && (values[end - 1] === null || values[end - 1] === undefined)) end--;
    values.length = end;
    // Sparse holes (a row with A and C but no B) read as null, never undefined.
    for (let index = 0; index < end; index++) {
      if (values[index] === undefined) values[index] = null;
    }
    return values;
  }

  /** Row count without parsing cells. */
  rowCount(index: number): number {
    return countRows(this.getSheetXml(index));
  }

  *rows(index: number, header: boolean): Generator<CellValue[] | RowObject> {
    const sharedStrings = this.loadSharedStrings();
    this.loadStyles();
    let keys: string[] | null = null;
    for (const raw of scanRows(this.getSheetXml(index))) {
      const values = this.toArray(raw.cells, sharedStrings);
      if (values.length === 0) continue;
      if (!header) {
        yield values;
        continue;
      }
      if (!keys) {
        keys = makeKeys(values);
        continue;
      }
      const row: RowObject = {};
      for (let i = 0; i < keys.length; i++) row[keys[i]] = values[i] ?? null;
      for (let i = keys.length; i < values.length; i++) row[`column_${i + 1}`] = values[i] ?? null;
      yield row;
    }
  }
}

// ---------------------------------------------------------------------------
// public API
// ---------------------------------------------------------------------------

export function readSheet(
  source: Source,
  options: ReadOptions & { header: true },
): AsyncGenerator<RowObject>;
export function readSheet(
  source: Source,
  options?: ReadOptions & { header?: false },
): AsyncGenerator<CellValue[]>;
export function readSheet(source: Source, options?: ReadOptions): AsyncGenerator<SheetRow>;
export async function* readSheet(source: Source, options: ReadOptions = {}): AsyncGenerator<SheetRow> {
  const reader = await WorkbookReader.open(source);
  const index = reader.resolveSheet(options.sheet);
  yield* reader.rows(index, Boolean(options.header)) as Generator<SheetRow>;
}

export interface Workbook {
  sheetNames: string[];
  /** Iterate one sheet, picked by name or index (default: the first). */
  sheet(sheet?: string | number, options?: { header?: boolean }): AsyncGenerator<SheetRow>;
  /** Stack of every sheet in order. */
  sheets(options?: { header?: boolean }): AsyncGenerator<SheetRow>[];
}

export async function readWorkbook(source: Source): Promise<Workbook> {
  const reader = await WorkbookReader.open(source);
  const stream = (index: number, header: boolean): AsyncGenerator<SheetRow> =>
    (async function* () {
      yield* reader.rows(index, header);
    })();
  return {
    sheetNames: reader.sheetNames,
    sheet(sheet = 0, options = {}) {
      return stream(reader.resolveSheet(sheet), Boolean(options.header));
    },
    sheets(options = {}) {
      return reader.sheetNames.map((_, index) => stream(index, Boolean(options.header)));
    },
  };
}

/** Sheet names plus row counts, without materialising cell values. */
export async function listSheets(source: Source): Promise<SheetInfo[]> {
  const reader = await WorkbookReader.open(source);
  return reader.sheetNames.map((name, index) => ({
    name,
    index,
    rowCount: reader.rowCount(index),
  }));
}
