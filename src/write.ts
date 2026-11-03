import { writeFileBytes } from './fs.js';
import { encodeUtf8, zipFiles } from './zip.js';
import { escapeXml } from './xml.js';
import type { CellValue, ColumnSpec, RowInput, SheetInput, WriteOptions } from './types.js';

const DAY_MS = 86_400_000;
const EXCEL_EPOCH_1900 = Date.UTC(1899, 11, 30);
const MAX_STRING = 32_767;

/** Reverse of `serialToDate`: a JS Date becomes the Excel serial number. */
export function dateToSerial(date: Date): number {
  const shifted = (date.getTime() - EXCEL_EPOCH_1900) / DAY_MS;
  // Excel's phantom 1900-02-29 shifts every serial below 61 by one day. Serial 61 is
  // 1900-03-01 and is NOT shifted, so the boundary is inclusive.
  if (shifted >= 61) return shifted;
  if (shifted >= 1) return shifted - 1;
  return shifted;
}

/** 0 -> `A`, 25 -> `Z`, 26 -> `AA`. */
export function columnLetter(index: number): string {
  let remaining = index + 1;
  let out = '';
  while (remaining > 0) {
    const rem = (remaining - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return out;
}

function normalizeColor(color: string): string {
  const hex = color.trim().replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{6}$/.test(hex)) return `FF${hex}`;
  if (/^[0-9A-F]{8}$/.test(hex)) return hex;
  throw new Error(`excel-lite: invalid colour "${color}" — use RRGGBB or AARRGGBB`);
}

function uniqueSheetName(raw: string, used: Set<string>): string {
  const cleaned =
    raw
      .replace(/[\\/?*[\]:]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim() || 'Sheet';
  const base = cleaned.slice(0, 31);
  let candidate = base;
  let suffix = 2;
  while (used.has(candidate.toLowerCase())) {
    const tail = ` (${suffix++})`;
    candidate = base.slice(0, 31 - tail.length) + tail;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function hasTimePart(date: Date): boolean {
  return (
    date.getUTCHours() !== 0 ||
    date.getUTCMinutes() !== 0 ||
    date.getUTCSeconds() !== 0 ||
    date.getUTCMilliseconds() !== 0
  );
}

function looksLikeDateFormat(format: string): boolean {
  const stripped = format
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[(?![hms])[^\]]*\]/gi, '');
  return /[ymdhs]/i.test(stripped);
}

// ---------------------------------------------------------------------------
// styles
// ---------------------------------------------------------------------------

interface Xf {
  numFmtId: number;
  bold: boolean;
  fill: string | null;
  align: string | null;
}

class StyleBook {
  private customFormats: string[] = [];
  private fills: string[] = [];
  private xfs: Xf[] = [{ numFmtId: 0, bold: false, fill: null, align: null }];
  private keys = new Map<string, number>([['0|false||', 0]]);

  numFmtId(format?: string): number {
    if (!format || /^general$/i.test(format.trim())) return 0;
    const existing = this.customFormats.indexOf(format);
    if (existing !== -1) return 164 + existing;
    this.customFormats.push(format);
    return 164 + this.customFormats.length - 1;
  }

  xfId(style: { numFmtId?: number; bold?: boolean; fill?: string; align?: string } = {}): number {
    const fill = style.fill ? normalizeColor(style.fill) : null;
    const xf: Xf = {
      numFmtId: style.numFmtId ?? 0,
      bold: Boolean(style.bold),
      fill,
      align: style.align ?? null,
    };
    const key = `${xf.numFmtId}|${xf.bold}|${fill ?? ''}|${xf.align ?? ''}`;
    const existing = this.keys.get(key);
    if (existing !== undefined) return existing;
    this.xfs.push(xf);
    const id = this.xfs.length - 1;
    this.keys.set(key, id);
    if (fill && !this.fills.includes(fill)) this.fills.push(fill);
    return id;
  }

  toXml(): string {
    const numFmts = this.customFormats.length
      ? `<numFmts count="${this.customFormats.length}">${this.customFormats
          .map((code, index) => `<numFmt numFmtId="${164 + index}" formatCode="${escapeXml(code)}"/>`)
          .join('')}</numFmts>`
      : '';

    const fonts =
      '<fonts count="2">' +
      '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
      '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
      '</fonts>';

    const fills =
      `<fills count="${2 + this.fills.length}">` +
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      this.fills
        .map(
          (argb) =>
            `<fill><patternFill patternType="solid"><fgColor rgb="${argb}"/><bgColor indexed="64"/></patternFill></fill>`,
        )
        .join('') +
      '</fills>';

    const borders = '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>';

    const cellXfs = this.xfs
      .map((xf) => {
        const attrs = [
          `numFmtId="${xf.numFmtId}"`,
          `fontId="${xf.bold ? 1 : 0}"`,
          `fillId="${xf.fill ? this.fills.indexOf(xf.fill) + 2 : 0}"`,
          'borderId="0"',
          'xfId="0"',
        ];
        if (xf.numFmtId) attrs.push('applyNumberFormat="1"');
        if (xf.bold) attrs.push('applyFont="1"');
        if (xf.fill) attrs.push('applyFill="1"');
        if (xf.align) attrs.push('applyAlignment="1"');
        const alignment = xf.align ? `<alignment horizontal="${xf.align}"/>` : '';
        return `<xf ${attrs.join(' ')}>${alignment}</xf>`;
      })
      .join('');

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      numFmts +
      fonts +
      fills +
      borders +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${cellXfs}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/>' +
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
      '</styleSheet>'
    );
  }
}

// ---------------------------------------------------------------------------
// sheet writer
// ---------------------------------------------------------------------------

/**
 * Shared string table. `count` is the total number of references and
 * `uniqueCount` the number of distinct strings — Excel is strict about the
 * difference, even though most parsers ignore it.
 */
class SharedStrings {
  private readonly byValue = new Map<string, number>();
  readonly values: string[] = [];
  references = 0;

  add(text: string): number {
    this.references++;
    const value = text.length > MAX_STRING ? text.slice(0, MAX_STRING) : text;
    const existing = this.byValue.get(value);
    if (existing !== undefined) return existing;
    const position = this.values.length;
    this.values.push(value);
    this.byValue.set(value, position);
    return position;
  }
}

function displayWidth(value: CellValue): number {
  if (value === null || value === undefined) return 0;
  if (value instanceof Date) return 19;
  if (typeof value === 'boolean') return 5;
  return String(value).length;
}

function parseFreeze(freeze: SheetInput['freeze']): { xSplit: number; ySplit: number; topLeft: string } | null {
  if (!freeze) return null;
  let xSplit: number;
  let ySplit: number;
  if (typeof freeze === 'string') {
    const match = /^([A-Za-z]+)(\d+)$/.exec(freeze.trim());
    if (!match) throw new Error(`excel-lite: freeze must look like "A2", got "${freeze}"`);
    const letters = match[1].toUpperCase();
    let column = 0;
    for (const char of letters) column = column * 26 + (char.charCodeAt(0) - 64);
    xSplit = column - 1;
    ySplit = Number(match[2]) - 1;
  } else {
    xSplit = freeze.col ?? 0;
    ySplit = freeze.row ?? 0;
  }
  if (xSplit <= 0 && ySplit <= 0) return null;
  return { xSplit, ySplit, topLeft: `${columnLetter(xSplit)}${ySplit + 1}` };
}

function cellXml(
  value: CellValue,
  ref: string,
  columnStyleId: number,
  column: ColumnSpec | undefined,
  styles: StyleBook,
  shared: SharedStrings,
): string {
  if (value === null || value === undefined) return '';

  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return '';
    const format =
      column?.format && looksLikeDateFormat(column.format)
        ? column.format
        : hasTimePart(value)
          ? 'yyyy-mm-dd hh:mm:ss'
          : 'yyyy-mm-dd';
    const styleId = styles.xfId({
      numFmtId: styles.numFmtId(format),
      bold: column?.bold,
      fill: column?.fill,
      align: column?.align,
    });
    return `<c r="${ref}" s="${styleId}"><v>${dateToSerial(value)}</v></c>`;
  }

  const styleAttr = columnStyleId ? ` s="${columnStyleId}"` : '';

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      const index = shared.add(String(value));
      return `<c r="${ref}"${styleAttr} t="s"><v>${index}</v></c>`;
    }
    return `<c r="${ref}"${styleAttr}><v>${value}</v></c>`;
  }

  if (typeof value === 'boolean') {
    return `<c r="${ref}"${styleAttr} t="b"><v>${value ? 1 : 0}</v></c>`;
  }

  const text = String(value);
  if (text === '') return '';
  const index = shared.add(text);
  return `<c r="${ref}"${styleAttr} t="s"><v>${index}</v></c>`;
}

function buildSheetXml(sheet: SheetInput, styles: StyleBook, shared: SharedStrings): string {
  const columns: ColumnSpec[] = sheet.columns ? [...sheet.columns] : [];
  const body: Array<Array<CellValue | undefined>> = [];

  if (sheet.rows) {
    for (const row of sheet.rows) {
      if (Array.isArray(row)) {
        body.push(row as Array<CellValue | undefined>);
      } else {
        if (columns.length === 0) {
          for (const key of Object.keys(row as Record<string, CellValue | undefined>)) columns.push({ key });
        }
        body.push(columns.map((column) => (row as Record<string, CellValue | undefined>)[column.key]));
      }
    }
  }

  const emitHeader = (sheet.header ?? columns.length > 0) && columns.length > 0;
  const columnCount = Math.max(
    columns.length,
    body.reduce((max, row) => Math.max(max, row.length), 0),
  );
  const widths = new Array<number>(columnCount).fill(0);
  const lines: string[] = [];
  let rowNumber = 0;

  if (emitHeader) {
    rowNumber = 1;
    const cells: string[] = [];
    columns.forEach((column, index) => {
      const header = column.header ?? column.key;
      widths[index] = Math.max(widths[index], displayWidth(header));
      const styleId = styles.xfId({ bold: true });
      cells.push(cellXml(header, `${columnLetter(index)}1`, styleId, { ...column, bold: false }, styles, shared));
    });
    lines.push(`<row r="1">${cells.join('')}</row>`);
  }

  for (const values of body) {
    rowNumber++;
    const cells: string[] = [];
    const limit = Math.max(values.length, columns.length);
    for (let index = 0; index < limit; index++) {
      const value = values[index];
      if (value === null || value === undefined) continue;
      const column = columns[index];
      const styleId = column
        ? styles.xfId({
            numFmtId: column.format ? styles.numFmtId(column.format) : 0,
            bold: column.bold,
            fill: column.fill,
            align: column.align,
          })
        : 0;
      const ref = `${columnLetter(index)}${rowNumber}`;
      const xml = cellXml(value, ref, styleId, column, styles, shared);
      if (xml) cells.push(xml);
      if (column?.width === undefined) widths[index] = Math.max(widths[index], displayWidth(value));
    }
    if (cells.length > 0) lines.push(`<row r="${rowNumber}">${cells.join('')}</row>`);
  }

  const lastCell = columnCount > 0 && rowNumber > 0 ? `${columnLetter(columnCount - 1)}${rowNumber}` : 'A1';
  const dimension = `A1:${lastCell}`;

  const columnsXml = (() => {
    const autoWidth = sheet.autoWidth !== false;
    const parts: string[] = [];
    for (let index = 0; index < columnCount; index++) {
      const explicit = columns[index]?.width;
      const width = explicit ?? (autoWidth ? Math.min(60, Math.max(8, (widths[index] ?? 0) + 2)) : undefined);
      if (!width) continue;
      const rounded = Math.round(width * 100) / 100;
      parts.push(`<col min="${index + 1}" max="${index + 1}" width="${rounded}" customWidth="1"/>`);
    }
    return parts.length ? `<cols>${parts.join('')}</cols>` : '';
  })();

  const panes = parseFreeze(sheet.freeze);
  const viewsXml = panes
    ? '<sheetViews><sheetView workbookViewId="0" tabSelected="1">' +
      `<pane${panes.xSplit ? ` xSplit="${panes.xSplit}"` : ''}${
        panes.ySplit ? ` ySplit="${panes.ySplit}"` : ''
      } topLeftCell="${panes.topLeft}" activePane="${
        panes.xSplit && panes.ySplit ? 'bottomRight' : panes.ySplit ? 'bottomLeft' : 'topRight'
      }" state="frozen"/>` +
      '</sheetView></sheetViews>'
    : '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>';

  const showFilter = (sheet.autoFilter ?? emitHeader) && emitHeader && columnCount > 0;
  const filterXml = showFilter ? `<autoFilter ref="A1:${columnLetter(columnCount - 1)}1"/>` : '';

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<dimension ref="${dimension}"/>` +
    viewsXml +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    columnsXml +
    `<sheetData>${lines.join('')}</sheetData>` +
    filterXml +
    '</worksheet>'
  );
}

// ---------------------------------------------------------------------------
// package parts
// ---------------------------------------------------------------------------

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function contentTypesXml(sheetCount: number): string {
  const overrides: string[] = [];
  for (let index = 1; index <= sheetCount; index++) {
    overrides.push(
      `<Override PartName="/xl/worksheets/sheet${index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    );
  }
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    overrides.join('') +
    '</Types>'
  );
}

function rootRelsXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Relationships xmlns="${RELS_NS}">` +
    `<Relationship Id="rId1" Type="${DOC_REL}/officeDocument" Target="xl/workbook.xml"/>` +
    '</Relationships>'
  );
}

function workbookXml(names: string[]): string {
  const sheets = names
    .map((name, index) => `<sheet name="${escapeXml(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<bookViews><workbookView activeTab="0"/></bookViews>' +
    `<sheets>${sheets}</sheets>` +
    '<calcPr calcId="191029"/>' +
    '</workbook>'
  );
}

function workbookRelsXml(sheetCount: number): string {
  const parts: string[] = [];
  for (let index = 1; index <= sheetCount; index++) {
    parts.push(`<Relationship Id="rId${index}" Type="${DOC_REL}/worksheet" Target="worksheets/sheet${index}.xml"/>`);
  }
  parts.push(`<Relationship Id="rId${sheetCount + 1}" Type="${DOC_REL}/styles" Target="styles.xml"/>`);
  parts.push(`<Relationship Id="rId${sheetCount + 2}" Type="${DOC_REL}/sharedStrings" Target="sharedStrings.xml"/>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}">${parts.join(
    '',
  )}</Relationships>`;
}

function sharedStringsXml(shared: SharedStrings): string {
  const items = shared.values.map((text) => `<si><t xml:space="preserve">${escapeXml(text)}</t></si>`).join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${shared.references}" uniqueCount="${shared.values.length}">` +
    items +
    '</sst>'
  );
}

// ---------------------------------------------------------------------------
// public API
// ---------------------------------------------------------------------------

/**
 * Build a workbook in memory and return the `.xlsx` bytes.
 *
 * Works in Node, browsers and workers — nothing here touches the filesystem.
 *
 * @param sheets One entry per sheet: rows, optional `columns`, `freeze`, `autoFilter`.
 * @throws If `sheets` is empty, a colour is malformed, or `freeze` is not an A1-style cell.
 *
 * @example
 * ```ts
 * const bytes = toXlsx([{ rows: [{ id: 1, name: 'Ada' }], columns: [{ key: 'id' }, { key: 'name' }] }]);
 * ```
 */
export function toXlsx(sheets: SheetInput[]): Uint8Array {
  if (!Array.isArray(sheets) || sheets.length === 0) {
    throw new Error('excel-lite: toXlsx needs at least one sheet');
  }
  const styles = new StyleBook();
  const shared = new SharedStrings();
  const usedNames = new Set<string>();

  const built = sheets.map((sheet, index) => ({
    name: uniqueSheetName(sheet.name ?? `Sheet${index + 1}`, usedNames),
    xml: buildSheetXml(sheet, styles, shared),
  }));

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': encodeUtf8(contentTypesXml(built.length)),
    '_rels/.rels': encodeUtf8(rootRelsXml()),
    'xl/workbook.xml': encodeUtf8(workbookXml(built.map((sheet) => sheet.name))),
    'xl/_rels/workbook.xml.rels': encodeUtf8(workbookRelsXml(built.length)),
    'xl/styles.xml': encodeUtf8(styles.toXml()),
    'xl/sharedStrings.xml': encodeUtf8(sharedStringsXml(shared)),
  };
  built.forEach((sheet, index) => {
    files[`xl/worksheets/sheet${index + 1}.xml`] = encodeUtf8(sheet.xml);
  });

  return zipFiles(files);
}

/**
 * Write a workbook to a file path (Node 22+) and return the bytes that were written.
 *
 * @example
 * ```ts
 * await writeWorkbook('book.xlsx', [
 *   { name: 'Report', rows, columns: [{ key: 'amount', format: '#,##0.00' }] },
 *   { name: 'Raw', rows: [[1, 2, 3]] },
 * ]);
 * ```
 */
export async function writeWorkbook(destination: string, sheets: SheetInput[]): Promise<Uint8Array> {
  const bytes = toXlsx(sheets);
  await writeFileBytes(destination, bytes);
  return bytes;
}

/**
 * One-sheet convenience wrapper around {@link writeWorkbook}.
 *
 * @example
 * ```ts
 * await writeSheet('report.xlsx', rows, {
 *   columns: [{ key: 'email', width: 32 }, { key: 'joined', format: 'yyyy-mm-dd' }],
 *   freeze: 'A2',
 * });
 * ```
 */
export async function writeSheet(
  destination: string,
  rows: Iterable<RowInput>,
  options: WriteOptions = {},
): Promise<Uint8Array> {
  return writeWorkbook(destination, [{ ...options, rows }]);
}
