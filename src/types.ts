/** Values that a spreadsheet cell can hold after reading. */
export type CellValue = string | number | boolean | Date | null;

/** A row read with `header: true`. */
export type RowObject = Record<string, CellValue>;

export type SheetRow = CellValue[] | RowObject;

/** Anything we can turn into bytes: file path (Node), Buffer, ArrayBuffer, Blob, or a byte stream. */
export type Source = string | Uint8Array | ArrayBuffer | Blob | AsyncIterable<Uint8Array>;

export interface ReadOptions {
  /** Sheet name or 0-based index. Defaults to the first sheet. */
  sheet?: string | number;
  /** Treat the first non-empty row as object keys. */
  header?: boolean;
}

export interface SheetInfo {
  name: string;
  index: number;
  rowCount: number;
}

export interface ColumnSpec {
  /** Object key in the input rows. */
  key: string;
  /** Header text written to row 1. Defaults to `key`. */
  header?: string;
  /** Column width in Excel units. */
  width?: number;
  /** Excel number format, e.g. `#,##0.00` or `yyyy-mm-dd`. */
  format?: string;
  bold?: boolean;
  /** Solid fill colour, `RRGGBB` or `AARRGGBB` (leading `#` is fine). */
  fill?: string;
  align?: 'left' | 'center' | 'right';
}

export type RowInput = CellValue[] | Record<string, CellValue | undefined>;

export interface SheetInput {
  /** Sheet tab name. Defaults to `Sheet1`, `Sheet2`, … */
  name?: string;
  rows: Iterable<RowInput>;
  /** Column order, headers, widths and number formats. */
  columns?: ColumnSpec[];
  /** Write the header row. Defaults to true when `columns` is given. */
  header?: boolean;
  /** Freeze panes: `'A2'` (top row) or `{ row, col }` counts. */
  freeze?: string | { row?: number; col?: number };
  /** Add a filter dropdown to the header row. Defaults to true when a header is written. */
  autoFilter?: boolean;
  /** Compute column widths from the data. Defaults to true when a column has no explicit width. */
  autoWidth?: boolean;
}

export interface WriteOptions extends Omit<SheetInput, 'rows' | 'name'> {
  name?: string;
}
