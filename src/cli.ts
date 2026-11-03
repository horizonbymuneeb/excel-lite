#!/usr/bin/env node
import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';
import { listSheets, readSheet } from './index.js';
import type { CellValue, SheetRow } from './types.js';

const USAGE = `excel-lite — read .xlsx without the ceremony

Usage:
  excel-lite info <file.xlsx>
  excel-lite to-json <file.xlsx> [options]
  excel-lite to-csv  <file.xlsx> [options]

Options:
  --sheet <name|index>   sheet to read (default: first)
  --header               use the first row as keys
  --out <file>           write to a file instead of stdout
  --pretty               to-json: emit one JSON array instead of NDJSON
  --delimiter <char>     to-csv: field separator (default: ,)
  -h, --help             show this help

Examples:
  excel-lite info report.xlsx
  excel-lite to-json report.xlsx --sheet Data --header > rows.ndjson
  excel-lite to-csv report.xlsx --out report.csv
`;

interface Options {
  sheet?: string | number;
  header: boolean;
  out?: string;
  pretty: boolean;
  delimiter: string;
}

function parseArgs(argv: string[]): { command: string; file?: string; options: Options } {
  const options: Options = { header: false, pretty: false, delimiter: ',' };
  let command = '';
  let file: string | undefined;

  for (let index = 0; index < argv.length; index++) {
    const arg = argv[index];
    if (arg === '-h' || arg === '--help') {
      command = 'help';
      continue;
    }
    if (arg === '--header') {
      options.header = true;
    } else if (arg === '--pretty') {
      options.pretty = true;
    } else if (arg === '--sheet') {
      const value = argv[++index];
      if (value === undefined) throw new Error('--sheet needs a value');
      options.sheet = /^\d+$/.test(value) ? Number(value) : value;
    } else if (arg === '--out') {
      const value = argv[++index];
      if (value === undefined) throw new Error('--out needs a value');
      options.out = value;
    } else if (arg === '--delimiter') {
      const value = argv[++index];
      if (value === undefined) throw new Error('--delimiter needs a value');
      options.delimiter = value;
    } else if (arg.startsWith('--')) {
      throw new Error(`unknown option ${arg}`);
    } else if (!command) {
      command = arg;
    } else if (!file) {
      file = arg;
    } else {
      throw new Error(`unexpected argument "${arg}"`);
    }
  }

  return { command, file, options };
}

function cellText(value: CellValue): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

function csvField(text: string, delimiter: string): string {
  const needsQuotes = text.includes(delimiter) || text.includes('"') || text.includes('\n') || text.includes('\r');
  const escaped = text.replace(/"/g, '""');
  return needsQuotes ? `"${escaped}"` : escaped;
}

function objectKeys(row: SheetRow, fallback: string[]): string[] {
  if (Array.isArray(row)) return fallback;
  return Object.keys(row);
}

function rowArray(row: SheetRow, keys: string[]): string[] {
  if (Array.isArray(row)) return row.map(cellText);
  return keys.map((key) => cellText(row[key] ?? null));
}

function ignoreBrokenPipe(stream: NodeJS.WriteStream | Writable): void {
  stream.on('error', (error: NodeJS.ErrnoException) => {
    if (error.code === 'EPIPE') process.exit(0);
    throw error;
  });
}

async function main(argv: string[]): Promise<void> {
  const { command, file, options } = parseArgs(argv);
  if (command === 'help' || command === '') {
    process.stdout.write(USAGE);
    return;
  }
  if (!file) throw new Error(`"${command}" needs a .xlsx file`);
  if (command !== 'info' && command !== 'to-json' && command !== 'to-csv') {
    throw new Error(`unknown command "${command}" (try --help)`);
  }

  if (command === 'info') {
    const sheets = await listSheets(file);
    const width = Math.max(...sheets.map((sheet) => sheet.name.length), 4);
    process.stdout.write(`${file}\n`);
    for (const sheet of sheets) {
      process.stdout.write(`  ${String(sheet.index).padStart(2)}  ${sheet.name.padEnd(width)}  ${sheet.rowCount} rows\n`);
    }
    return;
  }

  const sink = options.out ? createWriteStream(options.out, 'utf8') : (process.stdout as unknown as Writable);
  ignoreBrokenPipe(sink);

  let headerKeys: string[] | null = null;
  let first = true;
  const prettyRows: SheetRow[] = [];

  for await (const row of readSheet(file, { sheet: options.sheet, header: options.header })) {
    if (options.header) {
      const keys = objectKeys(row, []);
      if (keys.length > 0 && !Array.isArray(row)) headerKeys = keys;
    }

    if (command === 'to-csv') {
      if (first && headerKeys) {
        sink.write(`${headerKeys.map((key) => csvField(key, options.delimiter)).join(options.delimiter)}\n`);
      }
      const fields = Array.isArray(row) ? row.map(cellText) : headerKeys!.map((key) => cellText(row[key] ?? null));
      sink.write(`${fields.map((field) => csvField(field, options.delimiter)).join(options.delimiter)}\n`);
    } else if (options.pretty) {
      prettyRows.push(row);
    } else {
      sink.write(`${JSON.stringify(row)}\n`);
    }
    first = false;
  }

  if (command === 'to-json' && options.pretty) {
    sink.write(`${JSON.stringify(prettyRows, null, 2)}\n`);
  }

  if (sink !== process.stdout) {
    await new Promise<void>((resolve, reject) => {
      sink.end(() => resolve());
      sink.on('error', reject);
    });
  }
}

main(process.argv.slice(2)).catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`excel-lite: ${message}\n`);
  process.exitCode = 1;
});
