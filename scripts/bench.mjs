#!/usr/bin/env node
/**
 * Speed and peak-memory comparison against ExcelJS and SheetJS.
 *
 *   npm run bench            # 50k rows × 6 columns
 *   npm run bench -- 200000  # bigger file
 *
 * Every measurement runs in its own child process, so peak RSS belongs to the
 * operation and not to whatever ran before it. Read tests all use the same
 * SheetJS-written file; nobody reads a file their own writer shaped.
 */
import { spawnSync } from 'node:child_process';
import { statSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const argv = process.argv.slice(2);
const flag = (name) => argv.find((arg) => arg.startsWith(`--${name}=`))?.split('=')[1];
const ROWS = Number(argv.find((arg) => /^\d+$/.test(arg)) ?? 50_000);
const NAMES = ['Muneeb', 'Sara', 'Ali', 'Nina', 'Ünïcode ✓', 'Chen', 'Amara', 'Diego'];
const NOTES = ['invoice paid', 'pending review', 'follow up', 'duplicate?', 'a, "quoted" note'];

const COLUMNS = [
  { key: 'id', header: 'ID' },
  { key: 'name' },
  { key: 'qty' },
  { key: 'price', format: '#,##0.00' },
  { key: 'when', header: 'date', format: 'yyyy-mm-dd' },
  { key: 'note' },
];

function makeRows(count) {
  return Array.from({ length: count }, (_, index) => ({
    id: index + 1,
    name: NAMES[index % NAMES.length],
    qty: (index * 7) % 500,
    price: Math.round(((index * 1337) % 100000) / 100) + 0.99,
    when: new Date(Date.UTC(2020, 0, 1) + (index % 1500) * 86_400_000),
    note: NOTES[index % NOTES.length],
  }));
}

// ---------------------------------------------------------------------------
// child mode: one operation, one process
// ---------------------------------------------------------------------------

async function child(lib, mode, file) {
  // Only the write measurement needs the source data; building it for reads would inflate
  // their peak RSS with ~50 MB of harness objects.
  const rows = mode === 'write' ? makeRows(ROWS) : [];
  const started = performance.now();
  let produced = 0;

  if (mode === 'write') {
    if (lib === 'excel-lite') {
      const { writeSheet } = await import('excel-lite');
      await writeSheet(file, rows, { columns: COLUMNS, freeze: 'A2' });
    } else if (lib === 'ExcelJS') {
      const { default: ExcelJS } = await import('exceljs');
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Sheet1');
      sheet.columns = COLUMNS.map((column) => ({ header: column.header ?? column.key, key: column.key }));
      for (const row of rows) sheet.addRow(row);
      await workbook.xlsx.writeFile(file);
    } else {
      const XLSX = (await import('xlsx')).default;
      const sheet = XLSX.utils.json_to_sheet(rows, { cellDates: true });
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
      XLSX.writeFile(workbook, file);
    }
  } else if (lib === 'excel-lite') {
    const { readSheet } = await import('excel-lite');
    for await (const _row of readSheet(file)) produced++;
  } else if (lib === 'ExcelJS') {
    const { default: ExcelJS } = await import('exceljs');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(file);
    produced = workbook.getWorksheet(1).rowCount - 1;
  } else {
    const XLSX = (await import('xlsx')).default;
    const workbook = XLSX.readFile(file);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    produced = XLSX.utils.sheet_to_json(sheet, { header: 1 }).length - 1;
  }

  const ms = performance.now() - started;
  const rss = process.resourceUsage().maxRSS * 1024; // ru_maxrss is KiB on Linux
  let size = 0;
  try {
    size = statSync(file).size;
  } catch {
    /* read mode: the file already exists */
  }
  process.stdout.write(`${JSON.stringify({ ms, rss, size, rows: produced })}\n`);
}

// ---------------------------------------------------------------------------
// parent mode
// ---------------------------------------------------------------------------

function runChild(lib, mode, file) {
  const result = spawnSync(
    process.execPath,
    [new URL(import.meta.url).pathname, `--child=${lib}`, `--mode=${mode}`, `--file=${file}`, String(ROWS)],
    { encoding: 'utf8', maxBuffer: 8 * 1024 * 1024 },
  );
  if (result.status !== 0) {
    throw new Error(`${lib} ${mode} failed:\n${result.stderr}`);
  }
  const line = result.stdout.trim().split('\n').filter(Boolean).at(-1);
  return JSON.parse(line);
}

function median(values) {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)];
}

function baselineRss() {
  const result = spawnSync(process.execPath, ['-e', 'process.stdout.write(String(process.resourceUsage().maxRSS))'], {
    encoding: 'utf8',
  });
  return Number(result.stdout) * 1024;
}

function formatMs(ms) {
  return ms < 1000 ? `${Math.round(ms)} ms` : `${(ms / 1000).toFixed(2)} s`;
}

function main() {
  const dir = tmpdir();
  const files = {
    'excel-lite': join(dir, 'bench-excel-lite.xlsx'),
    ExcelJS: join(dir, 'bench-exceljs.xlsx'),
    SheetJS: join(dir, 'bench-sheetjs.xlsx'),
  };
  const libs = ['excel-lite', 'ExcelJS', 'SheetJS'];
  const rowsLabel = ROWS.toLocaleString();
  const base = baselineRss();
  const results = {};
  for (const file of Object.values(files)) rmSync(file, { force: true });

  console.log(`\n${rowsLabel} rows × 6 columns — Node ${process.version}, median of 3 runs\n`);

  // Every library writes first, then every library reads the SheetJS file.
  for (const lib of libs) {
    results[lib] = { read: null };
    const runs = Array.from({ length: 3 }, () => runChild(lib, 'write', files[lib]));
    results[lib].write = {
      ms: median(runs.map((run) => run.ms)),
      rss: median(runs.map((run) => run.rss)),
      size: runs[0].size,
      rows: runs[0].rows,
    };
  }
  for (const lib of libs) {
    const runs = Array.from({ length: 3 }, () => runChild(lib, 'read', files.SheetJS));
    results[lib].read = {
      ms: median(runs.map((run) => run.ms)),
      rss: median(runs.map((run) => run.rss)),
      size: runs[0].size,
      rows: runs[0].rows,
    };
  }

  for (const lib of libs) {
    const write = results[lib].write;
    const read = results[lib].read;
    console.log(
      `${lib.padEnd(11)} write ${formatMs(write.ms).padStart(8)}   read ${formatMs(read.ms).padStart(8)}` +
        `   peak RSS ${((read.rss - base) / 1024 / 1024).toFixed(0)} MB` +
        `   file ${(write.size / 1024 / 1024).toFixed(2)} MB`,
    );
  }

  console.log('\nMarkdown for the README:\n');
  console.log('| | excel-lite | ExcelJS | SheetJS |');
  console.log('|---|---|---|---|');
  const cell = (key) => libs.map((lib) => results[lib][key]);
  console.log(
    `| write ${rowsLabel} rows | ${cell('write')
      .map((r) => formatMs(r.ms))
      .join(' | ')} |`,
  );
  console.log(
    `| read ${rowsLabel} rows | ${cell('read')
      .map((r) => formatMs(r.ms))
      .join(' | ')} |`,
  );
  console.log(
    `| peak RSS (read) | ${cell('read')
      .map((r) => `${((r.rss - base) / 1024 / 1024).toFixed(0)} MB`)
      .join(' | ')} |`,
  );
  console.log(
    `| output file | ${cell('write')
      .map((r) => `${(r.size / 1024 / 1024).toFixed(2)} MB`)
      .join(' | ')} |`,
  );

  for (const file of Object.values(files)) rmSync(file, { force: true });
}

if (flag('child')) {
  await child(flag('child'), flag('mode'), flag('file'));
} else {
  main();
}
