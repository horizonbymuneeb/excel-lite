import { describe, expect, it } from 'vitest';
import { execFile } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';
import { toXlsx } from 'excel-lite';

const run = promisify(execFile);
const cli = join(process.cwd(), 'dist', 'cli.js');
const mixed = join(process.cwd(), 'fixtures', 'openpyxl-mixed.xlsx');

const tmp = mkdtempSync(join(tmpdir(), 'excel-lite-cli-'));
const temp = (name: string) => join(tmp, name);
const node = (args: string[]) => run(process.execPath, [cli, ...args]);

interface Failure {
  code: number;
  stdout: string;
  stderr: string;
}

async function expectFailure(args: string[]): Promise<Failure> {
  try {
    await node(args);
  } catch (error) {
    return error as Failure;
  }
  throw new Error(`expected ${args.join(' ')} to fail`);
}

// --- fixtures for the CSV/EPIPE cases -------------------------------------

const csvFile = temp('quoting.xlsx');
writeFileSync(
  csvFile,
  toXlsx([{ name: 'Csv', rows: [['a,b', 'q"q', 'line\nbreak', 'plain'], [null, '', 1.5]] }]),
);

const bigFile = temp('big.xlsx');
writeFileSync(
  bigFile,
  toXlsx([{ rows: Array.from({ length: 5000 }, (_, index) => [index, `row-${index}`]) }]),
);

const notXlsx = temp('not.xlsx');
writeFileSync(notXlsx, 'this is not a spreadsheet');

describe('cli', () => {
  it('prints sheet names and row counts', async () => {
    const { stdout } = await node(['info', mixed]);
    expect(stdout).toBe(`${mixed}\n   0  Data   4 rows\n   1  Notes  3 rows\n`);
  });

  it('streams NDJSON with --header', async () => {
    const { stdout } = await node(['to-json', mixed, '--header']);
    const lines = stdout.trim().split('\n');
    expect(lines).toHaveLength(3);
    expect(JSON.parse(lines[0]!)).toMatchObject({ name: 'Ali', qty: 3, active: true });
    expect(JSON.parse(lines[2]!)).toMatchObject({ name: 'Nina', qty: 1299 });
  });

  it('streams NDJSON arrays without --header', async () => {
    const { stdout } = await node(['to-json', mixed]);
    const lines = stdout.trim().split('\n');
    expect(lines).toHaveLength(4);
    expect(JSON.parse(lines[0]!)).toEqual(['name', 'qty', 'price', 'when', 'active', 'note']);
    expect(JSON.parse(lines[1]!)[0]).toBe('Ali');
  });

  it('serialises Date cells as ISO-8601 strings', async () => {
    const { stdout } = await node(['to-json', mixed, '--header']);
    const first = JSON.parse(stdout.trim().split('\n')[0]!);
    expect(first.when).toBe('2023-07-04T00:00:00.000Z');
  });

  it('writes one JSON array with --pretty', async () => {
    const { stdout } = await node(['to-json', mixed, '--pretty']);
    const parsed = JSON.parse(stdout);
    expect(Array.isArray(parsed)).toBe(true);
    expect(parsed).toHaveLength(4);
    expect(parsed[0]).toEqual(['name', 'qty', 'price', 'when', 'active', 'note']);
  });

  it('writes --pretty output to a file with --out', async () => {
    const out = temp('pretty.json');
    const { stdout } = await node(['to-json', mixed, '--pretty', '--out', out]);
    expect(stdout).toBe('');
    const parsed = JSON.parse(readFileSync(out, 'utf8'));
    expect(parsed).toHaveLength(4);
  });

  it('selects a sheet by index and by name', async () => {
    const byIndex = await node(['to-json', mixed, '--sheet', '1', '--header']);
    const byName = await node(['to-json', mixed, '--sheet', 'Notes', '--header']);
    const byCase = await node(['to-json', mixed, '--sheet', 'notes', '--header']);
    const expected = '{"key":"owner","value":"finance"}\n{"key":"updated","value":"2024-01-15T00:00:00.000Z"}\n';
    expect(byIndex.stdout).toBe(expected);
    expect(byName.stdout).toBe(expected);
    expect(byCase.stdout).toBe(expected);
  });

  it('emits CSV with proper quoting', async () => {
    const { stdout } = await node(['to-csv', mixed]);
    const lines = stdout.trimEnd().split('\n');
    expect(lines[0]).toBe('name,qty,price,when,active,note');
    expect(lines[2]).toContain('"comma, ""quote"""');
    expect(lines).toHaveLength(4);
  });

  it('quotes commas, quotes and newlines in CSV fields', async () => {
    const { stdout } = await node(['to-csv', csvFile]);
    expect(stdout).toBe('"a,b","q""q","line\nbreak",plain\n,,1.5\n');
  });

  it('supports a custom CSV delimiter', async () => {
    const { stdout } = await node(['to-csv', csvFile, '--delimiter', ';']);
    expect(stdout).toBe('a,b;"q""q";"line\nbreak";plain\n;;1.5\n');
  });

  it('writes CSV to a file with --out', async () => {
    const out = temp('report.csv');
    const { stdout } = await node(['to-csv', mixed, '--out', out]);
    expect(stdout).toBe('');
    const lines = readFileSync(out, 'utf8').trimEnd().split('\n');
    expect(lines).toHaveLength(4);
    expect(lines[0]).toBe('name,qty,price,when,active,note');
  });

  it('uses the header row as object keys when --header is set for CSV', async () => {
    const { stdout } = await node(['to-csv', mixed, '--header']);
    const lines = stdout.trimEnd().split('\n');
    expect(lines).toHaveLength(4);
    expect(lines[0]).toBe('name,qty,price,when,active,note');
    expect(lines[1]?.startsWith('Ali,3,19.99,')).toBe(true);
  });

  it('shows usage on --help and -h', async () => {
    for (const flag of ['--help', '-h']) {
      const { stdout } = await node([flag]);
      expect(stdout).toContain('excel-lite — read .xlsx without the ceremony');
      expect(stdout).toContain('to-json');
      expect(stdout).toContain('--delimiter');
    }
  });

  it('fails with exit code 1 on a missing file', async () => {
    const failure = await expectFailure(['info', '/nope/missing.xlsx']);
    expect(failure.code).toBe(1);
    expect(failure.stderr).toMatch(/ENOENT/);
  });

  it('fails with exit code 1 on a file that is not an xlsx', async () => {
    const failure = await expectFailure(['info', notXlsx]);
    expect(failure.code).toBe(1);
    expect(failure.stderr).toMatch(/not a zip/);
  });

  it('rejects an unknown command', async () => {
    const failure = await expectFailure(['bogus', mixed]);
    expect(failure.code).toBe(1);
    expect(failure.stderr).toContain('unknown command "bogus"');
    expect(failure.stderr).toContain('--help');
  });

  it('rejects an unknown option', async () => {
    const failure = await expectFailure(['info', mixed, '--bogus']);
    expect(failure.code).toBe(1);
    expect(failure.stderr).toContain('unknown option --bogus');
  });

  it('rejects a command without a file', async () => {
    const failure = await expectFailure(['to-json']);
    expect(failure.code).toBe(1);
    expect(failure.stderr).toContain('"to-json" needs a .xlsx file');
  });

  it('rejects a flag without its value', async () => {
    for (const [flag, message] of [
      ['--sheet', '--sheet needs a value'],
      ['--out', '--out needs a value'],
      ['--delimiter', '--delimiter needs a value'],
    ] as const) {
      const failure = await expectFailure(['to-json', mixed, flag]);
      expect(failure.code).toBe(1);
      expect(failure.stderr).toContain(message);
    }
  });

  it('exits 0 on success', async () => {
    const result = await node(['info', mixed]);
    expect(result.stderr).toBe('');
  });

  it('survives EPIPE when the consumer closes the pipe early', async () => {
    const { stdout } = await run('bash', [
      '-c',
      `set -o pipefail; "${process.execPath}" "${cli}" to-json "${bigFile}" | head -1`,
    ]);
    const lines = stdout.trim().split('\n');
    expect(lines).toHaveLength(1);
    expect(JSON.parse(lines[0]!)).toEqual([0, 'row-0']);
  });
});

process.on('exit', () => rmSync(tmp, { recursive: true, force: true }));
