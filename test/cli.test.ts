import { describe, expect, it } from 'vitest';
import { execFile } from 'node:child_process';
import { join } from 'node:path';
import { promisify } from 'node:util';

const run = promisify(execFile);
const cli = join(process.cwd(), 'dist', 'cli.js');
const mixed = join(process.cwd(), 'fixtures', 'openpyxl-mixed.xlsx');

describe('cli', () => {
  it('prints sheet names and row counts', async () => {
    const { stdout } = await run(process.execPath, [cli, 'info', mixed]);
    expect(stdout).toContain('Data');
    expect(stdout).toContain('Notes');
    expect(stdout).toMatch(/4 rows/);
  });

  it('streams NDJSON with --header', async () => {
    const { stdout } = await run(process.execPath, [cli, 'to-json', mixed, '--header']);
    const lines = stdout.trim().split('\n');
    expect(lines).toHaveLength(3);
    expect(JSON.parse(lines[0]!)).toMatchObject({ name: 'Ali', qty: 3, active: true });
    expect(JSON.parse(lines[2]!)).toMatchObject({ name: 'Nina', qty: 1299 });
  });

  it('writes one JSON array with --pretty', async () => {
    const { stdout } = await run(process.execPath, [cli, 'to-json', mixed, '--pretty']);
    const parsed = JSON.parse(stdout);
    expect(Array.isArray(parsed)).toBe(true);
    expect(parsed).toHaveLength(4);
    expect(parsed[0]).toEqual(['name', 'qty', 'price', 'when', 'active', 'note']);
  });

  it('emits CSV with proper quoting', async () => {
    const { stdout } = await run(process.execPath, [cli, 'to-csv', mixed]);
    const lines = stdout.trimEnd().split('\n');
    expect(lines[0]).toBe('name,qty,price,when,active,note');
    expect(lines[2]).toContain('"comma, ""quote"""');
    expect(lines).toHaveLength(4);
  });

  it('shows usage on --help', async () => {
    const { stdout } = await run(process.execPath, [cli, '--help']);
    expect(stdout).toContain('excel-lite');
    expect(stdout).toContain('to-json');
  });

  it('fails with a readable message on a missing file', async () => {
    await expect(run(process.execPath, [cli, 'info', '/nope/missing.xlsx'])).rejects.toMatchObject({ code: 1 });
  });
});
