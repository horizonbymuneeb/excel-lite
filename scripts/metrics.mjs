#!/usr/bin/env node
/**
 * Collect the numbers the README badges show, so they cannot go stale.
 *
 * Runs the test suite with coverage (which builds `dist` first, exactly like `npm test`),
 * measures the browser bundle with the same esbuild settings as `check:bundle`, and writes
 * site/metrics.json. The docs build turns that file into shields.io endpoint badges served
 * from GitHub Pages, so a badge changes only when the numbers in the repository change.
 *
 *   npm run metrics
 */
import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { build } from 'esbuild';

const ROOT = process.cwd();
const pretty = (command, args) => {
  const shown = [command, ...args].join(' ');
  console.log(`\n$ ${shown}`);
  const result = spawnSync(command, args, { cwd: ROOT, stdio: 'inherit', shell: false });
  if (result.status !== 0) throw new Error(`${shown} exited with ${result.status}`);
};

// 1. build + tests with coverage (the CLI tests run against dist)
pretty('npm', ['run', 'build']);
pretty('npx', [
  'vitest',
  'run',
  '-c',
  'vitest.config.coverage.ts',
  '--coverage',
  '--reporter=json',
  '--outputFile=.metrics-tests.json',
]);

const tests = JSON.parse(readFileSync(join(ROOT, '.metrics-tests.json'), 'utf8'));
const total = tests.numTotalTests ?? 0;
const passed = tests.numPassedTests ?? 0;
const failed = tests.numFailedTests ?? 0;

const summary = JSON.parse(readFileSync(join(ROOT, 'coverage', 'coverage-summary.json'), 'utf8')).total;

// 2. browser bundle size, same settings as the CI guard
const bundle = await build({
  stdin: {
    contents: "import { toXlsx, readSheet } from 'excel-lite';\nglobalThis.__probe = { toXlsx, readSheet };",
    resolveDir: ROOT,
    loader: 'js',
  },
  bundle: true,
  write: false,
  platform: 'browser',
  format: 'esm',
  target: 'es2022',
  minify: true,
  logLevel: 'silent',
});
const minified = bundle.outputFiles[0].text;
const { gzipSync } = await import('node:zlib');
const round = (bytes) => Math.round((bytes / 1024) * 10) / 10;
const bundleMinKb = round(Buffer.byteLength(minified));
const bundleGzipKb = round(gzipSync(Buffer.from(minified), { level: 9 }).length);

const version = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')).version;
const metrics = {
  version,
  tests: { total, passed, failed },
  coverage: {
    statements: summary.statements.pct,
    branches: summary.branches.pct,
    functions: summary.functions.pct,
    lines: summary.lines.pct,
  },
  bundle: { minKb: bundleMinKb, gzipKb: bundleGzipKb },
  generatedAt: new Date().toISOString(),
};

writeFileSync(join(ROOT, 'site', 'metrics.json'), `${JSON.stringify(metrics, null, 2)}\n`);
console.log(
  `\nmetrics: ${passed}/${total} tests, ${metrics.coverage.statements}% statements, ` +
    `${bundleMinKb} kB minified / ${bundleGzipKb} kB gzip browser bundle`,
);
