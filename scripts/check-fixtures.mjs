#!/usr/bin/env node
/**
 * Fixture drift checker.
 *
 * Compares the committed fixtures against freshly generated ones by *content*, not by bytes:
 * zip streams are deflate output, and different zlib versions (GitHub's Ubuntu image vs a
 * laptop) legally produce different bytes for identical input. Entry names and uncompressed
 * bytes must match exactly.
 *
 *   node scripts/check-fixtures.mjs fixtures /tmp/regenerated-fixtures
 */
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync } from 'fflate';

const [expectedDir = 'fixtures', actualDir = '/tmp/regenerated-fixtures'] = process.argv.slice(2);

function workbooks(dir) {
  return readdirSync(dir)
    .filter((name) => name.endsWith('.xlsx'))
    .sort();
}

const expected = workbooks(expectedDir);
const actual = workbooks(actualDir);
const problems = [];

if (expected.join(',') !== actual.join(',')) {
  problems.push(`fixture set differs:\n  committed:   ${expected.join(', ')}\n  regenerated: ${actual.join(', ')}`);
}

for (const name of expected) {
  if (!actual.includes(name)) continue;
  const committed = unzipSync(readFileSync(join(expectedDir, name)));
  const regenerated = unzipSync(readFileSync(join(actualDir, name)));
  const committedParts = Object.keys(committed).sort();
  const regeneratedParts = Object.keys(regenerated).sort();
  if (committedParts.join(',') !== regeneratedParts.join(',')) {
    problems.push(`${name}: zip parts differ (${committedParts.length} vs ${regeneratedParts.length})`);
    continue;
  }
  const decoder = new TextDecoder();
  for (const part of committedParts) {
    const a = decoder.decode(committed[part]);
    const b = decoder.decode(regenerated[part]);
    if (a !== b) {
      const line = a.split('\n').findIndex((text, index) => text !== b.split('\n')[index]);
      problems.push(`${name}: ${part} differs (first difference near line ${line + 1})`);
      break;
    }
  }
}

if (problems.length) {
  console.error('fixtures are stale compared to scripts/make-fixtures.py:');
  for (const problem of problems) console.error(` - ${problem}`);
  console.error('\nRun: python3 scripts/make-fixtures.py');
  process.exit(1);
}

console.log(`fixtures match the generator: ${expected.length} workbooks, content-identical`);
