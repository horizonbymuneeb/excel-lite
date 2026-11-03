#!/usr/bin/env node
/**
 * Docs link checker.
 *
 * Fails when a relative link points at a file that does not exist, and warns when an
 * in-page anchor does not match a heading. Also fails when a page in docs/ is not linked
 * from anywhere (an unlinked page is a page nobody will find).
 *
 *   node scripts/check-links.mjs
 */
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';

const ROOT = process.cwd();
const TARGETS = [
  'README.md',
  'CHANGELOG.md',
  'CONTRIBUTING.md',
  'CODE_OF_CONDUCT.md',
  'SECURITY.md',
  'SUPPORT.md',
  'ROADMAP.md',
  ...walk(join(ROOT, 'docs')).map((path) => relative(ROOT, path)),
  ...walk(join(ROOT, 'examples')).map((path) => relative(ROOT, path)),
];

function walk(directory) {
  let found = [];
  let entries;
  try {
    entries = readdirSync(directory, { withFileTypes: true });
  } catch {
    return found;
  }
  for (const entry of entries) {
    const full = join(directory, entry.name);
    if (entry.isDirectory()) found = found.concat(walk(full));
    else if (entry.name.endsWith('.md')) found.push(full);
  }
  return found;
}

/** GitHub's heading slug rules, close enough for link checking. */
function slugify(heading) {
  return heading
    .trim()
    .toLowerCase()
    .replace(/[`*_~]/g, '')
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/[^a-z0-9\u00c0-\uffff\s-]/g, '')
    .replace(/\s+/g, '-');
}

function headingsOf(path) {
  const text = readFileSync(path, 'utf8')
    .replace(/```[\s\S]*?```/g, '')
    .replace(/`[^`\n]*`/g, '');
  const slugs = new Set();
  for (const line of text.split('\n')) {
    const match = /^(#{1,6})\s+(.*)$/.exec(line);
    if (match) slugs.add(slugify(match[2]));
  }
  return slugs;
}

const LINK_RE = /!?\[[^\]]*\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g;
const errors = [];
const warnings = [];
const linkedFiles = new Set();

for (const file of TARGETS) {
  const absolute = join(ROOT, file);
  let exists = true;
  try {
    statSync(absolute);
  } catch {
    exists = false;
  }
  if (!exists) {
    errors.push(`${file}: missing file (listed in the docs set)`);
    continue;
  }

  const source = readFileSync(absolute, 'utf8').replace(/```[\s\S]*?```/g, '');
  let match;
  while ((match = LINK_RE.exec(source))) {
    const target = match[1];
    if (!target || target.startsWith('http://') || target.startsWith('https://') || target.startsWith('mailto:')) {
      continue;
    }
    if (target.includes('${') || target === '#') continue;

    const [pathPart, anchor] = target.split('#');
    const resolved = pathPart ? resolve(dirname(absolute), pathPart) : absolute;
    let stat;
    try {
      stat = statSync(resolved);
    } catch {
      errors.push(`${file}: broken link -> ${target}`);
      continue;
    }
    linkedFiles.add(relative(ROOT, resolved));
    if (anchor && stat.isFile() && resolved.endsWith('.md')) {
      if (!headingsOf(resolved).has(anchor)) {
        warnings.push(`${file}: anchor not found -> ${target}`);
      }
    }
  }
}

// Every docs page should be reachable from the docs index.
const docsIndex = join(ROOT, 'docs', 'README.md');
let indexed = '';
try {
  indexed = readFileSync(docsIndex, 'utf8');
} catch {
  errors.push('docs/README.md: missing docs index');
}

for (const page of walk(join(ROOT, 'docs')).map((path) => relative(join(ROOT, 'docs'), path))) {
  if (page === 'README.md') continue;
  if (!indexed.includes(page)) errors.push(`docs/README.md does not link to docs/${page}`);
}

for (const warning of warnings) console.warn(`warn  ${warning}`);
if (errors.length) {
  for (const error of errors) console.error(`error ${error}`);
  console.error(`\n${errors.length} problem(s), ${warnings.length} warning(s)`);
  process.exit(1);
}
console.log(`checked ${TARGETS.length} markdown files: 0 broken file links, ${warnings.length} anchor warning(s)`);
