#!/usr/bin/env node
/**
 * Documentation site builder.
 *
 * The markdown in docs/ + the repo's governance files are the single source of truth; this
 * script renders them into a static site (site/dist) that is deployed to GitHub Pages by
 * .github/workflows/docs.yml. No framework, no client-side routing — plain HTML that works
 * with JavaScript disabled except for search and the theme toggle.
 *
 *   node site/build.mjs
 */
import { mkdirSync, readFileSync, writeFileSync, existsSync, rmSync, copyFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import MarkdownIt from 'markdown-it';
import hljs from 'highlight.js/lib/core';
import bash from 'highlight.js/lib/languages/bash';
import javascript from 'highlight.js/lib/languages/javascript';
import typescript from 'highlight.js/lib/languages/typescript';
import json from 'highlight.js/lib/languages/json';
import python from 'highlight.js/lib/languages/python';
import sql from 'highlight.js/lib/languages/sql';
import xml from 'highlight.js/lib/languages/xml';
import yaml from 'highlight.js/lib/languages/yaml';
import ini from 'highlight.js/lib/languages/ini';

const ROOT = resolve(import.meta.dirname, '..');
const OUT = join(ROOT, 'site', 'dist');
const REPO_URL = 'https://github.com/horizonbymuneeb/excel-lite';
const VERSION = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')).version;
const REPO_BRANCH = 'main';

for (const [name, language] of Object.entries({
  bash,
  shell: bash,
  sh: bash,
  javascript,
  js: javascript,
  typescript,
  ts: typescript,
  json,
  python,
  sql,
  xml,
  html: xml,
  yaml,
  yml: yaml,
  ini,
  toml: ini,
})) {
  hljs.registerLanguage(name, language);
}

// ---------------------------------------------------------------------------
// pages
// ---------------------------------------------------------------------------

const PAGES = [
  { src: 'docs/README.md', out: 'index.html', title: 'Introduction', section: 'Manual', lead: true },
  { src: 'docs/getting-started.md', out: 'getting-started.html', title: 'Getting started', section: 'Manual' },
  { src: 'docs/reading.md', out: 'reading.html', title: 'Reading', section: 'Manual' },
  { src: 'docs/writing.md', out: 'writing.html', title: 'Writing', section: 'Manual' },
  { src: 'docs/api.md', out: 'api.html', title: 'API reference', section: 'Manual' },
  { src: 'docs/cli.md', out: 'cli.html', title: 'Command line', section: 'Manual' },
  { src: 'docs/recipes.md', out: 'recipes.html', title: 'Recipes', section: 'Manual' },
  { src: 'docs/performance.md', out: 'performance.html', title: 'Performance', section: 'Manual' },
  { src: 'docs/compatibility.md', out: 'compatibility.html', title: 'Compatibility', section: 'Manual' },
  { src: 'docs/migration.md', out: 'migration.html', title: 'Migration', section: 'Manual' },
  { src: 'docs/troubleshooting.md', out: 'troubleshooting.html', title: 'Troubleshooting', section: 'Manual' },
  { src: 'docs/architecture.md', out: 'architecture.html', title: 'Architecture', section: 'Manual' },
  { src: 'docs/browser.md', out: 'browser.html', title: 'Browser', section: 'Manual' },
  { src: 'README.md', out: 'overview.html', title: 'Overview', section: 'Project' },
  { src: 'ROADMAP.md', out: 'roadmap.html', title: 'Roadmap', section: 'Project' },
  { src: 'CONTRIBUTING.md', out: 'contributing.html', title: 'Contributing', section: 'Project' },
  { src: 'SECURITY.md', out: 'security.html', title: 'Security', section: 'Project' },
  { src: 'CHANGELOG.md', out: 'changelog.html', title: 'Changelog', section: 'Project' },
  { src: 'SUPPORT.md', out: 'support.html', title: 'Support', section: 'Project' },
  { src: 'CODE_OF_CONDUCT.md', out: 'code-of-conduct.html', title: 'Code of conduct', section: 'Project' },
];

const bySource = new Map(PAGES.map((page) => [page.src, page]));

// ---------------------------------------------------------------------------
// markdown
// ---------------------------------------------------------------------------

/** GitHub-compatible heading slugs, with de-duplication per page. */
function makeSlugger() {
  const seen = new Map();
  return (text) => {
    const base = text
      .trim()
      .toLowerCase()
      .replace(/[`*_~]/g, '')
      .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
      .replace(/[^\p{L}\p{N}\s-]/gu, '')
      .replace(/\s+/g, '-');
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base}-${count}`;
  };
}

function escapeHtml(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Rewrite repo-relative markdown links to site paths (or GitHub, for files that are not pages). */
function makeLinkRewriter(page) {
  return (href) => {
    if (!href) return href;
    if (/^(https?:|mailto:|tel:|data:)/.test(href) || href.startsWith('#')) return href;
    const [pathPart, anchor] = href.split('#');
    if (!pathPart) return href;
    const absolute = resolve(dirname(join(ROOT, page.src)), pathPart);
    const repoRelative = relative(ROOT, absolute);
    const target = bySource.get(repoRelative);
    if (target) return `${target.out}${anchor ? `#${anchor}` : ''}`;
    if (existsSync(absolute)) return `${REPO_URL}/blob/${REPO_BRANCH}/${repoRelative}`;
    return href;
  };
}

function renderCode(code, language) {
  const lang = language && hljs.getLanguage(language) ? language : null;
  const body = lang ? hljs.highlight(code, { language: lang, ignoreIllegals: true }).value : escapeHtml(code);
  return (
    `<div class="code">` +
    (lang ? `<span class="code-lang">${escapeHtml(lang)}</span>` : '') +
    `<button class="code-copy" type="button" aria-label="Copy code">Copy</button>` +
    `<pre><code class="hljs">${body}</code></pre>` +
    `</div>`
  );
}

function renderMarkdown(page) {
  const slug = makeSlugger();
  const headings = [];
  const rewrite = makeLinkRewriter(page);
  const md = new MarkdownIt({ html: false, linkify: false, typographer: false });

  md.renderer.rules.fence = (tokens, index) => {
    const token = tokens[index];
    return renderCode(token.content, token.info.trim().split(/\s+/)[0]);
  };
  md.renderer.rules.link_open = (tokens, index) => {
    const token = tokens[index];
    const original = token.attrGet('href');
    const rewritten = rewrite(original);
    if (rewritten) token.attrSet('href', rewritten);
    if (/^https?:/.test(rewritten ?? '')) {
      token.attrSet('target', '_blank');
      token.attrSet('rel', 'noreferrer noopener');
    }
    return md.renderer.renderToken(tokens, index);
  };
  md.renderer.rules.table_open = () => '<div class="table-wrap"><table>';
  md.renderer.rules.table_close = () => '</table></div>';
  md.renderer.rules.heading_open = (tokens, index) => {
    const token = tokens[index];
    const text = tokens[index + 1]?.content ?? '';
    const id = slug(text);
    token.attrSet('id', id);
    if (token.tag === 'h2' || token.tag === 'h3') headings.push({ level: Number(token.tag[1]), text, id });
    return md.renderer.renderToken(tokens, index);
  };

  const source = readFileSync(join(ROOT, page.src), 'utf8');
  let html = md.render(source);
  // Hover anchors live inside the heading so the `#` sits next to the text.
  html = html.replace(
    /<h([23]) id="([^"]+)">([\s\S]*?)<\/h\1>/g,
    (match, level, id, inner) =>
      `<h${level} id="${id}">${inner}<a class="anchor" href="#${id}" aria-hidden="true" tabindex="-1">#</a></h${level}>`,
  );
  return { html, headings };
}

// ---------------------------------------------------------------------------
// templates
// ---------------------------------------------------------------------------

function navGroups(activePage) {
  const groups = new Map();
  for (const page of PAGES) {
    if (!groups.has(page.section)) groups.set(page.section, []);
    groups.get(page.section).push(page);
  }
  let html = '';
  for (const [section, pages] of groups) {
    html += `<div class="nav-group"><p class="nav-label">${escapeHtml(section)}</p><ul>`;
    for (const page of pages) {
      const active = page.out === activePage.out ? ' class="active" aria-current="page"' : '';
      html += `<li><a href="${page.out}"${active}>${escapeHtml(page.title)}</a></li>`;
    }
    html += '</ul></div>';
  }
  return html;
}

function tocHtml(headings) {
  const items = headings.filter((heading) => heading.level === 2);
  if (items.length < 2) return '';
  return (
    `<aside class="toc" aria-label="On this page"><p class="nav-label">On this page</p><ul>` +
    items.map((heading) => `<li><a href="#${heading.id}">${escapeHtml(heading.text)}</a></li>`).join('') +
    '</ul></aside>'
  );
}

function layout(page, content, headings) {
  const index = PAGES.findIndex((candidate) => candidate.out === page.out);
  const previous = PAGES[index - 1];
  const next = PAGES[index + 1];
  const description = `${page.title} — excel-lite documentation: read and write .xlsx in Node and the browser.`;
  return `<!doctype html>
<html lang="en" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(page.title)} · excel-lite</title>
<meta name="description" content="${escapeHtml(description)}">
<link rel="icon" href="assets/favicon.svg" type="image/svg+xml">
<link rel="stylesheet" href="assets/theme.css">
<meta name="color-scheme" content="light dark">
</head>
<body>
<a class="skip-link" href="#content">Skip to content</a>
<header class="topbar">
  <button class="menu-button" type="button" aria-label="Open navigation" aria-expanded="false" aria-controls="sidebar">
    <svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true"><path d="M3 5h14M3 10h14M3 15h14" stroke="currentColor" stroke-width="1.6" fill="none" stroke-linecap="round"/></svg>
  </button>
  <a class="brand" href="index.html"><span class="brand-mark" aria-hidden="true">x</span><span>excel-lite</span><span class="brand-version">docs</span></a>
  <div class="topbar-tools">
    <div class="search" role="search">
      <input id="search-input" type="search" placeholder="Search docs…" aria-label="Search documentation" autocomplete="off" spellcheck="false">
      <kbd>/</kbd>
      <div id="search-results" class="search-results" hidden></div>
    </div>
    <button class="icon-button" id="theme-toggle" type="button" aria-label="Toggle colour theme">
      <svg viewBox="0 0 20 20" width="17" height="17" aria-hidden="true"><circle cx="10" cy="10" r="4.2" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M10 1.5v2M10 16.5v2M1.5 10h2M16.5 10h2M3.9 3.9l1.4 1.4M14.7 14.7l1.4 1.4M16.1 3.9l-1.4 1.4M5.3 14.7l-1.4 1.4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
    </button>
    <a class="icon-button" href="${REPO_URL}" aria-label="GitHub repository" target="_blank" rel="noreferrer noopener">
      <svg viewBox="0 0 16 16" width="17" height="17" aria-hidden="true"><path fill="currentColor" d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82a7.4 7.4 0 0 1 2-.27c.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z"/></svg>
    </a>
    <a class="icon-button npm-link" href="https://www.npmjs.com/package/excel-lite" aria-label="npm package" target="_blank" rel="noreferrer noopener">npm</a>
  </div>
</header>
<div class="shell">
  <nav class="sidebar" id="sidebar" aria-label="Documentation">
    ${navGroups(page)}
    <div class="nav-group nav-foot">
      <a href="https://www.npmjs.com/package/excel-lite" target="_blank" rel="noreferrer noopener">npm package</a>
      <a href="${REPO_URL}/issues" target="_blank" rel="noreferrer noopener">Report an issue</a>
    </div>
  </nav>
  <main id="content" class="content">
    <p class="eyebrow">${escapeHtml(page.section)}</p>
    ${content}
    <nav class="pager" aria-label="Previous and next page">
      ${previous ? `<a class="pager-item" href="${previous.out}"><span>Previous</span>${escapeHtml(previous.title)}</a>` : '<span></span>'}
      ${next ? `<a class="pager-item next" href="${next.out}"><span>Next</span>${escapeHtml(next.title)}</a>` : '<span></span>'}
    </nav>
    <footer class="page-foot">
      <a href="${REPO_URL}/edit/${REPO_BRANCH}/${page.src}" target="_blank" rel="noreferrer noopener">Edit this page on GitHub</a>
      <span>MIT licensed · excel-lite ${VERSION}</span>
    </footer>
  </main>
  ${tocHtml(headings)}
</div>
<div class="scrim" hidden></div>
<script src="assets/app.js" defer></script>
</body>
</html>
`;
}

// ---------------------------------------------------------------------------
// build
// ---------------------------------------------------------------------------

function build() {
  rmSync(OUT, { recursive: true, force: true });
  mkdirSync(join(OUT, 'assets'), { recursive: true });

  const search = [];
  for (const page of PAGES) {
    const { html, headings } = renderMarkdown(page);
    writeFileSync(join(OUT, page.out), layout(page, html, headings));
    search.push({
      title: page.title,
      url: page.out,
      section: page.section,
      headings: headings.map((heading) => ({ text: heading.text, id: heading.id })),
    });
  }

  writeFileSync(join(OUT, 'search-index.json'), JSON.stringify(search));

  // Shields.io endpoint badges, refreshed every time the docs are deployed.
  const badge = (label, message, color) => ({ schemaVersion: 1, label, message, color });
  mkdirSync(join(OUT, 'badges'), { recursive: true });
  writeFileSync(join(OUT, 'badges', 'docs.json'), JSON.stringify(badge('docs', `${PAGES.length} pages`, '0b6b5f')));
  let metrics = null;
  try {
    metrics = JSON.parse(readFileSync(join(ROOT, 'site', 'metrics.json'), 'utf8'));
  } catch {
    console.warn('site/metrics.json missing — run `npm run metrics` for the tests/coverage badges');
  }
  if (metrics) {
    const { tests, coverage, bundle } = metrics;
    const allPassing = tests.failed === 0 && tests.passed === tests.total;
    writeFileSync(
      join(OUT, 'badges', 'tests.json'),
      JSON.stringify(
        badge(
          'tests',
          allPassing ? `${tests.passed} passing` : `${tests.failed} failing`,
          allPassing ? '2ea44f' : 'd1242f',
        ),
      ),
    );
    const statements = coverage.statements;
    writeFileSync(
      join(OUT, 'badges', 'coverage.json'),
      JSON.stringify(
        badge('coverage', `${statements}%`, statements >= 90 ? '2ea44f' : statements >= 75 ? 'd4a72c' : 'd1242f'),
      ),
    );
    writeFileSync(
      join(OUT, 'badges', 'bundle.json'),
      JSON.stringify(badge('browser bundle', `${bundle.minKb} kB min`, '2f6f9f')),
    );
  }
  copyFileSync(join(ROOT, 'site', 'assets', 'theme.css'), join(OUT, 'assets', 'theme.css'));
  copyFileSync(join(ROOT, 'site', 'assets', 'app.js'), join(OUT, 'assets', 'app.js'));
  writeFileSync(join(OUT, '.nojekyll'), '');
  writeFileSync(
    join(OUT, '404.html'),
    `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Not found · excel-lite</title>` +
      `<link rel="stylesheet" href="assets/theme.css"></head><body class="not-found"><main class="content">` +
      `<p class="eyebrow">404</p><h1>Page not found</h1><p>That page does not exist. Try the ` +
      `<a href="index.html">introduction</a> or the <a href="api.html">API reference</a>.</p></main></body></html>`,
  );
  console.log(`docs site: ${PAGES.length} pages + search index -> ${relative(ROOT, OUT)}/`);
}

build();
