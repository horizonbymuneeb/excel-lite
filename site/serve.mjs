#!/usr/bin/env node
/** Minimal static server for previewing site/dist locally. `node site/serve.mjs [port]`. */
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname, join, normalize } from 'node:path';

const ROOT = join(import.meta.dirname, 'dist');
const PORT = Number(process.argv[2] ?? process.env.PORT ?? 4180);
const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.woff2': 'font/woff2',
};

createServer(async (request, response) => {
  const url = new URL(request.url, `http://localhost:${PORT}`);
  const clean = normalize(decodeURIComponent(url.pathname)).replace(/^(\.\.[/\\])+/, '');
  const file = join(ROOT, clean.endsWith('/') ? `${clean}index.html` : clean);
  try {
    const body = await readFile(file);
    response.writeHead(200, { 'content-type': TYPES[extname(file)] ?? 'application/octet-stream' });
    response.end(body);
  } catch {
    try {
      const body = await readFile(join(ROOT, '404.html'));
      response.writeHead(404, { 'content-type': TYPES['.html'] });
      response.end(body);
    } catch {
      response.writeHead(404).end('not found');
    }
  }
}).listen(PORT, () => console.log(`docs preview: http://localhost:${PORT}/`));
