# Browser and other runtimes

`excel-lite` works in browsers, web workers, service workers, Bun, Deno and edge runtimes — as
long as you pass **bytes**, not file paths. This page explains why that split exists, how to
configure the common bundlers, and gives a complete React upload/download example.

## On this page

- [What works where](#what-works-where)
- [Why there is no node:fs in the bundle](#why-there-is-no-nodefs-in-the-bundle)
- [Bundlers](#bundlers)
- [Reading a file the user picked](#reading-a-file-the-user-picked)
- [Complete Vite + React example](#complete-vite-react-example)
- [Bundle size](#bundle-size)
- [Workers and edge runtimes](#workers-and-edge-runtimes)
- [Browser limits](#browser-limits)

## What works where

| Feature | Browser / worker | Node 22.3+ | Bun / Deno |
|---|---|---|---|
| `toXlsx(sheets)` → `Uint8Array` | yes | yes | yes |
| `readSheet(Uint8Array \| ArrayBuffer \| Blob \| async iterable)` | yes | yes | yes |
| `readSheet('/path/file.xlsx')` | **no** — throws | yes | no |
| `writeSheet('/path', …)`, `writeWorkbook('/path', …)` | **no** — throws | yes | no |
| `readWorkbook`, `listSheets`, `WorkbookReader` | yes (bytes) | yes | yes |
| CLI (`excel-lite`) | no | yes | no |

A path in a browser throws a message that tells you exactly what to do instead:

```
excel-lite: reading or writing by file path needs Node 22+. In browsers and older runtimes pass bytes instead: readSheet(uint8Array) / toXlsx(sheets).
```

## Why there is no node:fs in the bundle

A library that statically imports `node:fs` cannot be bundled for the browser: bundlers refuse to
resolve `node:` specifiers, and polyfilling `fs` means shipping a broken shim. `excel-lite` needs
files only when you pass a path, so `src/fs.ts` gets the module at **runtime**:

```ts
function nodeFs() {
  const proc = globalThis.process as { getBuiltinModule?: (id: string) => unknown } | undefined;
  if (typeof proc?.getBuiltinModule !== 'function') return null;
  try {
    return proc.getBuiltinModule('fs/promises');
  } catch {
    return null;
  }
}
```

- In Node, `process.getBuiltinModule('fs/promises')` (Node 22.3+) returns the real module.
- In a browser or edge runtime, `process` or the function is missing, so `nodeFs()` returns `null`
  and path calls throw the friendly error above. Byte calls never touch it.
- Bundlers never see an `fs` import, so the graph is clean; `process` is referenced as a global only
  inside a runtime guard.

This is verified in CI: `npm run check:bundle` bundles the package for `platform: 'browser'`, fails
if the output contains `node:` builtins (or `require('node:…')`), fails above 120 kB, and then
**imports** the bundle, writes a workbook and reads it back.

## Bundlers

Nothing special is required. The package is ESM with a `exports` map and no side effects
(`"sideEffects": false`), so tree shaking works.

### Vite

```bash
npm create vite@latest my-app -- --template react
cd my-app && npm install excel-lite
```

```js
// src/lib/xlsx.js
import { toXlsx, readSheet } from 'excel-lite';
export { toXlsx, readSheet };
```

No `resolve.alias`, no `define`, no `optimizeDeps` entry. Vite's dev server and build both resolve
the `import` condition from the `exports` map. If you use Vite SSR, the same import works and paths
work on the server side.

### esbuild

```bash
npx esbuild src/index.js --bundle --platform=browser --format=esm --outfile=dist/app.js
```

`--platform=browser` is the whole configuration. Do **not** add
`--define:process.env.NODE_ENV=...` for this package; it does not read `process.env`.

### webpack 5

```js
// webpack.config.js
export default {
  target: ['web', 'es2022'],
  // no resolve.fallback needed: excel-lite never imports node builtins
};
```

If you see `node:fs` or `fs` in the stats graph, the importer is another dependency (an SDK, a
logger, a CLI helper), not `excel-lite` — see
[troubleshooting](troubleshooting.md#bundler-errors-about-nodefs).

### Other: Rollup, Parcel, Bun

- **Rollup**: works with `@rollup/plugin-node-resolve`; no node polyfills.
- **Parcel**: works with zero configuration; the `exports` map is honoured.
- **Bun**: `bun build --target=browser` works; `bun run script.ts` for Node-style scripts (bytes
  only — path access relies on Node's `process.getBuiltinModule`).

## Reading a file the user picked

`File` is a `Blob`; `arrayBuffer()` gives you bytes and works everywhere:

```js
input.addEventListener('change', async () => {
  const file = input.files[0];
  if (!file) return;

  const bytes = new Uint8Array(await file.arrayBuffer());
  for await (const row of readSheet(bytes, { header: true })) {
    console.log(row);
  }
});
```

In 0.1.0 the `Source` type does not include `Blob` even though the runtime accepts it
([known type gaps](api.md#known-type-gaps)) — `new Uint8Array(await file.arrayBuffer())` keeps
TypeScript happy and copies nothing extra.

## Complete Vite + React example

`npm install excel-lite react react-dom`, then:

```jsx
// src/WorkbookTool.jsx
import { useRef, useState } from 'react';
import { readSheet, toXlsx } from 'excel-lite';

export default function WorkbookTool() {
  const inputRef = useRef(null);
  const [rows, setRows] = useState([]);
  const [error, setError] = useState(null);

  async function onFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    setError(null);

    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const parsed = [];
      for await (const row of readSheet(bytes, { header: true })) parsed.push(row);
      setRows(parsed);
    } catch (err) {
      setRows([]);
      setError(err.message); // "not a zip", missing workbook.xml, 512 MiB guard, …
    }
  }

  function download() {
    const bytes = toXlsx([
      {
        name: 'Export',
        rows,
        columns: [
          { key: 'email', header: 'Email', width: 32 },
          { key: 'amount', header: 'Amount', format: '#,##0.00' },
          { key: 'joined', header: 'Joined', format: 'yyyy-mm-dd' },
        ],
        freeze: 'A2',
      },
    ]);

    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'export.xlsx';
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  return (
    <div>
      <input ref={inputRef} type="file" accept=".xlsx,.xlsm" onChange={onFile} />
      {error && <p role="alert">{error}</p>}
      <p>{rows.length} rows loaded</p>
      <button onClick={download} disabled={rows.length === 0}>
        Download .xlsx
      </button>
    </div>
  );
}
```

Notes:

- Nothing Node-specific is imported; the same component works in a Vite, CRA, Next.js (client
  component) or Parcel app.
- Errors from the reader are thrown while iterating, so the whole loop is inside `try`.
- `URL.revokeObjectURL` after `click()` prevents a memory leak; the download is triggered by a
  synthetic anchor, which works in all evergreen browsers.

## Bundle size

Measured on this repository with `esbuild` 0.27.7:

| Bundle | Size |
|---|---|
| library alone, `platform: browser`, ESM, unminified | **58,492 B** (57.1 KiB) |
| a consumer entry `import { readSheet, toXlsx } from 'excel-lite'`, unminified | **57,869 B** (56.5 KiB) |
| the same, minified | **31,248 B** (30.5 KiB) |
| the same, minified + gzip | **13,123 B** (12.8 KiB) |
| `npm run check:bundle` report | 57.1 kB, no Node builtins |

Reproduce:

```bash
# library bundle
npx esbuild dist/index.js --bundle --platform=browser --format=esm --outfile=/tmp/excel-lite.js
wc -c /tmp/excel-lite.js

# what an app actually pays (run from the package directory)
printf "import { readSheet, toXlsx } from 'excel-lite';\nconsole.log(typeof readSheet, typeof toXlsx);\n" \
  | npx esbuild --bundle --platform=browser --format=esm --minify --outfile=/tmp/app.js
wc -c /tmp/app.js
```

The whole cost is `excel-lite` plus `fflate` (the zip implementation). There is no Node polyfill,
no XML library, no ICU data. CI enforces a 120 kB ceiling on the browser bundle, so a dependency
that creeps in fails the build.

## Workers and edge runtimes

The same byte API works in a web worker — useful for parsing a large file off the main thread:

```js
// worker.js
import { readSheet } from 'excel-lite';

self.onmessage = async (event) => {
  const rows = [];
  for await (const row of readSheet(event.data.bytes, { header: true })) rows.push(row);
  self.postMessage(rows);
};
```

```js
const worker = new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });
const bytes = new Uint8Array(await file.arrayBuffer());
worker.postMessage({ bytes }, [bytes.buffer]); // transfer, do not copy
```

Cloudflare Workers, Deno Deploy, Vercel Edge and similar runtimes have `Uint8Array`, `Blob`,
`TextEncoder`/`TextDecoder` and `fetch` — enough for everything except paths. If `process` happens
to exist but `getBuiltinModule` does not, path calls still throw the same clear error.

## Browser limits

- **Memory.** The workbook is held in memory (the file bytes plus the inflated sheet XML, see
  [performance](performance.md)). A 50 MB upload in a tab can cost hundreds of MB; cap uploads with
  `file.size` before parsing.
- **No file paths.** There is no filesystem; use the File System Access API or a download as above.
- **No ZIP password support.** Encrypted workbooks are rejected as "not a zip".
- **ZIP64 edge cases.** `fflate` handles ZIP64; a single entry over 512 MiB is refused by design.
- **Streaming.** `fetch(...).body` can be passed as an async iterable, but it is collected into one
  buffer before parsing; it is not a streaming parser ([reading](reading.md#sources)).

## See also

- [Reading](reading.md#sources) — the byte sources in detail.
- [Recipes](recipes.md#8-browser-download) — download button in context.
- [Troubleshooting](troubleshooting.md#bundler-errors-about-nodefs) — when a bundler complains.
- [Architecture](architecture.md) — `src/fs.ts` and the bundle-safety check.
