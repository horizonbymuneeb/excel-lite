# Recipes

Complete, working patterns for things applications actually do with `.xlsx`. Each recipe names
the packages it needs beyond `excel-lite`.

Recipes that need a package this repository does not depend on (`express`, `multer`, `pg`,
`pg-copy-streams`, `zod`, `@aws-sdk/client-s3`) are complete sketches: the `excel-lite` calls in
them are the same ones the runnable examples execute, but the surrounding framework code was not
run here. The recipes with a script in [`examples/`](../examples/README.md) were executed against
this build.

## On this page

- [1. Express multipart upload → rows](#1-express-multipart-upload-to-rows)
- [2. Next.js route handler → rows](#2-nextjs-route-handler-to-rows)
- [3. Fetch body → rows](#3-fetch-body-to-rows)
- [4. CSV → xlsx](#4-csv-to-xlsx)
- [5. xlsx → CSV / NDJSON](#5-xlsx-to-csv-and-ndjson)
- [6. xlsx → Postgres COPY](#6-xlsx-to-postgres-copy)
- [7. Scheduled styled report](#7-scheduled-styled-report)
- [8. Browser download](#8-browser-download)
- [9. Stream from S3 / Cloudflare R2](#9-stream-from-s3-and-cloudflare-r2)
- [10. Validate rows with zod](#10-validate-rows-with-zod)
- [11. ExcelJS interop](#11-exceljs-interop)
- [12. openpyxl interop](#12-openpyxl-interop)

## 1. Express multipart upload to rows

Packages: `express`, `multer`.

```js
import express from 'express';
import multer from 'multer';
import { readSheet } from 'excel-lite';

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 }, // cap untrusted uploads (see SECURITY.md)
});

app.post('/import', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'file field is required' });

  const rows = [];
  const errors = [];

  try {
    for await (const row of readSheet(req.file.buffer, { header: true })) {
      if (typeof row.email !== 'string') {
        errors.push({ row: rows.length + 2, reason: 'email is not text' });
        continue;
      }
      rows.push({ email: row.email, amount: Number(row.amount ?? 0) });
    }
  } catch (error) {
    // not a zip, missing workbook.xml, bad sheet name, 512 MiB limit, …
    return res.status(422).json({ error: error.message });
  }

  res.json({ imported: rows.length, errors });
});

app.listen(3000);
```

Notes: `multer.memoryStorage()` hands you a `Buffer`, which `readSheet()` accepts directly — no
temp file. Parse errors are thrown while iterating, so keep the loop inside `try`. If you prefer
disk storage, `readSheet(req.file.path)` also works (Node 22.3+).

## 2. Next.js route handler to rows

Packages: none beyond Next.js.

```js
// app/api/import/route.js  (App Router, Node runtime)
import { readSheet, toXlsx } from 'excel-lite';

export const runtime = 'nodejs';

export async function POST(request) {
  const form = await request.formData();
  const file = form.get('file');
  if (!(file instanceof File)) {
    return Response.json({ error: 'file field is required' }, { status: 400 });
  }

  const bytes = new Uint8Array(await file.arrayBuffer());
  const good = [];
  const bad = [];

  try {
    for await (const row of readSheet(bytes, { header: true })) {
      const amount = Number(row.amount);
      if (!Number.isFinite(amount)) bad.push({ email: row.email, amount: row.amount });
      else good.push({ email: row.email, amount });
    }
  } catch (error) {
    return Response.json({ error: error.message }, { status: 422 });
  }

  // hand back the rejected rows as a workbook
  const report = toXlsx([
    { name: 'Rejected', rows: bad, columns: [{ key: 'email' }, { key: 'amount' }] },
  ]);

  return new Response(report, {
    headers: {
      'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'content-disposition': 'attachment; filename="rejected.xlsx"',
    },
  });
}
```

`File` is a `Blob`, but the cleanest path is `arrayBuffer()` → `Uint8Array`; that also keeps the
code working on edge runtimes where `fs` does not exist.

## 3. Fetch body to rows

Works in browsers, Node 18+, and edge runtimes: a fetch `Response.body` is an async iterable of
chunks.

```js
import { readSheet } from 'excel-lite';

const response = await fetch('https://files.example.com/report.xlsx');
if (!response.ok) throw new Error(`download failed: ${response.status}`);

// pass the body straight through; the chunks are collected into one buffer
for await (const row of readSheet(response.body, { header: true })) {
  console.log(row.id, row.amount);
}
```

If the server supports it, prefer `arrayBuffer()` — the body is collected anyway, so this just
skips one layer:

```js
for await (const row of readSheet(new Uint8Array(await response.arrayBuffer()))) { /* … */ }
```

## 4. CSV to xlsx

Packages: none (the parser below is deliberately small). For real CSV dialects use `csv-parse`.

```js
import { writeSheet } from 'excel-lite';

const csv = 'sku,price,note\nA-1,19.99,"widget, large"\nA-2,5,simple\n';

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows.filter((fields) => fields.some((value) => value.trim() !== ''));
}

const [head, ...data] = parseCsv(csv);
const rows = data.map((fields) =>
  Object.fromEntries(
    head.map((key, index) => {
      const raw = fields[index] ?? '';
      return [key, /^-?\d+(\.\d+)?$/.test(raw) ? Number(raw) : raw];
    }),
  ),
);

await writeSheet('from-csv.xlsx', rows, {
  columns: [
    { key: 'sku', width: 10 },
    { key: 'price', format: '#,##0.00' },
    { key: 'note', width: 24 },
  ],
  freeze: 'A2',
});
```

A verified version with assertions is [`examples/csv-bridge.mjs`](../examples/csv-bridge.mjs).

## 5. xlsx to CSV and NDJSON

Packages: none.

```js
import { readSheet } from 'excel-lite';

function csvField(value) {
  const text = value instanceof Date ? value.toISOString() : value == null ? '' : String(value);
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

async function toCsv(file) {
  let header = null;
  let out = '';
  for await (const row of readSheet(file, { header: true })) {
    if (!header) {
      header = Object.keys(row);
      out += header.map(csvField).join(',') + '\n';
    }
    out += header.map((key) => csvField(row[key])).join(',') + '\n';
  }
  return out;
}

async function toNdjson(file) {
  let out = '';
  for await (const row of readSheet(file, { header: true })) out += JSON.stringify(row) + '\n';
  return out;
}

console.log(await toCsv('report.xlsx'));
console.log(await toNdjson('report.xlsx'));
```

Both accumulate the whole file in a string. For big exports write chunks to a stream instead
(`process.stdout.write`/`res.write` inside the loop), or use the CLI (`to-csv --out`).

## 6. xlsx to Postgres COPY

Packages: `pg`, `pg-copy-streams` (or `pg-copy-streams`' `to` for the reverse direction).

Use Postgres's server-side `COPY … FROM STDIN` with CSV — it is far faster than
`INSERT … VALUES` per row, and `excel-lite` can stream CSV rows into it.

```sql
-- exact SQL: create the table once
CREATE TABLE report (
  name     text,
  qty      integer,
  price    numeric(12,2),
  "when"   timestamptz,
  active   boolean,
  note     text
);

-- the load itself
COPY report (name, qty, price, "when", active, note)
FROM STDIN
WITH (FORMAT csv, HEADER true, DELIMITER ',', ENCODING 'UTF8', NULL '');
```

```js
import { pipeline } from 'node:stream/promises';
import pg from 'pg';
import { from as copyFrom } from 'pg-copy-streams';
import { readSheet } from 'excel-lite';

const client = new pg.Client({ connectionString: process.env.DATABASE_URL });
await client.connect();

function csvField(value) {
  const text = value instanceof Date ? value.toISOString() : value == null ? '' : String(value);
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

// A readable stream of CSV lines in the exact column order of the COPY statement.
const source = async function* () {
  let header = null;
  const names = ['name', 'qty', 'price', 'when', 'active', 'note'];
  for await (const row of readSheet('report.xlsx', { header: true })) {
    if (!header) {
      header = true;
      yield names.map(csvField).join(',') + '\n'; // HEADER true
    }
    yield names.map((name) => csvField(row[name])).join(',') + '\n';
  }
};

const sink = client.query(copyFrom('COPY report (name, qty, price, "when", active, note) FROM STDIN WITH (FORMAT csv, HEADER true, NULL \'\')'));
await pipeline(source, sink);

console.log('copied', sink.rowCount);
await client.end();
```

Notes:

- The CSV column order must match the `COPY` column list; `--header` object rows are unordered,
  so always write an explicit list.
- `NULL ''` makes an empty field a SQL `NULL`. A literal empty string cannot be distinguished from
  `NULL` in CSV — `excel-lite` writes empty cells as empty fields and empty strings as no cell, so
  both land as `NULL`, which is usually what you want.
- Dates are ISO strings; a `timestamptz` column accepts them. For a `date` column Postgres
  accepts the date part of the ISO string.
- If you would rather not add `pg-copy-streams`, load a temp table with `FOR EACH ROW` inserts —
  it is the same data, slower.

The CLI equivalent, if the CSV column order already matches the table:

```bash
npx excel-lite to-csv report.xlsx --header \
  | psql "$DATABASE_URL" -c "COPY report FROM STDIN WITH (FORMAT csv, HEADER true)"
```

## 7. Scheduled styled report

Packages: none for the report; scheduling is your OS cron (or `node-cron` if you prefer).

`report.mjs`:

```js
import { mkdirSync } from 'node:fs';
import { writeWorkbook } from 'excel-lite';

// Replace this with your query; the rest of the script does not depend on it.
const rows = await loadRows();

const stamp = new Date().toISOString().slice(0, 10);
mkdirSync('/srv/reports', { recursive: true });

const file = `/srv/reports/revenue-${stamp}.xlsx`;
await writeWorkbook(file, [
  {
    name: 'Revenue',
    rows,
    freeze: 'A2',
    autoFilter: true,
    columns: [
      { key: 'region', header: 'Region', width: 16, bold: true, fill: '1F2937' },
      { key: 'units', header: 'Units', format: '#,##0', align: 'right' },
      { key: 'revenue', header: 'Revenue', format: '#,##0.00', align: 'right' },
      { key: 'margin', header: 'Margin', format: '0.0%', align: 'right' },
      { key: 'updated', header: 'Updated', format: 'yyyy-mm-dd' },
    ],
  },
  { name: 'Notes', rows: [{ note: `generated ${new Date().toISOString()}` }], columns: [{ key: 'note', width: 60 }] },
]);
console.log('wrote', file);
```

Crontab (07:00 on weekdays):

```cron
0 7 * * 1-5 cd /srv/app && /usr/bin/node report.mjs >> /var/log/reports.log 2>&1
```

Notes: the `updated` column needs `Date` values to get a real date format; a number in a
date-formatted column is written as `General` (see [writing](writing.md#dates)). Overwrites are
silent — write to a temp path and rename if a reader may be watching the file.

## 8. Browser download

Packages: your framework (React shown). No Node APIs.

```jsx
import { useState } from 'react';
import { toXlsx } from 'excel-lite';

export function ExportButton({ rows }) {
  const [busy, setBusy] = useState(false);

  function download() {
    setBusy(true);
    try {
      const bytes = toXlsx([
        {
          name: 'Export',
          rows,
          columns: [
            { key: 'email', header: 'Email', width: 32 },
            { key: 'amount', header: 'Amount', format: '#,##0.00' },
            { key: 'joined', header: 'Joined', format: 'yyyy-mm-dd' },
          ],
        },
      ]);

      const url = URL.createObjectURL(
        new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      );
      const link = document.createElement('a');
      link.href = url;
      link.download = 'export.xlsx';
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } finally {
      setBusy(false);
    }
  }

  return (
    <button onClick={download} disabled={busy}>
      {busy ? 'Building…' : 'Download .xlsx'}
    </button>
  );
}
```

`new Blob([bytes])` is the only browser API involved; the byte array is what `toXlsx()` already
returns. See [browser](browser.md) for the upload side and bundle notes.

## 9. Stream from S3 and Cloudflare R2

Packages: `@aws-sdk/client-s3` (`excel-lite` never talks to the network itself).

```js
import { S3Client, GetObjectCommand, PutObjectCommand } from '@aws-sdk/client-s3';
import { readSheet, toXlsx } from 'excel-lite';

// S3
const s3 = new S3Client({ region: process.env.AWS_REGION });

// Cloudflare R2: same SDK, different endpoint
// const r2 = new S3Client({
//   region: 'auto',
//   endpoint: `https://${process.env.R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
//   credentials: { accessKeyId: process.env.R2_ACCESS_KEY_ID, secretAccessKey: process.env.R2_SECRET_ACCESS_KEY },
// });

const object = await s3.send(new GetObjectCommand({ Bucket: 'reports', Key: '2026-01/report.xlsx' }));

const good = [];
for await (const row of readSheet(object.Body, { header: true })) {
  // object.Body is a Node Readable: an async iterable of Uint8Array chunks
  if (row.amount > 0) good.push(row);
}

await s3.send(
  new PutObjectCommand({
    Bucket: 'reports',
    Key: '2026-01/positive.xlsx',
    Body: toXlsx([{ name: 'Positive', rows: good }]),
    ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  }),
);
```

Notes: the whole object is downloaded and unzipped before rows are yielded — memory tracks the
decompressed sheet, not the S3 object ([performance](performance.md)). For very large objects,
download to disk first or use a real streaming reader. In the browser, use `GetObjectCommand`
with `@aws-sdk/client-s3` and pass `object.Body.transformToByteArray()` instead of the stream.

## 10. Validate rows with zod

Packages: `zod`.

```js
import { z } from 'zod';
import { readSheet, toXlsx, writeSheet } from 'excel-lite';

const Row = z.object({
  email: z.string().email(),
  amount: z.coerce.number().finite().nonnegative(),
  joined: z.date().optional(),
});

const accepted = [];
const rejected = [];

for await (const row of readSheet('upload.xlsx', { header: true })) {
  const parsed = Row.safeParse(row);
  if (parsed.success) accepted.push(parsed.data);
  else rejected.push({ ...row, error: parsed.error.issues.map((issue) => `${issue.path.join('.')}: ${issue.message}`).join('; ') });
}

await writeSheet('accepted.xlsx', accepted, {
  columns: [{ key: 'email' , width: 32 }, { key: 'amount', format: '#,##0.00' }, { key: 'joined', format: 'yyyy-mm-dd' }],
});

// Keep the rejected rows as text so no value is lost in the report.
await writeSheet('rejected.xlsx', rejected, {
  columns: [{ key: 'email' }, { key: 'amount' }, { key: 'joined' }, { key: 'error', width: 60 }],
});

const bytes = toXlsx([{ name: 'Rejected', rows: rejected }]); // or ship the bytes to the client
void bytes;
```

Notes: `row` values are `Date` objects for date-formatted cells, so `z.date()` is right for those
(not `z.string()`). Missing cells are `null`; use `.nullable()` or `z.coerce` deliberately. Keep
the original row in the rejection record — the parsed object is what zod rebuilt, not what the
user sent.

## 11. ExcelJS interop

Packages: `exceljs`.

```js
import ExcelJS from 'exceljs';
import { readSheet, writeSheet } from 'excel-lite';

const rows = [
  { name: 'Ada', amount: 1234.5 },
  { name: 'Linus', amount: 99 },
];

// excel-lite -> ExcelJS: read the values, add what excel-lite cannot write (borders, images, merges)
await writeSheet('base.xlsx', rows, {
  columns: [{ key: 'name', header: 'Name' }, { key: 'amount', header: 'Amount', format: '#,##0.00' }],
  freeze: 'A2',
});

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('base.xlsx');
const ws = wb.getWorksheet(1);
ws.getColumn(1).border = { bottom: { style: 'thin' } };
await wb.xlsx.writeFile('styled.xlsx');

// ExcelJS -> excel-lite: read values from a workbook ExcelJS produced
const source = new ExcelJS.Workbook();
await source.xlsx.readFile('from-exceljs.xlsx');
const data = source.getWorksheet(1).getSheetValues().filter(Boolean); // arrays, row 1 first

// equivalent, using excel-lite's reader directly:
for await (const row of readSheet('from-exceljs.xlsx', { header: true })) console.log(row);
```

Verified both directions in [`examples/exceljs-interop.mjs`](../examples/exceljs-interop.mjs)
(which also asserts the styles ExcelJS sees). excel-lite does not read formatting back, so use
ExcelJS when the workbook is the thing you are editing; use excel-lite when the rows are.

## 12. openpyxl interop

`excel-lite` reads fixtures generated by openpyxl in its own test suite, and its output is checked
with openpyxl by `scripts/verify-with-openpyxl.py`. In your own project:

**Python reads what excel-lite wrote:**

```python
# read_ours.py
from openpyxl import load_workbook

wb = load_workbook("report.xlsx")
print(wb.sheetnames)
for row in wb["Sheet1"].iter_rows(values_only=True):
    print(row)
```

**Python writes a file excel-lite reads** (dates as real dates, numbers as numbers):

```python
# write_for_js.py
import datetime as dt
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.append(["name", "amount", "when"])
ws.append(["Ada", 1234.5, dt.date(2024, 5, 17)])
ws["C2"].number_format = "yyyy-mm-dd"
wb.save("from-openpyxl.xlsx")
```

```js
import { readSheet } from 'excel-lite';

for await (const row of readSheet('from-openpyxl.xlsx', { header: true })) {
  console.log(row.name, row.amount, row.when); // Ada 1234.5 2024-05-17T00:00:00.000Z
}
```

Notes: openpyxl writes `datetime.date`/`datetime.datetime` as serials with date formats, which is
exactly what the reader expects. A `datetime.time` becomes a serial below 1 and reads back as a
`Date` on 1899-12-30 — read the UTC time parts ([reading](reading.md#dates-times-and-timezones)).

## See also

- [`examples/`](../examples/README.md) — runnable versions of several recipes.
- [Reading](reading.md), [Writing](writing.md) — the rules behind each behaviour.
- [Troubleshooting](troubleshooting.md) — when one of these throws.
