/**
 * File-system access without a static `node:fs` import.
 *
 * Bundlers targeting the browser refuse to resolve `node:` specifiers, so a
 * library that touches fs has to keep that import out of the graph entirely.
 * `process.getBuiltinModule('fs/promises')` is a Node API that needs no import
 * and simply does not exist in browsers — so the same bundle works in both,
 * and only file paths require Node.
 */

interface FsPromises {
  readFile(path: string): Promise<Uint8Array>;
  writeFile(path: string, data: Uint8Array): Promise<void>;
}

function nodeFs(): FsPromises | null {
  const proc = (globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } }).process;
  if (typeof proc?.getBuiltinModule !== 'function') return null;
  try {
    return (proc.getBuiltinModule('fs/promises') as FsPromises | undefined) ?? null;
  } catch {
    return null;
  }
}

function requireFs(): FsPromises {
  const fs = nodeFs();
  if (!fs) {
    throw new Error(
      'excel-lite: reading or writing by file path needs Node 22+. ' +
        'In browsers and older runtimes pass bytes instead: readSheet(uint8Array) / toXlsx(sheets).',
    );
  }
  return fs;
}

export async function readFileBytes(path: string): Promise<Uint8Array> {
  const fs = requireFs();
  const data = await fs.readFile(path);
  if (data instanceof Uint8Array) return data;
  return new Uint8Array(data as ArrayBuffer);
}

export async function writeFileBytes(path: string, bytes: Uint8Array): Promise<void> {
  await requireFs().writeFile(path, bytes);
}
