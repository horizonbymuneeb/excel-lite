import { unzipSync, zipSync } from 'fflate';

const decoder = new TextDecoder('utf-8');

const BOM = '\ufeff';

export function decodeUtf8(bytes: Uint8Array): string {
  const text = decoder.decode(bytes);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

export function encodeUtf8(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

/** Inflate only the entries we asked for; everything else stays compressed. */
export function unzipEntries(data: Uint8Array, names: readonly string[]): Record<string, Uint8Array> {
  const wanted = new Set(names);
  return unzipSync(data, {
    filter: (file) => wanted.has(file.name),
  });
}

export function zipFiles(files: Record<string, Uint8Array>): Uint8Array {
  return zipSync(files, { level: 6 });
}

export { BOM };
