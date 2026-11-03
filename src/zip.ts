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

/**
 * Hard ceiling for a single entry's uncompressed size. A workbook that claims to
 * inflate past this is either corrupt or a zip bomb; refusing is cheaper than an OOM.
 * See SECURITY.md.
 */
const MAX_ENTRY_BYTES = 512 * 1024 * 1024;

/** Inflate only the entries we asked for; everything else stays compressed. */
export function unzipEntries(data: Uint8Array, names: readonly string[]): Record<string, Uint8Array> {
  const wanted = new Set(names);
  return unzipSync(data, {
    filter: (file) => {
      if (!wanted.has(file.name)) return false;
      if (file.originalSize > MAX_ENTRY_BYTES) {
        const size = (file.originalSize / 1024 / 1024).toFixed(0);
        throw new Error(
          `excel-lite: ${file.name} inflates to ${size} MiB, above the 512 MiB safety limit — refusing to decompress (see SECURITY.md)`,
        );
      }
      return true;
    },
  });
}

export function zipFiles(files: Record<string, Uint8Array>): Uint8Array {
  return zipSync(files, { level: 6 });
}

export { BOM };
