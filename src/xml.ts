/**
 * Minimal XML reader for OOXML parts.
 *
 * OOXML files are machine generated, so a full XML implementation is
 * unnecessary. What we need is: a small DOM for the metadata parts
 * (workbook, rels, styles, sharedStrings) and a fast row scanner for
 * worksheet parts, which can be hundreds of megabytes.
 */

export interface XmlNode {
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

const NAMED_ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

const ENTITY_RE = /&(#[0-9]+|#x[0-9a-fA-F]+|[a-zA-Z]+);/g;

/** Decode the five XML entities plus numeric character references. */
export function decodeEntities(input: string): string {
  if (input.indexOf('&') === -1) return input;
  return input.replace(ENTITY_RE, (match, body: string) => {
    if (body.charCodeAt(0) === 35 /* # */) {
      const code =
        body.charCodeAt(1) === 120 || body.charCodeAt(1) === 88 /* xX */
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : match;
    }
    const named = NAMED_ENTITIES[body];
    return named === undefined ? match : named;
  });
}

/**
 * Escape a string for use as XML text or an attribute value.
 * Characters that XML 1.0 forbids are dropped: Excel refuses to open a
 * workbook that contains them.
 */
export function escapeXml(input: string): string {
  let out = '';
  for (let i = 0; i < input.length; i++) {
    const ch = input[i];
    switch (ch) {
      case '&':
        out += '&amp;';
        break;
      case '<':
        out += '&lt;';
        break;
      case '>':
        out += '&gt;';
        break;
      case '"':
        out += '&quot;';
        break;
      case "'":
        out += '&apos;';
        break;
      default: {
        const code = input.charCodeAt(i);
        if (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) break; // drop
        if (code === 0xfffe || code === 0xffff) break;
        out += ch;
      }
    }
  }
  return out;
}

/** `x:sheet` -> `sheet` (OOXML files usually have no prefix, but some tools do). */
export function localName(name: string): string {
  const colon = name.indexOf(':');
  return colon === -1 ? name : name.slice(colon + 1);
}

const ATTR_RE = /([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

export function parseAttrs(source: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (!source) return attrs;
  ATTR_RE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = ATTR_RE.exec(source))) {
    const value = match[2] !== undefined ? match[2] : match[3]!;
    attrs[localName(match[1])] = decodeEntities(value);
  }
  return attrs;
}

const START_TAG_RE = /<(\/?)([A-Za-z_][\w.:-]*)((?:\s+[\w.:-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;

/** Parse a (small) XML document into a node tree. */
export function parseXml(source: string): XmlNode {
  const root: XmlNode = { name: '#root', attrs: {}, children: [], text: '' };
  const stack: XmlNode[] = [root];
  let cursor = 0;

  while (cursor < source.length) {
    const lt = source.indexOf('<', cursor);
    if (lt === -1) {
      stack[stack.length - 1].text += decodeEntities(source.slice(cursor));
      break;
    }
    if (lt > cursor) {
      stack[stack.length - 1].text += decodeEntities(source.slice(cursor, lt));
    }

    if (source.startsWith('<!--', lt)) {
      const end = source.indexOf('-->', lt + 4);
      cursor = end === -1 ? source.length : end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', lt)) {
      const end = source.indexOf(']]>', lt + 9);
      const body = source.slice(lt + 9, end === -1 ? source.length : end);
      stack[stack.length - 1].text += body;
      cursor = end === -1 ? source.length : end + 3;
      continue;
    }
    if (source[lt + 1] === '?' || source[lt + 1] === '!') {
      const end = source.indexOf('>', lt + 1);
      cursor = end === -1 ? source.length : end + 1;
      continue;
    }

    START_TAG_RE.lastIndex = lt;
    const match = START_TAG_RE.exec(source);
    if (!match || match.index !== lt) {
      // Not a tag we understand (stray "<"), treat it as text and move on.
      stack[stack.length - 1].text += '<';
      cursor = lt + 1;
      continue;
    }

    const [, closing, name, attrSource, selfClosing] = match;
    if (closing) {
      if (stack.length > 1) stack.pop();
    } else {
      const node: XmlNode = { name, attrs: parseAttrs(attrSource ?? ''), children: [], text: '' };
      stack[stack.length - 1].children.push(node);
      if (!selfClosing) stack.push(node);
    }
    cursor = lt + match[0].length;
  }

  return root;
}

/** First direct child with the given local name. */
export function findChild(node: XmlNode | undefined, local: string): XmlNode | undefined {
  if (!node) return undefined;
  for (const child of node.children) {
    if (localName(child.name) === local) return child;
  }
  return undefined;
}

/** All direct children with the given local name. */
export function findChildren(node: XmlNode | undefined, local: string): XmlNode[] {
  if (!node) return [];
  return node.children.filter((child) => localName(child.name) === local);
}

/** Text of a node including all descendants (OOXML splits rich text into runs). */
export function textContent(node: XmlNode | undefined): string {
  if (!node) return '';
  let out = node.text;
  for (const child of node.children) out += textContent(child);
  return out;
}

// ---------------------------------------------------------------------------
// Worksheet row scanner
// ---------------------------------------------------------------------------

export interface RawCell {
  /** Cell reference such as `B7`, if the producer wrote one. */
  ref: string;
  /** `t` attribute: n (number), s (shared string), str, inlineStr, b, e, d. */
  type: string;
  /** Index into the workbook style table, 0 when absent. */
  style: number;
  /** Raw `<v>` text, or the concatenated `<t>` runs for inline strings. */
  value: string;
}

export interface RawRow {
  /** 1-based row number from the `r` attribute, or -1 when missing. */
  rowNumber: number;
  cells: RawCell[];
}

const ROW_RE = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
const CELL_RE = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
const VALUE_RE = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/;
const TEXT_RUN_RE = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;

/**
 * Lazily yield rows from a `<sheetData>` document. Regex driven on purpose:
 * building a DOM for a 400k row sheet would cost more memory than the file.
 */
export function* scanRows(xml: string): Generator<RawRow> {
  ROW_RE.lastIndex = 0;
  let rowMatch: RegExpExecArray | null;
  while ((rowMatch = ROW_RE.exec(xml))) {
    const rowAttrs = parseAttrs(rowMatch[1] ?? '');
    const body = rowMatch[2];
    const cells: RawCell[] = [];

    if (body) {
      CELL_RE.lastIndex = 0;
      let cellMatch: RegExpExecArray | null;
      while ((cellMatch = CELL_RE.exec(body))) {
        const cellAttrs = parseAttrs(cellMatch[1] ?? '');
        const inner = cellMatch[2] ?? '';
        const type = cellAttrs.t ?? 'n';

        let value = '';
        const valueMatch = inner ? VALUE_RE.exec(inner) : null;
        if (valueMatch) value = decodeEntities(valueMatch[1]);

        if (type === 'inlineStr' || (inner !== '' && inner.indexOf('<is') !== -1)) {
          const runs: string[] = [];
          TEXT_RUN_RE.lastIndex = 0;
          let run: RegExpExecArray | null;
          while ((run = TEXT_RUN_RE.exec(inner))) runs.push(decodeEntities(run[1]));
          value = runs.join('');
        }

        cells.push({
          ref: cellAttrs.r ?? '',
          type,
          style: cellAttrs.s ? Number(cellAttrs.s) || 0 : 0,
          value,
        });
      }
    }

    yield { rowNumber: rowAttrs.r ? Number(rowAttrs.r) : -1, cells };
  }
}

/** Count `<row>` elements without materialising cells (`<rowBreaks>` does not count). */
export function countRows(xml: string): number {
  let count = 0;
  const rowStart = /<row(?=[\s/>])/g;
  while (rowStart.exec(xml)) count++;
  return count;
}
