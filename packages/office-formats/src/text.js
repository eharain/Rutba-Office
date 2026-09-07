import { parseMarkdown, serializeMarkdown } from './markdown.js';

// The plain formats: delimited data, Markdown, plain text.
//
// They look trivial until you meet real files — a CSV whose quoted field spans
// three lines, a UTF-16 export from a bank, a Markdown table with no trailing
// pipe. These are the parts that decide whether "open anything" is true.

/** Strip a byte-order mark and decode, honouring UTF-16 when the BOM says so. */
export function decodeText(bytes) {
  if (typeof bytes === 'string') return bytes.replace(/^﻿/, '');
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (b.length >= 2 && b[0] === 0xff && b[1] === 0xfe) return new TextDecoder('utf-16le').decode(b.subarray(2));
  if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) return new TextDecoder('utf-16be').decode(b.subarray(2));
  if (b.length >= 3 && b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return new TextDecoder('utf-8').decode(b.subarray(3));
  // No BOM. A file with many NULs at even offsets is UTF-16 that forgot one.
  let nuls = 0;
  const probe = Math.min(b.length, 512);
  for (let i = 1; i < probe; i += 2) if (b[i] === 0) nuls++;
  if (probe > 16 && nuls / (probe / 2) > 0.7) return new TextDecoder('utf-16le').decode(b);
  const utf8 = new TextDecoder('utf-8', { fatal: false });
  const text = utf8.decode(b);
  // U+FFFD everywhere means it was never UTF-8; Windows-1252 is the safe guess.
  const bad = (text.match(/�/g) || []).length;
  if (bad > text.length * 0.02) return new TextDecoder('windows-1252').decode(b);
  return text;
}

/** Guess the delimiter from the first lines: the one with a consistent count. */
export function sniffDelimiter(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.length).slice(0, 20);
  if (!lines.length) return ',';
  let best = ',';
  let bestScore = -1;
  for (const d of [',', '\t', ';', '|']) {
    const counts = lines.map((l) => splitLine(l, d).length);
    const first = counts[0];
    if (first < 2) continue;
    const consistent = counts.filter((c) => c === first).length / counts.length;
    const score = consistent * 10 + Math.min(first, 20) / 20;
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

function splitLine(line, delim) {
  // Only used for sniffing, so quotes are approximated rather than parsed.
  const out = [];
  let cur = '';
  let q = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') q = !q;
    else if (c === delim && !q) {
      out.push(cur);
      cur = '';
    } else cur += c;
  }
  out.push(cur);
  return out;
}

/**
 * Parse delimited text into rows. Handles quoted fields containing the
 * delimiter, doubled quotes, and newlines inside quotes.
 * @returns {{ rows: string[][], delimiter: string }}
 */
export function readDelimited(input, { delimiter } = {}) {
  const text = decodeText(input);
  const d = delimiter || sniffDelimiter(text);
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  let i = 0;
  const n = text.length;
  while (i < n) {
    const c = text[i];
    if (quoted) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += c;
      i++;
      continue;
    }
    if (c === '"' && field === '') {
      quoted = true;
      i++;
      continue;
    }
    if (c === d) {
      row.push(field);
      field = '';
      i++;
      continue;
    }
    if (c === '\r') {
      i++;
      continue;
    }
    if (c === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      i++;
      continue;
    }
    field += c;
    i++;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  while (rows.length && rows[rows.length - 1].every((v) => v === '')) rows.pop();
  return { rows, delimiter: d };
}

/** Serialise rows back out, quoting only what needs it. */
export function writeDelimited(rows, { delimiter = ',', eol = '\r\n' } = {}) {
  const needsQuote = (v) => v.includes(delimiter) || v.includes('"') || v.includes('\n') || v.includes('\r');
  return rows
    .map((r) => r.map((v) => {
      const s = v == null ? '' : String(v);
      return needsQuote(s) ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(delimiter))
    .join(eol);
}

/**
 * Markdown to the same block model the other readers produce. A deliberately
 * small subset — headings, paragraphs, lists, quotes, code, tables, rules and
 * inline emphasis — because this exists to open notes, not to be a CommonMark
 * implementation.
 */
export function readMarkdown(input) {
  return parseMarkdown(decodeText(input));
}

/** Plain text: one paragraph per blank-line-separated chunk, hard breaks kept. */
export function readPlain(input) {
  const text = decodeText(input).replace(/\r\n?/g, '\n');
  const blocks = text.split(/\n{2,}/).map((chunk) => ({
    type: 'paragraph',
    text: chunk,
    runs: [{ text: chunk }],
  }));
  return { blocks: blocks.length ? blocks : [{ type: 'paragraph', text: '', runs: [] }] };
}

/** Blocks back to Markdown, for Save As. */
export function writeMarkdown(blocks, meta) {
  return serializeMarkdown(blocks, meta);
}

/** Blocks to plain text. */
export function writePlain(blocks) {
  return blocks
    .map((b) =>
      b.type === 'table'
        ? b.rows.map((r) => r.map((c) => c.text ?? '').join('\t')).join('\n')
        : b.type === 'list'
          ? b.items.map((i) => `- ${i.text ?? ''}`).join('\n')
          : b.text ?? (b.runs || []).map((r) => r.text).join('')
    )
    .join('\n\n') + '\n';
}
