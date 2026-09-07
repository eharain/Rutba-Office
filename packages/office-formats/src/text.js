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
  const text = decodeText(input).replace(/\r\n?/g, '\n');
  const lines = text.split('\n');
  const blocks = [];
  let i = 0;

  const inline = (s) => {
    const runs = [];
    const re = /(\*\*|__)(.+?)\1|(\*|_)(.+?)\3|`([^`]+)`|\[([^\]]+)\]\(([^)]+)\)/g;
    let last = 0;
    let m;
    while ((m = re.exec(s))) {
      if (m.index > last) runs.push({ text: s.slice(last, m.index) });
      if (m[2] != null) runs.push({ text: m[2], bold: true });
      else if (m[4] != null) runs.push({ text: m[4], italic: true });
      else if (m[5] != null) runs.push({ text: m[5], code: true });
      else if (m[6] != null) runs.push({ text: m[6], link: m[7] });
      last = re.lastIndex;
    }
    if (last < s.length) runs.push({ text: s.slice(last) });
    return runs.length ? runs : [{ text: s }];
  };

  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) {
      i++;
      continue;
    }
    let m;
    if ((m = /^(#{1,6})\s+(.*)$/.exec(line))) {
      blocks.push({ type: 'heading', level: m[1].length, runs: inline(m[2].trim()), text: m[2].trim() });
      i++;
      continue;
    }
    if (/^```/.test(line)) {
      const lang = line.slice(3).trim();
      const body = [];
      i++;
      while (i < lines.length && !/^```/.test(lines[i])) body.push(lines[i++]);
      i++;
      blocks.push({ type: 'code', language: lang, text: body.join('\n') });
      continue;
    }
    if (/^(\*\s*){3,}$|^(-\s*){3,}$|^(_\s*){3,}$/.test(line.trim())) {
      blocks.push({ type: 'rule' });
      i++;
      continue;
    }
    if (/^\s*\|.*\|\s*$/.test(line) && /^\s*\|?[\s:-]*\|[\s:|-]*$/.test(lines[i + 1] || '')) {
      const cells = (l) => l.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());
      const header = cells(line);
      i += 2;
      const rows = [header];
      while (i < lines.length && /\|/.test(lines[i]) && lines[i].trim()) rows.push(cells(lines[i++]));
      blocks.push({ type: 'table', header: true, rows: rows.map((r) => r.map((c) => ({ runs: inline(c), text: c }))) });
      continue;
    }
    if (/^\s*>\s?/.test(line)) {
      const body = [];
      while (i < lines.length && /^\s*>\s?/.test(lines[i])) body.push(lines[i++].replace(/^\s*>\s?/, ''));
      blocks.push({ type: 'quote', runs: inline(body.join(' ')), text: body.join(' ') });
      continue;
    }
    if (/^\s*([-*+]|\d+[.)])\s+/.test(line)) {
      const ordered = /^\s*\d/.test(line);
      const items = [];
      while (i < lines.length && /^\s*([-*+]|\d+[.)])\s+/.test(lines[i])) {
        const raw = lines[i].replace(/^\s*([-*+]|\d+[.)])\s+/, '');
        const indent = (/^\s*/.exec(lines[i])[0] || '').length;
        items.push({ text: raw, runs: inline(raw), level: Math.floor(indent / 2) });
        i++;
      }
      blocks.push({ type: 'list', ordered, items });
      continue;
    }
    const para = [];
    while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|```|\s*[-*+]\s|\s*\d+[.)]\s|\s*>)/.test(lines[i])) {
      para.push(lines[i++]);
    }
    const joined = para.join(' ').trim();
    blocks.push({ type: 'paragraph', runs: inline(joined), text: joined });
  }
  return { blocks };
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
export function writeMarkdown(blocks) {
  const runText = (runs) =>
    (runs || [])
      .map((r) => {
        let t = r.text ?? '';
        if (r.code) t = `\`${t}\``;
        if (r.bold) t = `**${t}**`;
        if (r.italic) t = `*${t}*`;
        if (r.link) t = `[${t}](${r.link})`;
        return t;
      })
      .join('');
  const out = [];
  for (const b of blocks) {
    if (b.type === 'heading') out.push(`${'#'.repeat(b.level || 1)} ${b.text ?? runText(b.runs)}`);
    else if (b.type === 'list') out.push(b.items.map((it, n) => `${b.ordered ? `${n + 1}.` : '-'} ${it.text ?? runText(it.runs)}`).join('\n'));
    else if (b.type === 'quote') out.push(`> ${b.text ?? runText(b.runs)}`);
    else if (b.type === 'code') out.push(`\`\`\`${b.language || ''}\n${b.text}\n\`\`\``);
    else if (b.type === 'rule') out.push('---');
    else if (b.type === 'table') {
      const rows = b.rows.map((r) => `| ${r.map((c) => c.text ?? runText(c.runs)).join(' | ')} |`);
      if (rows.length) rows.splice(1, 0, `| ${b.rows[0].map(() => '---').join(' | ')} |`);
      out.push(rows.join('\n'));
    } else out.push(b.text ?? runText(b.runs));
  }
  return out.join('\n\n') + '\n';
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
