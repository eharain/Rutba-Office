'use strict';

/**
 * An invisible text layer on the pages of a scanned PDF, written as an
 * incremental update: the words text recognition found are set over each
 * page in render mode 3 (neither filled nor stroked), where they cannot be
 * seen but can be found, selected and copied, as every reader does with a
 * recognised scan. The original bytes stay the exact start of the file and
 * nothing of a page is redrawn: each page named is redefined (the same
 * object number, later in the file) with one content stream more, and a
 * font for the layer is added to its resources.
 *
 * Unlike incremental.js's overlays this works from reader.js, so it takes
 * what scanners and modern tools write: cross-reference streams (the
 * update then ends with one too), object streams, page trees of any shape,
 * inherited resources and rotated pages (the words come in the page's own
 * coordinates, so a rotation is the page's business, not the layer's).
 * Encrypted files and files whose cross-references had to be rebuilt are
 * refused by name: an update chained to a table that was wrong would be
 * wrong too.
 *
 * The font is the one OCR layers have always used: a Type0 font whose
 * codes are the text's UTF-16 code units, mapped back to Unicode by its
 * ToUnicode CMap, over an embedded TrueType font with one empty glyph half
 * an em wide. Each word is placed with its own text matrix, stretched to
 * the width of the word on the page, so selecting it highlights the word.
 * Right-to-left words are written in the order they are drawn, as a PDF
 * writes Urdu and Arabic, and readers put them back in reading order.
 */

const zlib = require('zlib');
const { openPdf, Name, Ref, Dict, Stream } = require('./reader');
const { valueText } = require('./assemble');
const { visualText, hasRtl } = require('./bidi');

class TextLayerError extends Error {
  /** code: 'encrypted' | 'damaged' | 'page' */
  constructor(message, code) {
    super(message);
    this.name = 'PdfTextLayerError';
    this.code = code;
  }
}

// ── the font ───────────────────────────────────────────────────────────────

/** A TrueType font of two empty glyphs, each half an em (500 of 1000) wide. */
function glyphlessFont() {
  const tables = {};
  const u16 = (v) => { const b = Buffer.alloc(2); b.writeUInt16BE(v & 0xffff); return b; };
  const i16 = (v) => { const b = Buffer.alloc(2); b.writeInt16BE(v); return b; };
  const u32 = (v) => { const b = Buffer.alloc(4); b.writeUInt32BE(v >>> 0); return b; };
  tables.head = Buffer.concat([
    u32(0x00010000), u32(0x00010000), u32(0), u32(0x5f0f3cf5), u16(0x000b), u16(1000),
    Buffer.alloc(16), i16(0), i16(-200), i16(500), i16(800), u16(0), u16(8), i16(2), i16(0), i16(0),
  ]);
  tables.hhea = Buffer.concat([
    u32(0x00010000), i16(800), i16(-200), i16(0), u16(500), i16(0), i16(0), i16(500), i16(1), i16(0), i16(0),
    Buffer.alloc(8), i16(0), u16(2),
  ]);
  tables.maxp = Buffer.concat([u32(0x00010000), u16(2), u16(0), u16(0), u16(0), u16(0), u16(2), u16(0), u16(0), u16(0), u16(0), u16(0), u16(0), u16(0), u16(0)]);
  tables.hmtx = Buffer.concat([u16(500), i16(0), u16(500), i16(0)]);
  tables.loca = Buffer.concat([u16(0), u16(0), u16(0)]);
  tables.glyf = Buffer.alloc(0);
  // cmap: format 4 with only the closing segment; codes reach glyphs through the PDF's own map.
  tables.cmap = Buffer.concat([u16(0), u16(1), u16(3), u16(1), u32(12), u16(4), u16(24), u16(0), u16(2), u16(2), u16(0), u16(0), u16(0xffff), u16(0), u16(0xffff), u16(1), u16(0)]);
  tables.post = Buffer.concat([u32(0x00030000), u32(0), i16(-100), i16(50), u32(0), u32(0), u32(0), u32(0), u32(0)]);
  const tags = Object.keys(tables).sort();
  const header = Buffer.concat([u32(0x00010000), u16(tags.length), u16(64), u16(2), u16(tags.length * 16 - 64)]);
  let offset = 12 + tags.length * 16;
  const dir = [];
  const bodies = [];
  const sum = (b) => {
    const p = Buffer.concat([b, Buffer.alloc((4 - (b.length % 4)) % 4)]);
    let s = 0;
    for (let i = 0; i < p.length; i += 4) s = (s + p.readUInt32BE(i)) >>> 0;
    return s;
  };
  for (const tag of tags) {
    const body = tables[tag];
    dir.push(Buffer.concat([Buffer.from(tag, 'latin1'), u32(sum(body)), u32(offset), u32(body.length)]));
    const padded = Buffer.concat([body, Buffer.alloc((4 - (body.length % 4)) % 4)]);
    bodies.push(padded);
    offset += padded.length;
  }
  const font = Buffer.concat([header, ...dir, ...bodies]);
  // head.checkSumAdjustment: 0xB1B0AFBA less the sum of the whole font.
  const headAt = 12 + tags.length * 16 + bodies.slice(0, tags.indexOf('head')).reduce((n, b) => n + b.length, 0);
  font.writeUInt32BE((0xb1b0afba - sum(font)) >>> 0, headAt + 8);
  return font;
}

/** Every code from 0000 to FFFF is its own UTF-16 code unit. */
function toUnicodeCMap() {
  const ranges = [];
  for (let hi = 0; hi < 256; hi++) {
    if (hi >= 0xd8 && hi <= 0xdf) continue; // surrogates are never written alone
    const h = hi.toString(16).padStart(2, '0').toUpperCase();
    ranges.push(`<${h}00> <${h}FF> <${h}00>`);
  }
  const blocks = [];
  for (let i = 0; i < ranges.length; i += 100) {
    const part = ranges.slice(i, i + 100);
    blocks.push(`${part.length} beginbfrange\n${part.join('\n')}\nendbfrange`);
  }
  return `/CIDInit /ProcSet findresource begin
12 dict begin
begincmap
/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def
/CMapName /Adobe-Identity-UCS def
/CMapType 2 def
1 begincodespacerange
<0000> <FFFF>
endcodespacerange
${blocks.join('\n')}
endcmap
CMapName currentdict /CMap defineresource pop
end
end
`;
}

// ── the words ──────────────────────────────────────────────────────────────

const CJK = /[⺀-ヿ㐀-䶿一-鿿豈-﫿＀-￯]/u;
const num = (n) => {
  if (!Number.isFinite(n)) return '0';
  const s = (Math.round(n * 1000) / 1000).toString();
  return s === '-0' ? '0' : s;
};

/** UTF-16 code units of the text, as the layer's font codes. Characters outside the BMP are left out. */
function codes(text) {
  const out = [];
  for (const ch of text) {
    const cp = ch.codePointAt(0);
    if (cp <= 0xffff && !(cp >= 0xd800 && cp <= 0xdfff)) out.push(cp);
  }
  return out;
}

/**
 * The content stream for one page's words.
 * @param {{ text: string, x: number, y: number, width: number, size: number, angle?: number, space?: boolean }[]} words
 *   each word's baseline start (x, y) and width along the baseline in the
 *   page's default user space, its height `size`, the baseline's angle in
 *   degrees (0 along x), and whether a space follows it on its line
 */
function layerContent(words, fontName) {
  const ops = ['Q', 'BT', '3 Tr', `/${fontName} 1 Tf`];
  for (const w of words) {
    const text = String(w.text || '').trim();
    const shown = hasRtl(text) ? visualText(text) : text;
    const units = codes(shown);
    if (!units.length || !(w.width > 0) || !(w.size > 0)) continue;
    // Stretched so the word's own units, at half an em each, span its width; a following space goes past it.
    const scale = w.width / (units.length * 0.5);
    if (w.space) units.push(0x20);
    const a = ((w.angle || 0) * Math.PI) / 180;
    const cos = Math.cos(a), sin = Math.sin(a);
    ops.push(`${num(scale * cos)} ${num(scale * sin)} ${num(-w.size * sin)} ${num(w.size * cos)} ${num(w.x)} ${num(w.y)} Tm`);
    ops.push(`<${units.map((u) => u.toString(16).padStart(4, '0')).join('')}> Tj`);
  }
  ops.push('ET');
  return ops.join('\n');
}

/**
 * Lines of words (each in reading order) as the layer's words, a space after
 * every word but the last of its line and between Chinese or Japanese.
 */
function wordsOfLines(lines) {
  const out = [];
  for (const line of lines || []) {
    const ws = line.words || [];
    ws.forEach((w, i) => {
      const next = ws[i + 1];
      const space = Boolean(next) && !(CJK.test(String(w.text).slice(-1)) && CJK.test(String(next.text)[0]));
      out.push({ ...w, space });
    });
  }
  return out;
}

// ── the update ─────────────────────────────────────────────────────────────

/**
 * Add an invisible text layer to pages of a PDF.
 *
 * @param {Buffer|Uint8Array} original the whole file
 * @param {{ page: number, lines: { words: object[] }[] }[]} pages 0-based
 *   page numbers, and on each the lines of words (see layerContent for a
 *   word) in the page's default user space
 * @returns {{ bytes: Buffer, pages: number[] }} the file with the update
 *   appended, and the pages it wrote a layer on
 */
function addTextLayer(original, pages) {
  const buf = Buffer.isBuffer(original) ? original : Buffer.from(original.buffer, original.byteOffset, original.byteLength);
  const file = openPdf(buf);
  if (file.encrypted) throw new TextLayerError('A text layer cannot be added to a PDF protected with a password.', 'encrypted');
  if (file.repaired) throw new TextLayerError('This PDF\'s cross-references are damaged, so a text layer cannot be added to it as it is. Save a copy first, then recognise the copy.', 'damaged');

  const all = file.pages();
  const jobs = (pages || []).filter((p) => p && wordsOfLines(p.lines).length);
  for (const p of jobs) {
    if (!all[p.page]) throw new TextLayerError(`There is no page ${Number(p.page) + 1} in this PDF.`, 'page');
    if (!all[p.page].ref) throw new TextLayerError(`Page ${Number(p.page) + 1} cannot be redefined.`, 'page');
  }
  if (!jobs.length) return { bytes: buf, pages: [] };

  const trailerSize = Number(file.trailer.get('Size')) || 0;
  let next = Math.max(trailerSize, ...[...file.entries.keys()].map((n) => n + 1), 1);
  const objects = []; // { num, gen, body: Buffer }
  const add = (body, num = next++, gen = 0) => {
    objects.push({ num, gen, body: Buffer.isBuffer(body) ? body : Buffer.from(body, 'latin1') });
    return num;
  };
  const stream = (dict, data) => Buffer.concat([
    Buffer.from(`${dict.replace(/>>$/, '')} /Length ${data.length} >>\nstream\n`, 'latin1'), data, Buffer.from('\nendstream', 'latin1'),
  ]);

  // The font, once for every page.
  const program = glyphlessFont();
  const fontFileNum = add(stream(`<< /Filter /FlateDecode /Length1 ${program.length} >>`, zlib.deflateSync(program)));
  const descriptorNum = add(`<< /Type /FontDescriptor /FontName /GlyphLessFont /Flags 5 /FontBBox [0 -200 500 800] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 800 /StemV 80 /FontFile2 ${fontFileNum} 0 R >>`);
  const cidToGid = zlib.deflateSync(Buffer.alloc(65536 * 2, 0).fill(Buffer.from([0, 1])));
  const cidToGidNum = add(stream('<< /Filter /FlateDecode >>', cidToGid));
  const cidNum = add(`<< /Type /Font /Subtype /CIDFontType2 /BaseFont /GlyphLessFont /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${descriptorNum} 0 R /DW 500 /CIDToGIDMap ${cidToGidNum} 0 R >>`);
  const cmap = zlib.deflateSync(Buffer.from(toUnicodeCMap(), 'latin1'));
  const toUnicodeNum = add(stream('<< /Filter /FlateDecode >>', cmap));
  const fontNum = add(`<< /Type /Font /Subtype /Type0 /BaseFont /GlyphLessFont /Encoding /Identity-H /DescendantFonts [${cidNum} 0 R] /ToUnicode ${toUnicodeNum} 0 R >>`);
  // Saves the original's graphics state, so whatever its content leaves behind cannot move the layer.
  const saveNum = add(stream('<< >>', Buffer.from('q', 'latin1')));

  const written = [];
  for (const job of jobs) {
    const page = all[job.page];
    // The page's resources, its own or inherited, as a direct dictionary with the layer's font added.
    const resources = new Dict();
    const own = file.resolve(page.dict.has('Resources') ? page.dict.get('Resources') : page.inherited.Resources);
    if (own instanceof Dict) for (const [k, v] of own.entries()) resources.set(k, v);
    const fonts = new Dict();
    const ownFonts = file.resolve(resources.get('Font'));
    if (ownFonts instanceof Dict) for (const [k, v] of ownFonts.entries()) fonts.set(k, v);
    let fontName = 'ROCR';
    for (let i = 1; fonts.has(fontName); i++) fontName = `ROCR${i}`;
    fonts.set(fontName, new Ref(fontNum, 0));
    resources.set('Font', fonts);

    const content = zlib.deflateSync(Buffer.from(layerContent(wordsOfLines(job.lines), fontName), 'latin1'));
    const contentNum = add(stream('<< /Filter /FlateDecode >>', content));

    const dict = new Dict();
    for (const [k, v] of page.dict.entries()) dict.set(k, v);
    // What the page inherited is written on it, as assemble.js does, so the redefined page needs nothing above it.
    for (const key of ['MediaBox', 'CropBox', 'Rotate']) if (!dict.has(key) && page.inherited[key] !== undefined) dict.set(key, page.inherited[key]);
    const contents = file.resolve(page.dict.get('Contents'));
    const list = Array.isArray(contents) ? contents : page.dict.get('Contents') ? [page.dict.get('Contents')] : [];
    dict.set('Contents', [new Ref(saveNum, 0), ...list, new Ref(contentNum, 0)]);
    dict.set('Resources', resources);
    add(`${valueText(dict)}`, page.ref.num, page.ref.gen);
    written.push(job.page);
  }

  // The update: the objects, then cross-references in the form the file's last section has.
  const base = file.base;
  const prevAt = file.startXref() - base;
  const classic = buf.toString('latin1', file.startXref(), file.startXref() + 4) === 'xref';
  const parts = [buf];
  let length = buf.length;
  const push = (data) => {
    const b = Buffer.isBuffer(data) ? data : Buffer.from(data, 'latin1');
    parts.push(b);
    length += b.length;
  };
  if (buf[buf.length - 1] !== 0x0a && buf[buf.length - 1] !== 0x0d) push('\n');
  const offsets = new Map();
  for (const o of objects) {
    offsets.set(o.num, { offset: length - base, gen: o.gen });
    push(`${o.num} ${o.gen} obj\n`);
    push(o.body);
    push('\nendobj\n');
  }
  const trailer = new Dict();
  for (const key of ['Root', 'Info', 'ID']) if (file.trailer.has(key)) trailer.set(key, file.trailer.get(key));
  trailer.set('Prev', prevAt);

  // Runs of consecutive object numbers, as both forms of cross-reference list them.
  const runs = (nums) => {
    const out = [];
    for (const n of nums.sort((a, b) => a - b)) {
      const last = out[out.length - 1];
      if (last && last.first + last.count === n) last.count++;
      else out.push({ first: n, count: 1 });
    }
    return out;
  };

  if (classic) {
    const size = Math.max(trailerSize, next);
    trailer.set('Size', size);
    const xrefAt = length - base;
    let xref = 'xref\n';
    for (const run of runs([...offsets.keys()])) {
      xref += `${run.first} ${run.count}\n`;
      for (let n = run.first; n < run.first + run.count; n++) {
        const e = offsets.get(n);
        xref += `${String(e.offset).padStart(10, '0')} ${String(e.gen).padStart(5, '0')} n \n`;
      }
    }
    push(xref);
    push(`trailer\n${valueText(trailer)}\nstartxref\n${xrefAt}\n%%EOF\n`);
  } else {
    // A cross-reference stream: its own number last, rows of type, offset (4 bytes) and generation (2).
    const xrefNum = next++;
    const xrefAt = length - base;
    offsets.set(xrefNum, { offset: xrefAt, gen: 0 });
    const list = runs([...offsets.keys()]);
    const rows = [];
    for (const run of list) {
      for (let n = run.first; n < run.first + run.count; n++) {
        const e = offsets.get(n);
        const row = Buffer.alloc(7);
        row[0] = 1;
        row.writeUInt32BE(e.offset, 1);
        row.writeUInt16BE(e.gen, 5);
        rows.push(row);
      }
    }
    const data = zlib.deflateSync(Buffer.concat(rows));
    trailer.set('Type', new Name('XRef'));
    trailer.set('Size', Math.max(trailerSize, next));
    trailer.set('W', [1, 4, 2]);
    trailer.set('Index', list.flatMap((r) => [r.first, r.count]));
    trailer.set('Filter', new Name('FlateDecode'));
    trailer.set('Length', data.length);
    push(`${xrefNum} 0 obj\n${valueText(trailer)}\nstream\n`);
    push(data);
    push(`\nendstream\nendobj\nstartxref\n${xrefAt}\n%%EOF\n`);
  }
  return { bytes: Buffer.concat(parts, length), pages: written };
}

/** Whether a page already carries text a reader could find (its content shows text at all). */
function pageHasText(file, index) {
  return /\bBT\b[\s\S]*?\bT[jJ]\b/.test(file.pageContent(index).toString('latin1'));
}

module.exports = { addTextLayer, pageHasText, layerContent, wordsOfLines, glyphlessFont, toUnicodeCMap, TextLayerError };
