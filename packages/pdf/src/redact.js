'use strict';

/**
 * Redaction that removes what it covers, not a black box drawn over it.
 *
 * The areas to redact are the Redact annotations a person marked (each its
 * QuadPoints, or its Rect), or areas given. Applying them rewrites each page
 * that has one:
 *
 * - Its content is read operation by operation with the text state kept
 *   (fonts and their widths, character and word spacing, scaling, rise, the
 *   text and transformation matrices), so the box of every glyph drawn is
 *   known on the page. A glyph whose box falls in an area is taken out of
 *   the string that draws it, and the space it took is kept by a number in
 *   the TJ array, so the words either side of it stay where they were. This
 *   is as true of invisible text (a recognised scan's layer, render mode 3)
 *   as of text that shows.
 * - A picture drawn under an area has the pixels in the area painted out and
 *   is written again as a new image (pictures this module can decode: no
 *   compression, or Flate, 8 bits a sample, grey or RGB); a picture it
 *   cannot decode (JPEG, JPEG 2000, fax and JBIG2 scans) is painted out by
 *   pixels the caller decodes and hands back (`images`), and is otherwise
 *   left out of the page entirely. An inline picture under an area is left
 *   out. A form XObject under an area is redacted in the same way, as a new
 *   object, so other pages that draw the same form are left alone.
 * - Marked content that carried the words as /ActualText or /Alt loses them
 *   when glyphs inside it are removed.
 * - Every annotation that overlaps an area goes (comments over the words,
 *   links, form widgets), the Redact marks with them, and each area is
 *   filled black, as the mark's /IC asks (or its colour), with its overlay
 *   text if it has one.
 *
 * Nothing is written as an incremental update: an update keeps the old
 * bytes, the words among them. The file is written whole by assemble.js,
 * which copies only what the pages still reach, so the old content streams,
 * pictures and annotations are not in the new file at all.
 *
 * The words a mark covered are learnt as the glyphs go (each font's
 * /ToUnicode, or its single-byte encoding), with the words the mark itself
 * records (its /Contents: the words selected, or the words searched for),
 * and the rewrite takes them out of every other place a person could find
 * them (assemble.js's `scrub`, from `scrubber` here): the document's title,
 * subject, keywords, author and creator; XMP metadata; bookmarks; comments
 * and their replies; links; attached files; form values, tooltips and field
 * names; page labels. Each such thing goes whole (a comment, a bookmark, a
 * link) or is blanked (a form value, a label's prefix), and is counted by
 * kind for the window to list before anything is saved.
 *
 * What is not removed: vector line art under an area (it is covered by the
 * fill, not cut away), and the words where they stand unmarked in a page's
 * content (mark every match to take them all).
 */

const zlib = require('zlib');
const {
  Lexer, Name, Ref, Dict, Stream, PdfString, Keyword, parseValue, PdfReadError,
} = require('./reader');
const { assemble, valueText, textString } = require('./assemble');
const { readAnnotations } = require('./annotations');
const metrics = require('./metrics');
const { encode, literal } = require('./encoding');

class RedactError extends Error {
  constructor(message, code = 'redact') {
    super(message);
    this.name = 'PdfRedactError';
    this.code = code;
  }
}

const num = (v, fallback = 0) => (typeof v === 'number' && Number.isFinite(v) ? v : fallback);
const n = (v) => {
  if (!Number.isFinite(v)) return '0';
  const s = (Math.round(v * 1000) / 1000).toString();
  return s === '-0' ? '0' : s;
};

// ── geometry ──────────────────────────────────────────────────────────────

/** a × b, matrices as [a b c d e f]. */
const mul = (a, b) => [
  a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
  a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
  a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5],
];
const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
function invert(m) {
  const det = m[0] * m[3] - m[1] * m[2];
  if (!det) return null;
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det, (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
}
/** The box on the page of a unit-relative rectangle [x0, y0, x1, y1] under matrix m. */
function boxOf(m, x0, y0, x1, y1) {
  const pts = [apply(m, x0, y0), apply(m, x1, y0), apply(m, x0, y1), apply(m, x1, y1)];
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}
const area = (b) => Math.max(0, b[2] - b[0]) * Math.max(0, b[3] - b[1]);
const overlap = (a, b) => area([Math.max(a[0], b[0]), Math.max(a[1], b[1]), Math.min(a[2], b[2]), Math.min(a[3], b[3])]);
const intersects = (a, b) => a[0] < b[2] && b[0] < a[2] && a[1] < b[3] && b[1] < a[3];

/** Whether a glyph's box is in one of the areas: its centre inside one, or a fifth of it under one. */
function glyphHit(box, areas) {
  const cx = (box[0] + box[2]) / 2;
  const cy = (box[1] + box[3]) / 2;
  const a = area(box);
  for (const r of areas) {
    if (cx >= r[0] && cx <= r[2] && cy >= r[1] && cy <= r[3]) return true;
    if (a > 0 && overlap(box, r) >= a * 0.2) return true;
  }
  return false;
}

// ── fonts ─────────────────────────────────────────────────────────────────

/** Widths and code lengths of a font, enough to place its glyphs. */
/** A /ToUnicode CMap's codes and their text: bfchar and bfrange, both forms. */
function parseToUnicode(data) {
  const map = new Map();
  const text = data.toString('latin1');
  const utf16 = (h) => {
    let s = '';
    for (let i = 0; i + 3 < h.length; i += 4) s += String.fromCharCode(parseInt(h.slice(i, i + 4), 16));
    return s;
  };
  for (const block of text.matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const m of block[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]*)>/g)) map.set(parseInt(m[1], 16), utf16(m[2]));
  }
  for (const block of text.matchAll(/beginbfrange([\s\S]*?)endbfrange/g)) {
    for (const m of block[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(?:<([0-9a-fA-F]+)>|\[([^\]]*)\])/g)) {
      const a = parseInt(m[1], 16);
      const b = parseInt(m[2], 16);
      if (!(b >= a) || b - a > 65535) continue;
      if (m[3] !== undefined) {
        const head = m[3].slice(0, -4);
        const last = parseInt(m[3].slice(-4) || '0', 16);
        for (let c = a; c <= b; c++) map.set(c, utf16(head + (last + c - a).toString(16).padStart(4, '0')));
      } else {
        const items = [...m[4].matchAll(/<([0-9a-fA-F]*)>/g)].map((x) => utf16(x[1]));
        for (let c = a; c <= b && c - a < items.length; c++) map.set(c, items[c - a]);
      }
    }
  }
  return map;
}

// Windows-1252's 0x80 to 0x9F, as the reader reads them.
const HIGH = '\u20ac\u0081\u201a\u0192\u201e\u2026\u2020\u2021\u02c6\u2030\u0160\u2039\u0152\u008d\u017d\u008f\u0090\u2018\u2019\u201c\u201d\u2022\u2013\u2014\u02dc\u2122\u0161\u203a\u0153\u009d\u017e\u0178';

function fontInfo(file, fontDict) {
  const font = file.resolve(fontDict);
  const info = { twoByte: false, width: () => 0.5, scale: 0.001, type3: false, unicode: (code) => (code >= 0x80 && code <= 0x9f ? HIGH[code - 0x80] : String.fromCharCode(code)) };
  if (!(font instanceof Dict)) return info;
  const subtype = font.get('Subtype') instanceof Name ? font.get('Subtype').name : '';
  const cmap = file.resolve(font.get('ToUnicode'));
  if (cmap instanceof Stream) {
    try {
      const map = parseToUnicode(file.decode(cmap));
      const fallback = info.unicode;
      if (map.size) info.unicode = (code) => (map.has(code) ? map.get(code) : subtype === 'Type0' ? '' : fallback(code));
    } catch { /* the encoding's own letters */ }
  } else if (subtype === 'Type0') {
    info.unicode = () => '';
  }
  if (subtype === 'Type0') {
    info.twoByte = true;
    const desc = file.resolve((file.resolve(font.get('DescendantFonts')) || [])[0]);
    const dw = desc instanceof Dict ? num(file.resolve(desc.get('DW')), 1000) : 1000;
    const widths = new Map();
    const w = desc instanceof Dict ? file.resolve(desc.get('W')) : null;
    if (Array.isArray(w)) {
      for (let i = 0; i < w.length;) {
        const first = num(file.resolve(w[i]));
        const next = file.resolve(w[i + 1]);
        if (Array.isArray(next)) {
          next.forEach((x, k) => widths.set(first + k, num(file.resolve(x))));
          i += 2;
        } else {
          const last = num(next);
          const value = num(file.resolve(w[i + 2]));
          for (let c = first; c <= last && c - first < 65536; c++) widths.set(c, value);
          i += 3;
        }
      }
    }
    info.width = (code) => (widths.has(code) ? widths.get(code) : dw) / 1000;
    return info;
  }
  const first = num(file.resolve(font.get('FirstChar')), 0);
  const list = file.resolve(font.get('Widths'));
  const descriptor = file.resolve(font.get('FontDescriptor'));
  const missing = descriptor instanceof Dict ? num(file.resolve(descriptor.get('MissingWidth')), 0) : 0;
  let scale = 0.001;
  if (subtype === 'Type3') {
    const fm = file.resolve(font.get('FontMatrix'));
    scale = Array.isArray(fm) ? Math.abs(num(file.resolve(fm[0]), 0.001)) : 0.001;
    info.type3 = true;
  }
  if (Array.isArray(list) && list.length) {
    const widths = list.map((x) => num(file.resolve(x)));
    info.width = (code) => {
      const v = widths[code - first];
      return (v === undefined ? missing : v) * scale;
    };
    return info;
  }
  // A standard font without widths: those the package knows, else a fair average.
  const base = font.get('BaseFont') instanceof Name ? font.get('BaseFont').name.replace(/^[A-Z]{6}\+/, '') : '';
  const known = Object.keys(metrics.FONTS).find((k) => k === base) || (/^Courier/.test(base) ? 'Courier' : /^(Helvetica|Arial)/.test(base) ? (/Bold/.test(base) ? 'Helvetica-Bold' : 'Helvetica') : null);
  info.width = known ? (code) => metrics.byteWidth(known, code) / 1000 : () => 0.55;
  return info;
}

/** A string's codes in a font: one byte each, or two for a composite font. */
function codesOf(bytes, twoByte) {
  const out = [];
  if (twoByte) for (let i = 0; i + 1 < bytes.length; i += 2) out.push({ code: (bytes[i] << 8) | bytes[i + 1], bytes: bytes.subarray(i, i + 2) });
  else for (let i = 0; i < bytes.length; i++) out.push({ code: bytes[i], bytes: bytes.subarray(i, i + 1) });
  return out;
}

// ── content: read, and written again ─────────────────────────────────────

/** A content stream as operations: { op, operands } or { raw } for an inline image. */
function parseContent(data) {
  const lx = new Lexer(data, 0);
  const ops = [];
  let operands = [];
  for (let guard = 0; guard < 2e7; guard++) {
    lx.skipSpace();
    const at = lx.pos;
    const tok = lx.next();
    if (tok && tok.eof) break;
    if (tok === '[' || tok === '<<') {
      lx.pos = at;
      operands.push(parseValue(lx, false));
      continue;
    }
    if (typeof tok === 'string') continue; // a stray ']' or '>>'
    if (tok instanceof Keyword) {
      if (tok.word === 'BI') {
        // An inline image, kept as the bytes it is, to EI.
        const id = data.indexOf('ID', lx.pos, 'latin1');
        let end = id < 0 ? data.length : id + 3;
        for (;;) {
          end = data.indexOf('EI', end, 'latin1');
          if (end < 0) { end = data.length; break; }
          const before = data[end - 1];
          const after = data[end + 2];
          if ((before === 32 || before === 10 || before === 13 || before === 9) && (end + 2 >= data.length || after === 32 || after === 10 || after === 13 || after === 9)) break;
          end += 2;
        }
        lx.pos = Math.min(data.length, end + 2);
        ops.push({ op: 'BI', raw: data.subarray(at, lx.pos), operands: [] });
        operands = [];
        continue;
      }
      ops.push({ op: tok.word, operands });
      operands = [];
      continue;
    }
    operands.push(tok);
  }
  return ops;
}

function writeContent(ops) {
  const parts = [];
  for (const o of ops) {
    if (o.raw) { parts.push(Buffer.from(o.raw)); parts.push(Buffer.from('\n')); continue; }
    const words = o.operands.map((v) => valueText(v));
    words.push(o.op);
    parts.push(Buffer.from(`${words.join(' ')}\n`, 'latin1'));
  }
  return Buffer.concat(parts);
}

// ── pictures ──────────────────────────────────────────────────────────────

/** An image's samples, when this module can read them: { width, height, comps, data } or null. */
function readPixels(file, stream) {
  const d = stream.dict;
  const filters = file.resolve(d.get('Filter'));
  const list = (Array.isArray(filters) ? filters : filters ? [filters] : []).map((f) => (f instanceof Name ? f.name : ''));
  if (list.some((f) => !['FlateDecode', 'Fl', 'LZWDecode', 'LZW', 'ASCIIHexDecode', 'AHx', 'ASCII85Decode', 'A85', 'RunLengthDecode', 'RL'].includes(f))) return null;
  if (num(file.resolve(d.get('BitsPerComponent')), 8) !== 8 || file.resolve(d.get('ImageMask')) === true) return null;
  let cs = file.resolve(d.get('ColorSpace'));
  let comps = 0;
  if (cs instanceof Name) comps = { DeviceRGB: 3, RGB: 3, DeviceGray: 1, G: 1, DeviceCMYK: 4, CMYK: 4 }[cs.name] || 0;
  else if (Array.isArray(cs) && cs[0] instanceof Name && cs[0].name === 'ICCBased') {
    const icc = file.resolve(cs[1]);
    comps = icc instanceof Stream ? num(file.resolve(icc.dict.get('N')), 0) : 0;
  }
  if (!comps) return null;
  const width = num(file.resolve(d.get('Width')));
  const height = num(file.resolve(d.get('Height')));
  let data;
  try { data = file.decode(stream); } catch { return null; }
  if (!(width > 0) || !(height > 0) || data.length < width * height * comps) return null;
  return { width, height, comps, data: Buffer.from(data.subarray(0, width * height * comps)) };
}

/** The pixels of an image under the areas, in the image's own pixel grid: [[x0, y0, x1, y1], ...], rows from the top. */
function pixelRects(ctm, width, height, areas) {
  const inv = invert(ctm);
  if (!inv) return [];
  const out = [];
  for (const r of areas) {
    const b = boxOf(inv, r[0], r[1], r[2], r[3]); // in the unit square
    const x0 = Math.max(0, Math.floor(b[0] * width));
    const x1 = Math.min(width, Math.ceil(b[2] * width));
    const y0 = Math.max(0, Math.floor((1 - b[3]) * height));
    const y1 = Math.min(height, Math.ceil((1 - b[1]) * height));
    if (x1 > x0 && y1 > y0) out.push([x0, y0, x1, y1]);
  }
  return out;
}

/** An image's samples with rectangles of pixels painted out (black, or the redaction's colour). */
function paintOut(pix, rects, fill = [0, 0, 0]) {
  const { width, comps, data } = pix;
  const value = comps === 1 ? [Math.round(fill[0] * 255)] : comps === 4 ? [0, 0, 0, 255] : fill.map((c) => Math.round(c * 255));
  for (const [x0, y0, x1, y1] of rects) {
    for (let y = y0; y < y1; y++) {
      for (let x = x0; x < x1; x++) {
        const at = (y * width + x) * comps;
        for (let c = 0; c < comps; c++) data[at + c] = value[c];
      }
    }
  }
  return data;
}

function imageStream(file, original, pix, data) {
  const dict = new Dict();
  for (const [k, v] of original.dict.entries()) {
    if (['Filter', 'DecodeParms', 'Length', 'DL'].includes(k)) continue;
    dict.set(k, v);
  }
  dict.set('Width', pix.width);
  dict.set('Height', pix.height);
  dict.set('BitsPerComponent', 8);
  if (pix.colorSpace) dict.set('ColorSpace', new Name(pix.colorSpace));
  dict.set('Filter', new Name('FlateDecode'));
  return new Stream(dict, zlib.deflateSync(data));
}

// ── a page ────────────────────────────────────────────────────────────────

/**
 * Redact one content stream (a page's, or a form XObject's) under `areas`,
 * starting from matrix `base`. Returns { data, resources, removed }, with
 * `resources` a new dictionary when pictures or forms in it were replaced.
 */
function redactContent(ctx, data, resources, base, depth = 0) {
  const { file, areas, removed, plan, images } = ctx;
  const ops = parseContent(data);
  const res = file.resolve(resources);
  const fonts = res instanceof Dict ? file.resolve(res.get('Font')) : null;
  const xobjects = res instanceof Dict ? file.resolve(res.get('XObject')) : null;
  const newX = new Map(); // fresh name -> replacement (Stream)
  const gone = new Set(); // names a picture or form was drawn under, painted out or left out
  const fontCache = new Map();
  const fontOf = (name) => {
    if (!fontCache.has(name)) fontCache.set(name, fontInfo(file, fonts instanceof Dict ? fonts.get(name) : null));
    return fontCache.get(name);
  };

  let gs = { ctm: base.slice(), font: null, size: 0, tc: 0, tw: 0, th: 1, tl: 0, rise: 0 };
  const stack = [];
  let tm = [1, 0, 0, 1, 0, 0];
  let tlm = [1, 0, 0, 1, 0, 0];
  const marked = []; // BDC/BMC: { op, hit }
  const out = [];

  const show = (o, parts) => {
    // parts: PdfStrings and numbers, as in TJ. Returns the new array and whether a glyph went.
    const font = gs.font ? fontOf(gs.font) : fontInfo(file, null);
    const size = gs.size;
    const result = [];
    let cut = false;
    for (const part of parts) {
      if (typeof part === 'number') {
        result.push(part);
        tm = mul([1, 0, 0, 1, (-part / 1000) * size * gs.th, 0], tm);
        continue;
      }
      if (!(part instanceof PdfString)) continue;
      let kept = [];
      let pending = 0; // thousandths of text space to move for glyphs taken out
      const flush = () => {
        if (kept.length) { result.push(new PdfString(Buffer.concat(kept), part.hex)); kept = []; }
      };
      for (const g of codesOf(part.bytes, font.twoByte)) {
        const w0 = font.width(g.code);
        const space = !font.twoByte && g.code === 32 ? gs.tw : 0;
        const adv = (w0 * size + gs.tc + space) * gs.th;
        const trm = mul([size * gs.th, 0, 0, size, 0, gs.rise], mul(tm, gs.ctm));
        const box = boxOf(trm, 0, -0.22, Math.max(w0, 0.05), 0.88);
        if (size !== 0 && glyphHit(box, areas)) {
          if (pending === 0) flush();
          pending += size && gs.th ? (adv / (size * gs.th)) * 1000 : 0;
          removed.glyphs += 1;
          ctx.run += font.unicode(g.code);
          cut = true;
        } else {
          if (pending) { flush(); result.push(-Math.round(pending * 1000) / 1000); pending = 0; }
          if (ctx.run) { ctx.phrases.push(ctx.run); ctx.run = ''; }
          kept.push(g.bytes);
        }
        tm = mul([1, 0, 0, 1, adv, 0], tm);
      }
      if (pending) { flush(); result.push(-Math.round(pending * 1000) / 1000); }
      flush();
    }
    return { result, cut };
  };
  // Letters cut from one line and then the next are two words, not one:
  // a move to a new line or a new place ends the word being taken.
  const gap = () => { if (ctx.run && !ctx.run.endsWith(' ')) ctx.run += ' '; };

  for (const o of ops) {
    const a = o.operands;
    switch (o.op) {
      case 'q': stack.push({ ...gs, ctm: gs.ctm.slice() }); out.push(o); break;
      case 'Q': if (stack.length) gs = stack.pop(); out.push(o); break;
      case 'cm': if (a.length >= 6) gs.ctm = mul(a.slice(0, 6).map((x) => num(x)), gs.ctm); out.push(o); break;
      case 'BT': gap(); tm = [1, 0, 0, 1, 0, 0]; tlm = [1, 0, 0, 1, 0, 0]; out.push(o); break;
      case 'Tf': gs.font = a[0] instanceof Name ? a[0].name : null; gs.size = num(a[1]); out.push(o); break;
      case 'Tc': gs.tc = num(a[0]); out.push(o); break;
      case 'Tw': gs.tw = num(a[0]); out.push(o); break;
      case 'Tz': gs.th = num(a[0], 100) / 100; out.push(o); break;
      case 'TL': gs.tl = num(a[0]); out.push(o); break;
      case 'Ts': gs.rise = num(a[0]); out.push(o); break;
      case 'Td': gap(); tlm = mul([1, 0, 0, 1, num(a[0]), num(a[1])], tlm); tm = tlm.slice(); out.push(o); break;
      case 'TD': gap(); gs.tl = -num(a[1]); tlm = mul([1, 0, 0, 1, num(a[0]), num(a[1])], tlm); tm = tlm.slice(); out.push(o); break;
      case 'Tm': gap(); tlm = a.slice(0, 6).map((x) => num(x)); tm = tlm.slice(); out.push(o); break;
      case 'T*': gap(); tlm = mul([1, 0, 0, 1, 0, -gs.tl], tlm); tm = tlm.slice(); out.push(o); break;
      case 'Tj': case "'": case '"': case 'TJ': {
        if (o.op === "'" || o.op === '"') {
          if (o.op === '"') { gs.tw = num(a[0]); gs.tc = num(a[1]); out.push({ op: 'Tw', operands: [gs.tw] }, { op: 'Tc', operands: [gs.tc] }); }
          gap();
          tlm = mul([1, 0, 0, 1, 0, -gs.tl], tlm);
          tm = tlm.slice();
          out.push({ op: 'T*', operands: [] });
        }
        const parts = o.op === 'TJ' ? (Array.isArray(a[0]) ? a[0] : []) : [a[a.length - 1]];
        const { result, cut } = show(o, parts);
        if (cut) for (const m of marked) m.hit = true;
        out.push(cut || o.op !== 'Tj' ? { op: 'TJ', operands: [result] } : o);
        break;
      }
      case 'BDC': case 'BMC': {
        const entry = { op: o, hit: false };
        marked.push(entry);
        out.push(o);
        break;
      }
      case 'EMC': {
        const m = marked.pop();
        if (m && m.hit && m.op.op === 'BDC' && m.op.operands[1] instanceof Dict) {
          // Words the marked content carried for search and reading aloud go with the glyphs.
          const props = new Dict(m.op.operands[1].entries());
          for (const k of ['ActualText', 'Alt', 'E']) props.delete(k);
          m.op.operands = [m.op.operands[0], props];
          if (marked.length) marked[marked.length - 1].hit = true;
        }
        out.push(o);
        break;
      }
      case 'BI': {
        // An inline image in the unit square under the CTM.
        if (areas.some((r) => intersects(boxOf(gs.ctm, 0, 0, 1, 1), r))) { removed.images += 1; break; }
        out.push(o);
        break;
      }
      case 'Do': {
        const name = a[0] instanceof Name ? a[0].name : null;
        const ref = name && xobjects instanceof Dict ? xobjects.get(name) : null;
        const x = ref ? file.resolve(ref) : null;
        if (!(x instanceof Stream)) { out.push(o); break; }
        const sub = x.dict.get('Subtype') instanceof Name ? x.dict.get('Subtype').name : '';
        if (sub === 'Image') {
          const box = boxOf(gs.ctm, 0, 0, 1, 1);
          const under = areas.filter((r) => intersects(box, r));
          if (!under.length) { out.push(o); break; }
          const w = num(file.resolve(x.dict.get('Width')));
          const h = num(file.resolve(x.dict.get('Height')));
          const rects = pixelRects(gs.ctm, w, h, under);
          if (!rects.length) { out.push(o); break; }
          const key = `${ref instanceof Ref ? ref.num : name}@${gs.ctm.map((v) => Math.round(v * 100) / 100).join(',')}`;
          let replacement = null;
          const pix = readPixels(file, x);
          if (pix) replacement = imageStream(file, x, pix, paintOut(pix, rects, ctx.fill));
          else if (images && images[key]) {
            const given = images[key];
            const comps = given.data.length >= given.width * given.height * 3 ? 3 : 1;
            const p = { width: given.width, height: given.height, comps, data: Buffer.from(given.data), colorSpace: comps === 3 ? 'DeviceRGB' : 'DeviceGray' };
            const scaled = rects.map(([x0, y0, x1, y1]) => [Math.floor(x0 * p.width / w), Math.floor(y0 * p.height / h), Math.ceil(x1 * p.width / w), Math.ceil(y1 * p.height / h)]);
            replacement = imageStream(file, x, p, paintOut(p, scaled, ctx.fill));
            replacement.dict.delete('SMask');
            if (x.dict.get('SMask')) replacement.dict.set('SMask', x.dict.get('SMask'));
          } else if (plan) {
            plan.push({ key, ref: ref instanceof Ref ? ref.num : null, width: w, height: h, filter: (() => { const f = file.resolve(x.dict.get('Filter')); return f instanceof Name ? f.name : Array.isArray(f) ? f.map((q) => q.name).join(' ') : ''; })(), rects, data: x.data });
          }
          gone.add(name);
          if (replacement) {
            const fresh = `R${name}${ctx.counter++}`;
            newX.set(fresh, replacement);
            out.push({ op: 'Do', operands: [new Name(fresh)] });
            removed.pictures += 1;
          } else {
            // A picture that cannot be painted out leaves the page whole.
            removed.images += 1;
          }
          break;
        }
        if (sub === 'Form' && depth < 12) {
          const fm = file.resolve(x.dict.get('Matrix'));
          const m = Array.isArray(fm) && fm.length >= 6 ? fm.slice(0, 6).map((v) => num(file.resolve(v))) : [1, 0, 0, 1, 0, 0];
          const bb = file.resolve(x.dict.get('BBox'));
          const ctm = mul(m, gs.ctm);
          const box = Array.isArray(bb) && bb.length >= 4 ? boxOf(ctm, ...bb.slice(0, 4).map((v) => num(file.resolve(v)))) : null;
          if (box && !areas.some((r) => intersects(box, r))) { out.push(o); break; }
          const before = removed.glyphs + removed.images + removed.pictures;
          let decoded;
          try { decoded = file.decode(x); } catch { decoded = Buffer.alloc(0); }
          const inner = redactContent(ctx, decoded, x.dict.get('Resources') || resources, ctm, depth + 1);
          if (removed.glyphs + removed.images + removed.pictures === before) { out.push(o); break; }
          gone.add(name);
          const dict = new Dict();
          for (const [k, v] of x.dict.entries()) if (!['Filter', 'DecodeParms', 'Length'].includes(k)) dict.set(k, v);
          if (inner.resources) dict.set('Resources', inner.resources);
          dict.set('Filter', new Name('FlateDecode'));
          const fresh = `R${name}${ctx.counter++}`;
          newX.set(fresh, new Stream(dict, zlib.deflateSync(inner.data)));
          out.push({ op: 'Do', operands: [new Name(fresh)] });
          for (const m2 of marked) m2.hit = true;
          break;
        }
        out.push(o);
        break;
      }
      default:
        out.push(o);
    }
  }

  let newResources = null;
  if (newX.size || gone.size) {
    // The old picture or form leaves the resources unless something on the page still draws it.
    const used = new Set(out.filter((x) => x.op === 'Do' && x.operands[0] instanceof Name).map((x) => x.operands[0].name));
    newResources = new Dict(res instanceof Dict ? res.entries() : []);
    const xo = new Dict(xobjects instanceof Dict ? xobjects.entries() : []);
    for (const k of gone) if (!used.has(k)) xo.delete(k);
    for (const [k, v] of newX) xo.set(k, v);
    newResources.set('XObject', xo);
  }
  return { data: writeContent(out), resources: newResources, depth: stack.length };
}

/** The areas to redact on each page, from its Redact marks: { page: { rects, fills, overlays, marks } }. */
function marksOf(file) {
  const out = new Map();
  for (const a of readAnnotations(file)) {
    if (a.type !== 'Redact') continue;
    if (!out.has(a.page)) out.set(a.page, []);
    const rects = a.quads ? a.quads.map((q) => {
      const xs = [q[0], q[2], q[4], q[6]];
      const ys = [q[1], q[3], q[5], q[7]];
      return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
    }) : a.rect ? [a.rect] : [];
    out.get(a.page).push({ id: a.id, rects, fill: a.interior || [0, 0, 0], overlay: a.overlay || '', words: a.contents || '' });
  }
  return out;
}

/**
 * The pictures under the marks that this module cannot decode, for a caller
 * that can (the window, with the browser's decoders) to hand back painted
 * out: [{ page, key, width, height, filter, rects, data }]. `bytes`: the
 * file, or the file read (a protected one with its password).
 */
function redactionPlan(bytes, { areas = null } = {}) {
  const file = require('./reader').asPdf(bytes);
  const plan = [];
  run(file, { areas, plan, images: null, dryRun: true });
  return plan;
}

function run(file, { areas: given = null, plan = null, images = null }) {
  if (file.locked) throw new RedactError('A PDF protected with a password can be redacted only once it is opened with its password.', 'encrypted');
  const pages = file.pages();
  const marks = marksOf(file);
  const removed = { glyphs: 0, images: 0, pictures: 0, annotations: 0, pages: 0 };
  const items = [];
  const phrases = [];
  const under = {}; // kind -> count, of the annotations lying under a mark
  const unpainted = []; // { page, count }: pictures left off a page
  pages.forEach((page, p) => {
    const mine = marks.get(p) || [];
    const extra = given ? given.filter((g) => g.page === p).flatMap((g) => g.rects.map((r) => ({ rects: [r], fill: g.fill || [0, 0, 0], overlay: g.overlay || '' }))) : [];
    const all = [...mine, ...extra];
    const rects = all.flatMap((m) => m.rects.map((r) => [Math.min(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[0], r[2]), Math.max(r[1], r[3])]));
    if (!rects.length) { items.push({ file, page: p }); return; }
    removed.pages += 1;
    const resources = page.dict.has('Resources') ? page.dict.get('Resources') : page.inherited.Resources;
    const ctx = { file, areas: rects, removed, plan: plan ? [] : null, images, counter: 1, fill: [0, 0, 0], run: '', phrases };
    const content = file.pageContent(p);
    const imagesBefore = removed.images;
    const redacted = redactContent(ctx, content, resources, [1, 0, 0, 1, 0, 0]);
    if (ctx.run) phrases.push(ctx.run);
    for (const m of all) if (m.words) phrases.push(m.words);
    if (removed.images > imagesBefore) unpainted.push({ page: p, count: removed.images - imagesBefore });
    if (plan) for (const x of ctx.plan) plan.push({ page: p, ...x });
    // The fill over each area, and its overlay words.
    // Every save the page's content left open is closed, and the one around it, so the fills are drawn on the page as it stands.
    const paint = Array.from({ length: redacted.depth + 1 }, () => 'Q');
    for (const m of all) {
      const rgb = (m.fill || [0, 0, 0]).slice(0, 3);
      for (const r of m.rects) {
        const [x0, y0, x1, y1] = [Math.min(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[0], r[2]), Math.max(r[1], r[3])];
        paint.push(`${rgb.map(n).join(' ')} rg ${n(x0)} ${n(y0)} ${n(x1 - x0)} ${n(y1 - y0)} re f`);
      }
      if (m.overlay && m.rects[0]) {
        const r = m.rects[0];
        const h = Math.abs(r[3] - r[1]);
        const size = Math.max(4, Math.min(12, h * 0.7));
        paint.push(`BT 1 1 1 rg /RHelv ${n(size)} Tf ${n(Math.min(r[0], r[2]) + 2)} ${n(Math.min(r[1], r[3]) + (h - size * 0.7) / 2)} Td ${literal(encode(m.overlay))} Tj ET`);
      }
    }
    const data = Buffer.concat([Buffer.from('q\n'), redacted.data, Buffer.from(`\n${paint.join('\n')}\n`)]);
    let res = redacted.resources || file.resolve(resources);
    if (all.some((m) => m.overlay)) {
      res = new Dict(res instanceof Dict ? res.entries() : []);
      const fontsDict = new Dict(file.resolve(res.get('Font')) instanceof Dict ? file.resolve(res.get('Font')).entries() : []);
      fontsDict.set('RHelv', new Dict([['Type', new Name('Font')], ['Subtype', new Name('Type1')], ['BaseFont', new Name('Helvetica')], ['Encoding', new Name('WinAnsiEncoding')]]));
      res.set('Font', fontsDict);
    }
    // Annotations over an area go, the marks with them.
    const list = file.resolve(page.dict.get('Annots'));
    const kept = [];
    if (Array.isArray(list)) {
      for (const item of list) {
        const d = file.resolve(item);
        if (!(d instanceof Dict)) continue;
        const sub = d.get('Subtype') instanceof Name ? d.get('Subtype').name : '';
        const r = file.resolve(d.get('Rect'));
        const box = Array.isArray(r) && r.length >= 4 ? (() => { const v = r.map((x) => num(file.resolve(x))); return [Math.min(v[0], v[2]), Math.min(v[1], v[3]), Math.max(v[0], v[2]), Math.max(v[1], v[3])]; })() : null;
        if (sub === 'Redact' || (box && rects.some((a) => intersects(box, a)))) {
          if (sub !== 'Redact' && sub !== 'Popup') {
            removed.annotations += 1;
            const kind = sub === 'Link' ? 'link' : sub === 'FileAttachment' ? 'attached file' : sub === 'Widget' ? 'form field' : 'comment';
            under[kind] = (under[kind] || 0) + 1;
          }
          continue;
        }
        if (sub === 'Popup') {
          const parent = file.resolve(d.get('Parent'));
          const pr = parent instanceof Dict ? file.resolve(parent.get('Rect')) : null;
          if (Array.isArray(pr) && rects.some((a) => intersects([Math.min(pr[0], pr[2]), Math.min(pr[1], pr[3]), Math.max(pr[0], pr[2]), Math.max(pr[1], pr[3])], a))) continue;
        }
        kept.push(item);
      }
    }
    items.push({
      file,
      page: p,
      override: {
        Contents: new Stream(new Dict([['Filter', new Name('FlateDecode')]]), zlib.deflateSync(data)),
        Resources: res instanceof Dict ? res : new Dict(),
        Annots: kept.length ? kept : null,
        Thumb: null,
        PieceInfo: null,
        Metadata: null,
      },
    });
  });
  return { items, removed, phrases, under, unpainted };
}

/**
 * What a redaction's rewrite takes out besides the pages' content: every
 * place a marked word appears, found case and spacing aside. The words are
 * the marked phrases and, for a phrase of several words, each word of four
 * letters or more in it. `counts` says, by kind, what was taken.
 */
function scrubber(words) {
  const norm = (s) => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
  const needles = new Set();
  // The phrases as they were written, once each, for a window to show.
  const shown = new Map();
  for (const w of words || []) {
    const phrase = norm(w);
    const plain = String(w || '').replace(/\s+/g, ' ').trim();
    if (plain && !shown.has(phrase)) shown.set(phrase, plain);
    if (phrase.replace(/[^\p{L}\p{N}]/gu, '').length >= 2) needles.add(phrase);
    for (const tok of phrase.split(' ')) {
      const word = tok.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, '');
      if (word.length >= 4) needles.add(word);
    }
  }
  const counts = {};
  return {
    needles: [...needles],
    shown: [...shown.values()],
    counts,
    text(s) {
      if (!needles.size || !s) return false;
      const t = norm(s);
      for (const n of needles) if (t.includes(n)) return true;
      return false;
    },
    note(kind, by = 1) { counts[kind] = (counts[kind] || 0) + by; },
  };
}

/**
 * The redaction of one file's pages, without writing: each page's item for
 * assemble.js (with its new content), the words the marks covered, what lay
 * under the marks, and the pictures left off each page.
 */
function redactPages(file, { areas = null, images = null } = {}) {
  return run(file, { areas, images });
}

/**
 * Apply the redactions: the marks in the file (and any `areas` given as
 * [{ page, rects: [[x0, y0, x1, y1]], fill, overlay }]), written as a new
 * file whole. `images` hands back pictures the plan named, decoded by the
 * caller: { [key]: { width, height, data } } (RGB or grey samples).
 *
 * @returns {{ bytes: Buffer, removed: { glyphs, pictures, images, annotations, pages }, elsewhere: Object<string, number>, words: string[], unpainted: { page, count }[] }}
 *   `pictures` painted out, `images` left out whole (could not be painted out),
 *   `elsewhere` what else was taken out, by kind ('title', 'bookmark', 'comment', ...)
 */
function applyRedactions(bytes, { areas = null, images = null, title, words = [] } = {}) {
  const file = require('./reader').asPdf(bytes);
  const { items, removed, phrases, under, unpainted } = run(file, { areas, images });
  if (!removed.pages) throw new RedactError('Nothing is marked for redaction.', 'nothing');
  const scrub = scrubber([...phrases, ...words]);
  for (const [kind, n] of Object.entries(under)) scrub.note(kind, n);
  const out = assemble(items, { title, scrub });
  return { bytes: out, removed, elsewhere: scrub.counts, words: scrub.shown, unpainted };
}

/**
 * Every place a piece of text appears in a file's bytes, as a plain text
 * extraction finds it: in its content and every other stream inflated, as
 * single-byte text, as UTF-16 (big and little end), and as hex strings of
 * either. For proving a redaction removed what it covered.
 */
function bytesContain(bytes, word) {
  const file = require('./reader').openPdf(bytes);
  const needles = [
    Buffer.from(word, 'latin1'),
    Buffer.from(word, 'utf16le').swap16(),
    Buffer.from(word, 'utf16le'),
  ];
  const hexes = needles.map((b) => b.toString('hex'));
  const hay = [bytes];
  for (const [num] of file.entries) {
    let v;
    try { v = file.getObject(num); } catch { continue; }
    if (v instanceof Stream) { try { hay.push(file.decode(v)); } catch { /* left as it is */ } }
  }
  for (const h of hay) {
    for (const nd of needles) if (h.indexOf(nd) >= 0) return true;
    const lower = h.toString('latin1').toLowerCase().replace(/[\s<>]/g, '');
    for (const hx of hexes) if (lower.includes(hx)) return true;
  }
  return false;
}

module.exports = {
  applyRedactions, redactPages, scrubber, redactionPlan, bytesContain, RedactError, parseToUnicode, parseContent, writeContent, fontInfo, textString, PdfReadError,
};
