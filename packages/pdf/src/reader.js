'use strict';

/**
 * Reading a PDF: the file's structure, its objects and its pages.
 *
 * The writer beside this (document.js) makes PDFs; this reads anybody's,
 * well enough to work on them a page at a time without drawing anything:
 * which objects there are and where (cross-reference tables, cross-reference
 * streams, object streams, every incremental update chained by /Prev), the
 * trailer, and the page tree with the attributes a page inherits from the
 * nodes above it. assemble.js builds new files from the pages this finds.
 *
 * A file whose cross-references are wrong (offsets that miss, a startxref
 * that names nothing, a table cut short) is read again by scanning the bytes
 * for its objects, as every desktop reader does: a damaged file that opens
 * elsewhere opens here. An encrypted file is recognised and said to be, and
 * nothing here pretends to read what it cannot decrypt.
 *
 * Values, as this module hands them out:
 *   null, true/false, numbers, PdfString (bytes), Name, Ref, arrays,
 *   Dict (a Map of name -> value) and Stream (a Dict plus the stream's
 *   bytes exactly as they sit in the file, still encoded).
 */

const zlib = require('zlib');

class PdfReadError extends Error {
  /** code: 'not-pdf' | 'damaged' | 'encrypted' | 'unsupported' */
  constructor(message, code = 'damaged') {
    super(message);
    this.name = 'PdfReadError';
    this.code = code;
  }
}

class Name {
  constructor(name) { this.name = name; }
  toString() { return `/${this.name}`; }
}

class Ref {
  constructor(num, gen = 0) { this.num = num; this.gen = gen; }
  get key() { return `${this.num} ${this.gen}`; }
  toString() { return `${this.num} ${this.gen} R`; }
}

class PdfString {
  constructor(bytes, hex = false) { this.bytes = bytes; this.hex = hex; }
  /** The string as text: UTF-16 with its mark, UTF-8 with its mark, else PDFDocEncoding read as Latin-1. */
  text() {
    const b = this.bytes;
    if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) {
      let s = '';
      for (let i = 2; i + 1 < b.length; i += 2) s += String.fromCharCode((b[i] << 8) | b[i + 1]);
      return s;
    }
    if (b.length >= 3 && b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return b.subarray(3).toString('utf8');
    return b.toString('latin1');
  }
}

class Dict {
  constructor(entries) { this.map = new Map(entries || []); }
  get(key) { return this.map.get(key); }
  has(key) { return this.map.has(key); }
  set(key, value) { this.map.set(key, value); return this; }
  delete(key) { return this.map.delete(key); }
  keys() { return this.map.keys(); }
  entries() { return this.map.entries(); }
  get size() { return this.map.size; }
  /** The /Type, as a plain string, or ''. */
  get type() { const t = this.map.get('Type'); return t instanceof Name ? t.name : ''; }
}

class Stream {
  constructor(dict, data) { this.dict = dict; this.data = data; }
}

/** A keyword in the token stream: obj, endobj, R, stream, true, an operator. */
class Keyword {
  constructor(word) { this.word = word; }
}

const isName = (v, name) => v instanceof Name && (name === undefined || v.name === name);

// ── the lexer ─────────────────────────────────────────────────────────────

const WHITE = new Uint8Array(256);
for (const c of [0, 9, 10, 12, 13, 32]) WHITE[c] = 1;
const DELIM = new Uint8Array(256);
for (const c of '()<>[]{}/%') DELIM[c.charCodeAt(0)] = 1;
const HEX = new Int8Array(256).fill(-1);
for (let i = 0; i < 10; i++) HEX[48 + i] = i;
for (let i = 0; i < 6; i++) { HEX[65 + i] = 10 + i; HEX[97 + i] = 10 + i; }

const END_OF_DATA = { eof: true };

class Lexer {
  constructor(buf, pos = 0) { this.buf = buf; this.pos = pos; }

  skipSpace() {
    const b = this.buf;
    let p = this.pos;
    while (p < b.length) {
      const c = b[p];
      if (WHITE[c]) p += 1;
      else if (c === 37) { // % comment to the end of the line
        while (p < b.length && b[p] !== 10 && b[p] !== 13) p += 1;
      } else break;
    }
    this.pos = p;
  }

  /** The next token: a value, a delimiter string ('[', ']', '<<', '>>', '{', '}'), a Keyword, or END_OF_DATA. */
  next() {
    this.skipSpace();
    const b = this.buf;
    let p = this.pos;
    if (p >= b.length) return END_OF_DATA;
    const c = b[p];
    if (c === 40) return this.literal();
    if (c === 47) return this.name();
    if (c === 91) { this.pos = p + 1; return '['; }
    if (c === 93) { this.pos = p + 1; return ']'; }
    if (c === 123) { this.pos = p + 1; return '{'; }
    if (c === 125) { this.pos = p + 1; return '}'; }
    if (c === 60) {
      if (b[p + 1] === 60) { this.pos = p + 2; return '<<'; }
      return this.hexString();
    }
    if (c === 62) {
      if (b[p + 1] === 62) { this.pos = p + 2; return '>>'; }
      this.pos = p + 1;
      return new Keyword('>');
    }
    if (c === 41) { this.pos = p + 1; return new Keyword(')'); }
    // A number or a keyword: everything up to the next space or delimiter.
    const start = p;
    while (p < b.length && !WHITE[b[p]] && !DELIM[b[p]]) p += 1;
    if (p === start) { this.pos = p + 1; return new Keyword(String.fromCharCode(c)); }
    this.pos = p;
    const word = b.toString('latin1', start, p);
    if (/^[+-]*(\d+\.?\d*|\.\d+)$/.test(word)) {
      // Lenient about doubled signs ("--5"), as the readers people use are.
      const negative = (word.match(/-/g) || []).length % 2 === 1;
      const n = Number(word.replace(/^[+-]+/, ''));
      return negative ? -n : n;
    }
    if (word === 'true') return true;
    if (word === 'false') return false;
    if (word === 'null') return null;
    return new Keyword(word);
  }

  name() {
    const b = this.buf;
    let p = this.pos + 1;
    const bytes = [];
    while (p < b.length && !WHITE[b[p]] && !DELIM[b[p]]) {
      if (b[p] === 35 && HEX[b[p + 1]] >= 0 && HEX[b[p + 2]] >= 0) {
        bytes.push((HEX[b[p + 1]] << 4) | HEX[b[p + 2]]);
        p += 3;
      } else {
        bytes.push(b[p]);
        p += 1;
      }
    }
    this.pos = p;
    return new Name(Buffer.from(bytes).toString('latin1'));
  }

  literal() {
    const b = this.buf;
    let p = this.pos + 1;
    let depth = 1;
    const out = [];
    while (p < b.length) {
      const c = b[p];
      if (c === 40) { depth += 1; out.push(c); p += 1; continue; }
      if (c === 41) {
        depth -= 1;
        p += 1;
        if (depth === 0) break;
        out.push(c);
        continue;
      }
      if (c === 13) { // an end of line inside a string is a newline, whatever it was
        out.push(10);
        p += b[p + 1] === 10 ? 2 : 1;
        continue;
      }
      if (c !== 92) { out.push(c); p += 1; continue; }
      const e = b[p + 1];
      p += 2;
      switch (e) {
        case 110: out.push(10); break;
        case 114: out.push(13); break;
        case 116: out.push(9); break;
        case 98: out.push(8); break;
        case 102: out.push(12); break;
        case 40: case 41: case 92: out.push(e); break;
        case 13: if (b[p] === 10) p += 1; break; // a backslash at the end of a line continues it
        case 10: break;
        default:
          if (e >= 48 && e <= 55) {
            let v = e - 48;
            for (let k = 0; k < 2 && b[p] >= 48 && b[p] <= 55; k++) { v = v * 8 + (b[p] - 48); p += 1; }
            out.push(v & 0xff);
          } else if (e !== undefined) {
            out.push(e);
          }
      }
    }
    this.pos = p;
    return new PdfString(Buffer.from(out), false);
  }

  hexString() {
    const b = this.buf;
    let p = this.pos + 1;
    const out = [];
    let high = -1;
    while (p < b.length && b[p] !== 62) {
      const v = HEX[b[p]];
      p += 1;
      if (v < 0) continue;
      if (high < 0) high = v;
      else { out.push((high << 4) | v); high = -1; }
    }
    if (high >= 0) out.push(high << 4);
    this.pos = p + 1;
    return new PdfString(Buffer.from(out), true);
  }
}

// ── the parser: tokens into values ───────────────────────────────────────

/**
 * Parse one value from the lexer. `refs` turns "n g R" into a Ref; content
 * streams have no references and leave it off.
 */
function parseValue(lx, refs = true, depth = 0) {
  if (depth > 256) throw new PdfReadError('the objects are nested too deeply to be real');
  const tok = lx.next();
  return valueOf(lx, tok, refs, depth);
}

function valueOf(lx, tok, refs, depth) {
  if (tok === '[') {
    const arr = [];
    for (;;) {
      const t = lx.next();
      if (t === ']' || t === END_OF_DATA) break;
      if (t instanceof Keyword && (t.word === 'endobj' || t.word === 'stream')) { lx.pos -= t.word.length; break; }
      arr.push(valueOf(lx, t, refs, depth + 1));
    }
    return arr;
  }
  if (tok === '<<') {
    const dict = new Dict();
    for (;;) {
      const t = lx.next();
      if (t === '>>' || t === END_OF_DATA) break;
      if (t instanceof Keyword && (t.word === 'endobj' || t.word === 'stream')) { lx.pos -= t.word.length; break; }
      if (!(t instanceof Name)) continue; // a stray token where a key belongs is passed over
      const v = parseValue(lx, refs, depth + 1);
      if (v instanceof Keyword && (v.word === 'endobj' || v.word === 'stream')) { lx.pos -= v.word.length; break; }
      if (v === '>>') { dict.set(t.name, null); break; }
      dict.set(t.name, v);
    }
    return dict;
  }
  if (refs && typeof tok === 'number' && Number.isInteger(tok) && tok >= 0) {
    const save = lx.pos;
    const gen = lx.next();
    if (typeof gen === 'number' && Number.isInteger(gen) && gen >= 0) {
      const r = lx.next();
      if (r instanceof Keyword && r.word === 'R') return new Ref(tok, gen);
    }
    lx.pos = save;
    return tok;
  }
  return tok;
}

// ── filters ───────────────────────────────────────────────────────────────

function inflate(data) {
  try {
    return zlib.inflateSync(data);
  } catch {
    // A stream cut short or with a broken check value still gives what it
    // has; raw deflate covers a writer that left the zlib header off.
    try { return zlib.inflateSync(data, { finishFlush: zlib.constants.Z_SYNC_FLUSH }); } catch { /* next */ }
    try { return zlib.inflateRawSync(data, { finishFlush: zlib.constants.Z_SYNC_FLUSH }); } catch { /* next */ }
    throw new PdfReadError('a compressed stream in this file is damaged beyond reading');
  }
}

function unpredict(data, parms) {
  const predictor = num(parms && parms.get('Predictor'), 1);
  if (predictor < 2) return data;
  const colors = num(parms.get('Colors'), 1);
  const bpc = num(parms.get('BitsPerComponent'), 8);
  const columns = num(parms.get('Columns'), 1);
  const bpp = Math.max(1, Math.ceil((colors * bpc) / 8));
  const rowBytes = Math.ceil((colors * bpc * columns) / 8);
  if (predictor === 2) {
    if (bpc !== 8) return data;
    const out = Buffer.from(data);
    for (let row = 0; row * rowBytes < out.length; row++) {
      const base = row * rowBytes;
      for (let i = bpp; i < rowBytes && base + i < out.length; i++) out[base + i] = (out[base + i] + out[base + i - bpp]) & 0xff;
    }
    return out;
  }
  // PNG predictors: each row starts with its own filter type.
  const rows = Math.floor(data.length / (rowBytes + 1));
  const out = Buffer.alloc(rows * rowBytes);
  const prev = Buffer.alloc(rowBytes);
  for (let r = 0; r < rows; r++) {
    const src = r * (rowBytes + 1);
    const type = data[src];
    const row = out.subarray(r * rowBytes, (r + 1) * rowBytes);
    for (let i = 0; i < rowBytes; i++) {
      const raw = data[src + 1 + i];
      const left = i >= bpp ? row[i - bpp] : 0;
      const up = prev[i];
      const upLeft = i >= bpp ? prev[i - bpp] : 0;
      let v;
      switch (type) {
        case 0: v = raw; break;
        case 1: v = raw + left; break;
        case 2: v = raw + up; break;
        case 3: v = raw + ((left + up) >> 1); break;
        case 4: {
          const pa = Math.abs(up - upLeft);
          const pb = Math.abs(left - upLeft);
          const pc = Math.abs(left + up - 2 * upLeft);
          v = raw + (pa <= pb && pa <= pc ? left : pb <= pc ? up : upLeft);
          break;
        }
        default: v = raw;
      }
      row[i] = v & 0xff;
    }
    row.copy(prev);
  }
  return out;
}

function asciiHex(data) {
  const out = [];
  let high = -1;
  for (const c of data) {
    if (c === 62) break;
    const v = HEX[c];
    if (v < 0) continue;
    if (high < 0) high = v;
    else { out.push((high << 4) | v); high = -1; }
  }
  if (high >= 0) out.push(high << 4);
  return Buffer.from(out);
}

function ascii85(data) {
  const out = [];
  const group = [];
  for (let i = 0; i < data.length; i++) {
    const c = data[i];
    if (c === 126) break; // ~>
    if (WHITE[c]) continue;
    if (c === 122 && group.length === 0) { out.push(0, 0, 0, 0); continue; }
    if (c < 33 || c > 117) continue;
    group.push(c - 33);
    if (group.length === 5) {
      let v = 0;
      for (const g of group) v = v * 85 + g;
      out.push((v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255);
      group.length = 0;
    }
  }
  if (group.length > 1) {
    const n = group.length;
    while (group.length < 5) group.push(84);
    let v = 0;
    for (const g of group) v = v * 85 + g;
    const bytes = [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
    out.push(...bytes.slice(0, n - 1));
  }
  return Buffer.from(out);
}

function lzw(data, early = 1) {
  const out = [];
  let dict = [];
  const reset = () => { dict = []; for (let i = 0; i < 256; i++) dict.push([i]); dict.push(null, null); };
  reset();
  let bits = 9;
  let buffer = 0;
  let count = 0;
  let prev = null;
  for (let i = 0; i < data.length; i++) {
    buffer = (buffer << 8) | data[i];
    count += 8;
    while (count >= bits) {
      const code = (buffer >>> (count - bits)) & ((1 << bits) - 1);
      count -= bits;
      buffer &= (1 << count) - 1;
      if (code === 256) { reset(); bits = 9; prev = null; continue; }
      if (code === 257) return Buffer.from(out);
      let entry;
      if (code < dict.length && dict[code]) entry = dict[code];
      else if (prev) entry = [...prev, prev[0]];
      else return Buffer.from(out);
      out.push(...entry);
      if (prev) dict.push([...prev, entry[0]]);
      prev = entry;
      if (dict.length + early >= (1 << bits) && bits < 12) bits += 1;
    }
  }
  return Buffer.from(out);
}

function runLength(data) {
  const out = [];
  for (let i = 0; i < data.length;) {
    const n = data[i];
    if (n === 128) break;
    if (n < 128) { for (let k = 0; k <= n; k++) out.push(data[i + 1 + k]); i += n + 2; } else { for (let k = 0; k < 257 - n; k++) out.push(data[i + 1]); i += 2; }
  }
  return Buffer.from(out);
}

const num = (v, fallback = 0) => (typeof v === 'number' && Number.isFinite(v) ? v : fallback);

// ── the file ──────────────────────────────────────────────────────────────

const INHERITABLE = ['Resources', 'MediaBox', 'CropBox', 'Rotate'];

class PdfFile {
  /**
   * @param {Buffer|Uint8Array} bytes  the whole file
   */
  constructor(bytes) {
    this.buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const head = this.buf.indexOf('%PDF-');
    if (head < 0 || head > 1024) throw new PdfReadError('This is not a PDF: it does not begin as one.', 'not-pdf');
    // Bytes before the header (a mail gateway's, or a download's) shift every offset by as much.
    this.base = head;
    const version = /^%PDF-(\d\.\d)/.exec(this.buf.toString('latin1', head, head + 16));
    this.version = version ? version[1] : '1.4';
    this.entries = new Map();   // num -> { type: 0|1|2, offset, gen, stm, idx }
    this.cache = new Map();     // num -> value
    this.objStms = new Map();   // num -> { objects: Map(num -> value) }
    this.trailer = null;
    this.repaired = false;
    this._pages = null;
    try {
      this.readXrefChain();
      this.catalog(); // the file must at least lead to its catalog
    } catch (err) {
      if (err instanceof PdfReadError && err.code !== 'damaged') throw err;
      this.reconstruct();
    }
    this.encrypted = Boolean(this.trailer && this.trailer.get('Encrypt'));
    const catalogVersion = this.catalogOrNull() && this.catalogOrNull().get('Version');
    if (catalogVersion instanceof Name && /^\d\.\d$/.test(catalogVersion.name) && catalogVersion.name > this.version) this.version = catalogVersion.name;
  }

  // ── cross-references ────────────────────────────────────────────────────

  startXref() {
    const b = this.buf;
    const from = Math.max(0, b.length - 4096);
    const tail = b.toString('latin1', from);
    const at = tail.lastIndexOf('startxref');
    if (at < 0) throw new PdfReadError('the file does not say where its cross-references are');
    const m = /startxref\s+(\d+)/.exec(tail.slice(at));
    if (!m) throw new PdfReadError('startxref names no offset');
    return Number(m[1]) + this.base;
  }

  readXrefChain() {
    const seen = new Set();
    let offset = this.startXref();
    let guard = 0;
    while (offset !== null && !seen.has(offset) && guard++ < 1000) {
      seen.add(offset);
      const section = this.readXrefSection(offset);
      for (const [n, entry] of section.entries) if (!this.entries.has(n)) this.entries.set(n, entry);
      if (!this.trailer) this.trailer = section.trailer;
      const prev = section.trailer.get('Prev');
      offset = typeof prev === 'number' ? prev + this.base : null;
    }
    if (!this.trailer) throw new PdfReadError('no trailer');
    // Every in-use offset must land on its object: one that misses means
    // the table is not to be trusted, and the file is read by scanning.
    for (const [n, e] of this.entries) {
      if (e.type !== 1) continue;
      if (!this.objectHeaderAt(e.offset, n)) throw new PdfReadError(`object ${n} is not where the cross-references say`);
    }
  }

  objectHeaderAt(offset, n) {
    if (offset < 0 || offset >= this.buf.length) return false;
    const lx = new Lexer(this.buf, offset);
    const a = lx.next();
    const g = lx.next();
    const k = lx.next();
    return a === n && typeof g === 'number' && k instanceof Keyword && k.word === 'obj';
  }

  readXrefSection(offset) {
    const lx = new Lexer(this.buf, offset);
    const first = lx.next();
    if (first instanceof Keyword && first.word === 'xref') return this.readXrefTable(lx);
    if (typeof first === 'number') {
      lx.pos = offset;
      const { value } = this.readIndirectAt(offset);
      if (value instanceof Stream && value.dict.type === 'XRef') {
        return { entries: this.xrefStreamEntries(value), trailer: value.dict };
      }
    }
    throw new PdfReadError('the cross-references are not where startxref says');
  }

  readXrefTable(lx) {
    const entries = new Map();
    for (;;) {
      const t = lx.next();
      if (t instanceof Keyword && t.word === 'trailer') break;
      if (typeof t !== 'number') throw new PdfReadError('a cross-reference table is cut short');
      const start = t;
      const count = lx.next();
      if (typeof count !== 'number') throw new PdfReadError('a cross-reference subsection has no count');
      for (let i = 0; i < count; i++) {
        const off = lx.next();
        const gen = lx.next();
        const kind = lx.next();
        if (typeof off !== 'number' || typeof gen !== 'number' || !(kind instanceof Keyword)) {
          throw new PdfReadError('a cross-reference entry is unreadable');
        }
        const n = start + i;
        if (entries.has(n)) continue;
        if (kind.word === 'n') {
          // An in-use entry at offset 0 is a writer's slip; the scan finds it.
          entries.set(n, off > 0 ? { type: 1, offset: off + this.base, gen } : { type: 0, gen });
        } else entries.set(n, { type: 0, gen });
      }
    }
    const trailer = parseValue(lx);
    if (!(trailer instanceof Dict)) throw new PdfReadError('the trailer is not a dictionary');
    // A hybrid file: objects in object streams are listed as free in the
    // table for older readers, and found by the stream named here.
    const stm = trailer.get('XRefStm');
    if (typeof stm === 'number') {
      try {
        const { value } = this.readIndirectAt(stm + this.base);
        if (value instanceof Stream) {
          for (const [n, e] of this.xrefStreamEntries(value)) {
            const had = entries.get(n);
            if (!had || had.type === 0) entries.set(n, e);
          }
        }
      } catch { /* the table alone */ }
    }
    return { entries, trailer };
  }

  xrefStreamEntries(stream) {
    const dict = stream.dict;
    const data = this.decode(stream);
    const w = (dict.get('W') || []).map((x) => num(x));
    if (w.length < 3) throw new PdfReadError('a cross-reference stream has no field widths');
    const size = num(dict.get('Size'));
    const index = Array.isArray(dict.get('Index')) ? dict.get('Index').map((x) => num(x)) : [0, size];
    const entryLength = w[0] + w[1] + w[2];
    const entries = new Map();
    let p = 0;
    const field = (width, fallback) => {
      if (width === 0) return fallback;
      let v = 0;
      for (let i = 0; i < width; i++) v = v * 256 + (data[p + i] || 0);
      p += width;
      return v;
    };
    for (let s = 0; s + 1 < index.length; s += 2) {
      const start = index[s];
      const count = index[s + 1];
      for (let i = 0; i < count && p + entryLength <= data.length; i++) {
        const type = field(w[0], 1);
        const a = field(w[1], 0);
        const b = field(w[2], 0);
        const n = start + i;
        if (entries.has(n)) continue;
        if (type === 1) entries.set(n, a > 0 ? { type: 1, offset: a + this.base, gen: b } : { type: 0, gen: b });
        else if (type === 2) entries.set(n, { type: 2, stm: a, idx: b, gen: 0 });
        else entries.set(n, { type: 0, gen: b });
      }
    }
    return entries;
  }

  /** Read the file again by finding its objects, when its cross-references cannot be trusted. */
  reconstruct() {
    this.entries = new Map();
    this.cache = new Map();
    this.objStms = new Map();
    this.repaired = true;
    const text = this.buf.toString('latin1');
    const re = /(\d+)\s+(\d+)\s+obj\b/g;
    let m;
    while ((m = re.exec(text))) {
      // Only at the start of a line or after whitespace: "12 0 obj" inside a string is not an object.
      const before = m.index > 0 ? text.charCodeAt(m.index - 1) : 10;
      if (!WHITE[before] && before !== 62 /* > */ && before !== 125) continue;
      this.entries.set(Number(m[1]), { type: 1, offset: m.index, gen: Number(m[2]) });
    }
    if (!this.entries.size) throw new PdfReadError('This file has no objects a PDF reader can find.', 'damaged');
    // Objects kept inside object streams.
    for (const [n, e] of [...this.entries]) {
      if (e.type !== 1) continue;
      let value;
      try { value = this.readIndirectAt(e.offset).value; } catch { continue; }
      if (value instanceof Stream && value.dict.type === 'ObjStm') {
        try {
          const nums = this.objStmNumbers(value);
          nums.forEach((objNum, idx) => { if (!this.entries.has(objNum)) this.entries.set(objNum, { type: 2, stm: n, idx, gen: 0 }); });
        } catch { /* a damaged object stream is passed over */ }
      }
    }
    // The trailer: the last one written, else the newest cross-reference stream's dictionary.
    let trailer = null;
    let at = text.lastIndexOf('trailer');
    while (at >= 0 && !trailer) {
      try {
        const v = parseValue(new Lexer(this.buf, at + 7));
        if (v instanceof Dict && v.get('Root')) trailer = v;
      } catch { /* an earlier one */ }
      at = text.lastIndexOf('trailer', at - 1);
    }
    if (!trailer) {
      for (const [, e] of [...this.entries].reverse()) {
        if (e.type !== 1) continue;
        try {
          const { value } = this.readIndirectAt(e.offset);
          if (value instanceof Stream && value.dict.type === 'XRef' && value.dict.get('Root')) { trailer = value.dict; break; }
        } catch { /* next */ }
      }
    }
    if (!trailer) {
      // No trailer at all: the catalog is the object that says it is one.
      for (const [n, e] of this.entries) {
        let v;
        try { v = this.getObject(new Ref(n, e.gen)); } catch { continue; }
        if (v instanceof Dict && v.type === 'Catalog') { trailer = new Dict([['Root', new Ref(n, e.gen)]]); break; }
      }
    }
    if (!trailer) throw new PdfReadError('This PDF is damaged: no part of it says where its pages begin.', 'damaged');
    this.trailer = trailer;
    this.catalog();
  }

  // ── objects ─────────────────────────────────────────────────────────────

  /** "n g obj value [stream...]" at a byte offset → { num, gen, value }. */
  readIndirectAt(offset) {
    const lx = new Lexer(this.buf, offset);
    const n = lx.next();
    const g = lx.next();
    const k = lx.next();
    if (typeof n !== 'number' || typeof g !== 'number' || !(k instanceof Keyword && k.word === 'obj')) {
      throw new PdfReadError(`no object at offset ${offset}`);
    }
    let value = parseValue(lx);
    if (value instanceof Dict) {
      const save = lx.pos;
      const t = lx.next();
      if (t instanceof Keyword && t.word === 'stream') value = new Stream(value, this.streamData(lx.pos, value));
      else lx.pos = save;
    }
    return { num: n, gen: g, value };
  }

  streamData(afterKeyword, dict) {
    const b = this.buf;
    let start = afterKeyword;
    if (b[start] === 13 && b[start + 1] === 10) start += 2;
    else if (b[start] === 10 || b[start] === 13) start += 1;
    let length = dict.get('Length');
    if (length instanceof Ref) {
      try { length = this.getObject(length); } catch { length = null; }
    }
    if (typeof length === 'number' && length >= 0 && start + length <= b.length) {
      // Trust the length only if endstream follows it.
      const after = b.toString('latin1', start + length, Math.min(b.length, start + length + 32));
      if (/^\s*endstream/.test(after)) return b.subarray(start, start + length);
    }
    const end = b.indexOf('endstream', start, 'latin1');
    let stop = end < 0 ? b.length : end;
    if (b[stop - 1] === 10) stop -= 1;
    if (b[stop - 1] === 13) stop -= 1;
    return b.subarray(start, Math.max(start, stop));
  }

  /** An indirect object's value; null for one that is free or missing, as the format says. */
  getObject(ref) {
    const n = ref instanceof Ref ? ref.num : ref;
    if (this.cache.has(n)) return this.cache.get(n);
    const entry = this.entries.get(n);
    let value = null;
    if (entry && entry.type === 1) {
      const got = this.readIndirectAt(entry.offset);
      value = got.num === n ? got.value : null;
    } else if (entry && entry.type === 2) {
      value = this.fromObjStm(entry.stm, n, entry.idx);
    }
    this.cache.set(n, value);
    return value;
  }

  /** Follow references until a direct value. */
  resolve(value) {
    let v = value;
    for (let i = 0; i < 32 && v instanceof Ref; i++) v = this.getObject(v);
    return v instanceof Ref ? null : v;
  }

  objStmNumbers(stream) {
    const data = this.decode(stream);
    const n = num(stream.dict.get('N'));
    const lx = new Lexer(data, 0);
    const nums = [];
    for (let i = 0; i < n; i++) {
      const a = lx.next();
      lx.next();
      if (typeof a !== 'number') break;
      nums.push(a);
    }
    return nums;
  }

  fromObjStm(stmNum, n, idx) {
    let stm = this.objStms.get(stmNum);
    if (!stm) {
      const stream = this.getObject(stmNum);
      if (!(stream instanceof Stream)) throw new PdfReadError(`object stream ${stmNum} is missing`);
      const data = this.decode(stream);
      const count = num(stream.dict.get('N'));
      const first = num(stream.dict.get('First'));
      const lx = new Lexer(data, 0);
      const offsets = [];
      for (let i = 0; i < count; i++) {
        const a = lx.next();
        const off = lx.next();
        if (typeof a !== 'number' || typeof off !== 'number') break;
        offsets.push([a, off]);
      }
      stm = { data, first, offsets };
      this.objStms.set(stmNum, stm);
    }
    let at = stm.offsets[idx];
    if (!at || at[0] !== n) at = stm.offsets.find((o) => o[0] === n);
    if (!at) return null;
    return parseValue(new Lexer(stm.data, stm.first + at[1]));
  }

  /** A stream's bytes with its filters undone. Image codecs (DCT, JPX, CCITT, JBIG2) are left as they are. */
  decode(stream) {
    let data = stream.data;
    const dict = stream.dict;
    const filters = this.resolve(dict.get('Filter'));
    const parms = this.resolve(dict.get('DecodeParms'));
    const list = Array.isArray(filters) ? filters : filters ? [filters] : [];
    const parmList = Array.isArray(parms) ? parms : [parms];
    list.forEach((f, i) => {
      const name = f instanceof Name ? f.name : '';
      const p = this.resolve(parmList[i]) || null;
      if (name === 'FlateDecode' || name === 'Fl') data = unpredict(inflate(data), p);
      else if (name === 'LZWDecode' || name === 'LZW') data = unpredict(lzw(data, p ? num(p.get('EarlyChange'), 1) : 1), p);
      else if (name === 'ASCIIHexDecode' || name === 'AHx') data = asciiHex(data);
      else if (name === 'ASCII85Decode' || name === 'A85') data = ascii85(data);
      else if (name === 'RunLengthDecode' || name === 'RL') data = runLength(data);
    });
    return data;
  }

  // ── the document ────────────────────────────────────────────────────────

  catalogOrNull() {
    try { return this.catalog(); } catch { return null; }
  }

  catalog() {
    const root = this.resolve(this.trailer && this.trailer.get('Root'));
    if (!(root instanceof Dict)) throw new PdfReadError('the trailer names no catalog');
    return root;
  }

  /** The document information dictionary's words: title, author, subject, keywords, creator, producer. */
  info() {
    const dict = this.resolve(this.trailer.get('Info'));
    const out = {};
    if (!(dict instanceof Dict) || this.encrypted) return out;
    for (const key of ['Title', 'Author', 'Subject', 'Keywords', 'Creator', 'Producer']) {
      const v = this.resolve(dict.get(key));
      if (v instanceof PdfString) out[key.toLowerCase()] = v.text();
    }
    return out;
  }

  /**
   * The pages in reading order: each { ref, dict, inherited } where
   * `inherited` holds the Resources, MediaBox, CropBox and Rotate the page
   * takes from the nodes above it when it does not carry its own.
   */
  pages() {
    if (this._pages) return this._pages;
    const out = [];
    const seen = new Set();
    const visit = (ref, inherited, depth) => {
      if (depth > 64) throw new PdfReadError('the page tree is deeper than any real document');
      if (ref instanceof Ref) {
        if (seen.has(ref.key)) return;
        seen.add(ref.key);
      }
      const node = this.resolve(ref);
      if (!(node instanceof Dict)) return;
      const own = { ...inherited };
      for (const key of INHERITABLE) if (node.has(key)) own[key] = node.get(key);
      const kids = this.resolve(node.get('Kids'));
      const type = node.type;
      if (type === 'Pages' || (type !== 'Page' && Array.isArray(kids))) {
        for (const kid of kids || []) visit(kid, own, depth + 1);
        return;
      }
      out.push({ ref: ref instanceof Ref ? ref : null, dict: node, inherited });
    };
    visit(this.catalog().get('Pages'), {}, 0);
    this._pages = out;
    return out;
  }

  get pageCount() { return this.pages().length; }

  /** A page's box, rotation and size as a viewer shows it. */
  pageInfo(index) {
    const page = this.pages()[index];
    if (!page) throw new PdfReadError(`there is no page ${index + 1}`, 'unsupported');
    const attr = (key) => (page.dict.has(key) ? page.dict.get(key) : page.inherited[key]);
    const box = (key) => {
      const arr = this.resolve(attr(key));
      if (!Array.isArray(arr) || arr.length < 4) return null;
      const n = arr.map((x) => num(this.resolve(x)));
      return [Math.min(n[0], n[2]), Math.min(n[1], n[3]), Math.max(n[0], n[2]), Math.max(n[1], n[3])];
    };
    const media = box('MediaBox') || [0, 0, 612, 792];
    const crop = box('CropBox') || media;
    const rotate = normaliseRotation(num(this.resolve(attr('Rotate'))));
    return { mediaBox: media, cropBox: crop, width: crop[2] - crop[0], height: crop[3] - crop[1], rotate };
  }

  /** The page's content as one buffer: its streams decoded and joined. */
  pageContent(index) {
    const page = this.pages()[index];
    if (!page) return Buffer.alloc(0);
    const contents = this.resolve(page.dict.get('Contents'));
    const list = Array.isArray(contents) ? contents : contents ? [contents] : [];
    const parts = [];
    for (const item of list) {
      const s = this.resolve(item);
      if (s instanceof Stream) parts.push(this.decode(s), Buffer.from('\n'));
    }
    return Buffer.concat(parts);
  }

  /**
   * The words on a page, in the order they are drawn: text shown with
   * simple (one byte a character) fonts, a space where TJ leaves a gap and a
   * line break where a line moves. Enough to say which page is which; a
   * page set in a composite font comes back as its bytes, not its words.
   */
  pageText(index) {
    return contentText(this.pageContent(index));
  }
}

const normaliseRotation = (deg) => ((Math.round((Number(deg) || 0) / 90) * 90) % 360 + 360) % 360;

// Windows-1252's 0x80 to 0x9F, which PDF's WinAnsiEncoding shares.
const WIN_ANSI_HIGH = String.fromCharCode(0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0x8d, 0x17d, 0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178);
const winAnsi = (bytes) => {
  let s = '';
  for (const c of bytes) s += c >= 0x80 && c <= 0x9f ? WIN_ANSI_HIGH[c - 0x80] : String.fromCharCode(c);
  return s;
};

function contentText(content) {
  const lx = new Lexer(content, 0);
  const operands = [];
  let out = '';
  const add = (s) => { out += s; };
  for (let guard = 0; guard < 5e6; guard++) {
    const tok = lx.next();
    if (tok === END_OF_DATA) break;
    if (tok instanceof Keyword) {
      const op = tok.word;
      if (op === 'BI') { // an inline image: its data is skipped to EI
        const at = content.indexOf('ID', lx.pos, 'latin1');
        let end = at < 0 ? content.length : at + 3;
        for (;;) {
          end = content.indexOf('EI', end, 'latin1');
          if (end < 0) { end = content.length; break; }
          if (WHITE[content[end - 1]] && (end + 2 >= content.length || WHITE[content[end + 2]])) break;
          end += 2;
        }
        lx.pos = end + 2;
        operands.length = 0;
        continue;
      }
      if (op === 'Tj' || op === "'" || op === '"') {
        if (op !== 'Tj' && out && !out.endsWith('\n')) add('\n');
        const s = operands[operands.length - 1];
        if (s instanceof PdfString) add(winAnsi(s.bytes));
      } else if (op === 'TJ') {
        const arr = operands[operands.length - 1];
        if (Array.isArray(arr)) {
          for (const part of arr) {
            if (part instanceof PdfString) add(winAnsi(part.bytes));
            else if (typeof part === 'number' && part < -200 && !out.endsWith(' ')) add(' ');
          }
        }
      } else if (op === 'Td' || op === 'TD') {
        const ty = operands[1];
        if (typeof ty === 'number' && ty !== 0) { if (out && !out.endsWith('\n')) add('\n'); } else if (out && !/\s$/.test(out)) add(' ');
      } else if (op === 'T*' || op === 'Tm' || op === 'ET') {
        if (out && !out.endsWith('\n')) add('\n');
      }
      operands.length = 0;
      continue;
    }
    if (tok === '[' || tok === '<<') {
      operands.push(valueOf(lx, tok, false, 0));
      continue;
    }
    if (typeof tok === 'string') continue;
    operands.push(tok);
  }
  return out.replace(/[ \t]+\n/g, '\n').trim();
}

/** Open a PDF's bytes. Throws PdfReadError with a code a window can put in a sentence. */
function openPdf(bytes) {
  return new PdfFile(bytes);
}

module.exports = {
  openPdf, PdfFile, PdfReadError, Name, Ref, PdfString, Dict, Stream, Keyword, Lexer,
  parseValue, contentText, normaliseRotation, isName, INHERITABLE,
};
