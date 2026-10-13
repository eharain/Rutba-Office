'use strict';

/**
 * New PDFs from the pages of others: extract, reorder, rotate, delete,
 * insert from another file, merge and split, without drawing a thing.
 *
 * A page is copied with everything it reaches (its content streams, its
 * resources, fonts, images, forms and annotations) by walking the object
 * graph from the page and giving each object it meets a new number in the
 * new file. Streams travel as the bytes they were, filters and all, so a
 * page's text, fonts and pictures are what they were, byte for byte. What a
 * page inherited from the page tree (resources, boxes, rotation) is written
 * onto the page itself, because the tree it inherited from is not copied.
 *
 * References that lead out of the pages being kept are cut rather than
 * followed: to the page tree (/Parent), to the catalog, and to pages left
 * behind (a link to a deleted page loses its destination). A reference to a
 * page that IS kept is pointed at that page in the new file, so a link from
 * page 3 to page 1 still goes to page 1 after the two swap places.
 *
 * The file is written plainly: every object on its own, one classic
 * cross-reference table. Object streams in the source are unpacked on the
 * way; nothing is recompressed.
 */

const crypto = require('crypto');
const {
  PdfFile, PdfReadError, Name, Ref, PdfString, Dict, Stream, INHERITABLE, normaliseRotation,
} = require('./reader');

// Page keys left behind with the page tree: the parent, and the beads of
// article threads the new file does not carry.
const PAGE_DROP = new Set(['Parent', 'B']);

/**
 * Build a PDF from pages of one or more files.
 *
 * @param {Array<{ file: PdfFile, page: number, rotate?: number }>} items
 *   the new file's pages in order: which file, which page (0-based), and
 *   optionally the page's rotation in degrees (absolute, a multiple of 90)
 * @param {{ title?: string, producer?: string, date?: Date }} [options]
 * @returns {Buffer}
 */
function assemble(items, options = {}) {
  if (!Array.isArray(items) || !items.length) throw new PdfReadError('A PDF needs at least one page.', 'unsupported');
  for (const item of items) {
    if (!(item.file instanceof PdfFile)) throw new TypeError('each page names the PdfFile it comes from');
    if (item.file.encrypted) {
      throw new PdfReadError('This PDF is protected with a password, and changing the pages of a protected PDF is not built yet.', 'encrypted');
    }
    if (!item.file.pages()[item.page]) throw new PdfReadError(`There is no page ${item.page + 1} in that file.`, 'unsupported');
  }

  const objects = [];  // objects[n] = value, n from 1
  let next = 1;
  const catalogNum = next++;
  const pagesNum = next++;

  // Per source file: old reference -> new number, and its kept pages.
  const copiers = new Map();
  const copier = (file) => {
    if (!copiers.has(file)) copiers.set(file, { map: new Map(), pages: new Map() });
    return copiers.get(file);
  };

  // Numbers for the new pages first, so a link to a later page resolves.
  const pageNums = items.map((item) => {
    const n = next++;
    const page = item.file.pages()[item.page];
    const c = copier(item.file);
    if (page.ref && !c.pages.has(page.ref.key)) c.pages.set(page.ref.key, n);
    return n;
  });

  const queue = [];
  const refFor = (file, ref) => {
    const c = copier(file);
    if (c.pages.has(ref.key)) return new Ref(c.pages.get(ref.key), 0);
    if (c.map.has(ref.key)) {
      const n = c.map.get(ref.key);
      return n === null ? null : new Ref(n, 0);
    }
    let target;
    try { target = file.getObject(ref); } catch { target = null; }
    if (target === null || target === undefined) { c.map.set(ref.key, null); return null; }
    if (target instanceof Dict && (target.type === 'Pages' || target.type === 'Catalog' || target.type === 'Page')) {
      c.map.set(ref.key, null);
      return null;
    }
    const n = next++;
    c.map.set(ref.key, n);
    queue.push({ file, ref, n });
    return new Ref(n, 0);
  };

  const copy = (file, value, depth = 0) => {
    if (depth > 512) throw new PdfReadError('the objects are nested too deeply to copy', 'damaged');
    if (value instanceof Ref) return refFor(file, value);
    if (Array.isArray(value)) return value.map((v) => copy(file, v, depth + 1));
    if (value instanceof Stream) return new Stream(copyDict(file, value.dict, depth), value.data);
    if (value instanceof Dict) return copyDict(file, value, depth);
    return value;
  };
  const copyDict = (file, dict, depth) => {
    const out = new Dict();
    for (const [k, v] of dict.entries()) out.set(k, copy(file, v, depth + 1));
    return out;
  };

  items.forEach((item, i) => {
    const page = item.file.pages()[item.page];
    const dict = new Dict();
    dict.set('Type', new Name('Page'));
    dict.set('Parent', new Ref(pagesNum, 0));
    for (const [k, v] of page.dict.entries()) {
      if (PAGE_DROP.has(k) || k === 'Type') continue;
      dict.set(k, copy(item.file, v));
    }
    for (const key of INHERITABLE) {
      if (!dict.has(key) && page.inherited[key] !== undefined) dict.set(key, copy(item.file, page.inherited[key]));
    }
    if (!dict.has('MediaBox')) dict.set('MediaBox', [0, 0, 612, 792]);
    if (!dict.has('Resources')) dict.set('Resources', new Dict());
    const own = item.file.pageInfo(item.page).rotate;
    const rotate = item.rotate === undefined || item.rotate === null ? own : normaliseRotation(item.rotate);
    if (rotate) dict.set('Rotate', rotate);
    else dict.delete('Rotate');
    objects[pageNums[i]] = dict;
  });

  while (queue.length) {
    const { file, ref, n } = queue.shift();
    objects[n] = copy(file, file.getObject(ref));
  }

  // The catalog: the page tree, and the few document-wide things a page's
  // own content can depend on (its language, layers' visibility) or the
  // reader expects (how the window opens).
  const first = items[0].file;
  const catalog = new Dict([['Type', new Name('Catalog')], ['Pages', new Ref(pagesNum, 0)]]);
  const sourceCatalog = first.catalog();
  for (const key of ['Lang', 'ViewerPreferences', 'OCProperties']) {
    if (sourceCatalog.has(key)) {
      const v = copy(first, sourceCatalog.get(key));
      if (v !== null && v !== undefined) catalog.set(key, v);
    }
  }
  while (queue.length) {
    const { file, ref, n } = queue.shift();
    objects[n] = copy(file, file.getObject(ref));
  }
  objects[catalogNum] = catalog;
  objects[pagesNum] = new Dict([
    ['Type', new Name('Pages')],
    ['Kids', pageNums.map((n) => new Ref(n, 0))],
    ['Count', pageNums.length],
  ]);

  const words = first.info();
  const info = new Dict();
  const title = options.title !== undefined ? options.title : words.title;
  if (title) info.set('Title', textString(title));
  if (words.author) info.set('Author', textString(words.author));
  if (words.subject) info.set('Subject', textString(words.subject));
  if (words.keywords) info.set('Keywords', textString(words.keywords));
  if (words.creator) info.set('Creator', textString(words.creator));
  info.set('Producer', textString(options.producer || 'Rutba Office'));
  info.set('ModDate', textString(pdfDate(options.date || new Date())));
  const infoNum = next++;
  objects[infoNum] = info;

  const versions = [...copiers.keys()].map((f) => f.version).sort();
  const version = versions[versions.length - 1] > '1.4' ? versions[versions.length - 1] : '1.4';
  return write(objects, next, { root: catalogNum, info: infoNum, version });
}

/** A text string as PDF writes one: Latin-1 when it fits, UTF-16 with its mark when not. */
function textString(text) {
  const s = String(text);
  if (/^[\x20-\x7e\xa0-\xff\t\n\r]*$/.test(s)) return new PdfString(Buffer.from(s, 'latin1'));
  const bytes = Buffer.alloc(2 + s.length * 2);
  bytes[0] = 0xfe; bytes[1] = 0xff;
  for (let i = 0; i < s.length; i++) bytes.writeUInt16BE(s.charCodeAt(i), 2 + i * 2);
  return new PdfString(bytes, true);
}

function pdfDate(d) {
  const p = (n) => String(n).padStart(2, '0');
  return `D:${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`;
}

// ── writing ───────────────────────────────────────────────────────────────

const REGULAR = /^[!-~]$/;
const NAME_ESCAPE = new Set([...'()<>[]{}/%#']);

function nameText(name) {
  let s = '/';
  for (const ch of Buffer.from(name, 'latin1')) {
    const c = String.fromCharCode(ch);
    s += REGULAR.test(c) && !NAME_ESCAPE.has(c) ? c : `#${ch.toString(16).padStart(2, '0').toUpperCase()}`;
  }
  return s;
}

function numberText(n) {
  if (!Number.isFinite(n)) return '0';
  if (Number.isInteger(n)) return String(n);
  let s = n.toFixed(6).replace(/0+$/, '').replace(/\.$/, '');
  if (s === '-0') s = '0';
  return s;
}

function stringText(str) {
  if (str.hex) return `<${str.bytes.toString('hex')}>`;
  let s = '(';
  for (const c of str.bytes) {
    if (c === 40 || c === 41 || c === 92) s += `\\${String.fromCharCode(c)}`;
    else if (c === 10) s += '\\n';
    else if (c === 13) s += '\\r';
    else if (c < 32 || c > 126) s += `\\${c.toString(8).padStart(3, '0')}`;
    else s += String.fromCharCode(c);
  }
  return `${s})`;
}

function valueText(v) {
  if (v === null || v === undefined) return 'null';
  if (v === true) return 'true';
  if (v === false) return 'false';
  if (typeof v === 'number') return numberText(v);
  if (v instanceof Name) return nameText(v.name);
  if (v instanceof Ref) return `${v.num} ${v.gen} R`;
  if (v instanceof PdfString) return stringText(v);
  if (Array.isArray(v)) return `[${v.map(valueText).join(' ')}]`;
  if (v instanceof Dict) {
    const parts = [];
    for (const [k, x] of v.entries()) parts.push(`${nameText(k)} ${valueText(x)}`);
    return `<<${parts.join(' ')}>>`;
  }
  if (typeof v === 'string') return v; // never produced by the reader; a writer's own token
  return 'null';
}

function write(objects, size, { root, info, version }) {
  const chunks = [];
  let length = 0;
  const push = (data) => {
    const b = Buffer.isBuffer(data) ? data : Buffer.from(data, 'latin1');
    chunks.push(b);
    length += b.length;
  };
  push(`%PDF-${version}\n%\xE2\xE3\xCF\xD3\n`);
  const offsets = new Array(size).fill(0);
  for (let n = 1; n < size; n++) {
    const v = objects[n];
    offsets[n] = length;
    push(`${n} 0 obj\n`);
    if (v instanceof Stream) {
      const dict = new Dict(v.dict.entries());
      dict.set('Length', v.data.length);
      push(`${valueText(dict)}\nstream\n`);
      push(v.data);
      push('\nendstream\nendobj\n');
    } else {
      push(`${valueText(v === undefined ? null : v)}\nendobj\n`);
    }
  }
  const xrefAt = length;
  let xref = `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let n = 1; n < size; n++) xref += `${String(offsets[n]).padStart(10, '0')} 00000 n \n`;
  push(xref);
  const hash = crypto.createHash('md5');
  for (const c of chunks) hash.update(c);
  const id = hash.digest('hex');
  push(`trailer\n<</Size ${size} /Root ${root} 0 R /Info ${info} 0 R /ID [<${id}> <${id}>]>>\nstartxref\n${xrefAt}\n%%EOF\n`);
  return Buffer.concat(chunks, length);
}

// ── page lists ────────────────────────────────────────────────────────────

/**
 * Where a document of `count` pages splits into files of `every` pages:
 * [[0, 1, 2], [3, 4, 5], [6]].
 */
function splitEvery(count, every) {
  const size = Math.max(1, Math.floor(Number(every) || 1));
  const out = [];
  for (let i = 0; i < count; i += size) out.push(Array.from({ length: Math.min(size, count - i) }, (_, k) => i + k));
  return out;
}

/**
 * Where a document splits before each page named (0-based): splitting a
 * ten-page file before 3 and 7 gives [0..2], [3..6], [7..9].
 */
function splitBefore(count, starts) {
  const cuts = [...new Set((starts || []).map(Number).filter((n) => Number.isInteger(n) && n > 0 && n < count))].sort((a, b) => a - b);
  const out = [];
  let from = 0;
  for (const cut of [...cuts, count]) {
    out.push(Array.from({ length: cut - from }, (_, k) => from + k));
    from = cut;
  }
  return out.filter((r) => r.length);
}

/** "1-3, 5, 8-" in a document of `count` pages, as 0-based page numbers in order. */
function parsePageRange(text, count) {
  const out = [];
  for (const part of String(text || '').split(/[,;\s]+/).filter(Boolean)) {
    const m = /^(\d*)\s*-\s*(\d*)$/.exec(part);
    if (m) {
      const a = m[1] ? Number(m[1]) : 1;
      const b = m[2] ? Number(m[2]) : count;
      if (a < 1 || b < 1 || a > count || a > b) return null;
      for (let i = a; i <= Math.min(b, count); i++) out.push(i - 1);
    } else if (/^\d+$/.test(part)) {
      const n = Number(part);
      if (n < 1 || n > count) return null;
      out.push(n - 1);
    } else return null;
  }
  return out.length ? out : null;
}

module.exports = { assemble, splitEvery, splitBefore, parsePageRange, textString, valueText };
