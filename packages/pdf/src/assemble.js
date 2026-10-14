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
 * What helps a person find their way goes with the pages (navigation.js):
 * bookmarks whose pages are kept, pointing at those pages where they now
 * stand (a bookmark to a page left out goes, its own bookmarks moving up a
 * level); links and bookmarks to named destinations turned into explicit
 * ones, so they work in any file the page lands in; each page's label as it
 * was ("iv", "A-3"), written as the fewest ranges that say it; and a form's
 * fields whose widgets are on kept pages, a merge's second field of a name
 * renamed so the two stay two fields. A merge's bookmarks follow one
 * another, in the order their files' pages first come.
 *
 * The file is written plainly: every object on its own, one classic
 * cross-reference table. Object streams in the source are unpacked on the
 * way; nothing is recompressed.
 */

const crypto = require('crypto');
const {
  PdfFile, PdfReadError, Name, Ref, PdfString, Dict, Stream, INHERITABLE, normaliseRotation,
} = require('./reader');
const { resolveDest, isNamed, readOutline, pageLabels } = require('./navigation');

// Page keys left behind with the page tree: the parent, and the beads of
// article threads the new file does not carry.
const PAGE_DROP = new Set(['Parent', 'B']);

/**
 * Build a PDF from pages of one or more files.
 *
 * @param {Array<{ file: PdfFile, page: number, rotate?: number, override?: object }>} items
 *   the new file's pages in order: which file, which page (0-based), and
 *   optionally the page's rotation in degrees (absolute, a multiple of 90)
 *   and entries of the page replaced (a value of the reader's own, whose
 *   references are to the same file and whose streams become new objects)
 *   or taken away (null), as redaction replaces a page's contents
 * @param {{ title?: string, producer?: string, date?: Date, replace?: Map }} [options]
 *   `replace`: Map of file -> Map of object number -> a value written in
 *   that object's place (a picture made smaller), the old one left behind
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
    if (value instanceof Stream) {
      const stream = new Stream(copyDict(file, value.dict, depth), value.data);
      if (depth === 0) return stream;
      // A stream given inside a value (a page's new contents, say) is an object of its own.
      const n = next++;
      objects[n] = stream;
      return new Ref(n, 0);
    }
    if (value instanceof Dict) return copyDict(file, value, depth);
    return value;
  };
  const copyDict = (file, dict, depth) => {
    const out = new Dict();
    for (const [k, v] of dict.entries()) out.set(k, copy(file, v, depth + 1));
    // A destination: named ones made explicit, and one to a page left out taken away.
    const s = dict.get('S');
    for (const key of [dict.has('Dest') ? 'Dest' : null, s instanceof Name && s.name === 'GoTo' ? 'D' : null]) {
      if (!key) continue;
      let dest = out.get(key);
      if (isNamed(file, dict.get(key))) {
        const explicit = resolveDest(file, dict.get(key));
        dest = explicit ? copy(file, explicit, depth + 1) : null;
      }
      if (Array.isArray(dest) && dest[0] !== null && dest[0] !== undefined) out.set(key, dest);
      else if (Array.isArray(dest) || dest === null) out.delete(key);
    }
    return out;
  };

  items.forEach((item, i) => {
    const page = item.file.pages()[item.page];
    const dict = new Dict();
    dict.set('Type', new Name('Page'));
    dict.set('Parent', new Ref(pagesNum, 0));
    const replaced = item.override || {};
    for (const [k, v] of page.dict.entries()) {
      // What the page's override replaces is never copied: the old objects would be written all the same.
      if (PAGE_DROP.has(k) || k === 'Type' || k in replaced) continue;
      dict.set(k, copy(item.file, v));
    }
    for (const key of INHERITABLE) {
      if (!dict.has(key) && !(key in replaced) && page.inherited[key] !== undefined) dict.set(key, copy(item.file, page.inherited[key]));
    }
    for (const [k, v] of Object.entries(item.override || {})) {
      if (v === null || v === undefined) dict.delete(k);
      else dict.set(k, copy(item.file, v, 1));
    }
    if (!dict.has('MediaBox')) dict.set('MediaBox', [0, 0, 612, 792]);
    if (!dict.has('Resources')) dict.set('Resources', new Dict());
    const own = item.file.pageInfo(item.page).rotate;
    const rotate = item.rotate === undefined || item.rotate === null ? own : normaliseRotation(item.rotate);
    if (rotate) dict.set('Rotate', rotate);
    else dict.delete('Rotate');
    objects[pageNums[i]] = dict;
  });

  const source = (file, ref) => options.replace?.get(file)?.get(ref.num) || file.getObject(ref);
  while (queue.length) {
    const { file, ref, n } = queue.shift();
    objects[n] = copy(file, source(file, ref));
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
    objects[n] = copy(file, source(file, ref));
  }
  const files = [...copiers.keys()];
  const drain = () => {
    while (queue.length) {
      const { file, ref, n } = queue.shift();
      objects[n] = copy(file, source(file, ref));
    }
  };

  // Bookmarks: each file's, its kept pages' only, in the order the files first come.
  const built = [];
  const build = (file, nodes) => {
    const out = [];
    for (const node of nodes) {
      const kids = build(file, node.kids);
      let dest = null;
      if (node.dest) {
        dest = copy(file, node.dest, 1);
        if (!Array.isArray(dest) || dest[0] === null || dest[0] === undefined) { out.push(...kids); continue; }
      }
      const action = node.action ? copy(file, node.action, 1) : null;
      if (!dest && !action && !kids.length) continue;
      out.push({ ...node, dest, action, kids });
    }
    return out;
  };
  for (const file of files) built.push(...build(file, readOutline(file)));
  drain();
  if (built.length) {
    const visible = (node) => (node.open ? node.kids.reduce((sum, k) => sum + 1 + visible(k), 0) : 0);
    const place = (nodes, parent) => {
      const nums = nodes.map(() => next++);
      nodes.forEach((node, i) => {
        const d = new Dict([['Title', textString(node.title)], ['Parent', new Ref(parent, 0)]]);
        if (i > 0) d.set('Prev', new Ref(nums[i - 1], 0));
        if (i < nodes.length - 1) d.set('Next', new Ref(nums[i + 1], 0));
        if (node.dest) d.set('Dest', node.dest);
        else if (node.action) d.set('A', node.action);
        if (node.color) d.set('C', node.color);
        const flags = (node.italic ? 1 : 0) | (node.bold ? 2 : 0);
        if (flags) d.set('F', flags);
        if (node.kids.length) {
          const kidNums = place(node.kids, nums[i]);
          d.set('First', new Ref(kidNums[0], 0));
          d.set('Last', new Ref(kidNums[kidNums.length - 1], 0));
          const all = node.kids.reduce((sum, k) => sum + 1 + visible(k), 0);
          d.set('Count', node.open ? all : -all);
        }
        objects[nums[i]] = d;
      });
      return nums;
    };
    const rootNum = next++;
    const top = place(built, rootNum);
    objects[rootNum] = new Dict([
      ['Type', new Name('Outlines')], ['First', new Ref(top[0], 0)], ['Last', new Ref(top[top.length - 1], 0)],
      ['Count', built.reduce((sum, k) => sum + 1 + visible(k), 0)],
    ]);
    catalog.set('Outlines', new Ref(rootNum, 0));
    if (sourceCatalog.get('PageMode') instanceof Name) catalog.set('PageMode', sourceCatalog.get('PageMode'));
  }

  // Page labels: each page called what it was, as the fewest ranges.
  const labelled = new Map(files.map((f) => [f, pageLabels(f)]));
  if ([...labelled.values()].some(Boolean)) {
    const nums = [];
    let prev = null;
    items.forEach((item, i) => {
      const label = labelled.get(item.file)?.[item.page] || { style: 'D', prefix: '', number: item.page + 1 };
      const continues = prev && prev.style === label.style && prev.prefix === label.prefix && label.number === prev.number + 1;
      if (!continues) {
        const d = new Dict();
        if (label.style) d.set('S', new Name(label.style));
        if (label.prefix) d.set('P', textString(label.prefix));
        if (label.number !== 1) d.set('St', label.number);
        nums.push(i, d);
      }
      prev = label;
    });
    catalog.set('PageLabels', new Dict([['Nums', nums]]));
  }

  // The form: the fields whose widgets came with the pages.
  const fieldRoots = [];
  let form = null;
  const names = new Set();
  for (const file of files) {
    const own = file.resolve(file.catalog().get('AcroForm'));
    if (!(own instanceof Dict)) continue;
    const roots = file.resolve(own.get('Fields'));
    const c = copiers.get(file);
    const kept = (Array.isArray(roots) ? roots : []).filter((r) => r instanceof Ref && c.map.has(r.key) && c.map.get(r.key) !== null);
    if (!kept.length) continue;
    if (!form) {
      form = new Dict();
      for (const key of ['DA', 'Q', 'NeedAppearances', 'DR']) if (own.has(key)) form.set(key, copy(file, own.get(key), 1));
    } else if (own.get('DR')) {
      // Fonts the other form's fields draw with, beside the first's.
      const theirs = file.resolve(file.resolve(own.get('DR'))?.get('Font'));
      const dr = form.get('DR') instanceof Dict ? form.get('DR') : new Dict();
      const fonts = dr.get('Font') instanceof Dict ? dr.get('Font') : new Dict();
      if (theirs instanceof Dict) for (const [k, v] of theirs.entries()) if (!fonts.has(k)) fonts.set(k, copy(file, v, 1));
      dr.set('Font', fonts);
      form.set('DR', dr);
    }
    for (const r of kept) {
      const num = c.map.get(r.key);
      const field = objects[num];
      if (field instanceof Dict && field.get('T') instanceof PdfString) {
        let name = field.get('T').text();
        if (names.has(name)) {
          let k = 2;
          while (names.has(`${name}_${k}`)) k++;
          name = `${name}_${k}`;
          field.set('T', textString(name));
        }
        names.add(name);
      }
      fieldRoots.push(new Ref(num, 0));
    }
  }
  drain();
  if (form) {
    form.set('Fields', fieldRoots);
    catalog.set('AcroForm', form);
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
