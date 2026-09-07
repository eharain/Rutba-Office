'use strict';

/**
 * Incremental updates: append a page to an EXISTING PDF, or draw ON its
 * existing pages, without touching one original byte (sign S5).
 *
 * An incremental update is the format's own append-only mechanism: the new
 * objects, a new xref section and a new trailer chaining to the old one via
 * /Prev go AFTER the original bytes, so the original remains byte-for-byte
 * present — which is exactly the property a hash-pinned document needs. A
 * viewer reads the newest trailer; a verifier can still hash the prefix.
 *
 * == The honest subset ==================================================
 *
 * Supported: unencrypted PDFs whose cross-references are CLASSIC XREF
 * TABLES (what this package writes, what scanners and most business tools
 * write), whose page tree carries direct /Kids arrays and /Count.
 * Refused BY NAME: encrypted files, xref STREAMS (PDF 1.5+ compressed
 * xref — common from browsers and modern generators), page trees this
 * cannot safely walk, and — for overlays — rotated pages and pages whose
 * resources it cannot merge. A refusal here is a fallback for the caller
 * (the certificate still carries the record), never a corrupted file.
 *
 * Two operations, one assembler:
 *
 *   appendExecutionPage(original, draw)   one new last page (the execution
 *                                         page); redefines the page-tree root
 *   overlayPages(original, drawByPage)    draws on chosen EXISTING pages by
 *                                         redefining each page object: its
 *                                         content array gains a `q` stream
 *                                         before the original streams and a
 *                                         `Q` + overlay stream after them, so
 *                                         whatever graphics state the original
 *                                         left behind cannot displace a mark
 *
 * Redefinition is the other half of what incremental updates are for: the
 * same object number at a new offset, the old body still in the file.
 */

const zlib = require('zlib');
const { Page } = require('./document');
const { FONTS } = require('./metrics');

const A4 = [595.28, 841.89];

class Unsupported extends Error {
  constructor(message) {
    super(message);
    this.name = 'PdfUnsupportedError';
  }
}

const latin = (buf, from, to) => buf.toString('latin1', from, to);

/** The offset the LAST startxref names. */
function lastStartXref(buffer) {
  const tail = latin(buffer, Math.max(0, buffer.length - 2048), buffer.length);
  const at = tail.lastIndexOf('startxref');
  if (at === -1) throw new Unsupported('no startxref - not a well-formed PDF');
  const match = /startxref\s+(\d+)/.exec(tail.slice(at));
  if (!match) throw new Unsupported('startxref names no offset');
  return Number(match[1]);
}

/** Parse one classic xref section + its trailer dict (raw text). */
function parseXrefSection(buffer, offset) {
  if (latin(buffer, offset, offset + 4) !== 'xref') {
    throw new Unsupported('cross-reference stream (PDF 1.5+) - only classic xref tables are supported');
  }
  const text = latin(buffer, offset, Math.min(buffer.length, offset + 4 * 1024 * 1024));
  const trailerAt = text.indexOf('trailer');
  if (trailerAt === -1) throw new Unsupported('xref section has no trailer');

  const entries = new Map(); // objNum -> { offset, gen, type }
  const body = text.slice(4, trailerAt);
  const lines = body.split(/\r\n|\r|\n/).map((l) => l.trim()).filter(Boolean);
  let i = 0;
  while (i < lines.length) {
    const header = /^(\d+)\s+(\d+)$/.exec(lines[i]);
    if (!header) throw new Unsupported(`unreadable xref subsection header: "${lines[i]}"`);
    const first = Number(header[1]);
    const count = Number(header[2]);
    i += 1;
    for (let n = 0; n < count; n++, i++) {
      const entry = /^(\d{10})\s(\d{5})\s([nf])/.exec(lines[i] || '');
      if (!entry) throw new Unsupported('unreadable xref entry');
      const objNum = first + n;
      if (!entries.has(objNum)) {
        entries.set(objNum, { offset: Number(entry[1]), gen: Number(entry[2]), type: entry[3] });
      }
    }
  }

  const dictStart = text.indexOf('<<', trailerAt);
  const dictRaw = balancedDict(text, dictStart);
  return { entries, trailer: dictRaw };
}

/** The raw text of a balanced << ... >> dictionary starting at `at`. */
function balancedDict(text, at) {
  if (text.slice(at, at + 2) !== '<<') throw new Unsupported('expected a dictionary');
  let depth = 0;
  for (let i = at; i < text.length - 1; i++) {
    if (text[i] === '<' && text[i + 1] === '<') { depth += 1; i += 1; continue; }
    if (text[i] === '>' && text[i + 1] === '>') {
      depth -= 1; i += 1;
      if (depth === 0) return text.slice(at, i + 1);
    }
  }
  throw new Unsupported('unterminated dictionary');
}

/** The raw text of a balanced [ ... ] array starting at `at`. */
function balancedArray(text, at) {
  if (text[at] !== '[') throw new Unsupported('expected an array');
  let depth = 0;
  for (let i = at; i < text.length; i++) {
    if (text[i] === '[') depth += 1;
    else if (text[i] === ']') {
      depth -= 1;
      if (depth === 0) return text.slice(at, i + 1);
    }
  }
  throw new Unsupported('unterminated array');
}

const dictRef = (dict, key) => {
  const m = new RegExp(`\\/${key}\\s+(\\d+)\\s+(\\d+)\\s+R`).exec(dict);
  return m ? { num: Number(m[1]), gen: Number(m[2]) } : null;
};

/**
 * The raw value of a TOP-LEVEL key in a dict: a nested dict, an array, a
 * reference, or a bare token — with the span it occupies, so a caller can
 * replace it. Nested dictionaries are skipped while scanning so a /Font
 * inside /Resources is not mistaken for a page-level key.
 */
function dictValue(dict, key) {
  let depth = 0;
  for (let i = 0; i < dict.length - 1; i++) {
    if (dict[i] === '<' && dict[i + 1] === '<') { depth += 1; i += 1; continue; }
    if (dict[i] === '>' && dict[i + 1] === '>') { depth -= 1; i += 1; continue; }
    if (depth !== 1 || dict[i] !== '/') continue;
    const name = /^\/([^\s/[\]<>(){}%]+)/.exec(dict.slice(i));
    if (!name) continue;
    if (name[1] !== key) { i += name[0].length - 1; continue; }
    let at = i + name[0].length;
    while (/\s/.test(dict[at] || '')) at += 1;
    let raw;
    if (dict.slice(at, at + 2) === '<<') raw = balancedDict(dict, at);
    else if (dict[at] === '[') raw = balancedArray(dict, at);
    else {
      const ref = /^(\d+)\s+(\d+)\s+R/.exec(dict.slice(at));
      raw = ref ? ref[0] : (/^[^\s/[\]<>]+/.exec(dict.slice(at)) || [''])[0];
    }
    return { raw, start: i, end: at + raw.length };
  }
  return null;
}

const asRef = (raw) => {
  const m = /^(\d+)\s+(\d+)\s+R$/.exec(String(raw || '').trim());
  return m ? { num: Number(m[1]), gen: Number(m[2]) } : null;
};

/** Merge the xref chain (newest wins) and collect trailer facts. */
function readStructure(buffer) {
  if (latin(buffer, 0, 5) !== '%PDF-') throw new Unsupported('not a PDF');
  const entries = new Map();
  let offset = lastStartXref(buffer);
  let root = null;
  let info = null;
  let size = 0;
  let guard = 0;
  const firstXref = offset;
  while (offset !== null && guard++ < 64) {
    const section = parseXrefSection(buffer, offset);
    for (const [num, entry] of section.entries) {
      if (!entries.has(num)) entries.set(num, entry);
    }
    if (/\/Encrypt\b/.test(section.trailer)) throw new Unsupported('encrypted PDFs are not appended to');
    if (!root) root = dictRef(section.trailer, 'Root');
    if (!info) info = dictRef(section.trailer, 'Info');
    const sizeMatch = /\/Size\s+(\d+)/.exec(section.trailer);
    if (sizeMatch) size = Math.max(size, Number(sizeMatch[1]));
    const prev = /\/Prev\s+(\d+)/.exec(section.trailer);
    offset = prev ? Number(prev[1]) : null;
  }
  if (!root) throw new Unsupported('the trailer names no /Root');
  return { buffer, entries, root, info, size, lastXrefOffset: firstXref };
}

/** The raw body text of object `num` (between "N G obj" and "endobj"). */
function objectBody(buffer, entries, num) {
  const entry = entries.get(num);
  if (!entry || entry.type !== 'n') throw new Unsupported(`object ${num} is not in the xref table`);
  const window = latin(buffer, entry.offset, Math.min(buffer.length, entry.offset + 512 * 1024));
  const open = new RegExp(`^\\s*${num}\\s+\\d+\\s+obj`).exec(window);
  if (!open) throw new Unsupported(`object ${num} is not at its recorded offset`);
  const end = window.indexOf('endobj', open[0].length);
  if (end === -1) throw new Unsupported(`object ${num} has no endobj in reach`);
  return { body: window.slice(open[0].length, end).trim(), gen: entry.gen };
}

/** A raw dict value, following ONE reference if it is one. */
function resolveDict(structure, raw) {
  const ref = asRef(raw);
  if (!ref) return String(raw || '');
  const { body } = objectBody(structure.buffer, structure.entries, ref.num);
  return body;
}

/**
 * The page objects in document order, each with the attributes a page
 * inherits from its ancestors (MediaBox, Resources, Rotate) resolved.
 */
function walkPages(structure) {
  const { buffer, entries, root } = structure;
  const rootObj = objectBody(buffer, entries, root.num);
  const pagesRef = dictRef(rootObj.body, 'Pages');
  if (!pagesRef) throw new Unsupported('the catalog names no /Pages');
  const pages = [];
  const visit = (num, inherited, depth) => {
    if (depth > 64) throw new Unsupported('page tree too deep');
    const obj = objectBody(buffer, entries, num);
    const own = {
      mediaBox: (dictValue(obj.body, 'MediaBox') || {}).raw || inherited.mediaBox,
      resources: (dictValue(obj.body, 'Resources') || {}).raw || inherited.resources,
      rotate: (dictValue(obj.body, 'Rotate') || {}).raw || inherited.rotate,
    };
    const isNode = /\/Type\s*\/Pages\b/.test(obj.body)
      || (!/\/Type\s*\/Page\b/.test(obj.body) && dictValue(obj.body, 'Kids'));
    if (isNode) {
      const kids = dictValue(obj.body, 'Kids');
      if (!kids || !kids.raw.startsWith('[')) {
        throw new Unsupported('a page-tree node has no direct /Kids array - not walked by guesswork');
      }
      for (const m of kids.raw.matchAll(/(\d+)\s+(\d+)\s+R/g)) visit(Number(m[1]), own, depth + 1);
      return;
    }
    pages.push({ num, gen: obj.gen, body: obj.body, ...own });
  };
  visit(pagesRef.num, { mediaBox: null, resources: null, rotate: null }, 0);
  return { pages, pagesRef };
}

// ── the drawing side: a Page plus registrars for fonts and images ────────

function drawingContext(size) {
  const fonts = new Map();
  const images = [];
  const fakeDoc = {
    _fontRef(name) {
      if (!fonts.has(name)) fonts.set(name, `RSF${fonts.size + 1}`);
      return fonts.get(name);
    },
  };
  const page = new Page(fakeDoc, size);
  const addImage = ({ width, height, colorSpace, bitsPerComponent = 8, data, alpha = null }) => {
    const ref = `RSIm${images.length + 1}`;
    images.push({ ref, width, height, colorSpace, bitsPerComponent, data, alpha });
    return ref;
  };
  return { page, addImage, fonts, images };
}

/** Font and image objects for one drawing context → { objects, fontIds, imageIds }. */
function resourceObjects(ctx, allocate) {
  const objects = [];
  const fontIds = new Map();
  for (const [name, ref] of ctx.fonts) {
    const id = allocate();
    fontIds.set(ref, id);
    objects.push({
      num: id, gen: 0,
      body: `<< /Type /Font /Subtype /Type1 /BaseFont /${FONTS[name].base} /Encoding /WinAnsiEncoding >>`,
    });
  }
  const imageIds = new Map();
  for (const image of ctx.images) {
    let smaskId = null;
    if (image.alpha) {
      smaskId = allocate();
      const alphaBytes = zlib.deflateSync(image.alpha);
      objects.push({
        num: smaskId, gen: 0,
        body: `<< /Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} `
          + `/ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode /Length ${alphaBytes.length} >>`
          + `\nstream\n${alphaBytes.toString('latin1')}\nendstream`,
      });
    }
    const id = allocate();
    imageIds.set(image.ref, id);
    const pixelBytes = zlib.deflateSync(image.data);
    objects.push({
      num: id, gen: 0,
      body: `<< /Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} `
        + `/ColorSpace /${image.colorSpace} /BitsPerComponent ${image.bitsPerComponent} `
        + `/Filter /FlateDecode${smaskId ? ` /SMask ${smaskId} 0 R` : ''} /Length ${pixelBytes.length} >>`
        + `\nstream\n${pixelBytes.toString('latin1')}\nendstream`,
    });
  }
  return { objects, fontIds, imageIds };
}

const streamObject = (num, content) => ({
  num, gen: 0,
  body: `<< /Length ${Buffer.byteLength(content, 'latin1')} >>\nstream\n${content}\nendstream`,
});

const resourceEntries = (fontIds, imageIds) => ({
  font: [...fontIds].map(([r, id]) => `/${r} ${id} 0 R`).join(' '),
  xobject: [...imageIds].map(([r, id]) => `/${r} ${id} 0 R`).join(' '),
});

/**
 * Merge our /Font and /XObject entries into an existing resources dict
 * (raw text, references followed one level), producing a new DIRECT dict.
 * An existing sub-dictionary that is itself a reference is read and
 * copied, so nothing the original content streams relied on is hidden
 * behind ours.
 */
function mergeResources(structure, resourcesRaw, entries) {
  let dict = resourcesRaw ? resolveDict(structure, resourcesRaw).trim() : '<< >>';
  if (!dict.startsWith('<<')) throw new Unsupported('page resources are not a dictionary');
  const mergeSub = (key, ours) => {
    if (!ours) return;
    const existing = dictValue(dict, key);
    if (!existing) {
      dict = `${dict.slice(0, dict.length - 2).trimEnd()} /${key} << ${ours} >> >>`;
      return;
    }
    const sub = resolveDict(structure, existing.raw).trim();
    if (!sub.startsWith('<<')) throw new Unsupported(`page /${key} resources are not a dictionary`);
    const inner = sub.slice(2, sub.length - 2).trim();
    const merged = `<< ${inner}${inner ? ' ' : ''}${ours} >>`;
    dict = `${dict.slice(0, existing.start)}/${key} ${merged}${dict.slice(existing.end)}`;
  };
  mergeSub('Font', entries.font);
  mergeSub('XObject', entries.xobject);
  return dict;
}

/** Write the update section after the original bytes and return the file. */
function assemble(structure, objects, nextNum) {
  const { buffer, root, info, lastXrefOffset } = structure;
  let out = latin(buffer, 0, buffer.length);
  if (!out.endsWith('\n')) out += '\n';
  const offsets = new Map();
  for (const object of objects) {
    offsets.set(object.num, Buffer.byteLength(out, 'latin1'));
    out += `${object.num} ${object.gen} obj\n${object.body}\nendobj\n`;
  }

  const xrefAt = Buffer.byteLength(out, 'latin1');
  const pad = (n, w) => String(n).padStart(w, '0');
  out += 'xref\n';
  // Contiguous runs of object numbers, one subsection each - redefined
  // objects sit anywhere in the numbering, new ones at the end.
  const sorted = [...objects].sort((a, b) => a.num - b.num);
  let run = [];
  const flush = () => {
    if (!run.length) return;
    out += `${run[0].num} ${run.length}\n`;
    for (const object of run) out += `${pad(offsets.get(object.num), 10)} ${pad(object.gen, 5)} n \n`;
    run = [];
  };
  for (const object of sorted) {
    if (run.length && object.num !== run[run.length - 1].num + 1) flush();
    run.push(object);
  }
  flush();

  out += `trailer\n<< /Size ${nextNum} /Root ${root.num} ${root.gen} R `
    + `${info ? `/Info ${info.num} ${info.gen} R ` : ''}/Prev ${lastXrefOffset} >>\n`
    + `startxref\n${xrefAt}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

/**
 * Append one drawn page. `draw({ page, addImage })` receives a Page (A4,
 * top-down coordinates like every page in this package) and an image
 * registrar taking images.js-decoded pixels; whatever it draws becomes the
 * document's new last page.
 *
 * Returns the WHOLE new file as a Buffer whose prefix is the original,
 * byte for byte.
 */
function appendExecutionPage(original, draw) {
  const buffer = Buffer.isBuffer(original) ? original : Buffer.from(original);
  const structure = readStructure(buffer);
  const { entries, root, size } = structure;

  const rootObj = objectBody(buffer, entries, root.num);
  const pagesRef = dictRef(rootObj.body, 'Pages');
  if (!pagesRef) throw new Unsupported('the catalog names no /Pages');
  const pagesObj = objectBody(buffer, entries, pagesRef.num);
  if (!/\/Kids\s*\[/.test(pagesObj.body) || !/\/Count\s+\d+/.test(pagesObj.body)) {
    throw new Unsupported('the page tree root has no direct /Kids and /Count - not rewritten by guesswork');
  }

  const ctx = drawingContext(A4);
  draw({ page: ctx.page, addImage: ctx.addImage });

  let nextNum = size;
  const allocate = () => nextNum++;
  const { objects, fontIds, imageIds } = resourceObjects(ctx, allocate);
  const contentId = allocate();
  objects.push(streamObject(contentId, ctx.page.content));

  const { font, xobject } = resourceEntries(fontIds, imageIds);
  const resourceParts = [];
  if (font) resourceParts.push(`/Font << ${font} >>`);
  if (xobject) resourceParts.push(`/XObject << ${xobject} >>`);
  const pageId = allocate();
  objects.push({
    num: pageId, gen: 0,
    body: `<< /Type /Page /Parent ${pagesRef.num} ${pagesRef.gen} R `
      + `/MediaBox [0 0 ${A4[0]} ${A4[1]}] `
      + `/Resources << ${resourceParts.join(' ')} >> /Contents ${contentId} 0 R >>`,
  });

  // The one redefinition: the page-tree root, with our page in its Kids.
  const newPagesBody = pagesObj.body
    .replace(/\/Kids\s*\[([\s\S]*?)\]/, (m, kids) => `/Kids [${kids.trim()} ${pageId} 0 R]`)
    .replace(/\/Count\s+(\d+)/, (m, count) => `/Count ${Number(count) + 1}`);
  objects.push({ num: pagesRef.num, gen: pagesObj.gen, body: newPagesBody });

  return assemble(structure, objects, nextNum);
}

/**
 * Draw on existing pages. `drawByPage` maps a 1-based page number to a
 * `draw({ page, addImage, width, height })` callback whose Page is sized to
 * that page's MediaBox, top-down coordinates in points like everything
 * here. Pages not named are untouched; a named page beyond the document is
 * a refusal, as is a rotated page (a mark placed by top-left coordinates on
 * a page the viewer turns would land where nobody meant it).
 *
 * Returns the whole new file, original bytes as its exact prefix. Compose
 * with appendExecutionPage freely - each update chains to the last.
 */
function overlayPages(original, drawByPage) {
  const buffer = Buffer.isBuffer(original) ? original : Buffer.from(original);
  const structure = readStructure(buffer);
  const { pages } = walkPages(structure);
  const wanted = Object.entries(drawByPage || {})
    .map(([n, draw]) => ({ index: Number(n), draw }))
    .filter((w) => typeof w.draw === 'function');
  if (!wanted.length) return buffer;

  let nextNum = structure.size;
  const allocate = () => nextNum++;
  const objects = [];
  // One `q` stream shared by every overlaid page: a content array is
  // concatenated, so this brackets the original streams' graphics state.
  const qId = allocate();
  objects.push(streamObject(qId, 'q'));

  for (const { index, draw } of wanted) {
    const target = pages[index - 1];
    if (!target) throw new Unsupported(`page ${index} does not exist (the document has ${pages.length})`);
    const rotate = Number(String(target.rotate || '0').trim()) || 0;
    if (rotate % 360 !== 0) {
      throw new Unsupported(`page ${index} is rotated (${rotate}) - overlays are not placed on rotated pages`);
    }
    const box = target.mediaBox ? resolveDict(structure, target.mediaBox) : null;
    const nums = box ? [...box.matchAll(/-?\d+(?:\.\d+)?/g)].map((m) => Number(m[0])) : [];
    if (nums.length !== 4) throw new Unsupported(`page ${index} has no readable /MediaBox`);
    const x0 = Math.min(nums[0], nums[2]);
    const y0 = Math.min(nums[1], nums[3]);
    const width = Math.max(nums[0], nums[2]) - x0;
    const height = Math.max(nums[1], nums[3]) - y0;

    const ctx = drawingContext([width, height]);
    draw({ page: ctx.page, addImage: ctx.addImage, width, height });
    const drawn = ctx.page.content;
    if (!drawn.trim()) continue;

    const { objects: resObjects, fontIds, imageIds } = resourceObjects(ctx, allocate);
    objects.push(...resObjects);
    // Q restores what the original streams did; the translation maps our
    // origin onto a MediaBox that does not start at 0,0.
    const translate = (x0 || y0) ? `1 0 0 1 ${x0} ${y0} cm\n` : '';
    const overlayId = allocate();
    objects.push(streamObject(overlayId, `Q\nq\n${translate}${drawn}\nQ`));

    // The redefined page: contents bracketed, resources merged.
    const contents = dictValue(target.body, 'Contents');
    const originalRefs = contents
      ? (contents.raw.startsWith('[') ? contents.raw.slice(1, -1).trim() : contents.raw.trim())
      : '';
    const newContents = `[${qId} 0 R ${originalRefs}${originalRefs ? ' ' : ''}${overlayId} 0 R]`;
    const merged = mergeResources(structure, target.resources, resourceEntries(fontIds, imageIds));
    let body = target.body;
    body = contents
      ? `${body.slice(0, contents.start)}/Contents ${newContents}${body.slice(contents.end)}`
      : `${body.slice(0, body.length - 2).trimEnd()} /Contents ${newContents} >>`;
    const res = dictValue(body, 'Resources');
    body = res
      ? `${body.slice(0, res.start)}/Resources ${merged}${body.slice(res.end)}`
      : `${body.slice(0, body.length - 2).trimEnd()} /Resources ${merged} >>`;
    objects.push({ num: target.num, gen: target.gen, body });
  }

  if (objects.length === 1) return buffer; // every named page drew nothing
  return assemble(structure, objects, nextNum);
}

/** How many pages, and each page's size - for placement UIs and anchors. */
function describePages(original) {
  const buffer = Buffer.isBuffer(original) ? original : Buffer.from(original);
  const structure = readStructure(buffer);
  const { pages } = walkPages(structure);
  return pages.map((p, i) => {
    const box = p.mediaBox ? resolveDict(structure, p.mediaBox) : '';
    const nums = [...box.matchAll(/-?\d+(?:\.\d+)?/g)].map((m) => Number(m[0]));
    return {
      page: i + 1,
      width: nums.length === 4 ? Math.abs(nums[2] - nums[0]) : null,
      height: nums.length === 4 ? Math.abs(nums[3] - nums[1]) : null,
      rotate: Number(String(p.rotate || '0').trim()) || 0,
    };
  });
}

module.exports = { appendExecutionPage, overlayPages, describePages, Unsupported };
