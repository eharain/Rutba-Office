'use strict';

/**
 * The PDF file itself: objects, pages, a cross-reference table.
 *
 * Small enough to read in one sitting, which is the point. A PDF is a list
 * of numbered objects, a table saying what byte each one starts at, and a
 * trailer pointing at the table. The only thing it is genuinely fussy about
 * is that those byte offsets are RIGHT - a file with a good xref opens
 * everywhere, and a file with a bad one opens in some viewers, opens
 * differently in others, and is rejected by the strict ones. So offsets are
 * measured from the assembled bytes at the end rather than predicted while
 * writing.
 *
 * == Two deliberate omissions ===========================================
 *
 * **No compression.** An invoice is a few kilobytes of text; deflating it
 * would save an email nothing worth measuring and would make every stream
 * opaque to `grep`. Being able to read a generated file - and to assert on
 * its contents in a test without a parser - is worth more here than the
 * bytes. The seam for it is one function if that ever stops being true.
 *
 * **No font embedding.** See `metrics.js`: the base-14 fonts are the
 * viewer's, which is what makes our measurements exact.
 *
 * == Coordinates ========================================================
 *
 * PDF puts the origin at the bottom-left with y increasing upwards, and
 * every human laying out a page thinks top-down. Rather than let both
 * conventions circulate, THIS FILE IS THE ONLY PLACE THEY MEET: every method
 * below takes y as a distance DOWN from the top of the page and flips it
 * once, on the way into the content stream. Nothing above here ever sees a
 * PDF coordinate.
 */

const zlib = require('zlib');
const { pdfString, encode } = require('./encoding');
const { widthOfBytes, fontOrThrow, FONTS } = require('./metrics');

/** Paper, in points. 72pt to the inch. */
const SIZES = {
  A4: [595.28, 841.89],
  A5: [419.53, 595.28],
  Letter: [612, 792],
  Legal: [612, 1008],
};

const num = (v) => {
  const n = Number(v);
  if (!Number.isFinite(n)) throw new Error(`${v} is not a usable coordinate.`);
  // Three decimals is a thousandth of a point - far finer than any output
  // device - and keeps the file free of 0.30000000000000004.
  return String(Math.round(n * 1000) / 1000);
};

/** A colour as PDF wants it: three components, 0..1. */
function colour(value) {
  if (value === null || value === undefined) return null;
  if (Array.isArray(value)) return value.map((c) => num(Math.min(1, Math.max(0, Number(c) || 0))));
  const hex = String(value).replace('#', '');
  const full = hex.length === 3 ? hex.split('').map((c) => c + c).join('') : hex;
  if (!/^[0-9a-fA-F]{6}$/.test(full)) throw new Error(`"${value}" is not a colour this can draw with.`);
  return [0, 2, 4].map((i) => num(parseInt(full.slice(i, i + 2), 16) / 255));
}

class Page {
  constructor(doc, [width, height]) {
    this.doc = doc;
    this.width = width;
    this.height = height;
    this.ops = [];
    // Drawn after everything else, so a footer written once the page count
    // is finally known still lands on top of nothing.
    this.lateOps = [];
  }

  _use(font) {
    fontOrThrow(font);
    return this.doc._fontRef(font);
  }

  /**
   * One line of text, with `y` the distance from the top of the page to the
   * BASELINE. Baselines rather than box tops because that is what the format
   * positions by, and converting in two places is how text drifts by an
   * ascent.
   */
  text(value, x, y, { font = 'Helvetica', size = 10, colour: fill = null, late = false } = {}) {
    const name = this._use(font);
    const bytes = encode(value);
    if (!bytes.length) return 0;
    const ops = [];
    const rgb = colour(fill);
    if (rgb) ops.push('q', `${rgb.join(' ')} rg`);
    ops.push(
      'BT',
      `/${name} ${num(size)} Tf`,
      `1 0 0 1 ${num(x)} ${num(this.height - y)} Tm`,
      `${pdfString(value)} Tj`,
      'ET',
    );
    if (rgb) ops.push('Q');
    (late ? this.lateOps : this.ops).push(ops.join('\n'));
    return widthOfBytes(bytes, font, size);
  }

  /** Text whose RIGHT edge sits at x - what every money column wants. */
  textRight(value, x, y, opts = {}) {
    const width = widthOfBytes(encode(value), opts.font || 'Helvetica', opts.size || 10);
    return this.text(value, x - width, y, opts);
  }

  /** Text centred on x. */
  textCentre(value, x, y, opts = {}) {
    const width = widthOfBytes(encode(value), opts.font || 'Helvetica', opts.size || 10);
    return this.text(value, x - width / 2, y, opts);
  }

  line(x1, y1, x2, y2, { width = 0.5, colour: stroke = '#000000', late = false } = {}) {
    const rgb = colour(stroke);
    (late ? this.lateOps : this.ops).push([
      'q', `${num(width)} w`, `${rgb.join(' ')} RG`,
      `${num(x1)} ${num(this.height - y1)} m`,
      `${num(x2)} ${num(this.height - y2)} l`,
      'S', 'Q',
    ].join('\n'));
  }

  /**
   * An image, given its TOP-left corner like everything else here. `ref`
   * comes from `PdfDocument.addImage`. The cm matrix places the unit image
   * square at the bottom-left, so the flip happens here - once - like text's.
   */
  image(ref, x, y, w, h, { late = false } = {}) {
    (late ? this.lateOps : this.ops).push([
      'q',
      `${num(w)} 0 0 ${num(h)} ${num(x)} ${num(this.height - y - h)} cm`,
      `/${ref} Do`,
      'Q',
    ].join('\n'));
  }

  /** A rectangle given its TOP-left corner, like everything else here. */
  rect(x, y, w, h, { fill = null, stroke = null, width = 0.5, late = false } = {}) {
    if (!fill && !stroke) return;
    const ops = ['q'];
    if (fill) ops.push(`${colour(fill).join(' ')} rg`);
    if (stroke) ops.push(`${colour(stroke).join(' ')} RG`, `${num(width)} w`);
    ops.push(`${num(x)} ${num(this.height - y - h)} ${num(w)} ${num(h)} re`);
    ops.push(fill && stroke ? 'B' : (fill ? 'f' : 'S'));
    ops.push('Q');
    (late ? this.lateOps : this.ops).push(ops.join('\n'));
  }

  get content() {
    return this.ops.concat(this.lateOps).join('\n');
  }
}

class PdfDocument {
  constructor({
    size = 'A4', title = '', author = '', subject = '', created = null, landscape = false,
  } = {}) {
    const paper = Array.isArray(size) ? size : SIZES[size];
    if (!paper) throw new Error(`${size} is not a paper size this knows (${Object.keys(SIZES).join(', ')}).`);
    this.size = landscape ? [paper[1], paper[0]] : paper.slice();
    this.title = title;
    this.author = author;
    this.subject = subject;
    // Injectable so a test can produce the same bytes twice. Defaulting to
    // "now" is right for a real document and useless for comparing two.
    this.created = created ? new Date(created) : new Date();
    this.pages = [];
    this.fonts = new Map();
    this.images = [];
  }

  _fontRef(name) {
    if (!this.fonts.has(name)) this.fonts.set(name, `F${this.fonts.size + 1}`);
    return this.fonts.get(name);
  }

  /**
   * Register decoded image samples (see images.js) and get back the name a
   * page draws with. Deflation happens at assembly - the one compressed
   * stream kind in the file, because raw RGB at even signature size would
   * triple the document, and pixels were never greppable anyway.
   */
  addImage({ width, height, colorSpace, bitsPerComponent = 8, data, alpha = null }) {
    if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) {
      throw new Error('an image needs integer pixel dimensions');
    }
    if (colorSpace !== 'DeviceRGB' && colorSpace !== 'DeviceGray') {
      throw new Error(`${colorSpace} is not an image colour space this writes`);
    }
    if (!Buffer.isBuffer(data)) throw new Error('image data must be a Buffer of raw samples');
    const ref = `Im${this.images.length + 1}`;
    this.images.push({ ref, width, height, colorSpace, bitsPerComponent, data, alpha });
    return ref;
  }

  addPage(size = null) {
    const paper = size ? (Array.isArray(size) ? size : SIZES[size]) : this.size;
    const page = new Page(this, paper);
    this.pages.push(page);
    return page;
  }

  /** Width of a string if it were drawn - the whole measurement story. */
  widthOf(value, { font = 'Helvetica', size = 10 } = {}) {
    return widthOfBytes(encode(value), font, size);
  }

  _date() {
    const d = this.created;
    const p = (n) => String(n).padStart(2, '0');
    return `D:${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}`
      + `${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`;
  }

  /**
   * Assemble the file.
   *
   * Objects go out in a fixed order - catalog, page tree, fonts, then a page
   * and its content stream in pairs - and offsets are taken from the bytes
   * as they are appended, so the xref cannot drift from the body no matter
   * what the body turned out to contain.
   */
  toBuffer() {
    if (!this.pages.length) throw new Error('A PDF with no pages is not a document. Add a page first.');

    const objects = [];      // 1-based; objects[i] is object i+1
    const add = (body) => { objects.push(body); return objects.length; };

    const catalogId = add(null);   // patched once the page tree is numbered
    const pagesId = add(null);

    const fontIds = new Map();
    for (const [name, ref] of this.fonts) {
      const base = FONTS[name].base;
      fontIds.set(ref, add(
        `<< /Type /Font /Subtype /Type1 /BaseFont /${base} /Encoding /WinAnsiEncoding >>`
      ));
    }
    const imageIds = new Map();
    for (const image of this.images) {
      let smaskId = null;
      if (image.alpha) {
        const alphaBytes = zlib.deflateSync(image.alpha);
        smaskId = add(
          `<< /Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} `
          + `/ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode `
          + `/Length ${alphaBytes.length} >>\nstream\n${alphaBytes.toString('latin1')}\nendstream`
        );
      }
      const pixelBytes = zlib.deflateSync(image.data);
      imageIds.set(image.ref, add(
        `<< /Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} `
        + `/ColorSpace /${image.colorSpace} /BitsPerComponent ${image.bitsPerComponent} `
        + `/Filter /FlateDecode${smaskId ? ` /SMask ${smaskId} 0 R` : ''} `
        + `/Length ${pixelBytes.length} >>\nstream\n${pixelBytes.toString('latin1')}\nendstream`
      ));
    }

    const parts = [];
    if (fontIds.size) {
      parts.push(`/Font << ${[...fontIds].map(([ref, id]) => `/${ref} ${id} 0 R`).join(' ')} >>`);
    }
    if (imageIds.size) {
      parts.push(`/XObject << ${[...imageIds].map(([ref, id]) => `/${ref} ${id} 0 R`).join(' ')} >>`);
    }
    const resources = parts.length ? `<< ${parts.join(' ')} >>` : '<< >>';

    const pageIds = [];
    for (const page of this.pages) {
      const content = page.content;
      const streamId = add(
        `<< /Length ${Buffer.byteLength(content, 'latin1')} >>\nstream\n${content}\nendstream`
      );
      pageIds.push(add(
        `<< /Type /Page /Parent ${pagesId} 0 R `
        + `/MediaBox [0 0 ${num(page.width)} ${num(page.height)}] `
        + `/Resources ${resources} /Contents ${streamId} 0 R >>`
      ));
    }

    objects[catalogId - 1] = `<< /Type /Catalog /Pages ${pagesId} 0 R >>`;
    objects[pagesId - 1] =
      `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pageIds.length} >>`;

    const infoId = add([
      '<<',
      this.title ? ` /Title ${pdfString(this.title)}` : '',
      this.author ? ` /Author ${pdfString(this.author)}` : '',
      this.subject ? ` /Subject ${pdfString(this.subject)}` : '',
      ` /Producer ${pdfString('Rutba')}`,
      ` /CreationDate ${pdfString(this._date())}`,
      '>>',
    ].filter(Boolean).join(''));

    let out = '%PDF-1.4\n';
    // A binary comment in the header is the convention that tells a
    // transport this is not a text file. Without it, something well-meaning
    // rewrites the line endings and every offset below is then wrong.
    out += '%\xE2\xE3\xCF\xD3\n';
    const offsets = [];
    objects.forEach((body, i) => {
      offsets[i] = Buffer.byteLength(out, 'latin1');
      out += `${i + 1} 0 obj\n${body}\nendobj\n`;
    });

    const xrefAt = Buffer.byteLength(out, 'latin1');
    out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
    for (const offset of offsets) out += `${String(offset).padStart(10, '0')} 00000 n \n`;

    const id = fingerprint(out);
    out += `trailer\n<< /Size ${objects.length + 1} /Root ${catalogId} 0 R /Info ${infoId} 0 R `
      + `/ID [<${id}> <${id}>] >>\nstartxref\n${xrefAt}\n%%EOF\n`;

    return Buffer.from(out, 'latin1');
  }

  toBase64() { return this.toBuffer().toString('base64'); }
}

/**
 * A 16-byte file id. Not a security hash and not claimed to be one - the
 * format wants a value that distinguishes one revision of a file from
 * another, and this is derived from the content so that identical content
 * gets an identical id and a test can compare two runs.
 */
function fingerprint(text) {
  let a = 0x811c9dc5;
  let b = 0x01000193;
  for (let i = 0; i < text.length; i++) {
    a = Math.imul(a ^ text.charCodeAt(i), 0x01000193) >>> 0;
    b = Math.imul(b + text.charCodeAt(i) + i, 0x85ebca6b) >>> 0;
  }
  const half = (n) => (n >>> 0).toString(16).padStart(8, '0');
  return (half(a) + half(b) + half(a ^ b) + half(Math.imul(a, b))).toUpperCase();
}

module.exports = { PdfDocument, Page, SIZES, colour, fingerprint };
