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
 * **Font embedding only where it must.** See `metrics.js`: the base-14
 * fonts are the viewer's, which is what makes our measurements exact, and
 * they are used for everything they can say. Text they cannot — Arabic,
 * Hebrew, Greek, Cyrillic — is drawn in a TrueType font the caller hands
 * over (`unicodeFont`, see `truetype.js`), subset to the glyphs used,
 * Arabic joined (`shaping.js`) and right-to-left text put in drawing order
 * (`bidi.js`). Without one, such text prints as '?' as it always did.
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
const { pdfString, encode, fitsWinAnsi } = require('./encoding');
const { widthOfBytes, fontOrThrow, FONTS } = require('./metrics');
const { readTrueType, subsetTrueType } = require('./truetype');
const { shapeArabic, hasArabic, baseLetters } = require('./shaping');
const { visualText, hasRtl } = require('./bidi');

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
  text(value, x, y, { font = 'Helvetica', size = 10, colour: fill = null, late = false, rotate = 0, stroke = null, strokeWidth = 0.5, rtl = null, visual = false } = {}) {
    // What the base-14 fonts cannot say goes out in the embedded font.
    // (A page an incremental update draws on belongs to no PdfDocument, and has no embedded font.)
    const embedded = typeof this.doc._embeddedFor === 'function' ? this.doc._embeddedFor(value, font) : null;
    if (embedded) return this._textEmbedded(embedded, value, x, y, { font, size, fill, late, rotate, stroke, strokeWidth, rtl, visual });
    const name = this._use(font);
    const bytes = encode(value);
    if (!bytes.length) return 0;
    const ops = [];
    const rgb = colour(fill);
    // A stroke outlines the glyphs instead of filling them — Word's Outline
    // text effect, hollow letters with just their edge inked. Render mode 1
    // strokes only; 2 would fill and stroke too, which nothing here asks for.
    const strokeRgb = colour(stroke);
    const wrapped = Boolean(rgb || strokeRgb);
    if (wrapped) ops.push('q');
    if (rgb) ops.push(`${rgb.join(' ')} rg`);
    if (strokeRgb) ops.push(`${strokeRgb.join(' ')} RG`, `${num(strokeWidth)} w`);
    // `rotate` turns the text counter-clockwise about its own origin, in
    // degrees — a watermark rising across the page. The text matrix carries
    // the rotation; nothing else on the page is touched.
    const rad = (Number(rotate) || 0) * (Math.PI / 180);
    const c = Math.cos(rad);
    const s = Math.sin(rad);
    ops.push(
      'BT',
      `/${name} ${num(size)} Tf`,
      ...(strokeRgb ? [`${rgb ? 2 : 1} Tr`] : []),
      `${num(c)} ${num(s)} ${num(-s)} ${num(c)} ${num(x)} ${num(this.height - y)} Tm`,
      `${pdfString(value)} Tj`,
      'ET',
    );
    if (wrapped) ops.push('Q');
    (late ? this.lateOps : this.ops).push(ops.join('\n'));
    return widthOfBytes(bytes, font, size);
  }

  /**
   * A line in the embedded font: joined and reordered unless the caller
   * already did (`visual`), each glyph written by its number. Bold and
   * italic, which the one embedded face does not have, are drawn as Word
   * draws a face it lacks: the glyphs thickened with their own stroke, and
   * slanted.
   */
  _textEmbedded(embedded, value, x, y, { font, size, fill, late, rotate, stroke, strokeWidth, rtl, visual }) {
    const glyphs = embedded.glyphsOf(value, { rtl, visual });
    if (!glyphs.length) return 0;
    const name = embedded.ref;
    const rgb = colour(fill);
    const strokeRgb = colour(stroke);
    const bold = /Bold/.test(font);
    const slant = /Oblique|Italic/.test(font) ? 0.2 : 0;
    const ops = ['q'];
    if (rgb) ops.push(`${rgb.join(' ')} rg`);
    if (strokeRgb) ops.push(`${strokeRgb.join(' ')} RG`, `${num(strokeWidth)} w`);
    else if (bold) ops.push(`${(rgb || ['0', '0', '0']).join(' ')} RG`, `${num(size * 0.035)} w`);
    const rad = (Number(rotate) || 0) * (Math.PI / 180);
    const c = Math.cos(rad);
    const s = Math.sin(rad);
    const mode = strokeRgb ? (rgb ? 2 : 1) : bold ? 2 : 0;
    ops.push(
      'BT',
      `/${name} ${num(size)} Tf`,
      ...(mode ? [`${mode} Tr`] : []),
      `${num(c)} ${num(s)} ${num(slant * c - s)} ${num(slant * s + c)} ${num(x)} ${num(this.height - y)} Tm`,
      `<${glyphs.map((g) => g.toString(16).padStart(4, '0')).join('')}> Tj`,
      'ET',
      'Q',
    );
    (late ? this.lateOps : this.ops).push(ops.join('\n'));
    return embedded.widthOfGlyphs(glyphs, size);
  }

  _widthOf(value, opts) {
    if (typeof this.doc.widthOf === 'function') return this.doc.widthOf(value, opts);
    return widthOfBytes(encode(value), opts.font || 'Helvetica', opts.size || 10);
  }

  /** Text whose RIGHT edge sits at x - what every money column wants. */
  textRight(value, x, y, opts = {}) {
    const width = this._widthOf(value, opts);
    return this.text(value, x - width, y, opts);
  }

  /** Text centred on x. */
  textCentre(value, x, y, opts = {}) {
    const width = this._widthOf(value, opts);
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

  /**
   * Turn (and mirror) what is drawn next about a centre given from the
   * page's TOP-left, like everything else here — `deg` clockwise, as a
   * drawing's rotation is — until `restore()`. The graphics state is saved,
   * so nothing drawn after the restore is turned.
   */
  turn(cx, cy, deg = 0, { flipH = false, flipV = false } = {}) {
    const phi = (-(Number(deg) || 0) * Math.PI) / 180;
    const sx = flipH ? -1 : 1;
    const sy = flipV ? -1 : 1;
    const X = cx;
    const Y = this.height - cy;
    const a = Math.cos(phi) * sx;
    const b = Math.sin(phi) * sx;
    const c = -Math.sin(phi) * sy;
    const d = Math.cos(phi) * sy;
    const e = X - (a * X + c * Y);
    const f = Y - (b * X + d * Y);
    this.ops.push(['q', `${num(a)} ${num(b)} ${num(c)} ${num(d)} ${num(e)} ${num(f)} cm`].join('\n'));
  }

  /** The end of a `turn`. */
  restore() {
    this.ops.push('Q');
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

/**
 * One embedded TrueType font: the glyphs a page asks for, their widths, and
 * at assembly the font program subset to them with the dictionaries a PDF
 * names it by — a Type 0 font over a CID font, glyphs numbered as the font
 * numbers them, and a ToUnicode map so the text can be copied back out.
 */
class EmbeddedFont {
  constructor(font, ref) {
    this.font = font;
    this.ref = ref;
    // glyph number -> the text it stands for
    this.used = new Map();
  }

  /** The glyph numbers that draw `value`, in drawing order. */
  glyphsOf(value, { rtl = null, visual = false, record = true } = {}) {
    let text = value === null || value === undefined ? '' : String(value);
    text = text.replace(/[\t\n\r]/g, ' ').replace(/[\u00AD\u200B\u200C\u200D\u200E\u200F]/g, '');
    if (!visual) {
      if (hasArabic(text)) text = shapeArabic(text).text;
      if (hasRtl(text)) text = visualText(text, { rtl });
    }
    const out = [];
    for (const ch of text) {
      const cp = ch.codePointAt(0);
      let gid = this.font.glyphOf(cp);
      const letters = baseLetters(cp);
      // A ligature drawn in a right-to-left line is read back reversed by a
      // viewer copying the text, so its letters are given reversed too.
      let meaning = String.fromCodePoint(...(letters.length > 1 ? [...letters].reverse() : letters));
      // A presentation form the font lacks: the letter itself, unjoined.
      if (!gid && meaning !== ch) gid = this.font.glyphOf(meaning.codePointAt(0));
      if (!gid) { gid = this.font.glyphOf(0x3F); meaning = '?'; }
      out.push(gid);
      if (record && !this.used.has(gid)) this.used.set(gid, meaning);
    }
    return out;
  }

  widthOfGlyphs(glyphs, size) {
    let units = 0;
    for (const g of glyphs) units += this.font.advance(g);
    return (units / this.font.unitsPerEm) * size;
  }

  /** Write the font's objects through `add`; answers the Type 0 font's object number. */
  write(add) {
    const f = this.font;
    const scale = (v) => Math.round((v * 1000) / f.unitsPerEm);
    const gids = [...this.used.keys()].sort((a, b) => a - b);
    // Six letters from the glyphs used: the subset tag the format asks for.
    let h = 0;
    for (const g of gids) h = (Math.imul(h, 31) + g) >>> 0;
    const tag = Array.from({ length: 6 }, (_, i) => String.fromCharCode(65 + ((h >>> (i * 5)) % 26))).join('');
    const baseFont = `${tag}+${f.postscriptName}`;
    const program = subsetTrueType(f, new Set(gids));
    const packed = zlib.deflateSync(program);
    const fileId = add(`<< /Length ${packed.length} /Length1 ${program.length} /Filter /FlateDecode >>\nstream\n${packed.toString('latin1')}\nendstream`);
    const descriptorId = add(
      `<< /Type /FontDescriptor /FontName /${baseFont} /Flags 32 `
      + `/FontBBox [${f.bbox.map(scale).join(' ')}] /ItalicAngle ${num(f.italicAngle)} `
      + `/Ascent ${scale(f.ascent)} /Descent ${scale(f.descent)} /CapHeight ${scale(f.capHeight)} /StemV 80 `
      + `/FontFile2 ${fileId} 0 R >>`
    );
    const widths = gids.map((g) => `${g} [${scale(f.advance(g))}]`).join(' ');
    const cidId = add(
      `<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${baseFont} `
      + `/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> `
      + `/FontDescriptor ${descriptorId} 0 R /DW 1000 /W [${widths}] /CIDToGIDMap /Identity >>`
    );
    const hex4 = (n) => n.toString(16).padStart(4, '0').toUpperCase();
    const utf16 = (s) => Array.from(s).map((ch) => {
      const cp = ch.codePointAt(0);
      if (cp <= 0xFFFF) return hex4(cp);
      const v = cp - 0x10000;
      return hex4(0xD800 + (v >> 10)) + hex4(0xDC00 + (v & 0x3FF));
    }).join('');
    const entries = gids.map((g) => `<${hex4(g)}> <${utf16(this.used.get(g))}>`);
    const blocks = [];
    for (let i = 0; i < entries.length; i += 100) {
      const chunk = entries.slice(i, i + 100);
      blocks.push(`${chunk.length} beginbfchar\n${chunk.join('\n')}\nendbfchar`);
    }
    const cmap = [
      '/CIDInit /ProcSet findresource begin',
      '12 dict begin',
      'begincmap',
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
      '/CMapName /Adobe-Identity-UCS def',
      '/CMapType 2 def',
      '1 begincodespacerange',
      '<0000> <FFFF>',
      'endcodespacerange',
      ...blocks,
      'endcmap',
      'CMapName currentdict /CMap defineresource pop',
      'end',
      'end',
    ].join('\n');
    const cmapId = add(`<< /Length ${Buffer.byteLength(cmap, 'latin1')} >>\nstream\n${cmap}\nendstream`);
    return add(
      `<< /Type /Font /Subtype /Type0 /BaseFont /${baseFont} /Encoding /Identity-H `
      + `/DescendantFonts [${cidId} 0 R] /ToUnicode ${cmapId} 0 R >>`
    );
  }
}

class PdfDocument {
  constructor({
    size = 'A4', title = '', author = '', subject = '', created = null, landscape = false, unicodeFont = null,
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
    this.embedded = null;
    if (unicodeFont) this.useUnicodeFont(unicodeFont);
  }

  /**
   * The TrueType font (its bytes, or `{ bytes, face }` for a face of a
   * collection) that draws what the base-14 fonts cannot. Answers whether
   * it was taken: a font whose licence forbids embedding, or one this
   * cannot read, is not.
   */
  useUnicodeFont(spec) {
    try {
      const bytes = Buffer.isBuffer(spec) || spec instanceof Uint8Array ? spec : spec.bytes;
      const font = readTrueType(bytes, { face: spec && spec.face ? spec.face : null });
      if (!font.embeddable) return false;
      this.embedded = new EmbeddedFont(font, `F${this.fonts.size + 100}`);
      return true;
    } catch {
      return false;
    }
  }

  /** The embedded font, when `value` needs it — text WinAnsi cannot say, or text to join or reorder. */
  _embeddedFor(value, font) {
    if (!this.embedded) return null;
    const text = value === null || value === undefined ? '' : String(value);
    return fitsWinAnsi(text) ? null : this.embedded;
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
  widthOf(value, { font = 'Helvetica', size = 10, rtl = null, visual = false } = {}) {
    const embedded = this._embeddedFor(value, font);
    if (embedded) return embedded.widthOfGlyphs(embedded.glyphsOf(value, { rtl, visual, record: false }), size);
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
    if (this.embedded && this.embedded.used.size) fontIds.set(this.embedded.ref, this.embedded.write(add));
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
