/**
 * A document VIEW over a package — the `.docx` counterpart of `workbook.js`,
 * and the first thing `workspace.docs` has ever had.
 *
 * Same discipline, same reasons:
 *
 * 1. Parts we do not edit are never rewritten (the package layer guarantees it).
 * 2. Inside `word/document.xml`, everything outside the runs we touch is held
 *    verbatim — `w:sectPr`, `w:fldSimple`, `w:bookmarkStart`, `w:commentRangeStart`,
 *    `w:numPr`, tables, `w:pStyle`. These are how a real template carries its
 *    intelligence: lose the field code and the reference number stops updating;
 *    lose `w:sectPr` and the letterhead margins go.
 * 3. A run we DO edit keeps its `w:rPr`, so changing text never changes how it
 *    is formatted.
 *
 * CONTENT CONTROLS ARE THE BINDING TARGET. `w:sdt` with a `w:tag` is the
 * document equivalent of a spreadsheet's named range: a stable, named,
 * user-invisible anchor that survives editing and round trips through Word. That
 * is what lets a bank-confirmation template carry a live AR balance without us
 * needing to understand the rest of the letter.
 *
 * What this deliberately does not do: layout, pagination, rendering, styles
 * resolution, or track-changes semantics. It edits named anchors and paragraph
 * text in a file it otherwise leaves alone. See FORMAT-FIDELITY.md.
 */
import { OoxmlPackage, attrs, esc } from './package.js';
import { unesc } from './workbook.js';
// The creator's chart-part writer, reused by insert-chart: the editor still
// cannot create a PACKAGE, but a chart part is a part like numbering.xml.
import { chartPartXml } from './build.js';
import { parseTable, parseSection, childElements, firstElement, headBefore } from './table.js';
import { readHeadersAndFooters } from './headers.js';
import { readParagraphStyles, readCharacterStyles, readNumberingDefs, readThemeFonts, readThemeColours, STANDARD_STYLES_XML } from './docstyles.js';

/**
 * Media bytes -> data URI. Deliberately duplicated from `@rutba/drawing` rather
 * than imported: the drawing package's OOXML adapter depends on THIS layer being
 * format-only, and importing drawing here would run the arrow backwards.
 */
const MEDIA_TYPES = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp',
};
function toDataUri(bytes, partName) {
  const type = MEDIA_TYPES[String(partName ?? '').split('.').pop().toLowerCase()];
  // EMF and WMF are common in Office files and not web image types — no href,
  // rather than a broken one; the caller can still say what it was.
  if (!type || !bytes) return null;
  return 'data:' + type + ';base64,' + Buffer.from(bytes).toString('base64');
}

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export {
  textOf, parseRuns, hasToggle, withToggle, renderRuns, renderRun, firstRunProps, RPR_RE,
} from './runs.js';
import {
  textOf, parseRuns, renderRuns, renderRun, firstRunProps, RPR_RE,
} from './runs.js';

/**
 * An explicit page break. `w:pageBreakBefore` is a paragraph property; a
 `w:br w:type="page"` is a run-level break sitting inside the paragraph. Word
 * honours both, so both have to be seen or a report breaks in the wrong places.
 */
const PAGE_BREAK_BEFORE = /<w:pageBreakBefore\b(?![^>]*w:val="(?:0|false)")/;
const EXPLICIT_BREAK = /<w:br\b[^>]*w:type="page"/;
/**
 * The two "keep" properties a paginator honours: a heading that stays with
 * the paragraph after it, and a paragraph whose lines are not split across
 * pages. Both are toggles that may be written as w:val="0" to turn off.
 */
const KEEP_NEXT = /<w:keepNext\b(?![^>]*w:val="(?:0|false)")/;
const KEEP_LINES = /<w:keepLines\b(?![^>]*w:val="(?:0|false)")/;

/**
 * A paragraph's own `<w:spacing>` — the DIRECT formatting, distinct from the
 * spacing its style resolves (docstyles reads that). Neutral units on the way
 * out: pixels for the gaps, a multiplier for `lineRule="auto"` (240 twentieths
 * = single), pixels for an exact/atLeast line. Null when the paragraph sets
 * nothing, so a consumer can fall back to the style's answer.
 */
const twipsToPx = (tw) => Number(tw) * (96 / 1440);
function readDirectSpacing(pPr) {
  // "Don't add space between paragraphs of the same style": the flag a list
  // item carries so the list sits tight while the list keeps its space from
  // the prose around it. Resolved against the neighbours by the view.
  const contextual = /<w:contextualSpacing\b(?![^>]*w:val="(?:0|false)")/.test(pPr || '');
  const spacing = /<w:spacing\b([^>]*?)\/?>/.exec(pPr || '');
  if (!spacing) return contextual ? { contextual: true } : null;

  const attr = (name) => {
    const m = new RegExp('\\b' + name + '="([^"]*)"').exec(spacing[1]);
    return m ? m[1] : undefined;
  };
  const out = {};
  if (attr('w:before') !== undefined) out.beforePx = twipsToPx(attr('w:before'));
  if (attr('w:after') !== undefined) out.afterPx = twipsToPx(attr('w:after'));
  const line = attr('w:line');
  if (line !== undefined) {
    const rule = attr('w:lineRule') ?? 'auto';
    if (rule === 'auto') out.lineFactor = Number(line) / 240;
    else out.lineExactPx = twipsToPx(line);
  }
  if (contextual) out.contextual = true;
  return Object.keys(out).length ? out : null;
}


/**
 * The rest of a paragraph's DIRECT formatting that the page has to draw: its
 * tab stops, its shading, its borders, and the two indents `w:ind` carries
 * besides the left one. These are the corpus's most-used features after
 * numbering and tables — a form is "Name:<tab>______", a heading band is
 * shading, a rule under a title is a bottom border — and a page that drops
 * them looks nothing like Word's, however right the text is.
 *
 * The paragraph-mark run properties (`<w:rPr>` inside `<w:pPr>`) are cut out
 * first: a shaded pilcrow is not a shaded paragraph.
 */
const eighthsToPx = (sz) => Math.max(1, Math.round((Number(sz) / 8) * (96 / 72)));

/** The two notes parts: where they live, what they are, how they are styled. */
const NOTE_PARTS = {
  footnote: {
    part: 'word/footnotes.xml', target: 'footnotes.xml',
    ct: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
    textStyle: 'FootnoteText', refStyle: 'FootnoteReference',
  },
  endnote: {
    part: 'word/endnotes.xml', target: 'endnotes.xml',
    ct: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
    textStyle: 'EndnoteText', refStyle: 'EndnoteReference',
  },
};

/**
 * An empty notes part as Word writes one: the separator (the short rule
 * between body and notes) and the continuation separator, at ids −1 and 0,
 * before any note. A part without them is one Word repairs on open.
 */
function emptyNotesXml(kind) {
  const sep = (type, id, el) =>
    '<w:' + kind + ' w:type="' + type + '" w:id="' + id + '"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:' + el + '/></w:r></w:p></w:' + kind + '>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:' + kind + 's xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + sep('separator', -1, 'separator') + sep('continuationSeparator', 0, 'continuationSeparator')
    + '</w:' + kind + 's>';
}

/** A paragraph's XML without the text boxes it anchors (and their VML twins). */
const stripTextBoxes = (xml) => String(xml)
  .replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '')
  .replace(/<w:txbxContent\b[^>]*>[\s\S]*?<\/w:txbxContent>/g, '');

/**
 * A DrawingML colour in a fragment, as CSS: `<a:srgbClr val="7E97AD"/>` as it
 * is, `<a:schemeClr val="accent1"/>` through the theme, with the lumMod/lumOff
 * children that make "Accent 1, lighter 60%" applied in HSL the way Word
 * applies them. Null for no fill or nothing recognisable.
 */
function colourOf(fragment, colours) {
  const xml = String(fragment || '');
  if (!xml || /^\s*$/.test(xml)) return null;
  const solid = /<a:solidFill>([\s\S]*?)<\/a:solidFill>/.exec(xml);
  const inner = solid ? solid[1] : xml;
  if (!solid && /<a:noFill\b/.test(xml)) return null;
  let hex = null;
  const srgb = /<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(inner);
  const scheme = /<a:schemeClr\b[^>]*\bval="([^"]*)"/.exec(inner);
  if (srgb) hex = srgb[1];
  else if (scheme) hex = colours?.[SCHEME_ALIAS[scheme[1]] ?? scheme[1]] ?? null;
  if (!hex) return null;
  const mod = /<a:lumMod\b[^>]*\bval="(\d+)"/.exec(inner);
  const off = /<a:lumOff\b[^>]*\bval="(-?\d+)"/.exec(inner);
  if (!mod && !off) return '#' + hex.toUpperCase();
  const [h, s, l] = hexToHsl(hex);
  const l2 = Math.max(0, Math.min(1, l * (mod ? Number(mod[1]) / 100000 : 1) + (off ? Number(off[1]) / 100000 : 0)));
  return hslToHex(h, s, l2);
}
const SCHEME_ALIAS = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
function hexToHsl(hex) {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}
function hslToHex(h, s, l) {
  const f = (n) => {
    const k = (n + h * 12) % 12;
    const a = s * Math.min(l, 1 - l);
    const v = l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1));
    return Math.round(v * 255).toString(16).padStart(2, '0');
  };
  return ('#' + f(0) + f(8) + f(4)).toUpperCase();
}

function readParagraphDecor(pPrXml) {
  const pPr = String(pPrXml || '').replace(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/g, '');
  const out = {};

  // <w:tabs><w:tab w:val="left" w:pos="720" w:leader="dot"/></w:tabs>
  const tabs = /<w:tabs\b[^>]*>([\s\S]*?)<\/w:tabs>/.exec(pPr);
  if (tabs) {
    const stops = [];
    for (const m of tabs[1].matchAll(/<w:tab\b([^>]*)\/>/g)) {
      const a = attrs(m[1]);
      if (a['w:val'] === 'clear' || a['w:pos'] === undefined) continue;
      stops.push({ align: a['w:val'] || 'left', posPx: twipsToPx(a['w:pos']), leader: a['w:leader'] && a['w:leader'] !== 'none' ? a['w:leader'] : null });
    }
    if (stops.length) out.tabs = stops.sort((x, y) => x.posPx - y.posPx);
  }

  // <w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/> — the fill is the colour.
  const shd = /<w:shd\b([^>]*)\/>/.exec(pPr);
  if (shd) {
    const a = attrs(shd[1]);
    const fill = a['w:fill'];
    if (fill && fill !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(fill)) out.shading = '#' + fill.toUpperCase();
  }

  // <w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>
  const pBdr = /<w:pBdr\b[^>]*>([\s\S]*?)<\/w:pBdr>/.exec(pPr);
  if (pBdr) {
    const borders = {};
    for (const m of pBdr[1].matchAll(/<w:(top|left|bottom|right|between|bar)\b([^>]*)\/>/g)) {
      const a = attrs(m[2]);
      const style = a['w:val'] || 'single';
      if (style === 'nil' || style === 'none') continue;
      const colour = a['w:color'] && a['w:color'] !== 'auto' ? '#' + a['w:color'].toUpperCase() : '#000000';
      borders[m[1]] = { style, widthPx: eighthsToPx(a['w:sz'] ?? 4), colour, spacePt: Number(a['w:space'] ?? 0) || 0 };
    }
    if (Object.keys(borders).length) out.borders = borders;
  }

  // The indents beyond `left`: a first-line indent, a hanging one, the right.
  const ind = /<w:ind\b([^>]*?)\/?>/.exec(pPr);
  if (ind) {
    const a = attrs(ind[1]);
    if (a['w:firstLine'] !== undefined) out.firstLinePx = twipsToPx(a['w:firstLine']);
    if (a['w:hanging'] !== undefined) out.hangingPx = twipsToPx(a['w:hanging']);
    const right = a['w:right'] ?? a['w:end'];
    if (right !== undefined) out.rightPx = twipsToPx(right);
  }

  return Object.keys(out).length ? out : null;
}

// The two numbering definitions the editor's list button can create. Each is a
// single-level `<w:abstractNum>` — level 0 is all the write side ever sets — with
// the hanging indent Word gives a fresh list. Kept deliberately minimal, in the
// spirit of the rest of this file: model what we write, nothing more.
const NUMBERING_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const NUMBERING_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const STYLES_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const STYLES_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const HYPERLINK_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const HEADER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const FOOTER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const HEADER_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const FOOTER_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
/** What makes a band un-editable as plain text — flattening these loses them. */
const BAND_STRUCTURE = /<w:(fldSimple|fldChar|sdt)\b/;
const COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const COMMENTS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const CHART_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const EMPTY_COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>';
/** The image formats the insert accepts — what every Word since 2007 renders. */
const IMAGE_EXTENSIONS = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif' };
/** CSS pixels to EMUs, the unit DrawingML measures in. 1px at 96dpi = 9525. */
const PX_TO_EMU = 9525;

const bulletAbstract = (id) =>
  '<w:abstractNum w:abstractNumId="' + id + '">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>' +
  '<w:lvlText w:val="&#61623;"/><w:lvlJc w:val="left"/>' +
  '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>' +
  '</w:lvl></w:abstractNum>';

const numberAbstract = (id) =>
  '<w:abstractNum w:abstractNumId="' + id + '">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
  '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>' +
  '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
  '</w:lvl></w:abstractNum>';

const numDef = (numId, abstractId) =>
  '<w:num w:numId="' + numId + '"><w:abstractNumId w:val="' + abstractId + '"/></w:num>';

export class Document {
  constructor(pkg) {
    this.pkg = pkg;
    this.mainPart = pkg.mainDocument();
    if (!this.mainPart.startsWith('word/')) {
      throw new Error('not a word processing package: ' + this.mainPart);
    }
    this.xml = pkg.text(this.mainPart);
    this.dirty = false;
    // What the package's main part currently HOLDS. `save()` moves it; undo
    // compares against it — see `restore()` for the bug this closes.
    this._flushed = this.xml;
    // SIDE parts (headers, footers, comments) an edit has touched or is
    // about to touch. Snapshots capture these so undo can put one back;
    // empty until the first such edit, so ordinary typing pays nothing.
    this._undoParts = new Set();
  }

  static open(buf) { return new Document(OoxmlPackage.read(buf)); }

  /** Body content, with the surrounding document element held verbatim. */
  _body() {
    const open = /<w:body(\s[^>]*)?>/.exec(this.xml);
    if (!open) throw new Error('document has no <w:body>');
    const close = this.xml.lastIndexOf('</w:body>');
    return {
      prefix: this.xml.slice(0, open.index + open[0].length),
      body: this.xml.slice(open.index + open[0].length, close),
      suffix: this.xml.slice(close),
    };
  }

  // ---- reading -------------------------------------------------------------

  /**
   * Top-level paragraphs. Paragraphs nested inside tables and content controls
   * are deliberately excluded — addressing them by a flat index would be a trap,
   * because inserting a table row would silently renumber every later index.
   */
  paragraphs() {
    // Cached against the XML it was read from.
    //
    // Without this, `paragraph(index)` rescans the entire body, and a caller
    // building a block list calls it once per paragraph — quadratic in the
    // length of the document. It cost 2.6 seconds a keystroke on a
    // fifty-eight-page report before anyone noticed, because every document
    // anyone had tested with was short.
    //
    // The key is the XML string itself rather than a dirty flag: every edit
    // replaces it, so a stale cache is not expressible.
    if (this._paragraphsFor === this.xml) return this._paragraphs;

    const { body } = this._body();
    const out = [];
    let depth = 0;
    const re = /<w:(tbl|sdt)\b[^>]*?(\/?)>|<\/w:(tbl|sdt)>|<w:p\b[^>]*?(\/?)>|<\/w:p>/g;
    let m;
    let pStart = -1;
    while ((m = re.exec(body))) {
      const tag = m[0];
      if (/^<w:(tbl|sdt)\b/.test(tag) && !tag.endsWith('/>')) { depth += 1; continue; }
      if (/^<\/w:(tbl|sdt)>/.test(tag)) { depth -= 1; continue; }
      if (depth > 0) continue;
      if (/^<w:p\b/.test(tag)) {
        if (tag.endsWith('/>')) {
          out.push({ index: out.length, xml: tag, start: m.index, end: m.index + tag.length, text: '' });
        } else {
          pStart = m.index;
        }
        continue;
      }
      if (tag === '</w:p>' && pStart >= 0) {
        const xml = body.slice(pStart, m.index + tag.length);
        out.push({ index: out.length, xml, start: pStart, end: m.index + tag.length, text: textOf(xml) });
        pStart = -1;
      }
    }
    this._paragraphsFor = this.xml;
    this._paragraphs = out;
    return out;
  }

  text() { return this.paragraphs().map((p) => p.text).join('\n'); }

  /**
   * EVERY editable paragraph in document order — top-level AND inside table
   * cells — with a `container` naming the cell each nested one lives in.
   *
   * This is the EDIT address space, the one the editor's caret walks since
   * tables became editable. `paragraphs()` above keeps its original contract
   * (top-level only) for the callers that address a document's prose —
   * `text()`, `setParagraphText`, the fidelity tests.
   *
   * `container` is `null` for a body paragraph, or a path key such as
   * `t1024:r2:c0` — the table's start offset in the body, then row and cell
   * ordinals, extended again for a nested table (`t1024:r1:c1:t2048:r0:c0`).
   * The key is stable within one version of the XML, which is all the editor
   * compares: every edit replaces the XML and the list is recomputed. Two
   * blocks share a container exactly when they live in the same cell, and that
   * equality is what the view's boundary guards are built on.
   *
   * A `w:sdt` subtree is opaque here, exactly as in `paragraphs()`: a content
   * control's paragraphs are structure the editor must not rebuild, wherever
   * the control sits. `cellBlocks` in table.js flags the same paragraphs
   * `inSdt` so the two scans agree about what is addressable.
   */
  editParagraphs() {
    if (this._editParagraphsFor === this.xml) return this._editParagraphs;

    const { body } = this._body();
    const out = [];
    const re = /<w:(tbl|tr|tc|sdt|txbxContent)\b[^>]*?(\/?)>|<\/w:(tbl|tr|tc|sdt|txbxContent)>|<w:p\b[^>]*?(\/?)>|<\/w:p>/g;
    const stack = [];
    let sdtDepth = 0;
    // A table inside a body-level content control: its paragraphs stay opaque
    // (cellBlocks flags them `inSdt` and the zip in _assignBlockIndices skips
    // them), so they must not be listed here either.
    let sdtTableDepth = 0;
    // A text box's paragraphs belong to the paragraph that anchors the box —
    // `_paragraphTextBoxes` reads them from there. Listing them here made a
    // cover page's text box lines into body paragraphs and LOST the paragraph
    // that carried the box.
    let boxDepth = 0;
    let pStart = -1;
    let key = null; // the container key for the innermost open cell, or null

    const rebuildKey = () => {
      key = null;
      const parts = [];
      for (const f of stack) {
        if (f.tag === 'tbl') parts.push('t' + f.id);
        else if (f.tag === 'tr') parts.push('r' + f.index);
        else if (f.tag === 'tc') parts.push('c' + f.index);
      }
      if (parts.length) key = parts.join(':');
    };

    let m;
    while ((m = re.exec(body))) {
      const tag = m[0];
      if (m[1] || m[3]) { // open/self-close or close of tbl|tr|tc|sdt|txbxContent
        const name = m[1] ?? m[3];
        if (name === 'txbxContent') {
          if (m[1] && m[2] !== '/') boxDepth += 1;
          else if (m[3]) boxDepth = Math.max(0, boxDepth - 1);
          continue;
        }
        if (boxDepth > 0) continue;
        if (name === 'sdt') {
          if (m[1] && m[2] !== '/') sdtDepth += 1;
          else if (m[3]) sdtDepth = Math.max(0, sdtDepth - 1);
          continue;
        }
        if (sdtDepth > 0) { // structure inside a content control is opaque
          if (name === 'tbl') {
            if (m[1] && m[2] !== '/') sdtTableDepth += 1;
            else if (m[3]) sdtTableDepth = Math.max(0, sdtTableDepth - 1);
          }
          continue;
        }
        if (m[1]) {
          if (m[2] === '/') continue; // an empty element — nothing to enter
          if (name === 'tbl') stack.push({ tag: 'tbl', id: m.index, nextRow: 0 });
          else if (name === 'tr') {
            const top = stack[stack.length - 1];
            stack.push({ tag: 'tr', index: top?.tag === 'tbl' ? top.nextRow++ : 0, nextCell: 0 });
          } else {
            const top = stack[stack.length - 1];
            // A vMerge continuation cell exists in the file but not on the
            // page (the restart above covers it), so its paragraphs are
            // flagged: a caret must not be walked into an invisible cell.
            const head = (() => {
              const firstP = body.indexOf('<w:p', m.index);
              return body.slice(m.index, firstP === -1 ? m.index + 400 : firstP);
            })();
            const hidden = /<w:vMerge\b(?![^>]*w:val="restart")/.test(head);
            stack.push({ tag: 'tc', index: top?.tag === 'tr' ? top.nextCell++ : 0, hidden });
          }
        } else {
          for (let i = stack.length - 1; i >= 0; i--) {
            if (stack[i].tag === name) { stack.length = i; break; }
          }
        }
        rebuildKey();
        continue;
      }

      // a <w:p …> or </w:p>
      if (boxDepth > 0) continue;
      // A BODY-LEVEL content control's paragraphs are listed, read-only: a
      // cover page, a table of contents, a bound field are all sdt, and a
      // page that skipped them opened a fifteen-page tender on its second
      // page. Inside a table cell the control stays opaque, as cellBlocks
      // expects; a table inside a body-level control stays opaque too.
      if (sdtDepth > 0 && (stack.length > 0 || sdtTableDepth > 0)) continue;
      const inSdt = sdtDepth > 0 || undefined;
      const top = stack[stack.length - 1];
      if (top && top.tag !== 'tc') continue; // a stray paragraph under tbl/tr — not a valid home
      const hiddenCell = stack.some((f) => f.tag === 'tc' && f.hidden) || undefined;
      if (/^<w:p\b/.test(tag)) {
        if (tag.endsWith('/>')) {
          out.push({ index: out.length, xml: tag, start: m.index, end: m.index + tag.length, text: '', container: key, ...(hiddenCell ? { hiddenCell } : {}), ...(inSdt ? { inSdt } : {}) });
        } else {
          pStart = m.index;
        }
        continue;
      }
      if (tag === '</w:p>' && pStart >= 0) {
        const xml = body.slice(pStart, m.index + tag.length);
        out.push({ index: out.length, xml, start: pStart, end: m.index + tag.length, text: textOf(xml), container: key, ...(hiddenCell ? { hiddenCell } : {}), ...(inSdt ? { inSdt } : {}) });
        pStart = -1;
      }
    }

    this._editParagraphsFor = this.xml;
    this._editParagraphs = out;
    return out;
  }

  /**
   * The body in DOCUMENT ORDER — paragraphs and tables interleaved as the author
   * wrote them.
   *
   * Separate from `paragraphs()` on purpose. That method's flat index is the
   * editor's address space and must keep counting top-level paragraphs only;
   * renumbering it to make room for tables would move every caret in the
   * document. So this is a second, read-only view for the renderer: a paragraph
   * block carries its paragraph index, a table block carries the parsed table.
   *
   * @returns {Array<{kind:'paragraph', paragraphIndex:number}|{kind:'table', table:object}>}
   */
  blocks() {
    const { body } = this._body();
    // Paragraph entries carry EDIT-space indices — the renderer resolves them
    // against the editor's block list, which is that space. The body subset of
    // editParagraphs() has the same paragraphs at the same offsets as
    // paragraphs(); only the numbering differs once a table has contributed
    // its cells.
    const byStart = new Map(
      this.editParagraphs().filter((p) => p.container === null).map((p) => [p.start, p.index]),
    );
    const out = [];

    // One scan, tracking tbl/sdt/txbxContent depth exactly as editParagraphs()
    // does so the two views cannot disagree about which paragraphs are top
    // level: a body-level content control's paragraphs are (read-only), a
    // table inside one is not, and a text box's are never.
    const re = /<w:(tbl|sdt|txbxContent)\b[^>]*?(\/?)>|<\/w:(tbl|sdt|txbxContent)>|<w:p\b[^>]*?(\/?)>|<\/w:p>/g;
    let depth = 0;
    let sdtDepth = 0;
    let boxDepth = 0;
    let tblStart = -1;
    let m;
    while ((m = re.exec(body))) {
      const tag = m[0];
      if (/^<w:txbxContent\b/.test(tag) && !tag.endsWith('/>')) { boxDepth += 1; continue; }
      if (tag === '</w:txbxContent>') { boxDepth = Math.max(0, boxDepth - 1); continue; }
      if (boxDepth > 0) continue;
      if (/^<w:tbl\b/.test(tag) && !tag.endsWith('/>')) {
        if (depth === 0 && sdtDepth === 0) tblStart = m.index;
        depth += 1;
        continue;
      }
      if (/^<w:sdt\b/.test(tag) && !tag.endsWith('/>')) { sdtDepth += 1; continue; }
      if (tag === '</w:sdt>') { sdtDepth = Math.max(0, sdtDepth - 1); continue; }
      if (tag === '</w:tbl>') {
        depth -= 1;
        if (depth === 0 && sdtDepth === 0 && tblStart >= 0) {
          const table = parseTable(body.slice(tblStart, m.index + tag.length));
          this._assignBlockIndices(table, tblStart, m.index + tag.length);
          out.push({ kind: 'table', table });
          tblStart = -1;
        }
        continue;
      }
      if (depth > 0) continue;
      if (/^<w:p\b/.test(tag) && byStart.has(m.index)) {
        out.push({ kind: 'paragraph', paragraphIndex: byStart.get(m.index) });
      }
    }
    return out;
  }

  /**
   * Give each editable cell paragraph in a parsed table its EDIT-SPACE index,
   * so the renderer can address it — a `data-block` on a cell paragraph is what
   * lets a caret live there.
   *
   * The pairing is a zip in document order: `editParagraphs()` lists this
   * table's paragraphs in XML order, and walking the parse tree — each cell's
   * blocks in sequence, recursing into a nested table at its position — visits
   * the same paragraphs in the same order. Both scans treat a `w:sdt` subtree
   * as opaque (`inSdt` on the parse side), so the two sequences agree by
   * construction. If they ever disagree — a shape neither scan foresaw — the
   * table is left unindexed and renders read-only, which degrades honestly
   * instead of putting a caret in the wrong cell.
   */
  _assignBlockIndices(table, from, to) {
    const inside = this.editParagraphs().filter(
      (p) => p.container !== null && p.start >= from && p.end <= to,
    );
    const seq = [];
    const walkBlocks = (blocks) => {
      for (const b of blocks) {
        if (b.inSdt) continue;
        if (b.kind === 'table') { walkTable(b.table); continue; }
        seq.push(b);
      }
    };
    const walkTable = (t) => {
      for (const row of t.rows) for (const cell of row.cells) walkBlocks(cell.blocks);
    };
    walkTable(table);
    if (seq.length !== inside.length) return;
    seq.forEach((b, i) => { b.blockIndex = inside[i].index; });
  }

  /** Page size and margins from `w:sectPr`, in CSS pixels. */
  section() { return parseSection(this._body().body); }

  /**
   * Pictures embedded in one paragraph, as data URIs.
   *
   * `w:drawing` wraps either `wp:inline` (flows with the text) or `wp:anchor`
   * (floats at a position). Both are rendered as a BLOCK under the paragraph's
   * text — an honest simplification, stated rather than hidden: true float
   * layout means wrapping text around a box, and nearly every image in a real
   * business letter is a logo or a signature scan sitting in a paragraph of its
   * own, where the difference is invisible. The XML is preserved verbatim
   * either way, so a file we cannot lay out perfectly still round-trips.
   */
  _paragraphImages(paragraphXml) {
    if (!paragraphXml.includes('<w:drawing')) return [];
    const out = [];
    const rels = new Map(this.pkg.rels(this.mainPart).map((r) => [r.Id, r.Target]));
    for (const m of String(paragraphXml).matchAll(/<w:drawing\b[^>]*>([\s\S]*?)<\/w:drawing>/g)) {
      const inner = m[1];
      const extent = /<wp:extent\b([^>]*)\/>/.exec(inner);
      const ext = extent ? attrs(extent[1]) : {};
      const blip = /<a:blip\b([^>]*)\/?>/.exec(inner);
      // No blip means this drawing is not a PICTURE — a chart or a shape,
      // which `_paragraphRichDrawings` owns. Reporting it here as an image
      // with a null href made every consumer carry a ghost to filter out.
      if (!blip) continue;
      const relId = attrs(blip[1])['r:embed'] ?? attrs(blip[1]).embed ?? null;
      const namePr = /<wp:docPr\b([^>]*)\/?>/.exec(inner);
      const target = relId ? rels.get(relId) : null;
      const part = target ? OoxmlPackage.resolveTarget(this.mainPart, target) : null;
      const bytes = part && this.pkg.has(part) ? this.pkg.read(part) : null;
      out.push({
        name: namePr ? attrs(namePr[1]).name ?? null : null,
        part,
        // EMU -> px; a missing extent falls back to something visible.
        widthPx: ext.cx ? Number(ext.cx) / 9525 : 96,
        heightPx: ext.cy ? Number(ext.cy) / 9525 : 96,
        href: bytes ? toDataUri(bytes, part) : null,
        anchored: inner.includes('<wp:anchor') || m[0].includes('<wp:anchor'),
      });
    }
    return out;
  }

  /** Paragraph styles resolved from word/styles.xml, flattened, in CSS px. */
  paragraphStyles() {
    const part = 'word/styles.xml';
    return readParagraphStyles(this.pkg.has(part) ? this.pkg.text(part) : null, this.themeFonts());
  }

  /** Character styles resolved from word/styles.xml — what a `w:rStyle` gives a run. */
  characterStyles() {
    const part = 'word/styles.xml';
    return readCharacterStyles(this.pkg.has(part) ? this.pkg.text(part) : null, this.themeFonts());
  }

  /**
   * The theme's heading and body faces. Word names nearly every font through
   * these two slots, so a reader that ignores the theme shows every document
   * in Calibri. Cached: no editing operation writes the theme part.
   */
  themeFonts() {
    if (this._themeFonts) return this._themeFonts;
    const part = this.pkg.partNames().find((n) => /^word\/theme\/theme\d*\.xml$/.test(n));
    this._themeFonts = readThemeFonts(part ? this.pkg.text(part) : null);
    return this._themeFonts;
  }

  /** Numbering definitions from word/numbering.xml: numId -> levels. */
  numberingDefs() {
    const part = 'word/numbering.xml';
    return readNumberingDefs(this.pkg.has(part) ? this.pkg.text(part) : null);
  }

  /**
   * Ensure a bullet and a decimal-number list definition exist, returning the
   * two numIds a paragraph's `<w:numPr>` points at.
   *
   * This is the WRITE half of numbering — `numberingDefs()` above is the read
   * half. The list button in the editor toggles a paragraph into a list, and a
   * list needs a definition to point at: this makes one. It is careful to be a
   * no-op-shaped operation whenever it can — a document that already carries a
   * suitable bullet or decimal definition reuses it rather than growing a
   * duplicate, so toggling the same list on and off never bloats numbering.xml.
   *
   * Only called when a list is actually applied, so a document that never gets a
   * list keeps no numbering.xml and round-trips byte-identical. When the part is
   * absent it is created — with its content-type override and a numbering
   * relationship on the main document — exactly as Word would have written them.
   *
   * @returns {{bullet: string, number: string}} the numIds to write into numPr
   */
  ensureListNumbering() {
    const part = 'word/numbering.xml';
    const has = this.pkg.has(part);
    const xml = has ? this.pkg.text(part) : null;
    const found = this._classifyNumbering(xml);

    let bulletId = found.bullet;
    let numberId = found.number;
    if (bulletId && numberId) return { bullet: bulletId, number: numberId };

    let nextAbstract = found.maxAbstract + 1;
    let nextNum = found.maxNum + 1;
    const abstracts = [];
    const nums = [];

    if (!bulletId) {
      const aId = nextAbstract++;
      const nId = nextNum++;
      abstracts.push(bulletAbstract(aId));
      nums.push(numDef(nId, aId));
      bulletId = String(nId);
    }
    if (!numberId) {
      const aId = nextAbstract++;
      const nId = nextNum++;
      abstracts.push(numberAbstract(aId));
      nums.push(numDef(nId, aId));
      numberId = String(nId);
    }

    if (!has) {
      const body =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:numbering xmlns:w="' + WORD_NS + '">' + abstracts.join('') + nums.join('') + '</w:numbering>';
      this.pkg.addPart(part, body, NUMBERING_CT);
      this._ensureNumberingRel();
    } else {
      this.pkg.write_(part, this._spliceNumbering(xml, abstracts.join(''), nums.join('')));
    }
    return { bullet: bulletId, number: numberId };
  }

  /**
   * Which numIds already carry a bullet or a decimal level-0 definition, plus the
   * highest ids in use — everything `ensureListNumbering` needs to decide whether
   * to reuse or to mint fresh ids that cannot collide.
   */
  _classifyNumbering(xml) {
    const out = { bullet: null, number: null, maxNum: 0, maxAbstract: -1 };
    if (!xml) return out;

    // A `<w:num>` points at a `<w:abstractNum>`; the abstract's level-0 numFmt
    // (its first, since level 0 is written first) is what says bullet vs number.
    const absFmt = new Map();
    for (const m of xml.matchAll(/<w:abstractNum\b([^>]*)>([\s\S]*?)<\/w:abstractNum>/g)) {
      const id = attrs(m[1])['w:abstractNumId'];
      if (id !== undefined) out.maxAbstract = Math.max(out.maxAbstract, Number(id));
      const fmt = /<w:numFmt\b[^>]*\bw:val="([^"]*)"/.exec(m[2]);
      absFmt.set(String(id), fmt ? fmt[1] : 'decimal');
    }
    for (const m of xml.matchAll(/<w:num\b([^>]*)>([\s\S]*?)<\/w:num>/g)) {
      const numId = attrs(m[1])['w:numId'];
      if (numId === undefined) continue;
      out.maxNum = Math.max(out.maxNum, Number(numId));
      const absId = /<w:abstractNumId\b[^>]*\bw:val="([^"]*)"/.exec(m[2]);
      const fmt = absId ? absFmt.get(absId[1]) : undefined;
      if (fmt === 'bullet') { if (!out.bullet) out.bullet = String(numId); }
      else if (fmt === 'decimal') { if (!out.number) out.number = String(numId); }
    }
    return out;
  }

  /** Splice new abstractNums (before the first `<w:num>`) and nums (at the end). */
  _spliceNumbering(xml, abstractsXml, numsXml) {
    let out = xml;
    if (abstractsXml) {
      const firstNum = /<w:num\b/.exec(out);
      out = firstNum
        ? out.slice(0, firstNum.index) + abstractsXml + out.slice(firstNum.index)
        : out.replace('</w:numbering>', abstractsXml + '</w:numbering>');
    }
    if (numsXml) out = out.replace('</w:numbering>', numsXml + '</w:numbering>');
    return out;
  }

  /** Point the main document at word/numbering.xml, creating its .rels if needed. */
  _ensureNumberingRel() {
    const relsPath = OoxmlPackage.relsPathFor(this.mainPart);
    const xml = this.pkg.has(relsPath) ? this.pkg.text(relsPath) : null;
    if (xml && (/Target="numbering\.xml"/.test(xml) || xml.includes('Type="' + NUMBERING_REL_TYPE + '"'))) return;
    this._addRel(NUMBERING_REL_TYPE, 'numbering.xml');
  }

  /**
   * Add one relationship from the main document, creating its .rels part if
   * the file never had one, and return the fresh rId. The numbering, styles
   * and image writers all reach parts through this — a target with no
   * relationship is a part Word will never look at.
   */
  _addRel(type, target, { external = false } = {}) {
    const relsPath = OoxmlPackage.relsPathFor(this.mainPart);
    const mode = external ? ' TargetMode="External"' : '';
    if (!this.pkg.has(relsPath)) {
      this.pkg.addPart(relsPath,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="' + type + '" Target="' + esc(target) + '"' + mode + '/>' +
        '</Relationships>');
      return 'rId1';
    }
    const xml = this.pkg.text(relsPath);
    const ids = [...xml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
    const next = 'rId' + ((ids.length ? Math.max(...ids) : 0) + 1);
    const rel = '<Relationship Id="' + next + '" Type="' + type + '" Target="' + esc(target) + '"' + mode + '/>';
    this.pkg.write_(relsPath, xml.replace('</Relationships>', rel + '</Relationships>'));
    return next;
  }

  /**
   * Mint a link token for a URL — the hyperlink relationship (External, the
   * way a docx points outside itself) plus the attribute string a
   * `<w:hyperlink>` wrapper carries. The token is what a run's `link` field
   * holds; `renderRuns` writes it back verbatim.
   */
  makeLink(url) {
    const target = String(url);
    if (!/^(https?:\/\/|mailto:)/i.test(target)) {
      throw new Error('a link must start http://, https:// or mailto: — got ' + target);
    }
    const rId = this._addRel(HYPERLINK_REL_TYPE, target, { external: true });
    return ' r:id="' + rId + '" w:history="1"';
  }

  /**
   * What a link token points at, for display: the relationship's URL, or
   * `#anchor` for an internal `w:anchor` link. Null when the token resolves
   * to nothing — a dangling id in a damaged file.
   */
  linkTarget(token) {
    const anchor = /\bw:anchor="([^"]*)"/.exec(String(token));
    if (anchor) return '#' + unesc(anchor[1]);
    const id = /\br:id="([^"]*)"/.exec(String(token));
    if (!id) return null;
    const rel = this.pkg.rels(this.mainPart).find((r) => r.Id === id[1]);
    return rel ? rel.Target : null;
  }

  /**
   * Give a document with NO styles part the standard catalogue, so picking
   * "Heading 1" has somewhere to point. A file that already has a styles part
   * is left completely alone — its catalogue is its own, and rewriting an
   * existing styles.xml restyles every paragraph that references it (the
   * append-never-rewrite rule). Returns whether anything was written.
   */
  ensureParagraphStyles() {
    const part = 'word/styles.xml';
    if (this.pkg.has(part)) return false;
    this.pkg.addPart(part, STANDARD_STYLES_XML, STYLES_CT);
    this._addRel(STYLES_REL_TYPE, 'styles.xml');
    return true;
  }

  /**
   * Page setup — orientation, paper size, margin presets — by editing the
   * body's `w:sectPr` in place. Everything not named rides through: header
   * and footer distances, the gutter, section type, header references.
   *
   *   orientation  'portrait' | 'landscape'
   *   size         'A4' | 'Letter' | 'Legal'
   *   margins      'normal' | 'narrow' | 'wide', or {top,right,bottom,left} twips
   *
   * Width and height are stored already-swapped for landscape (with
   * `w:orient` as the label), which is how Word writes it — the parser
   * reports the same way, so the two agree by construction.
   */
  setPageSetup({ orientation, size, margins } = {}) {
    const PAPER = { A4: [11906, 16838], Letter: [12240, 15840], Legal: [12240, 20160] };
    const MARGIN_PRESETS = {
      normal: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      narrow: { top: 720, right: 720, bottom: 720, left: 720 },
      wide: { top: 1440, right: 2880, bottom: 1440, left: 2880 },
    };

    const { prefix, body, suffix } = this._body();
    const at = /<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(body);
    if (!at) throw new Error('this document has no section properties to set up');
    let sectPr = at[0];
    if (/^<w:sectPr\b[^>]*\/>$/.test(sectPr)) {
      sectPr = sectPr.replace(/\/>$/, '>') + '</w:sectPr>';
    }

    const setChildAttrs = (tag, changes, { before = null } = {}) => {
      const one = new RegExp('<w:' + tag + '\\b([^>]*?)\\/?>');
      const m = one.exec(sectPr);
      let attrsXml = m ? m[1] : '';
      for (const [name, value] of Object.entries(changes)) {
        const attrRe = new RegExp('\\s*\\b' + name + '="[^"]*"');
        attrsXml = attrsXml.replace(attrRe, '');
        if (value !== null) attrsXml += ' ' + name + '="' + value + '"';
      }
      const element = '<w:' + tag + attrsXml + '/>';
      if (m) { sectPr = sectPr.replace(one, element); return; }
      const anchor = before ? new RegExp('<w:' + before + '\\b').exec(sectPr) : null;
      sectPr = anchor
        ? sectPr.slice(0, anchor.index) + element + sectPr.slice(anchor.index)
        : sectPr.replace('</w:sectPr>', element + '</w:sectPr>');
    };

    if (orientation !== undefined || size !== undefined) {
      // Current dimensions, normalised to portrait, so size and orientation
      // compose in either order and idempotently.
      const current = parseSection(body);
      const currentTwips = [Math.round(current.widthPx * 15), Math.round(current.heightPx * 15)];
      const portrait = size !== undefined
        ? (PAPER[size] ?? (() => { throw new Error('unknown paper size: ' + size); })())
        : [Math.min(...currentTwips), Math.max(...currentTwips)];
      const landscape = (orientation ?? current.orientation) === 'landscape';
      setChildAttrs('pgSz', {
        'w:w': String(landscape ? portrait[1] : portrait[0]),
        'w:h': String(landscape ? portrait[0] : portrait[1]),
        'w:orient': landscape ? 'landscape' : null,
      }, { before: 'pgMar' });
    }

    if (margins !== undefined) {
      const preset = typeof margins === 'string' ? MARGIN_PRESETS[margins] : margins;
      if (!preset) throw new Error('unknown margin preset: ' + margins);
      const changes = {};
      for (const side of ['top', 'right', 'bottom', 'left']) {
        if (preset[side] !== undefined) changes['w:' + side] = String(Math.max(0, Math.round(preset[side])));
      }
      setChildAttrs('pgMar', changes);
    }

    this.xml = prefix + body.slice(0, at.index) + sectPr + body.slice(at.index + at[0].length) + suffix;
    this.dirty = true;
    return this;
  }

  /**
   * Charts and preset shapes in one paragraph — raw. A picture rides
   * `_paragraphImages`; everything else a `w:drawing` can hold that we know
   * how to SHOW comes out here as xml plus its box, for the doc-view backend
   * to turn into paint (this layer stays format-only, so the drawing
   * package's arrow keeps pointing the right way).
   */
  _paragraphRichDrawings(paragraphXml) {
    if (!paragraphXml.includes('<w:drawing')) return [];
    const out = [];
    const rels = new Map(this.pkg.rels(this.mainPart).map((r) => [r.Id, r.Target]));
    for (const m of String(paragraphXml).matchAll(/<w:drawing\b[^>]*>([\s\S]*?)<\/w:drawing>/g)) {
      const inner = m[1];
      if (/<a:blip\b/.test(inner)) continue; // a picture — _paragraphImages has it
      const extent = /<wp:extent\b([^>]*)\/>/.exec(inner);
      const ext = extent ? attrs(extent[1]) : {};
      const widthPx = ext.cx ? Math.round(Number(ext.cx) / 9525) : null;
      const heightPx = ext.cy ? Math.round(Number(ext.cy) / 9525) : null;
      const namePr = /<wp:docPr\b([^>]*)\/?>/.exec(inner);
      const name = namePr ? (attrs(namePr[1])['name'] ?? null) : null;

      const chartRef = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(inner);
      if (chartRef) {
        const target = rels.get(chartRef[1]);
        const part = target ? OoxmlPackage.resolveTarget(this.mainPart, target) : null;
        if (part && this.pkg.has(part)) {
          out.push({ kind: 'chart', name, widthPx, heightPx, chartXml: this.pkg.text(part) });
        }
        continue;
      }
      const wsp = /<wps:wsp\b[\s\S]*?<\/wps:wsp>/.exec(inner);
      // A shape with a text box is a TEXT box — `_paragraphTextBoxes` draws
      // it with its words; painting the frame here too would double it.
      if (wsp && !/<wps:txbx\b/.test(wsp[0])) out.push({ kind: 'shape', name, widthPx, heightPx, shapeXml: wsp[0] });
    }
    return out;
  }

  /**
   * The text boxes anchored in a paragraph: each one's frame — size, fill,
   * outline, horizontal alignment — and its paragraphs, decorated like the
   * body's (style, alignment, indents, shading, runs, pictures) but not
   * addressable: a caret has no business in a box the engine cannot rebuild.
   *
   * Only the `mc:Choice` of an AlternateContent is read; the VML fallback
   * repeats the same words for older Words and would show them twice.
   */
  _paragraphTextBoxes(paragraphXml) {
    if (!paragraphXml.includes('<w:txbxContent')) return [];
    const xml = String(paragraphXml).replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '');
    const colours = this.themeColours();
    const out = [];
    for (const m of xml.matchAll(/<w:drawing\b[^>]*>([\s\S]*?)<\/w:drawing>|<w:pict\b[^>]*>([\s\S]*?)<\/w:pict>/g)) {
      const inner = m[1] ?? m[2] ?? '';
      const content = /<w:txbxContent\b[^>]*>([\s\S]*?)<\/w:txbxContent>/.exec(inner);
      if (!content) continue;
      const extent = /<wp:extent\b([^>]*)\/>/.exec(inner);
      const ext = extent ? attrs(extent[1]) : {};
      let widthPx = ext.cx ? Math.round(Number(ext.cx) / 9525) : null;
      let heightPx = ext.cy ? Math.round(Number(ext.cy) / 9525) : null;
      if (widthPx === null) {
        // VML: <v:shape style="width:451.3pt;height:38.4pt">
        const style = /<v:(?:shape|rect)\b[^>]*\bstyle="([^"]*)"/.exec(inner);
        const dim = (name) => { const d = new RegExp('(?:^|;)\\s*' + name + ':\\s*([\\d.]+)(pt|px|in|cm)?').exec(style?.[1] ?? ''); return d ? Math.round(Number(d[1]) * ({ pt: 96 / 72, px: 1, in: 96, cm: 96 / 2.54 })[d[2] || 'pt']) : null; };
        widthPx = dim('width');
        heightPx = dim('height');
      }
      const namePr = /<wp:docPr\b([^>]*)\/?>/.exec(inner);
      const hAlign = /<wp:positionH\b[^>]*>[\s\S]*?<wp:align>([^<]*)<\/wp:align>/.exec(inner)?.[1] ?? null;
      const spPr = /<wps:spPr\b[^>]*>([\s\S]*?)<\/wps:spPr>/.exec(inner)?.[1] ?? '';
      const line = /<a:ln\b[^>]*>([\s\S]*?)<\/a:ln>/.exec(spPr);
      const fill = colourOf(spPr.replace(/<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/, ''), colours);
      const vmlFill = /<v:(?:shape|rect)\b[^>]*\bfillcolor="([^"]*)"/.exec(inner);
      const paragraphs = this._liteParagraphs(content[1]);
      out.push({
        name: namePr ? (attrs(namePr[1])['name'] ?? null) : null,
        widthPx, heightPx, hAlign,
        fill: fill ?? (vmlFill ? vmlFill[1] : null),
        line: line ? colourOf(line[1], colours) : null,
        paragraphs,
      });
    }
    return out;
  }

  /**
   * The paragraphs of a fragment that is DISPLAYED, never edited — a text
   * box, a footnote — decorated the way body blocks are (style, alignment,
   * indent, spacing, decor, runs, pictures) but with no index and no
   * address: the caret cannot go there, so nothing here needs to rebuild.
   */
  _liteParagraphs(fragment) {
    const paragraphs = [];
    for (const p of String(fragment).matchAll(/<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)) {
      const px = p[0];
      const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>|<w:pPr\b[^>]*\/>/.exec(px);
      const style = /<w:pStyle\b[^>]*w:val="([^"]*)"/.exec(pPr ? pPr[0] : '');
      const jc = /<w:jc\b[^>]*w:val="([^"]*)"/.exec(pPr ? pPr[0] : '');
      const ind = /<w:ind\b([^>]*?)\/?>/.exec(pPr ? pPr[0] : '');
      const left = ind ? (/\bw:left="(-?\d+)"/.exec(ind[1]) || /\bw:start="(-?\d+)"/.exec(ind[1])) : null;
      paragraphs.push({
        text: textOf(px),
        style: style ? style[1] : null,
        align: jc ? jc[1] : null,
        indentPx: left ? twipsToPx(Number(left[1])) : null,
        spacing: readDirectSpacing(pPr ? pPr[0] : ''),
        decor: readParagraphDecor(pPr ? pPr[0] : ''),
        runs: parseRuns(px),
        images: this._paragraphImages(px),
      });
    }
    return paragraphs;
  }

  /**
   * The document's footnotes and endnotes, by id, each as displayed
   * paragraphs. Word's separator "notes" (type separator/continuationSeparator)
   * are the rule between body and notes, not notes, and are left out. The
   * NUMBERS are not here: a note is numbered by where its reference falls in
   * the body, which is the view's business. Cached — no edit writes the part.
   */
  notes() {
    // Cached against the parts' own text, so an inserted note, an edited
    // one, or an undo that puts a part back is seen the next time.
    const key = ['word/footnotes.xml', 'word/endnotes.xml'].map((p) => (this.pkg.has(p) ? this.pkg.text(p) : '')).join(' ');
    if (this._notes && this._notesKey === key) return this._notes;
    this._notesKey = key;
    const read = (part, tag) => {
      if (!this.pkg.has(part)) return [];
      const xml = this.pkg.text(part);
      const out = [];
      for (const m of xml.matchAll(new RegExp('<w:' + tag + '\\b([^>]*)>([\\s\\S]*?)</w:' + tag + '>', 'g'))) {
        const a = attrs(m[1]);
        if (a['w:type'] && a['w:type'] !== 'normal') continue;
        if (a['w:id'] === undefined) continue;
        out.push({ id: String(a['w:id']), paragraphs: this._liteParagraphs(m[2]) });
      }
      return out;
    };
    this._notes = { footnotes: read('word/footnotes.xml', 'footnote'), endnotes: read('word/endnotes.xml', 'endnote') };
    return this._notes;
  }

  /**
   * Make the notes part real (with Word's two separator entries) and put it
   * on the undo list — BEFORE the edit's snapshot, so undoing the very first
   * footnote restores an empty part rather than leaving an orphaned note.
   */
  registerNoteUndo(kind) {
    const spec = NOTE_PARTS[kind];
    if (!spec) throw new Error('a note is a footnote or an endnote');
    if (!this.pkg.has(spec.part)) {
      this.pkg.addPart(spec.part, emptyNotesXml(kind), spec.ct);
      this._addRel(spec.rel, spec.target);
    }
    this._undoParts.add(spec.part);
    return this;
  }

  /**
   * Add a footnote or endnote — its words, in the part — and return its id.
   * The REFERENCE in the body is the view's to place: it is one run in one
   * paragraph, and the view knows where the caret is. The note's paragraph
   * wears the note style, and its mark is superscript outright, so it reads
   * right even in a file whose stylesheet never heard of footnotes.
   */
  addNote(kind, text) {
    const body = String(text ?? '').trim();
    if (!body) throw new Error('a ' + kind + ' needs some words');
    this.registerNoteUndo(kind);
    const spec = NOTE_PARTS[kind];
    const xml = this.pkg.text(spec.part);
    let id = 0;
    for (const m of xml.matchAll(new RegExp('<w:' + kind + '\\b[^>]*\\bw:id="(-?\\d+)"', 'g'))) id = Math.max(id, Number(m[1]));
    id += 1;
    const entry = '<w:' + kind + ' w:id="' + id + '"><w:p><w:pPr><w:pStyle w:val="' + spec.textStyle + '"/></w:pPr>'
      + '<w:r><w:rPr><w:rStyle w:val="' + spec.refStyle + '"/><w:vertAlign w:val="superscript"/></w:rPr><w:' + kind + 'Ref/></w:r>'
      + renderRun(null, ' ' + body) + '</w:p></w:' + kind + '>';
    this.pkg.write_(spec.part, xml.replace('</w:' + kind + 's>', entry + '</w:' + kind + 's>'));
    this.dirty = true;
    return String(id);
  }

  /**
   * Replace a note's words. The first paragraph keeps its properties and its
   * mark; everything after the mark is the new text, as one plain run.
   */
  setNoteText(kind, id, text) {
    const body = String(text ?? '').trim();
    if (!body) throw new Error('a ' + kind + ' needs some words');
    this.registerNoteUndo(kind);
    const spec = NOTE_PARTS[kind];
    const xml = this.pkg.text(spec.part);
    const m = new RegExp('(<w:' + kind + '\\b[^>]*\\bw:id="' + String(id).replace(/[^-\d]/g, '') + '"[^>]*>)([\\s\\S]*?)(</w:' + kind + '>)').exec(xml);
    if (!m) throw new Error('no ' + kind + ' with id ' + id);
    const firstP = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/.exec(m[2]);
    const pPr = (firstP && /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/.exec(firstP[1])?.[0]) || '<w:pPr><w:pStyle w:val="' + spec.textStyle + '"/></w:pPr>';
    const mark = (firstP && new RegExp('<w:r\\b[^>]*>(?:(?!</w:r>)[\\s\\S])*?<w:' + kind + 'Ref\\b[^>]*/>[\\s\\S]*?</w:r>').exec(firstP[1])?.[0])
      || '<w:r><w:rPr><w:rStyle w:val="' + spec.refStyle + '"/><w:vertAlign w:val="superscript"/></w:rPr><w:' + kind + 'Ref/></w:r>';
    const rebuilt = m[1] + '<w:p>' + pPr + mark + renderRun(null, ' ' + body) + '</w:p>' + m[3];
    this.pkg.write_(spec.part, xml.replace(m[0], rebuilt));
    this.dirty = true;
    return this;
  }

  /**
   * The theme's colour scheme — accent1 and friends — resolved to hex, for a
   * shape fill named by scheme slot. Cached like the fonts.
   */
  themeColours() {
    if (this._themeColours) return this._themeColours;
    const part = this.pkg.partNames().find((n) => /^word\/theme\/theme\d*\.xml$/.test(n));
    this._themeColours = readThemeColours(part ? this.pkg.text(part) : null);
    return this._themeColours;
  }

  /**
   * Insert a chart AFTER a table, drawn from that table's figures — the data
   * lives in the chart part itself (`literal` caches), because a document
   * has no workbook to reference. The part, its content type, its
   * relationship and the inline drawing all wire here.
   */
  insertChartAfterTable(tableStart, { kind = 'column', title = null, categories = [], series = [], widthPx = 480, heightPx = 288 } = {}) {
    if (!['column', 'bar', 'line', 'area', 'pie', 'doughnut'].includes(kind)) throw new Error('unknown chart kind: ' + kind);
    if (!series.length || !series.some((s) => (s.values ?? []).some((v) => Number.isFinite(v)))) {
      throw new Error('a chart needs at least one numeric series');
    }
    const parts = this._tableParts(tableStart);

    let n = 1;
    while (this.pkg.has('word/charts/rutba-chart' + n + '.xml')) n += 1;
    const partName = 'word/charts/rutba-chart' + n + '.xml';
    this.pkg.addPart(partName, chartPartXml({
      kind, literal: true, title: title || undefined,
      categories: { values: categories },
      series: series.map((s) => ({ name: s.name, values: s.values })),
    }), CHART_CT);
    const rId = this._addRel(CHART_REL_TYPE, 'charts/rutba-chart' + n + '.xml');

    const cx = Math.max(1, Math.round(widthPx)) * 9525;
    const cy = Math.max(1, Math.round(heightPx)) * 9525;
    const id = this._nextDrawingId();
    const drawing =
      '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:docPr id="' + id + '" name="' + esc(String(title ?? 'Chart')) + '"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="' + rId + '"/>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing>';
    this._spliceBody(parts.end, parts.end, '<w:p><w:r>' + drawing + '</w:r></w:p>');
    return this;
  }

  /**
   * Insert a preset shape as its own paragraph — the same honest placement
   * pictures take. Word draws the real `prstGeom`; our painter draws the
   * same preset through the drawing commons, so the two agree on what was
   * asked for even where they rasterise differently.
   */
  insertShapeParagraph(index, { preset, widthPx = 160, heightPx = 100, fill = '1F5F8B', outline = null } = {}) {
    const PRESETS = [
      'rect', 'roundRect', 'ellipse', 'triangle', 'diamond',
      'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
      'pentagon', 'hexagon', 'star5', 'plus', 'chevron', 'parallelogram', 'trapezoid', 'line',
    ];
    if (!PRESETS.includes(preset)) throw new Error('unknown shape preset: ' + preset);
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    const hex = (v) => {
      const h = String(v).replace(/^#/, '').toUpperCase();
      if (!/^[0-9A-F]{6}$/.test(h)) throw new Error('a shape colour must be #RRGGBB: ' + v);
      return h;
    };

    const cx = Math.max(1, Math.round(widthPx)) * 9525;
    const cy = Math.max(1, Math.round(heightPx)) * 9525;
    const id = this._nextDrawingId();
    const fillXml = fill == null
      ? '<a:noFill/>'
      : '<a:solidFill><a:srgbClr val="' + hex(fill) + '"/></a:solidFill>';
    const lineXml = outline == null
      ? (preset === 'line' ? '<a:ln w="19050"><a:solidFill><a:srgbClr val="' + hex(fill ?? '1F5F8B') + '"/></a:solidFill></a:ln>' : '')
      : '<a:ln w="9525"><a:solidFill><a:srgbClr val="' + hex(outline) + '"/></a:solidFill></a:ln>';
    const drawing =
      '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:docPr id="' + id + '" name="' + esc(preset) + '"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:cNvSpPr/>' +
      '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="' + preset + '"><a:avLst/></a:prstGeom>' +
      (preset === 'line' ? '<a:noFill/>' : fillXml) +
      lineXml +
      '</wps:spPr><wps:bodyPr/></wps:wsp>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing>';
    this._spliceBody(p.end, p.end, '<w:p><w:r>' + drawing + '</w:r></w:p>');
    return this;
  }

  /** The next free `wp:docPr` id — Word wants them unique within the document. */
  _nextDrawingId() {
    let max = 0;
    for (const m of this.xml.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)) {
      max = Math.max(max, Number(m[1]));
    }
    return max + 1;
  }

  /**
   * Insert a picture as its OWN paragraph after an edit-space paragraph —
   * the ribbon's Insert picture.
   *
   * Its own paragraph rather than an inline run at the caret, deliberately:
   * the painter has always drawn images as a block under the paragraph's
   * text, so an inline insert would LOOK block-placed anyway, and a run the
   * caret cannot land in would complicate every offset in the address space.
   * This way the picture behaves like the logo in a real letter — a
   * paragraph of its own that typing cannot damage (the rebuilders keep
   * drawing runs).
   *
   * The bytes become `word/media/rutbaN.ext` with a relationship and a
   * per-part content type; the run is a standard `wp:inline` picture with
   * its namespaces declared inline, so it lands correctly even in a minimal
   * document whose root declares only `xmlns:w`.
   */
  insertImageParagraph(index, { name = 'Picture', contentType, data, widthPx, heightPx }) {
    const ext = IMAGE_EXTENSIONS[contentType];
    if (!ext) {
      throw new Error('unsupported image type: ' + contentType + ' (png, jpeg or gif)');
    }
    const bytes = Buffer.isBuffer(data) ? data : Buffer.from(String(data), 'base64');
    if (!bytes.length) throw new Error('the image has no bytes');
    const w = Math.round(Number(widthPx));
    const h = Math.round(Number(heightPx));
    if (!(w > 0) || !(h > 0)) throw new Error('an image needs a positive width and height');

    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);

    let n = 1;
    while (this.pkg.has('word/media/rutba' + n + '.' + ext)) n += 1;
    this.pkg.addPart('word/media/rutba' + n + '.' + ext, bytes, contentType);
    const rId = this._addRel(IMAGE_REL_TYPE, 'media/rutba' + n + '.' + ext);

    const cx = w * PX_TO_EMU;
    const cy = h * PX_TO_EMU;
    const id = this._nextDrawingId();
    const label = esc(String(name));
    const drawing =
      '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:docPr id="' + id + '" name="' + label + '"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="' + id + '" name="' + label + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rId + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>' +
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
    this._spliceBody(p.end, p.end, '<w:p><w:r>' + drawing + '</w:r></w:p>');
    return this;
  }

  /**
   * Headers and footers, by reference type.
   *
   * Read from the section properties rather than by globbing `word/header*.xml`:
   * a part with no reference to it belongs to a section that was deleted, and
   * painting it would show the reader a header the document does not use.
   */
  headerFooters() {
    const section = this.section();
    return readHeadersAndFooters(this.pkg, this.mainPart, section.sectPrXml);
  }

  /**
   * The document's comments, read-only, each anchored to the edit-space
   * paragraph its range starts in (−1 when the anchor is nowhere the caret
   * can go — inside a content control, say). Comments are margin voices, not
   * body content: nothing here edits them, and paragraphs carrying their
   * anchors stay structural so an edit cannot orphan a range.
   */
  comments() {
    if (!this.pkg.has('word/comments.xml')) return [];
    const xml = this.pkg.text('word/comments.xml');
    const out = [];
    for (const m of xml.matchAll(/<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g)) {
      const a = attrs(m[1]);
      if (a['w:id'] === undefined) continue;
      out.push({
        id: String(a['w:id']),
        author: unesc(a['w:author'] ?? ''),
        date: a['w:date'] ?? null,
        text: textOf(m[2]),
      });
    }
    if (out.length) {
      const paras = this.editParagraphs();
      for (const c of out) {
        const anchor = new RegExp('<w:comment(?:RangeStart|Reference)\\b[^>]*\\bw:id="' + c.id + '"');
        c.blockIndex = paras.findIndex((p) => anchor.test(p.xml));
      }
    }
    return out;
  }

  /**
   * Make the comments part real (empty if need be) and put it on the undo
   * list — BEFORE the edit's snapshot is taken, so undoing the very first
   * comment restores an empty part rather than leaving an orphaned voice.
   */
  registerCommentUndo() {
    if (!this.pkg.has('word/comments.xml')) {
      this.pkg.addPart('word/comments.xml', EMPTY_COMMENTS_XML, COMMENTS_CT);
      this._addRel(COMMENTS_REL_TYPE, 'comments.xml');
    }
    this._undoParts.add('word/comments.xml');
    return this;
  }

  /**
   * Add a comment anchored to the END of one edit-space paragraph — a point
   * comment (`w:commentReference` with no range), which is what keeps the
   * paragraph EDITABLE: a range start would make it structural, and a
   * comment that freezes the text it discusses would hardly be worth adding.
   * The reference run has no `w:t`, so every rebuild carries it whole.
   *
   * Commenting is allowed on structural paragraphs too: nothing here
   * rebuilds them — the reference splices in before the closing tag — and a
   * locked field paragraph is exactly the kind a colleague wants to remark
   * on.
   */
  addComment(index, { author = '', text }) {
    const body = String(text ?? '').trim();
    if (!body) throw new Error('a comment needs some words');
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);

    this.registerCommentUndo();
    const part = 'word/comments.xml';
    const xml = this.pkg.text(part);
    let id = 0;
    for (const m of xml.matchAll(/<w:comment\b[^>]*\bw:id="(\d+)"/g)) id = Math.max(id, Number(m[1]));
    id += 1;

    const stamp = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
    const entry = '<w:comment w:id="' + id + '" w:author="' + esc(String(author)) + '" w:date="' + stamp + '">'
      + '<w:p>' + renderRun(null, body) + '</w:p></w:comment>';
    this.pkg.write_(part, xml.replace('</w:comments>', entry + '</w:comments>'));

    const reference = '<w:r><w:commentReference w:id="' + id + '"/></w:r>';
    if (/^<w:p\b[^>]*\/>$/.test(p.xml)) {
      this._spliceBody(p.start, p.end, p.open + reference + '</w:p>');
    } else {
      this._spliceBody(p.end - '</w:p>'.length, p.end - '</w:p>'.length, reference);
    }
    return this;
  }

  /**
   * The DEFAULT header and footer as an editing panel sees them: whether one
   * exists, its paragraph texts, and whether it carries structure (a PAGE
   * field, a content control) that text editing would flatten — in which case
   * it can be looked at but not edited this way.
   */
  bandInfo() {
    const bands = this.headerFooters();
    const info = (bucket) => {
      const band = bucket.default ?? null;
      if (!band) return { exists: false, lines: [], structural: false };
      return {
        exists: true,
        lines: band.paragraphs.map((p) => p.text),
        structural: BAND_STRUCTURE.test(this.pkg.text(band.part)),
      };
    };
    return { header: info(bands.headers), footer: info(bands.footers) };
  }

  /**
   * Pre-register a band for undo: the part's CURRENT xml must be in the
   * snapshot taken before the edit, and the snapshot is taken before
   * `setBandText` runs — so the caller announces the target first.
   */
  registerBandUndo(which) {
    const bands = this.headerFooters();
    const band = (which === 'header' ? bands.headers : bands.footers).default;
    if (band) this._undoParts.add(band.part);
    return this;
  }

  /**
   * Replace the default header's or footer's text, line per paragraph —
   * creating the band, its content type, its relationship and its
   * `sectPr` reference when the document never had one.
   *
   * An existing band keeps each paragraph's `w:pPr` and its first run's
   * formatting, index-aligned, so a centred bold letterhead stays centred
   * and bold when its wording changes; extra lines take the last
   * paragraph's shape. A band carrying a field or a control refuses: the
   * page-number field a text edit would flatten is the band's most common
   * intelligence.
   */
  setBandText(which, lines) {
    if (which !== 'header' && which !== 'footer') {
      throw new Error("a band is 'header' or 'footer', not " + which);
    }
    const texts = (Array.isArray(lines) ? lines : [String(lines ?? '')]).map((l) => String(l));
    const bands = this.headerFooters();
    const existing = (which === 'header' ? bands.headers : bands.footers).default ?? null;

    if (existing) {
      const xml = this.pkg.text(existing.part);
      if (BAND_STRUCTURE.test(xml)) {
        throw new Error('This ' + which + ' carries a field or a content control — editing it as text would flatten that.');
      }
      const olds = [...xml.matchAll(/<w:p\b[^>]*?(?:\/>|>[\s\S]*?<\/w:p>)/g)].map((m) => m[0]);
      const shape = (i) => {
        const src = olds[Math.min(i, olds.length - 1)] ?? null;
        if (!src) return { pPr: '', rPr: null };
        const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>|<w:pPr\b[^>]*\/>/.exec(src);
        return { pPr: pPr ? pPr[0] : '', rPr: firstRunProps(src) };
      };
      const body = texts.map((t, i) => {
        const s = shape(i);
        return '<w:p>' + s.pPr + (t === '' ? '' : renderRun(s.rPr, t)) + '</w:p>';
      }).join('') || '<w:p/>';
      const root = /^([\s\S]*?<w:(?:hdr|ftr)\b[^>]*>)[\s\S]*(<\/w:(?:hdr|ftr)>[\s\S]*)$/.exec(xml);
      if (!root) throw new Error('unrecognised ' + which + ' part: ' + existing.part);
      this.pkg.write_(existing.part, root[1] + body + root[2]);
      this._undoParts.add(existing.part);
      return this;
    }

    // No band yet: a part, a content type, a relationship, a reference.
    const tag = which === 'header' ? 'hdr' : 'ftr';
    let n = 1;
    while (this.pkg.has('word/' + which + n + '.xml')) n += 1;
    const partName = 'word/' + which + n + '.xml';
    const body = texts.map((t) => '<w:p>' + (t === '' ? '' : renderRun(null, t)) + '</w:p>').join('') || '<w:p/>';
    this.pkg.addPart(partName,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:' + tag + ' xmlns:w="' + WORD_NS + '">' + body + '</w:' + tag + '>',
      which === 'header' ? HEADER_CT : FOOTER_CT);
    const rId = this._addRel(which === 'header' ? HEADER_REL_TYPE : FOOTER_REL_TYPE, which + n + '.xml');

    // The reference leads the sectPr's children by schema order; the r:
    // namespace is declared inline so a minimal document accepts it.
    const ref = '<w:' + which + 'Reference w:type="default" r:id="' + rId + '"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>';
    const { prefix, body: docBody, suffix } = this._body();
    const at = /<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>/.exec(docBody);
    if (!at) throw new Error('this document has no section properties to hang a ' + which + ' on');
    if (at[0].endsWith('/>')) {
      const expanded = at[0].replace(/\/>$/, '>') + ref + '</w:sectPr>';
      this.xml = prefix + docBody.slice(0, at.index) + expanded + docBody.slice(at.index + at[0].length) + suffix;
    } else {
      const openEnd = at.index + at[0].length;
      this.xml = prefix + docBody.slice(0, openEnd) + ref + docBody.slice(openEnd) + suffix;
    }
    this.dirty = true;
    this._undoParts.add(partName);
    return this;
  }

  /**
   * The whole main part, as it stands — one undo step.
   *
   * A document is one XML string and every edit rewrites it, so there is nothing
   * to be clever about: the snapshot IS the string. The workbook has to be
   * choosier because a workbook is many parts and only one changes.
   */
  snapshot() {
    const snap = { xml: this.xml, dirty: this.dirty };
    if (this._undoParts.size) {
      snap.bands = {};
      for (const name of this._undoParts) {
        if (this.pkg.has(name)) snap.bands[name] = this.pkg.text(name);
      }
    }
    return snap;
  }

  /**
   * Put a snapshot back. Everything is re-read from the XML on next access.
   *
   * The dirty flag is NOT taken from the snapshot on faith: a snapshot from
   * before an edit says `dirty: false`, but if a save happened in between,
   * the package's main part now holds the LATER xml — and a save that trusts
   * the flag would skip the write and emit a file that does not match the
   * document. Open → type → save → undo → save produced exactly that before
   * this check: the second save carried the typing the undo had removed.
   */
  restore(snapshot) {
    if (!snapshot || typeof snapshot.xml !== 'string') return this;
    this.xml = snapshot.xml;
    this.dirty = Boolean(snapshot.dirty) || this._flushed !== this.xml;
    // Band parts the snapshot captured go back too. A band CREATED after the
    // snapshot has no entry here; restoring the main xml removes its
    // reference, and the orphaned part is harmless — the same stance the
    // numbering and media parts take.
    if (snapshot.bands) {
      for (const [name, xml] of Object.entries(snapshot.bands)) {
        if (this.pkg.has(name) && this.pkg.text(name) !== xml) this.pkg.write_(name, xml);
      }
    }
    return this;
  }

  /**
   * Content controls, keyed by `w:tag`. These are the binding anchors.
   * @returns {Array<{tag: string, alias: string|null, id: string|null, text: string}>}
   */
  contentControls() {
    const out = [];
    const re = /<w:sdt\b[^>]*>([\s\S]*?)<\/w:sdt>/g;
    let m;
    while ((m = re.exec(this.xml))) {
      const inner = m[1];
      const pr = /<w:sdtPr\b[^>]*>([\s\S]*?)<\/w:sdtPr>/.exec(inner);
      const content = /<w:sdtContent\b[^>]*>([\s\S]*?)<\/w:sdtContent>/.exec(inner);
      const tag = pr ? /<w:tag\b([^>]*)\/>/.exec(pr[1]) : null;
      const alias = pr ? /<w:alias\b([^>]*)\/>/.exec(pr[1]) : null;
      const id = pr ? /<w:id\b([^>]*)\/>/.exec(pr[1]) : null;
      out.push({
        tag: tag ? attrs(tag[1])['w:val'] : null,
        alias: alias ? attrs(alias[1])['w:val'] : null,
        id: id ? attrs(id[1])['w:val'] : null,
        text: content ? textOf(content[1]) : '',
      });
    }
    return out;
  }

  /** Bookmark names, the older anchor mechanism. Read-only for now. */
  bookmarks() {
    const out = [];
    for (const m of this.xml.matchAll(/<w:bookmarkStart\b([^>]*)\/>/g)) {
      const a = attrs(m[1]);
      if (a['w:name'] && !a['w:name'].startsWith('_')) out.push({ id: a['w:id'], name: a['w:name'] });
    }
    return out;
  }

  // ---- writing -------------------------------------------------------------

  /**
   * Replace the text inside a content control, keeping its properties and the
   * formatting of the run that was there.
   *
   * This is the whole `workspace.docs` binding story: a template author marks
   * the spot in Word, we fill it, and nothing else in their letter is touched.
   */
  setContentControlText(tag, value) {
    const escapedTag = String(tag).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const sdtRe = new RegExp(
      '<w:sdt\\b[^>]*>(?:(?!<w:sdt\\b)[\\s\\S])*?<w:tag\\b[^>]*w:val="' + escapedTag + '"[\\s\\S]*?<\\/w:sdt>',
    );
    const found = sdtRe.exec(this.xml);
    if (!found) throw new Error('no content control tagged "' + tag + '"');

    const sdt = found[0];
    const contentRe = /<w:sdtContent\b[^>]*>([\s\S]*?)<\/w:sdtContent>/;
    const content = contentRe.exec(sdt);
    if (!content) throw new Error('content control "' + tag + '" has no <w:sdtContent>');

    const rPr = firstRunProps(content[1]);
    const replaced = sdt.replace(
      contentRe,
      '<w:sdtContent>' + renderRun(rPr, String(value)) + '</w:sdtContent>',
    );
    this.xml = this.xml.replace(sdt, replaced);
    this.dirty = true;
    return this;
  }

  /** Fill many at once; unknown tags are reported rather than silently skipped. */
  fillContentControls(values) {
    const known = new Set(this.contentControls().map((c) => c.tag));
    const missing = Object.keys(values).filter((t) => !known.has(t));
    if (missing.length) {
      throw new Error('no content control tagged: ' + missing.join(', ') +
        '; document has: ' + [...known].filter(Boolean).join(', '));
    }
    for (const [tag, value] of Object.entries(values)) this.setContentControlText(tag, value);
    return this;
  }

  /**
   * Replace a top-level paragraph's text, keeping its `w:pPr` and the first
   * run's `w:rPr`.
   *
   * Refuses paragraphs carrying structure we would destroy — a field code, a
   * bookmark, a comment anchor or a content control. Silently flattening those
   * is exactly the damage this whole design exists to avoid, so it is an error
   * rather than a best effort.
   */
  setParagraphText(index, value) {
    const { prefix, body, suffix } = this._body();
    const paragraphs = this.paragraphs();
    const p = paragraphs[index];
    if (!p) throw new Error('no paragraph at index ' + index);

    const structural = ['w:fldSimple', 'w:fldChar', 'w:bookmarkStart', 'w:commentRangeStart', 'w:sdt', 'w:ins', 'w:del'];
    // \b, not a bare prefix: '<w:ins' is the start of '<w:instrText' too.
    const present = structural.filter((tag) => new RegExp('<' + tag + '\\b').test(p.xml));
    if (present.length) {
      throw new Error(
        'paragraph ' + index + ' carries ' + present.join(', ') +
        '; replacing its text would destroy that. Edit the content control or run instead.',
      );
    }

    const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>|<w:pPr\b[^>]*\/>/.exec(p.xml);
    const open = /<w:p\b[^>]*?>/.exec(p.xml)[0];
    const rebuilt = open + (pPr ? pPr[0] : '') + renderRun(firstRunProps(p.xml), String(value)) + '</w:p>';

    this.xml = prefix + body.slice(0, p.start) + rebuilt + body.slice(p.end) + suffix;
    this.dirty = true;
    return this;
  }

  // ---- structured paragraph editing --------------------------------------

  /**
   * One paragraph, decomposed into the parts an editor needs.
   *
   * `structural` is the important flag: a paragraph carrying a field code,
   * bookmark, comment anchor or content control cannot be rebuilt from its runs
   * without destroying that structure, so the editor treats it as read-only text
   * and edits the content control instead. Refusing loudly beats flattening a
   * template's intelligence silently.
   */
  paragraph(index) {
    const p = this.paragraphs()[index];
    return p ? this._decorate(p) : null;
  }

  /**
   * One paragraph of the EDIT address space, same decomposition, plus its
   * `container`. This is what the editor's backend reads: `paragraph(index)`
   * above keeps addressing prose only, for the callers that predate editable
   * tables.
   */
  editParagraph(index) {
    const p = this.editParagraphs()[index];
    return p ? this._decorate(p) : null;
  }

  editParagraphCount() { return this.editParagraphs().length; }

  _decorate(p) {
    // w:ins and w:del joined the list 2026-08-27: a rebuild reassembles text
    // runs and cannot re-attribute them, so a paragraph mid-review is
    // read-only here until its changes are accepted or rejected in Word —
    // and its tracked state is SHOWN (see `tracked` below) rather than
    // silently flattened, which is what editing used to do.
    // w:txbxContent is on the list because a rebuild reassembles the runs
    // and would drop the box the paragraph anchors, words and all.
    // A note reference is NOT on the list: it is a run of its own with one
    // character of text, and the rebuild writes the element back from it.
    const structural = ['w:fldSimple', 'w:fldChar', 'w:bookmarkStart', 'w:commentRangeStart', 'w:sdt', 'w:ins', 'w:del', 'w:txbxContent']
      .filter((tag) => new RegExp('<' + tag + '\\b').test(p.xml)); // \b: w:ins is a prefix of w:instrText
    // A paragraph INSIDE a body-level content control carries no sdt tag of
    // its own; it is read-only for the same reason one that does is.
    if (p.inSdt && !structural.includes('w:sdt')) structural.push('w:sdt');
    // The paragraph's OWN words: a text box's are the box's (`textBoxes`
    // below), not the anchor's — counted once, drawn once.
    const own = p.xml.includes('<w:txbxContent') ? stripTextBoxes(p.xml) : p.xml;
    const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>|<w:pPr\b[^>]*\/>/.exec(p.xml);
    const style = /<w:pStyle\b[^>]*w:val="([^"]*)"/.exec(p.xml);
    // `open` is normalised to a real opening tag: a self-closing `<w:p/>` (an
    // empty cell, a blank line) must not be used as a prefix and then closed
    // AGAIN by the rebuilders — `<w:p/>…</w:p>` is how typing into a fresh
    // table cell once produced XML no scan could read.
    const rawOpen = /<w:p\b[^>]*?>/.exec(p.xml)?.[0] ?? '<w:p>';
    return {
      index: p.index,
      container: p.container ?? null,
      ...(p.hiddenCell ? { hiddenCell: true } : {}),
      xml: p.xml,
      start: p.start,
      end: p.end,
      text: own === p.xml ? p.text : textOf(own),
      open: rawOpen.endsWith('/>') ? rawOpen.replace(/\/>$/, '>') : rawOpen,
      pPr: pPr ? pPr[0] : null,
      style: style ? style[1] : null,
      // An explicit page break is the author's instruction, not a suggestion —
      // the paginator must not decide it knows better.
      pageBreakBefore: PAGE_BREAK_BEFORE.test(pPr ? pPr[0] : '') || EXPLICIT_BREAK.test(p.xml),
      keepNext: KEEP_NEXT.test(pPr ? pPr[0] : ''),
      keepLines: KEEP_LINES.test(pPr ? pPr[0] : ''),
      // Direct paragraph spacing, if the paragraph sets any — the paginator
      // lets it beat the style's spacing, exactly as Word does.
      spacing: readDirectSpacing(pPr ? pPr[0] : ''),
      // Direct alignment and indent, the same rule: the paragraph's own `w:jc`
      // and `w:ind` beat the style's. Alignment stays in the docx vocabulary
      // ('both', 'end') — the same one table cells and resolved styles carry,
      // so the painter learns each spelling exactly once.
      align: (() => {
        const jc = /<w:jc\b[^>]*w:val="([^"]*)"/.exec(pPr ? pPr[0] : '');
        return jc ? jc[1] : null;
      })(),
      indentPx: (() => {
        const ind = /<w:ind\b([^>]*?)\/?>/.exec(pPr ? pPr[0] : '');
        if (!ind) return null;
        const left = /\bw:left="(-?\d+)"/.exec(ind[1]) || /\bw:start="(-?\d+)"/.exec(ind[1]);
        return left ? twipsToPx(Number(left[1])) : null;
      })(),
      // Tab stops, shading, borders and the other indents — direct formatting
      // the page draws. Null when the paragraph sets none of them.
      decor: readParagraphDecor(pPr ? pPr[0] : ''),
      // A list paragraph names its numbering; the LABEL is computed by the view,
      // because "3." depends on the two list items before it, not on this XML.
      numbering: (() => {
        const numPr = /<w:numPr\b[^>]*>([\s\S]*?)<\/w:numPr>/.exec(pPr ? pPr[0] : '');
        if (!numPr) return null;
        const numId = /<w:numId\b[^>]*w:val="([^"]*)"/.exec(numPr[1]);
        const ilvl = /<w:ilvl\b[^>]*w:val="([^"]*)"/.exec(numPr[1]);
        return numId ? { numId: numId[1], level: Number(ilvl?.[1] ?? 0) || 0 } : null;
      })(),
      images: this._paragraphImages(own),
      // Charts and shapes, RAW — the xml and the box, no interpretation.
      // This layer is format-only; turning a chart part into something
      // paintable is the doc-view backend's job, with @rutba/drawing.
      ...((() => {
        const drawings = this._paragraphRichDrawings(p.xml);
        return drawings.length ? { drawings } : {};
      })()),
      // Text boxes anchored in this paragraph, their paragraphs read the way
      // the body's are — a cover page is nothing but these.
      ...((() => {
        const textBoxes = this._paragraphTextBoxes(p.xml);
        return textBoxes.length ? { textBoxes } : {};
      })()),
      ...(p.inSdt ? { inSdt: true } : {}),
      runs: parseRuns(own),
      structural: structural.length > 0,
      structuralTags: structural,
      // Tracked changes, summarised for the margin: who touched this
      // paragraph, how many times, and what a deletion removed — so the
      // review is VISIBLE here even though resolving it belongs to Word.
      tracked: (() => {
        if (!/<w:(ins|del)\b/.test(p.xml)) return null;
        const authors = new Set();
        let inserted = 0;
        let deleted = 0;
        const removals = [];
        for (const m of p.xml.matchAll(/<w:(ins|del)\b([^>]*)>([\s\S]*?)<\/w:\1>/g)) {
          const a = attrs(m[2]);
          if (a['w:author']) authors.add(unesc(a['w:author']));
          if (m[1] === 'ins') inserted += 1;
          else {
            deleted += 1;
            for (const t of m[3].matchAll(/<w:delText\b[^>]*>([\s\S]*?)<\/w:delText>/g)) removals.push(unesc(t[1]));
          }
        }
        return {
          inserted,
          deleted,
          authors: [...authors],
          deletedText: removals.join('').slice(0, 200) || null,
        };
      })(),
    };
  }

  paragraphCount() { return this.paragraphs().length; }

  _spliceBody(start, end, replacement) {
    const { prefix, body, suffix } = this._body();
    this.xml = prefix + body.slice(0, start) + replacement + body.slice(end) + suffix;
    this.dirty = true;
  }

  _assertEditable(p, what) {
    if (!p) throw new Error('no paragraph at that index');
    if (p.structural) {
      throw new Error(
        'paragraph ' + p.index + ' carries ' + p.structuralTags.join(', ') +
        '; ' + what + ' would destroy that. Edit the content control instead.',
      );
    }
  }

  /**
   * What a paragraph rebuild must NOT throw away.
   *
   * The rebuilders below reassemble a paragraph as `open + pPr + text runs`,
   * and for a long time that was the whole story — until a paragraph carrying
   * an inline image met a keystroke: the drawing run has no `<w:t>`, the text
   * runs were all that was rebuilt, and the company logo silently left the
   * file. So every rebuild now carries the KEPT fragments too:
   *
   *   - a top-level run with no text (a drawing, an explicit `w:br` page
   *     break) survives whole;
   *   - a drawing/object/pict nested INSIDE a text run is lifted into its own
   *     run — the text half is being rewritten, the picture half must not go
   *     with it;
   *   - any other top-level child (`m:oMath`, a proofErr range) rides along
   *     verbatim. Structural paragraphs never reach a rebuild, so bookmarks
   *     and field codes are not this method's problem.
   *
   * Kept fragments are appended AFTER the text runs. For an inline image that
   * can move it within the paragraph's XML — an accepted simplification that
   * matches the painter, which has always drawn images as a block under the
   * paragraph's text rather than at their inline position.
   */
  _keptFragments(p) {
    if (/^<w:p\b[^>]*\/>$/.test(p.xml)) return '';
    const openRaw = /<w:p\b[^>]*?>/.exec(p.xml)[0];
    const inner = p.xml.slice(openRaw.length, p.xml.length - '</w:p>'.length);

    // Top-level children, depth-aware: a run wraps arbitrary subtrees and a
    // flat regex would mistake grandchildren for siblings.
    const kept = [];
    const re = /<([A-Za-z0-9]+:[A-Za-z0-9]+)\b[^>]*?(\/?)>|<\/([A-Za-z0-9]+:[A-Za-z0-9]+)\s*>/g;
    let depth = 0;
    let start = -1;
    let tag = null;
    let m;
    const push = (childTag, chunk) => {
      if (childTag === 'w:pPr') return;
      // The rule is about CONTENT, not tag names: a chunk carrying `<w:t>`
      // anywhere is text the model owns — parseRuns read it and the rebuild
      // rewrites it — so keeping the chunk whole would DOUBLE the text (a
      // `w:ins` wrapper did exactly that before this rule; those paragraphs
      // are structural now, and this stays as the defence in depth). Only
      // what a rewrite cannot express is lifted out. A chunk with no text —
      // a break run, a hyperlink wrapping only an image, a proofErr marker —
      // survives verbatim.
      if (/<w:t\b/.test(chunk)) {
        for (const sub of chunk.matchAll(/<w:(drawing|object|pict)\b[^>]*(?:\/>|>[\s\S]*?<\/w:\1>)/g)) {
          kept.push('<w:r>' + sub[0] + '</w:r>');
        }
      } else if (/<w:(?:footnote|endnote)Reference\b/.test(chunk)) {
        // A note reference is a run the model OWNS — one character of text,
        // written back by renderRuns from the run that carries it. Keeping
        // the chunk too put a second reference at the end of the paragraph.
      } else {
        kept.push(chunk);
      }
    };
    while ((m = re.exec(inner))) {
      if (m[1]) {
        if (m[2] === '/') {
          if (depth === 0) push(m[1], m[0]);
        } else if (depth === 0) {
          start = m.index;
          tag = m[1];
          depth = 1;
        } else {
          depth += 1;
        }
      } else if (m[3] && depth > 0) {
        depth -= 1;
        if (depth === 0) {
          push(tag, inner.slice(start, m.index + m[0].length));
          start = -1;
          tag = null;
        }
      }
    }
    return kept.join('');
  }

  _setRuns(p, runs) {
    const rebuilt = p.open + (p.pPr ?? '') + renderRuns(runs) + this._keptFragments(p) + '</w:p>';
    this._spliceBody(p.start, p.end, rebuilt);
  }

  _split(p, runIndex, offset) {
    const before = [];
    const after = [];
    p.runs.forEach((run, i) => {
      if (i < runIndex) before.push(run);
      else if (i > runIndex) after.push(run);
      else {
        before.push({ ...run, text: run.text.slice(0, offset) });
        after.push({ ...run, text: run.text.slice(offset) });
      }
    });
    // The kept fragments stay with the FIRST half: a split is Enter at the
    // caret, and the image the paragraph carried does not follow the caret
    // onto the new line.
    const first = p.open + (p.pPr ?? '') + renderRuns(before) + this._keptFragments(p) + '</w:p>';
    const second = p.open + (p.pPr ?? '') + renderRuns(after) + '</w:p>';
    this._spliceBody(p.start, p.end, first + second);
  }

  _merge(p, next) {
    this._assertEditable(p, 'merging it');
    this._assertEditable(next, 'merging into it');
    // Adjacent in the LIST is not adjacent in the FILE: a table or a block
    // content control can sit between two paragraphs of the same address
    // space, and splicing across it would delete it silently. That is the
    // exact damage this design exists to refuse — before this guard, a
    // backspace at the start of the paragraph after a table swallowed the
    // whole table.
    const { body } = this._body();
    if (/\S/.test(body.slice(p.end, next.start))) {
      throw new Error(
        'these paragraphs have content between them (a table or a content control); merging would delete it',
      );
    }
    const merged = p.open + (p.pPr ?? '')
      + renderRuns([...p.runs, ...next.runs])
      + this._keptFragments(p) + this._keptFragments(next)
      + '</w:p>';
    this._spliceBody(p.start, next.end, merged);
  }

  /** Replace a paragraph's runs, keeping its opening tag and `w:pPr`. */
  setParagraphRuns(index, runs) {
    const p = this.paragraph(index);
    this._assertEditable(p, 'replacing its runs');
    this._setRuns(p, runs);
    return this;
  }

  /**
   * Split a paragraph at a run/offset boundary — what Enter does.
   * The new paragraph inherits the original's `w:pPr`, so pressing Enter in a
   * list item produces another list item.
   */
  splitParagraph(index, runIndex, offset) {
    const p = this.paragraph(index);
    this._assertEditable(p, 'splitting it');
    this._split(p, runIndex, offset);
    return this;
  }

  /** Merge a paragraph with the one after it — what Backspace at offset 0 does. */
  mergeWithNext(index) {
    this._merge(this.paragraph(index), this.paragraph(index + 1));
    return this;
  }

  /** Insert a new paragraph after `index`, inheriting its paragraph properties. */
  insertParagraphAfter(index, runs = [], { inheritStyle = true } = {}) {
    const p = this.paragraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    const added = p.open + (inheritStyle ? (p.pPr ?? '') : '') + renderRuns(runs) + '</w:p>';
    this._spliceBody(p.end, p.end, added);
    return this;
  }

  removeParagraph(index) {
    const p = this.paragraph(index);
    this._assertEditable(p, 'removing it');
    if (this.paragraphCount() <= 1) throw new Error('a document must keep at least one paragraph');
    this._spliceBody(p.start, p.end, '');
    return this;
  }

  // ---- the edit address space's write half --------------------------------
  //
  // The same five operations, addressed by `editParagraphs()` index — the
  // space that includes table-cell paragraphs. The splices are shared with the
  // top-level methods above; what these add is the boundary discipline:
  // a merge cannot cross a cell wall, and a cell keeps at least one paragraph,
  // because `<w:tc>` without a `<w:p>` is not a valid cell.

  setEditParagraphRuns(index, runs) {
    const p = this.editParagraph(index);
    this._assertEditable(p, 'replacing its runs');
    this._setRuns(p, runs);
    return this;
  }

  splitEditParagraph(index, runIndex, offset) {
    const p = this.editParagraph(index);
    this._assertEditable(p, 'splitting it');
    this._split(p, runIndex, offset);
    return this;
  }

  mergeEditWithNext(index) {
    const p = this.editParagraph(index);
    const next = this.editParagraph(index + 1);
    if (p && next && (p.container ?? null) !== (next.container ?? null)) {
      throw new Error('cannot merge paragraphs across a table boundary');
    }
    this._merge(p, next);
    return this;
  }

  insertEditParagraphAfter(index, runs = [], { inheritStyle = true } = {}) {
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    const added = p.open + (inheritStyle ? (p.pPr ?? '') : '') + renderRuns(runs) + '</w:p>';
    this._spliceBody(p.end, p.end, added);
    return this;
  }

  removeEditParagraph(index) {
    const p = this.editParagraph(index);
    this._assertEditable(p, 'removing it');
    const siblings = this.editParagraphs().filter((e) => (e.container ?? null) === (p.container ?? null));
    if (siblings.length <= 1) {
      throw new Error(p.container
        ? 'a table cell must keep at least one paragraph'
        : 'a document must keep at least one paragraph');
    }
    // A cell must END with a paragraph — `<w:tc>` closing straight after a
    // nested `</w:tbl>` is a file Word refuses to open. When removing this
    // paragraph would leave the cell like that, keep it and empty it instead,
    // which is exactly the paragraph Word itself insists on there.
    if (p.container) {
      const { body } = this._body();
      if (/<\/w:tbl>\s*$/.test(body.slice(0, p.start)) && /^\s*<\/w:tc>/.test(body.slice(p.end))) {
        this._setRuns(p, []);
        return this;
      }
    }
    this._spliceBody(p.start, p.end, '');
    return this;
  }

  // ---- creating table structure -------------------------------------------

  /** The width text can occupy, in twips — page size minus margins. */
  _contentWidthTwips() {
    const { body } = this._body();
    const sectPr = firstElement(body, 'w:sectPr');
    const sz = sectPr ? attrs(firstElement(sectPr, 'w:pgSz') ?? '') : {};
    const mar = sectPr ? attrs(firstElement(sectPr, 'w:pgMar') ?? '') : {};
    const width = Number(sz['w:w']) || 11906;
    const at = (name, fallback) => (mar['w:' + name] === undefined ? fallback : Number(mar['w:' + name]));
    return Math.max(1440, width - at('left', 1440) - at('right', 1440) - at('gutter', 0));
  }

  /**
   * Insert a fresh table after an edit-space paragraph — the ribbon's Insert
   * table. Everything about it is the minimum Word writes for the same
   * gesture: single half-point borders, an equal grid filling the text column,
   * one empty paragraph per cell. A brand-new element, so the append-only
   * discipline is trivially kept — nothing existing is rewritten.
   *
   * Nested creation is refused for now: editing INSIDE an existing table is
   * one thing, growing a table inside a cell is a decision for the day
   * somebody needs it.
   */
  insertTableAfter(index, rows, cols) {
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    if (p.container) throw new Error('cannot insert a table inside a table — place the caret outside it first');
    const r = Math.max(1, Math.min(50, Math.floor(Number(rows) || 0)));
    const c = Math.max(1, Math.min(12, Math.floor(Number(cols) || 0)));

    const colW = Math.max(240, Math.floor(this._contentWidthTwips() / c));
    const borders = '<w:tblBorders>'
      + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
        .map((side) => '<w:' + side + ' w:val="single" w:sz="4" w:space="0" w:color="auto"/>').join('')
      + '</w:tblBorders>';
    const grid = '<w:tblGrid>' + ('<w:gridCol w:w="' + colW + '"/>').repeat(c) + '</w:tblGrid>';
    const cell = '<w:tc><w:tcPr><w:tcW w:w="' + colW + '" w:type="dxa"/></w:tcPr><w:p/></w:tc>';
    const row = '<w:tr>' + cell.repeat(c) + '</w:tr>';
    const tbl = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + borders + '</w:tblPr>'
      + grid + row.repeat(r) + '</w:tbl>';

    // Two tables back to back read as one in Word, and a table with nothing
    // after it leaves the caret nowhere to land below — both get a spacer.
    const { body } = this._body();
    const after = body.slice(p.end);
    const spacer = after === '' || /^<w:(tbl|sectPr)\b/.test(after) ? '<w:p/>' : '';
    this._spliceBody(p.end, p.end, tbl + spacer);
    return this;
  }

  /**
   * Append one row to the table starting at `tableStart` (the `t…` component
   * of a container key) — what Tab in the last cell does. The new row copies
   * the LAST row's cell properties so widths and shading carry on, minus
   * `w:vMerge` (a fresh row must not continue a merge above it) and minus the
   * old row's `w:trPr` (repeating a `w:tblHeader` would make the new row a
   * header too).
   */
  appendTableRow(tableStart) {
    const { body } = this._body();
    const at = Number(tableStart);
    if (!Number.isInteger(at) || !/^<w:tbl\b/.test(body.slice(at))) {
      throw new Error('no table at that position');
    }
    // Find this table's own closing tag, depth-aware for nested tables.
    const re = /<w:tbl\b[^>]*?(\/?)>|<\/w:tbl>/g;
    re.lastIndex = at;
    let depth = 0;
    let end = -1;
    let m;
    while ((m = re.exec(body))) {
      if (m[0].startsWith('</')) {
        depth -= 1;
        if (depth === 0) { end = m.index; break; }
      } else if (m[1] !== '/') {
        depth += 1;
      }
    }
    if (end < 0) throw new Error('unterminated table');

    // Scan the table's CONTENTS, not the element — the outer open tag would
    // otherwise count against the nested-table guard and hide every row.
    const open = /^<w:tbl\b[^>]*>/.exec(body.slice(at))[0];
    const inner = body.slice(at + open.length, end);
    const rows = childElements(inner, 'w:tr', 'w:tbl');
    if (!rows.length) throw new Error('the table has no rows to copy');
    const last = rows[rows.length - 1];

    const cells = childElements(last, 'w:tc', 'w:tbl').map((tcXml) => {
      const tcPr = firstElement(headBefore(tcXml, /<w:(p|tbl)\b/), 'w:tcPr');
      const kept = tcPr ? tcPr.replace(/<w:vMerge\b[^>]*\/?>/g, '') : '';
      return '<w:tc>' + kept + '<w:p/></w:tc>';
    });
    this._spliceBody(end, end, '<w:tr>' + cells.join('') + '</w:tr>');
    return this;
  }

  // ---- restructuring an existing table ------------------------------------

  /**
   * The spans of one table's moving parts — absolute body offsets, so every
   * operation splices precisely and untouched rows keep their bytes. Nested
   * tables are guarded out: `rows` are THIS table's rows only.
   */
  _tableParts(tableStart) {
    const { body } = this._body();
    const at = Number(tableStart);
    if (!Number.isInteger(at) || !/^<w:tbl\b/.test(body.slice(at))) {
      throw new Error('no table at that position');
    }
    const scan = /<w:tbl\b[^>]*?(\/?)>|<\/w:tbl>/g;
    scan.lastIndex = at;
    let depth = 0;
    let end = -1;
    let m;
    while ((m = scan.exec(body))) {
      if (m[0].startsWith('</')) {
        depth -= 1;
        if (depth === 0) { end = m.index; break; }
      } else if (m[1] !== '/') depth += 1;
    }
    if (end < 0) throw new Error('unterminated table');

    const open = /^<w:tbl\b[^>]*>/.exec(body.slice(at))[0];
    const innerStart = at + open.length;
    const inner = body.slice(innerStart, end);

    const rows = [];
    let grid = null;
    const re = /<w:(tbl|tr|tblGrid)\b[^>]*?(\/?)>|<\/w:(tbl|tr|tblGrid)>/g;
    let tblDepth = 0;
    let openAt = -1;
    let openTag = null;
    while ((m = re.exec(inner))) {
      const name = m[1] ?? m[3];
      const opening = Boolean(m[1]);
      if (name === 'tbl') {
        if (opening && m[2] !== '/') tblDepth += 1;
        else if (!opening) tblDepth -= 1;
        continue;
      }
      if (tblDepth > 0) continue;
      if (opening && m[2] === '/') continue;
      if (opening) {
        if (openAt < 0) { openAt = m.index; openTag = name; }
      } else if (openTag === name) {
        const span = { start: innerStart + openAt, end: innerStart + m.index + m[0].length };
        if (name === 'tr') rows.push(span);
        else grid = span;
        openAt = -1;
        openTag = null;
      }
    }
    return { at, end: end + '</w:tbl>'.length, innerStart, inner, rows, grid };
  }

  /** Top-level `w:tc` spans within one row's XML — offsets relative to it. */
  _rowCellSpans(rowXml) {
    const out = [];
    const re = /<w:(tbl|tc)\b[^>]*?(\/?)>|<\/w:(tbl|tc)>/g;
    let tblDepth = 0;
    let depth = 0;
    let start = -1;
    let m;
    while ((m = re.exec(rowXml))) {
      const name = m[1] ?? m[3];
      const opening = Boolean(m[1]);
      if (name === 'tbl') {
        if (opening && m[2] !== '/') tblDepth += 1;
        else if (!opening) tblDepth -= 1;
        continue;
      }
      if (tblDepth > 0) continue;
      if (opening && m[2] === '/') continue;
      if (opening) {
        if (depth === 0) start = m.index;
        depth += 1;
      } else {
        depth -= 1;
        if (depth === 0) out.push({ start, end: m.index + m[0].length });
      }
    }
    return out;
  }

  /**
   * Merged ranges make SOME arithmetic lie, and the guards are per-axis:
   * a vertical merge spans rows, so inserting or deleting a row through it
   * would orphan continuations — but whole-column operations pass through a
   * vertical merge cleanly (the restart and its continuations leave, or gain
   * a neighbour, together). A horizontal gridSpan makes a cell index no
   * longer a column index, so column operations refuse under it — while a
   * whole row, spans and all, can be copied or deleted safely.
   */
  _refuseVerticalMerge(parts) {
    if (/<w:vMerge\b/.test(parts.inner)) {
      throw new Error('This table has vertically merged cells — rows through a merge cannot be restructured yet.');
    }
  }

  _refuseHorizontalMerge(parts) {
    if (/<w:gridSpan\b/.test(parts.inner)) {
      throw new Error('This table has horizontally merged cells — its columns cannot be restructured yet.');
    }
  }

  /** An empty cell copying `fromCellXml`'s properties, minus any merge. */
  _templateCell(fromCellXml) {
    const tcPr = firstElement(headBefore(fromCellXml, /<w:(p|tbl)\b/), 'w:tcPr');
    const kept = tcPr ? tcPr.replace(/<w:vMerge\b[^>]*\/?>/g, '') : '';
    return '<w:tc>' + kept + '<w:p/></w:tc>';
  }

  /** A fresh empty row shaped like the row at `rowIndex`. */
  _templateRow(parts, rowIndex) {
    const { body } = this._body();
    const row = parts.rows[rowIndex] ?? parts.rows[parts.rows.length - 1];
    if (!row) throw new Error('the table has no rows to copy');
    const rowXml = body.slice(row.start, row.end);
    const cells = this._rowCellSpans(rowXml).map((c) => this._templateCell(rowXml.slice(c.start, c.end)));
    return '<w:tr>' + cells.join('') + '</w:tr>';
  }

  insertTableRow(tableStart, rowIndex, where = 'below') {
    const parts = this._tableParts(tableStart);
    this._refuseVerticalMerge(parts);
    const row = parts.rows[rowIndex];
    if (!row) throw new Error('no row ' + rowIndex + ' in this table');
    const fresh = this._templateRow(parts, rowIndex);
    const insertAt = where === 'above' ? row.start : row.end;
    this._spliceBody(insertAt, insertAt, fresh);
    return this;
  }

  deleteTableRow(tableStart, rowIndex) {
    const parts = this._tableParts(tableStart);
    this._refuseVerticalMerge(parts);
    const row = parts.rows[rowIndex];
    if (!row) throw new Error('no row ' + rowIndex + ' in this table');
    // Deleting the only row deletes the table — an empty <w:tbl> is a file
    // Word refuses to open.
    if (parts.rows.length === 1) return this.deleteTable(tableStart);
    this._spliceBody(row.start, row.end, '');
    return this;
  }

  insertTableColumn(tableStart, cellIndex, where = 'right') {
    const parts = this._tableParts(tableStart);
    this._refuseHorizontalMerge(parts);
    const { body } = this._body();

    // Rows sit after the grid, so splicing them bottom-up leaves both the
    // earlier rows' and the grid's offsets valid.
    for (const row of [...parts.rows].reverse()) {
      const rowXml = body.slice(row.start, row.end);
      const cells = this._rowCellSpans(rowXml);
      const target = cells[cellIndex] ?? cells[cells.length - 1];
      if (!target) continue;
      const fresh = this._templateCell(rowXml.slice(target.start, target.end));
      const insertAt = row.start + (where === 'left' ? target.start : target.end);
      this._spliceBody(insertAt, insertAt, fresh);
    }

    if (parts.grid) {
      const { body: after } = this._body();
      const gridXml = after.slice(parts.grid.start, parts.grid.end);
      const cols = [...gridXml.matchAll(/<w:gridCol\b[^>]*\/>/g)];
      const target = cols[cellIndex] ?? cols[cols.length - 1];
      if (target) {
        const insertAt = parts.grid.start + (where === 'left' ? target.index : target.index + target[0].length);
        this._spliceBody(insertAt, insertAt, target[0]);
      }
    }
    return this;
  }

  deleteTableColumn(tableStart, cellIndex) {
    const parts = this._tableParts(tableStart);
    this._refuseHorizontalMerge(parts);
    const { body } = this._body();
    const firstRowCells = parts.rows.length
      ? this._rowCellSpans(body.slice(parts.rows[0].start, parts.rows[0].end)).length
      : 0;
    if (firstRowCells <= 1) return this.deleteTable(tableStart);

    for (const row of [...parts.rows].reverse()) {
      const rowXml = body.slice(row.start, row.end);
      const cells = this._rowCellSpans(rowXml);
      const target = cells[cellIndex];
      if (!target) continue;
      this._spliceBody(row.start + target.start, row.start + target.end, '');
    }
    if (parts.grid) {
      const { body: after } = this._body();
      const gridXml = after.slice(parts.grid.start, parts.grid.end);
      const cols = [...gridXml.matchAll(/<w:gridCol\b[^>]*\/>/g)];
      const target = cols[cellIndex];
      if (target) {
        this._spliceBody(parts.grid.start + target.index, parts.grid.start + target.index + target[0].length, '');
      }
    }
    return this;
  }

  /** A cell's declared horizontal span — 1 unless `w:gridSpan` says more. */
  _cellSpan(cellXml) {
    const g = /<w:gridSpan\b[^>]*\bw:val="(\d+)"/.exec(cellXml);
    return g ? Math.max(1, Number(g[1])) : 1;
  }

  /** Split one `<w:tc>` into open tag, tcPr (or null), content, for rebuilding. */
  _cellParts(cellXml) {
    const open = /^<w:tc\b[^>]*>/.exec(cellXml)[0];
    const tcPr = firstElement(headBefore(cellXml, /<w:(p|tbl)\b/), 'w:tcPr');
    const content = cellXml.slice(open.length + (tcPr ? tcPr.length : 0), cellXml.length - '</w:tc>'.length);
    return { open, tcPr: tcPr ?? null, content };
  }

  /**
   * Merge the cell at `cellIndex` with its right-hand neighbour — Word's
   * "Merge Cells" for the caret-and-one case. The left cell keeps its
   * properties, gains the combined `gridSpan` and the summed width, and both
   * cells' contents sit side by side in it, every paragraph intact. Refused
   * when either cell is part of a vertical merge: widening one row of a
   * vertical range would tear the range.
   */
  mergeTableCellRight(tableStart, rowIndex, cellIndex) {
    const parts = this._tableParts(tableStart);
    const { body } = this._body();
    const row = parts.rows[rowIndex];
    if (!row) throw new Error('no row ' + rowIndex + ' in this table');
    const rowXml = body.slice(row.start, row.end);
    const cells = this._rowCellSpans(rowXml);
    const left = cells[cellIndex];
    const right = cells[cellIndex + 1];
    if (!left || !right) throw new Error('There is no cell to the right to merge with.');
    const leftXml = rowXml.slice(left.start, left.end);
    const rightXml = rowXml.slice(right.start, right.end);
    if (/<w:vMerge\b/.test(leftXml) || /<w:vMerge\b/.test(rightXml)) {
      throw new Error('These cells are part of a vertical merge — widening one row of it would tear the range.');
    }

    const a = this._cellParts(leftXml);
    const b = this._cellParts(rightXml);
    const span = this._cellSpan(leftXml) + this._cellSpan(rightXml);

    // The left tcPr carries on, its gridSpan replaced and its width grown by
    // the right cell's when both speak dxa.
    let tcPr = a.tcPr ?? '<w:tcPr></w:tcPr>';
    if (/^<w:tcPr\b[^>]*\/>$/.test(tcPr)) tcPr = tcPr.replace(/\/>$/, '>') + '</w:tcPr>';
    tcPr = tcPr.replace(/<w:gridSpan\b[^>]*\/>/g, '');
    const wA = /<w:tcW\b[^>]*\bw:w="(\d+)"[^>]*\bw:type="dxa"/.exec(a.tcPr ?? '');
    const wB = /<w:tcW\b[^>]*\bw:w="(\d+)"[^>]*\bw:type="dxa"/.exec(b.tcPr ?? '');
    if (wA && wB) {
      tcPr = tcPr.replace(/<w:tcW\b[^>]*\/>/, '<w:tcW w:w="' + (Number(wA[1]) + Number(wB[1])) + '" w:type="dxa"/>');
    }
    const gridSpan = '<w:gridSpan w:val="' + span + '"/>';
    // gridSpan follows tcW in the schema's order; with no tcW it leads.
    tcPr = /<w:tcW\b[^>]*\/>/.test(tcPr)
      ? tcPr.replace(/(<w:tcW\b[^>]*\/>)/, '$1' + gridSpan)
      : tcPr.replace(/^(<w:tcPr\b[^>]*>)/, '$1' + gridSpan);

    const merged = a.open + tcPr + a.content + b.content + '</w:tc>';
    this._spliceBody(row.start + left.start, row.start + right.end, merged);
    return this;
  }

  /** tcPr with every merge attribute stripped; '' when nothing else remains. */
  _unmergedPr(tcPr, { keepWidth = true } = {}) {
    let out = tcPr ?? '';
    out = out.replace(/<w:gridSpan\b[^>]*\/>/g, '').replace(/<w:vMerge\b[^>]*\/>/g, '');
    if (!keepWidth) out = out.replace(/<w:tcW\b[^>]*\/>/g, '');
    return /<w:tcPr\b[^>]*>\s*<\/w:tcPr>/.test(out) ? '' : out;
  }

  /**
   * Split a merged cell back into its grid — BOTH directions. A horizontal
   * span becomes its columns (content in the first, the rest alike and
   * empty); a vertical range unhides every continuation row; a rectangle
   * does both. Splitting must start from the merge's FIRST cell, which is
   * the only one a caret can be in anyway — the continuations are hidden.
   */
  splitTableCell(tableStart, rowIndex, cellIndex) {
    const parts = this._tableParts(tableStart);
    const { body } = this._body();
    const row = parts.rows[rowIndex];
    if (!row) throw new Error('no row ' + rowIndex + ' in this table');
    const rowXml = body.slice(row.start, row.end);
    const cells = this._rowCellSpans(rowXml);
    const target = cells[cellIndex];
    if (!target) throw new Error('no cell ' + cellIndex + ' in this row');
    const cellXml = rowXml.slice(target.start, target.end);
    const span = this._cellSpan(cellXml);
    const merge = /<w:vMerge\b([^>]*)\/>/.exec(cellXml);
    const isRestart = Boolean(merge && /w:val="restart"/.test(merge[1]));
    if (merge && !isRestart) {
      throw new Error('Split a vertical merge from its first cell.');
    }
    if (span <= 1 && !isRestart) throw new Error('This cell is not merged — there is nothing to split.');

    // Every row the merge covers: the restart row, then continuations below
    // in the same column position.
    const affected = [{ rowIndex, keepContent: true }];
    if (isRestart) {
      for (let r = rowIndex + 1; r < parts.rows.length; r += 1) {
        const rXml = body.slice(parts.rows[r].start, parts.rows[r].end);
        const rCells = this._rowCellSpans(rXml);
        const c = rCells[cellIndex];
        if (!c) break;
        const cXml = rXml.slice(c.start, c.end);
        const m = /<w:vMerge\b([^>]*)\/>/.exec(cXml);
        if (!m || /w:val="restart"/.test(m[1])) break;
        affected.push({ rowIndex: r, keepContent: false });
      }
    }

    // Bottom-up, so earlier spans stay valid as lower rows are rewritten.
    for (const { rowIndex: r, keepContent } of [...affected].reverse()) {
      const rSpan = parts.rows[r];
      const { body: current } = this._body();
      const rXml = current.slice(rSpan.start, rSpan.end);
      const c = this._rowCellSpans(rXml)[cellIndex];
      const cXml = rXml.slice(c.start, c.end);
      const { open, tcPr, content } = this._cellParts(cXml);
      const first = open + this._unmergedPr(tcPr) + (keepContent ? content : '<w:p/>') + '</w:tc>';
      const clone = '<w:tc>' + this._unmergedPr(tcPr, { keepWidth: false }) + '<w:p/></w:tc>';
      this._spliceBody(rSpan.start + c.start, rSpan.start + c.end, first + clone.repeat(span - 1));
    }
    return this;
  }

  /**
   * Merge a RECTANGLE of cells — Word's Merge Cells over a selection. Every
   * row in the range keeps one cell spanning its columns; the top one shows
   * the whole range's contents in reading order and the lower ones continue
   * it vertically (`w:vMerge`). Refused when the range already contains a
   * merge: split first, then merge the shape you mean.
   */
  mergeTableCells(tableStart, r1, c1, r2, c2) {
    const parts = this._tableParts(tableStart);
    const { body } = this._body();
    const top = Math.min(r1, r2);
    const bottom = Math.max(r1, r2);
    const left = Math.min(c1, c2);
    const right = Math.max(c1, c2);
    if (top === bottom && left === right) {
      throw new Error('Select across at least two cells to merge.');
    }

    // Gather the range first: guards, contents, widths — before any splice.
    const contents = [];
    let width = 0;
    let widthKnown = true;
    const rows = [];
    for (let r = top; r <= bottom; r += 1) {
      const rSpan = parts.rows[r];
      if (!rSpan) throw new Error('the selection reaches past the table\'s last row');
      const rXml = body.slice(rSpan.start, rSpan.end);
      // An existing merge in any touched row outranks the bounds complaint:
      // under a gridSpan the missing column IS the merge, and "split them
      // first" is the answer that helps.
      if (/<w:(vMerge|gridSpan)\b/.test(rXml)) {
        throw new Error('The selection already contains merged cells — split them first.');
      }
      const cells = this._rowCellSpans(rXml);
      if (!cells[right]) throw new Error('the selection reaches past the table\'s last column');
      for (let c = left; c <= right; c += 1) {
        const cXml = rXml.slice(cells[c].start, cells[c].end);
        if (/<w:(vMerge|gridSpan)\b/.test(cXml)) {
          throw new Error('The selection already contains merged cells — split them first.');
        }
        const { tcPr, content } = this._cellParts(cXml);
        if (content !== '<w:p/>' && content !== '') contents.push(content);
        if (r === top) {
          const w = /<w:tcW\b[^>]*\bw:w="(\d+)"[^>]*\bw:type="dxa"/.exec(tcPr ?? '');
          if (w) width += Number(w[1]);
          else widthKnown = false;
        }
      }
      rows.push({ rSpan, cells });
    }

    const spanW = right - left + 1;
    const spanH = bottom - top + 1;
    const decorate = (tcPr, isTop) => {
      let out = this._unmergedPr(tcPr);
      if (out === '') out = '<w:tcPr></w:tcPr>';
      if (/^<w:tcPr\b[^>]*\/>$/.test(out)) out = out.replace(/\/>$/, '>') + '</w:tcPr>';
      if (widthKnown && width > 0) {
        out = /<w:tcW\b[^>]*\/>/.test(out)
          ? out.replace(/<w:tcW\b[^>]*\/>/, '<w:tcW w:w="' + width + '" w:type="dxa"/>')
          : out.replace(/^(<w:tcPr\b[^>]*>)/, '$1<w:tcW w:w="' + width + '" w:type="dxa"/>');
      }
      const marks = (spanW > 1 ? '<w:gridSpan w:val="' + spanW + '"/>' : '')
        + (spanH > 1 ? (isTop ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>') : '');
      return /<w:tcW\b[^>]*\/>/.test(out)
        ? out.replace(/(<w:tcW\b[^>]*\/>)/, '$1' + marks)
        : out.replace(/^(<w:tcPr\b[^>]*>)/, '$1' + marks);
    };

    // Bottom-up, splicing each row's range down to one cell.
    for (let i = rows.length - 1; i >= 0; i -= 1) {
      const r = top + i;
      const { rSpan, cells } = rows[i];
      const rXml = body.slice(rSpan.start, rSpan.end);
      const anchor = this._cellParts(rXml.slice(cells[left].start, cells[left].end));
      const merged = anchor.open
        + decorate(anchor.tcPr, r === top)
        + (r === top ? (contents.join('') || '<w:p/>') : '<w:p/>')
        + '</w:tc>';
      this._spliceBody(rSpan.start + cells[left].start, rSpan.start + cells[right].end, merged);
    }
    return this;
  }

  /**
   * Set one column's width, in twips — the grid's `gridCol` and every row's
   * `tcW` together, so measurement and paint agree. Refused under a
   * horizontal merge, where a cell index stops being a column index.
   */
  setTableColumnWidth(tableStart, cellIndex, twips) {
    const parts = this._tableParts(tableStart);
    this._refuseHorizontalMerge(parts);
    const width = Math.round(Number(twips));
    if (!(width >= 144 && width <= 20000)) {
      throw new Error('a column width must be between 0.25cm and 35cm');
    }
    const { body } = this._body();

    const setTcW = (cellXml) => {
      const { open, tcPr, content } = this._cellParts(cellXml);
      const el = '<w:tcW w:w="' + width + '" w:type="dxa"/>';
      let next = tcPr ?? '<w:tcPr>' + el + '</w:tcPr>';
      if (tcPr) {
        if (/^<w:tcPr\b[^>]*\/>$/.test(next)) next = next.replace(/\/>$/, '>') + '</w:tcPr>';
        next = /<w:tcW\b[^>]*\/>/.test(next)
          ? next.replace(/<w:tcW\b[^>]*\/>/, el)
          : next.replace(/^(<w:tcPr\b[^>]*>)/, '$1' + el); // tcW leads tcPr by schema order
      }
      return open + next + content + '</w:tc>';
    };

    for (const row of [...parts.rows].reverse()) {
      const rowXml = body.slice(row.start, row.end);
      const cells = this._rowCellSpans(rowXml);
      const target = cells[cellIndex];
      if (!target) continue;
      this._spliceBody(row.start + target.start, row.start + target.end, setTcW(rowXml.slice(target.start, target.end)));
    }
    if (parts.grid) {
      const { body: after } = this._body();
      const gridXml = after.slice(parts.grid.start, parts.grid.end);
      const cols = [...gridXml.matchAll(/<w:gridCol\b[^>]*\/>/g)];
      const target = cols[cellIndex];
      if (target) {
        this._spliceBody(
          parts.grid.start + target.index,
          parts.grid.start + target.index + target[0].length,
          '<w:gridCol w:w="' + width + '"/>',
        );
      }
    }
    return this;
  }

  deleteTable(tableStart) {
    const parts = this._tableParts(tableStart);
    // A document must keep a paragraph for the caret to land in; if the
    // table was the whole body, one takes its place.
    const replacement = this.paragraphs().length ? '' : '<w:p/>';
    this._spliceBody(parts.at, parts.end, replacement);
    return this;
  }

  /** Flush and return package bytes. */
  save() {
    if (this.dirty) { this.pkg.write_(this.mainPart, this.xml); this._flushed = this.xml; }
    return this.pkg.write();
  }

  modifiedParts() {
    if (this.dirty) { this.pkg.write_(this.mainPart, this.xml); this._flushed = this.xml; }
    return this.pkg.modifiedParts();
  }
}

export { WORD_NS };
