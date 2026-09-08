/**
 * OOXML backend — Workspace's `.docx`.
 *
 * Nearly a pass-through, because `@rutba/ooxml`'s `Document` was already shaped
 * around paragraphs and runs. That is the point: the port was extracted from
 * working code rather than designed in the abstract, so it fits.
 *
 * This is the ONLY file in `@rutba/doc-view` that imports `@rutba/ooxml`. Mail
 * imports the HTML backend instead and never pulls the format layer in.
 */
import { Document, withToggle, esc, unesc, STANDARD_PARAGRAPH_STYLES } from '@rutba/ooxml';
import { parseChartXml, parseShapeXml, buildChart, buildShape, svgDataUri, scene } from '@rutba/drawing';
import { DocView } from '../view.js';

export class OoxmlBackend {
  constructor(doc) { this.doc = doc; }

  static open(buf) { return new OoxmlBackend(Document.open(buf)); }

  // The editor's paragraph index is the EDIT address space — every editable
  // paragraph in document order, table-cell paragraphs included, each block
  // carrying the `container` key of the cell it lives in (null for prose).
  // `Document.paragraphs()` still exists for prose-only callers; the editor
  // stopped being one the day a caret could enter a table.
  paragraphCount() { return this.doc.editParagraphCount(); }
  paragraph(index) { return withDrawingsPainted(this.doc.editParagraph(index)); }

  setParagraphRuns(index, runs) { this.doc.setEditParagraphRuns(index, runs); return this; }
  splitParagraph(index, runIndex, offset) { this.doc.splitEditParagraph(index, runIndex, offset); return this; }
  mergeWithNext(index) { this.doc.mergeEditWithNext(index); return this; }
  removeParagraph(index) { this.doc.removeEditParagraph(index); return this; }
  insertParagraphAfter(index, runs, opts) { this.doc.insertEditParagraphAfter(index, runs, opts); return this; }

  /** Insert a fresh table after a body paragraph — the ribbon's Insert table. */
  insertTable(index, rows, cols) { this.doc.insertTableAfter(index, rows, cols); return this; }
  /** Insert a picture as its own paragraph — the ribbon's Insert picture. */
  insertImage(index, spec) { this.doc.insertImageParagraph(index, spec); return this; }
  /** A chart after a table, its data literal in the part; a preset shape. */
  insertChartAfterTable(tableStart, spec) { this.doc.insertChartAfterTable(tableStart, spec); return this; }
  insertShapeParagraph(index, spec) { this.doc.insertShapeParagraph(index, spec); return this; }
  /** Orientation, paper size and margins — the Layout tab. */
  setPageSetup(spec) { this.doc.setPageSetup(spec); return this; }
  /** Mint a link token for a URL — a relationship plus wrapper attributes. */
  makeLink(url) { return this.doc.makeLink(url); }
  /** What a link token points at, for the toolbar and the painter. */
  linkTarget(token) { return this.doc.linkTarget(token); }
  /** Append a row to the table whose container key starts `t<tableStart>`. */
  appendTableRow(tableStart) { this.doc.appendTableRow(tableStart); return this; }
  /** Restructure an existing table — the contextual Table tab. */
  insertTableRow(tableStart, rowIndex, where) { this.doc.insertTableRow(tableStart, rowIndex, where); return this; }
  deleteTableRow(tableStart, rowIndex) { this.doc.deleteTableRow(tableStart, rowIndex); return this; }
  insertTableColumn(tableStart, cellIndex, where) { this.doc.insertTableColumn(tableStart, cellIndex, where); return this; }
  deleteTableColumn(tableStart, cellIndex) { this.doc.deleteTableColumn(tableStart, cellIndex); return this; }
  deleteTable(tableStart) { this.doc.deleteTable(tableStart); return this; }
  mergeTableCellRight(tableStart, rowIndex, cellIndex) { this.doc.mergeTableCellRight(tableStart, rowIndex, cellIndex); return this; }
  mergeTableCells(tableStart, r1, c1, r2, c2) { this.doc.mergeTableCells(tableStart, r1, c1, r2, c2); return this; }
  splitTableCell(tableStart, rowIndex, cellIndex) { this.doc.splitTableCell(tableStart, rowIndex, cellIndex); return this; }
  setTableColumnWidth(tableStart, cellIndex, twips) { this.doc.setTableColumnWidth(tableStart, cellIndex, twips); return this; }

  /**
   * Formatting is a verbatim `<w:rPr>` edit — fonts and colours ride along.
   * The editor's tag vocabulary is its own; only this file knows that its
   * `s` is stored as `<w:strike>` (never to be confused with `w:dstrike`,
   * which rides through untouched like every unmodelled property).
   */
  toggleRunFormat(rPr, tag, on) {
    const docxTag = { b: 'b', i: 'i', u: 'u', s: 'strike' }[tag];
    if (!docxTag) throw new Error('unknown format: ' + tag);
    return withToggle(rPr, docxTag, on);
  }

  /**
   * Set or clear a VALUE run property — font family, size (points) or colour.
   *
   * The counterpart to `toggleRunFormat` for the properties that carry a value
   * rather than being on/off. A non-null value writes the docx element; `null`
   * strips it. Everything else in `<w:rPr>` is preserved exactly, the same way
   * the toggles preserve fonts and colours around them.
   */
  setRunProp(rPr, prop, value) { return withRunProp(rPr, prop, value); }

  /** Read family/size/colour back off an rPr — the toolbar's caret state. */
  readRunProps(rPr) { return readRunProps(rPr, typeof this.doc.themeFonts === 'function' ? this.doc.themeFonts() : null); }

  /**
   * Read a paragraph's alignment, left indentation and named style — the
   * paragraph-level counterpart of `readRunProps`. Alignment comes back in the
   * editor's neutral vocabulary ('left'|'center'|'right'|'justify'), never the
   * docx spelling, so the format-free view above never has to learn that
   * justify is stored as "both". Indentation is in twips (1/20 point), the unit
   * the port speaks. `style` is the `<w:pStyle>` id, or null for a paragraph
   * riding the document default.
   */
  getParagraphProps(index) {
    const p = this.doc.editParagraph(index);
    if (!p) return null;
    const props = readParagraphProps(p.pPr);
    // List membership is a paragraph property too, but its VOCABULARY is neutral
    // ('bullet'|'number'|null) rather than the numId a docx stores: the numId is
    // meaningless outside this file, and classifying it here means the view — and
    // an HTML backend supplying the same shape for an `<ol>`/`<ul>` — never has to
    // learn that a bullet is a numbering definition whose level-0 numFmt is
    // "bullet". Anything that is not a bullet counts as a number.
    props.listType = this._listTypeOf(p.numbering);
    return props;
  }

  /** A paragraph's `numbering:{numId,level}` -> 'bullet' | 'number' | null. */
  _listTypeOf(numbering) {
    if (!numbering) return null;
    const defs = typeof this.doc.numberingDefs === 'function' ? this.doc.numberingDefs() : null;
    const levels = defs ? defs[numbering.numId] : null;
    const format = levels ? (levels[numbering.level]?.format ?? levels[0]?.format) : null;
    return format === 'bullet' ? 'bullet' : 'number';
  }

  /**
   * Set or clear ONE paragraph property — alignment or left indentation — by
   * splicing `<w:jc>`/`<w:ind>` into the paragraph's `<w:pPr>`. The paragraph
   * counterpart of `setRunProp`: every other pPr child (numbering, style,
   * spacing) is preserved, an emptied pPr collapses to nothing, and the runs are
   * spliced back verbatim so formatting a paragraph never rewrites a run.
   *
   *   prop 'align'          value 'left'|'center'|'right'|'justify' or null to clear
   *   prop 'indentTwips'    value a twip count (0 or null clears)
   *   prop 'style'          value a paragraph style id from the document's own
   *                         catalogue, or null to clear back to the default
   *   prop 'lineSpacing'    value a multiplier (1, 1.15, 1.5, 2…), null clears
   *   prop 'spaceBeforePts' value points of space above, null clears
   *   prop 'spaceAfterPts'  value points of space below, null clears
   *   prop 'pageBreakBefore' true starts the paragraph on a fresh page, falsy clears
   */
  setParagraphProp(index, prop, value) {
    const p = this.doc.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    if (prop === 'style' && value != null) {
      // Validate FIRST — a refused id must not leave a new part behind. The
      // catalogue offers the standard set when the file has no styles part,
      // so a standard id passes here and the ensure below gives it somewhere
      // to point; an unknown id refuses before anything is written.
      if (!this.paragraphStyleCatalogue().some((s) => s.id === value)) {
        // Writing a pStyle the file cannot resolve would be a dangling reference
        // dressed up as formatting — Word shows nothing and so would we. The
        // ribbon never offers one (its dropdown is built FROM the catalogue), so
        // arriving here means a programmatic action got the id wrong: refuse
        // loudly rather than store it silently.
        throw new Error('this document has no paragraph style "' + value + '"');
      }
      // A document with NO styles part gets the standard catalogue written in
      // now, so the Heading a person just picked has somewhere to point —
      // this is how a legacy blank document reaches a heading at all. A file
      // WITH a styles part is never touched here.
      if (typeof this.doc.ensureParagraphStyles === 'function') this.doc.ensureParagraphStyles();
    }
    return this._replacePPr(p, withParagraphProp(p.pPr, prop, value));
  }

  /**
   * The named styles a person can PICK — every paragraph style the document's
   * `word/styles.xml` resolves, as `{ id, name }` in the file's own order. The
   * ribbon's Style dropdown is built from this, which is what keeps it honest:
   * a document with no styles part yields an empty list and the dropdown is
   * simply absent (the no-dead-controls rule), never a menu of styles the file
   * could not resolve. `*default*` is docstyles' synthetic entry for a
   * paragraph with no pStyle — an implementation detail, not a pickable style,
   * so it is filtered here rather than taught to the view.
   */
  paragraphStyleCatalogue() {
    const resolved = this.doc.paragraphStyles();
    const own = Object.keys(resolved)
      .filter((id) => id !== '*default*')
      .map((id) => ({ id, name: resolved[id].name ?? id }));
    // A file with no styles part offers the STANDARD set instead of nothing:
    // picking one writes the catalogue into the file (see setParagraphProp),
    // so these are honest controls, not dead ones — the pick both creates the
    // style and applies it.
    return own.length ? own : STANDARD_PARAGRAPH_STYLES.map((s) => ({ ...s }));
  }

  /**
   * Make a paragraph a bulleted or numbered list, or take it out of a list.
   *
   * The paragraph sibling of `setParagraphProp`, for the one paragraph property
   * that needs the rest of the document: a list item points at a numbering
   * definition, so `'bullet'`/`'number'` first asks `@rutba/ooxml` to ensure the
   * matching definition exists (reusing one if the file already has it), then
   * splices `<w:numPr>` into `<w:pPr>` at level 0. `null` drops the `<w:numPr>`
   * again. Every other pPr child and all the runs ride through untouched, exactly
   * as alignment and indentation do.
   */
  setParagraphList(index, listType) {
    const p = this.doc.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    let numId = null;
    if (listType != null) {
      const ids = this.doc.ensureListNumbering();
      numId = listType === 'bullet' ? ids.bullet : ids.number;
    }
    return this._replacePPr(p, withNumPr(p.pPr, numId));
  }

  /**
   * Splice a rebuilt `<w:pPr>` back into a paragraph, preserving its opening tag
   * and its runs. A self-closing `<w:p/>` has no place to hang a pPr, so it is
   * opened up first; the normal path keeps everything after the old pPr — the
   * runs — exactly as it was.
   */
  _replacePPr(p, nextPPr) {
    if (nextPPr === (p.pPr ?? null)) return this;   // nothing actually changed
    const selfClosing = /^<w:p\b[^>]*\/>$/.test(p.xml);
    const openTag = selfClosing ? p.open.replace(/\/>$/, '>') : p.open;
    const tail = selfClosing ? '</w:p>' : p.xml.slice(p.open.length + (p.pPr ? p.pPr.length : 0));
    this.doc._spliceBody(p.start, p.end, openTag + (nextPPr ?? '') + tail);
    return this;
  }

  setContentControlText(tag, value) { this.doc.setContentControlText(tag, value); return this; }
  contentControls() { return this.doc.contentControls(); }

  /** The body in document order, so a table appears where the author put it. */
  blocks() { return this.doc.blocks(); }
  /** Page size and margins, so the editor can draw a page rather than a ribbon. */
  section() { return this.doc.section(); }
  /** Headers and footers, by reference type — pagination gives them a home. */
  headerFooters() { return this.doc.headerFooters(); }
  /** The document's comments, anchored to edit-space paragraphs. */
  comments() { return this.doc.comments(); }
  registerCommentUndo() { this.doc.registerCommentUndo(); return this; }
  addComment(index, spec) { this.doc.addComment(index, spec); return this; }
  /** The default bands as an editing panel sees them, and their edit path. */
  bandInfo() { return this.doc.bandInfo(); }
  registerBandUndo(which) { this.doc.registerBandUndo(which); return this; }
  setBandText(which, lines) { this.doc.setBandText(which, lines); return this; }
  /** Paragraph styles resolved from styles.xml — flattened, CSS pixels. */
  paragraphStyles() { return this.doc.paragraphStyles(); }
  /** Character styles by id — what a run's `w:rStyle` gives it. */
  characterStyles() { return this.doc.characterStyles(); }
  /** Footnotes and endnotes by id, displayed paragraphs each — see `Document.notes`. */
  notes() { return this.doc.notes(); }
  registerNoteUndo(kind) { return this.doc.registerNoteUndo(kind); }
  /** The note's words into the part; returns the id the reference must carry. */
  addNote(kind, text) { return this.doc.addNote(kind, text); }
  setNoteText(kind, id, text) { return this.doc.setNoteText(kind, id, text); }
  /** Numbering definitions: numId -> levels. Labels are the view's to count. */
  numberingDefs() { return this.doc.numberingDefs(); }

  /** One undo step: the main part as it stands. */
  snapshot() { return this.doc.snapshot(); }
  restore(state) { this.doc.restore(state); return this; }

  save() { return this.doc.save(); }
  modifiedParts() { return this.doc.modifiedParts(); }
}

export { Document };

// ---- value run properties: font family, size, colour -----------------------
//
// b/i/u are toggles, handled by `@rutba/ooxml`'s `withToggle`. Family, size and
// colour carry a VALUE, so their `<w:rPr>` editing lives here — in the one file
// allowed to know what a `.docx` run property looks like. Each is a
// self-contained element (`<w:rFonts/>`, `<w:sz/>`, `<w:color/>`) spliced into
// the run's `<w:rPr>` or lifted back out, with everything around it preserved
// verbatim, exactly as `withToggle` preserves everything around a toggle.

/** Points -> half-points, the unit a docx stores font size in. 12pt -> "24". */
const ptToHalfPoints = (pt) => String(Math.round(Number(pt) * 2));

/** docx colour: six hex digits, uppercase, no alpha, no leading '#'. */
function normaliseColour(hex) {
  const h = String(hex).replace(/^#/, '').trim().toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(h)) {
    throw new Error('colour must be #RRGGBB (six hex digits): ' + hex);
  }
  return h;
}

/**
 * The highlighter's palette is NAMED, not hex — `<w:highlight w:val="yellow"/>`
 * — and the names are the spec's ST_HighlightColor list, closed. Writing a name
 * Word cannot read would be a corrupt file dressed up as formatting, so an
 * unknown one refuses loudly. (`none` is the spec's clear; ours is `null`,
 * which strips the element rather than storing a no-op.)
 */
const HIGHLIGHT_NAMES = new Set([
  'black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen',
  'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta',
  'red', 'white', 'yellow',
]);
function normaliseHighlight(name) {
  const n = String(name).trim();
  if (!HIGHLIGHT_NAMES.has(n)) throw new Error('unknown highlight colour: ' + name);
  return n;
}

/**
 * Splice one self-contained element (matched by its tag) into an rPr, replace an
 * existing one, or remove it — preserving the rest. Mirrors `withToggle`'s
 * contract: an rPr left empty collapses back to `null`.
 */
function withElement(rPr, tag, element) {
  const one = new RegExp('<w:' + tag + '\\b[^>]*/?>');
  if (element === null) {
    if (!rPr) return null;
    const stripped = rPr.replace(new RegExp('<w:' + tag + '\\b[^>]*/?>', 'g'), '');
    return /<w:rPr\b[^>]*>\s*<\/w:rPr>/.test(stripped) ? null : stripped;
  }
  if (!rPr) return '<w:rPr>' + element + '</w:rPr>';
  if (/<w:rPr\b[^>]*\/>/.test(rPr)) return '<w:rPr>' + element + '</w:rPr>';
  if (one.test(rPr)) return rPr.replace(one, element);
  // A value element rides near the end of rPr; Word reads it there just as it
  // reads the toggles `withToggle` writes at the front.
  return rPr.replace('</w:rPr>', element + '</w:rPr>');
}

/**
 * Set (or clear) the font family, keeping any other `<w:rFonts>` attributes —
 * an eastAsia face, a hint — that the run already carried. Word stores the Latin
 * family in three attributes (ascii/hAnsi/cs); all three follow the choice.
 */
function withFontName(rPr, name) {
  const FAMILY = ['w:ascii', 'w:hAnsi', 'w:cs'];
  const existing = /<w:rFonts\b([^>]*)\/>/.exec(rPr || '');

  if (name === null) {
    if (!existing) return rPr ?? null;
    let attrs = existing[1];
    for (const k of FAMILY) attrs = attrs.replace(new RegExp('\\s+' + k + '="[^"]*"'), '');
    const rebuilt = /\S/.test(attrs) ? '<w:rFonts' + attrs + '/>' : '';
    const stripped = rPr.replace(/<w:rFonts\b[^>]*\/>/, rebuilt);
    return /<w:rPr\b[^>]*>\s*<\/w:rPr>/.test(stripped) ? null : stripped;
  }

  const value = esc(String(name));
  if (existing) {
    let attrs = existing[1];
    for (const k of FAMILY) {
      const attrRe = new RegExp('\\s+' + k + '="[^"]*"');
      attrs = attrRe.test(attrs)
        ? attrs.replace(attrRe, ' ' + k + '="' + value + '"')
        : attrs + ' ' + k + '="' + value + '"';
    }
    return rPr.replace(/<w:rFonts\b[^>]*\/>/, '<w:rFonts' + attrs + '/>');
  }

  const element = '<w:rFonts w:ascii="' + value + '" w:hAnsi="' + value + '" w:cs="' + value + '"/>';
  if (!rPr) return '<w:rPr>' + element + '</w:rPr>';
  if (/<w:rPr\b[^>]*\/>/.test(rPr)) return '<w:rPr>' + element + '</w:rPr>';
  // rFonts belongs at the FRONT of rPr by the schema's ordering.
  return rPr.replace(/^(<w:rPr\b[^>]*>)/, '$1' + element);
}

/** Set or clear one value run property. `null` clears; anything else sets. */
function withRunProp(rPr, prop, value) {
  if (prop === 'fontName') return withFontName(rPr, value == null ? null : value);
  if (prop === 'fontSize') {
    return withElement(rPr, 'sz', value == null ? null : '<w:sz w:val="' + ptToHalfPoints(value) + '"/>');
  }
  if (prop === 'fontColour') {
    return withElement(rPr, 'color', value == null ? null : '<w:color w:val="' + normaliseColour(value) + '"/>');
  }
  if (prop === 'highlight') {
    return withElement(rPr, 'highlight', value == null ? null : '<w:highlight w:val="' + normaliseHighlight(value) + '"/>');
  }
  if (prop === 'vertAlign') {
    if (value != null && value !== 'superscript' && value !== 'subscript') throw new Error('vertAlign is superscript, subscript or null');
    return withElement(rPr, 'vertAlign', value == null ? null : '<w:vertAlign w:val="' + value + '"/>');
  }
  throw new Error('unknown run property: ' + prop);
}

/** Family/size/colour/highlight off an rPr — the inverse of `withRunProp`. */
function readRunProps(rPr, themeFonts = null) {
  const out = { fontName: null, fontSize: null, fontColour: null, highlight: null };
  if (!rPr) return out;
  const font = /<w:rFonts\b[^>]*\bw:ascii="([^"]*)"/.exec(rPr);
  if (font) out.fontName = unesc(font[1]);
  else if (themeFonts) {
    // Named through the theme instead — the heading or the body face.
    const slot = /<w:rFonts\b[^>]*\bw:asciiTheme="([^"]*)"/.exec(rPr);
    if (slot) out.fontName = slot[1].startsWith('major') ? themeFonts.major : themeFonts.minor;
  }
  // Superscript and subscript — a footnote reference, a squared metre.
  const vert = /<w:vertAlign\b[^>]*\bw:val="([^"]*)"/.exec(rPr);
  if (vert && vert[1] !== 'baseline') out.vertAlign = vert[1];
  if (/<w:caps\b(?![^>]*w:val="(?:0|false)")/.test(rPr)) out.caps = true;
  if (/<w:smallCaps\b(?![^>]*w:val="(?:0|false)")/.test(rPr)) out.smallCaps = true;
  const sz = /<w:sz\b[^>]*\bw:val="([^"]*)"/.exec(rPr);
  if (sz && /^\d+$/.test(sz[1])) out.fontSize = Number(sz[1]) / 2;
  const colour = /<w:color\b[^>]*\bw:val="([^"]*)"/.exec(rPr);
  if (colour && /^[0-9A-Fa-f]{6}$/.test(colour[1])) out.fontColour = colour[1].toUpperCase();
  const highlight = /<w:highlight\b[^>]*\bw:val="([^"]*)"/.exec(rPr);
  if (highlight && highlight[1] !== 'none') out.highlight = highlight[1];
  return out;
}

// ---- paragraph properties: alignment and indentation -----------------------
//
// The RUN helpers above splice `<w:sz>`/`<w:color>` into a `<w:rPr>`. These are
// their PARAGRAPH-level siblings: alignment (`<w:jc>`) and left indentation
// (`<w:ind>`) live in `<w:pPr>`, the first child of a `<w:p>`. This is the one
// file that knows that shape, so everything about it stays here — the view only
// ever asks for `{ align, indentTwips }` in neutral terms.
//
// Two invariants carry the fidelity guarantee: only jc and ind are ever
// touched, so numbering, styles and spacing ride through untouched; and a pPr we
// empty collapses back to nothing, so a paragraph that had no explicit
// properties returns to having none.

/** Editor alignment -> the docx `w:jc` value. Justify is stored as "both". */
const TO_DOCX_ALIGN = { left: 'left', center: 'center', right: 'right', justify: 'both', both: 'both' };
/** The docx `w:jc` value -> the editor's neutral vocabulary. */
const FROM_DOCX_ALIGN = { left: 'left', center: 'center', right: 'right', both: 'justify', start: 'left', end: 'right' };

/** Alignment, left indent and style id off a `<w:pPr>` — the inverse of the writers below. */
function readParagraphProps(pPr) {
  const out = {
    align: null, indentTwips: null, style: null,
    lineSpacing: null, spaceBeforePts: null, spaceAfterPts: null,
  };
  if (!pPr) return out;
  const pStyle = /<w:pStyle\b[^>]*\bw:val="([^"]*)"/.exec(pPr);
  if (pStyle) out.style = unesc(pStyle[1]);
  const jc = /<w:jc\b[^>]*\bw:val="([^"]*)"/.exec(pPr);
  if (jc) out.align = FROM_DOCX_ALIGN[jc[1]] ?? jc[1];
  const ind = /<w:ind\b([^>]*?)\/?>/.exec(pPr);
  if (ind) {
    // `w:left` is the classic attribute; `w:start` is its bidi-aware successor.
    const left = /\bw:left="(-?\d+)"/.exec(ind[1]) || /\bw:start="(-?\d+)"/.exec(ind[1]);
    if (left) out.indentTwips = Number(left[1]);
  }
  const spacing = /<w:spacing\b([^>]*?)\/?>/.exec(pPr);
  if (spacing) {
    const before = /\bw:before="(\d+)"/.exec(spacing[1]);
    if (before) out.spaceBeforePts = Number(before[1]) / 20;
    const after = /\bw:after="(\d+)"/.exec(spacing[1]);
    if (after) out.spaceAfterPts = Number(after[1]) / 20;
    const line = /\bw:line="(\d+)"/.exec(spacing[1]);
    const rule = /\bw:lineRule="([^"]*)"/.exec(spacing[1]);
    // Only the multiplier form reads back as "line spacing" — an exact height
    // is a different thing, and reporting 1.15 for it would invite the ribbon
    // to overwrite a deliberate layout with a factor.
    if (line && (!rule || rule[1] === 'auto')) out.lineSpacing = Number(line[1]) / 240;
  }
  return out;
}

/**
 * The top-level children of a `<w:pPr>`'s inner content, with their tag names
 * and spans. A depth-aware scan, because a child such as `<w:numPr>` or
 * `<w:rPr>` wraps its own elements and a flat regex would mistake those for
 * siblings.
 */
function pPrChildren(inner) {
  const out = [];
  const re = /<w:([A-Za-z]+)\b[^>]*?(\/?)>|<\/w:([A-Za-z]+)\s*>/g;
  let depth = 0;
  let start = -1;
  let tag = null;
  let m;
  while ((m = re.exec(inner))) {
    if (m[1]) {
      const selfClose = m[2] === '/';
      if (depth === 0) {
        if (selfClose) out.push({ tag: m[1], start: m.index, end: m.index + m[0].length });
        else { start = m.index; tag = m[1]; depth = 1; }
      } else if (!selfClose) {
        depth += 1;
      }
    } else if (m[3] && depth > 0) {
      depth -= 1;
      if (depth === 0) { out.push({ tag, start, end: m.index + m[0].length }); start = -1; tag = null; }
    }
  }
  return out;
}

// The schema order of the `<w:pPr>` children we might sit among. A new element
// goes in front of the first existing child that outranks it, which is what
// keeps `<w:ind>` before `<w:jc>` and both before a paragraph-mark `<w:rPr>`.
const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
  'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
  'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct',
  'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
  'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap',
  'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl',
  'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
];
const rankOf = (tag) => { const i = PPR_ORDER.indexOf(tag); return i === -1 ? PPR_ORDER.length : i; };

/** Insert `element` (a `<w:TAG.../>`) into pPr inner content at its schema slot. */
function insertOrdered(inner, tag, element) {
  const rank = rankOf(tag);
  for (const child of pPrChildren(inner)) {
    if (rankOf(child.tag) > rank) return inner.slice(0, child.start) + element + inner.slice(child.start);
  }
  return inner + element;
}

/** Split a `<w:pPr>` into its opening tag, inner content and close. */
function splitPPr(pPr) {
  if (!pPr) return { open: '<w:pPr>', inner: '', close: '</w:pPr>' };
  if (/^<w:pPr\b[^>]*\/>$/.test(pPr)) return { open: pPr.replace(/\/>$/, '>'), inner: '', close: '</w:pPr>' };
  const open = /^<w:pPr\b[^>]*>/.exec(pPr)[0];
  return { open, inner: pPr.slice(open.length, pPr.length - '</w:pPr>'.length), close: '</w:pPr>' };
}

/** Reassemble a pPr, collapsing an emptied one back to nothing. */
const joinPPr = (open, inner, close) => (inner === '' ? null : open + inner + close);

/**
 * Set or clear `<w:numPr>` in a pPr, preserving every other child. A numId sets
 * the paragraph's list membership at level 0; `null` removes it. numPr sits at
 * its own schema slot (after pStyle, before jc/ind), which `insertOrdered`
 * places for us, so numbering, alignment and indentation coexist correctly.
 */
function withNumPr(pPr, numId) {
  const { open, inner, close } = splitPPr(pPr);
  const existing = pPrChildren(inner).find((c) => c.tag === 'numPr');
  const without = existing ? inner.slice(0, existing.start) + inner.slice(existing.end) : inner;
  if (numId == null) return joinPPr(open, without, close);
  const element = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + numId + '"/></w:numPr>';
  return joinPPr(open, insertOrdered(without, 'numPr', element), close);
}

/**
 * Set or clear `<w:pStyle>` in a pPr, preserving every other child. A style id
 * writes the reference; `null` removes it, dropping the paragraph back to the
 * document default. pStyle is the FIRST child of `<w:pPr>` by the schema —
 * `insertOrdered` ranks it ahead of numPr, ind and jc — and every direct
 * property already there rides through untouched, because that is exactly what
 * pStyle plus direct formatting means in a docx: the style is the base and the
 * paragraph's own jc/ind/numPr override it.
 */
function withPStyle(pPr, styleId) {
  const { open, inner, close } = splitPPr(pPr);
  const existing = pPrChildren(inner).find((c) => c.tag === 'pStyle');
  const without = existing ? inner.slice(0, existing.start) + inner.slice(existing.end) : inner;
  if (styleId == null) return joinPPr(open, without, close);
  const element = '<w:pStyle w:val="' + esc(String(styleId)) + '"/>';
  return joinPPr(open, insertOrdered(without, 'pStyle', element), close);
}

/** Set, replace or remove `<w:jc>` in a pPr, preserving every other child. */
function withJc(pPr, docxValue) {
  const { open, inner, close } = splitPPr(pPr);
  const existing = pPrChildren(inner).find((c) => c.tag === 'jc');
  const without = existing ? inner.slice(0, existing.start) + inner.slice(existing.end) : inner;
  const next = docxValue === null ? without : insertOrdered(without, 'jc', '<w:jc w:val="' + docxValue + '"/>');
  return joinPPr(open, next, close);
}

/**
 * Set or clear the left indent in `<w:ind>`, keeping any other indent attributes
 * (a first-line indent, a right indent) the paragraph already carried. Clearing
 * the only attribute removes the now-empty `<w:ind>`.
 */
function withInd(pPr, twips) {
  const { open, inner, close } = splitPPr(pPr);
  const existing = pPrChildren(inner).find((c) => c.tag === 'ind');
  const set = twips != null && twips > 0;

  if (!set) {
    if (!existing) return joinPPr(open, inner, close);
    const stripped = stripAttr(inner.slice(existing.start, existing.end), 'w:left');
    return joinPPr(open, inner.slice(0, existing.start) + stripped + inner.slice(existing.end), close);
  }
  const element = setIndLeft(existing ? inner.slice(existing.start, existing.end) : null, twips);
  if (existing) return joinPPr(open, inner.slice(0, existing.start) + element + inner.slice(existing.end), close);
  return joinPPr(open, insertOrdered(inner, 'ind', element), close);
}

/** Add or replace `w:left` on a `<w:ind>` element, leaving its other attrs be. */
function setIndLeft(indXml, twips) {
  if (!indXml) return '<w:ind w:left="' + twips + '"/>';
  const attrs = /^<w:ind\b([^>]*?)\s*\/?>$/.exec(indXml)[1];
  const next = /\bw:left="[^"]*"/.test(attrs)
    ? attrs.replace(/\bw:left="[^"]*"/, 'w:left="' + twips + '"')
    : attrs + ' w:left="' + twips + '"';
  return '<w:ind' + next + '/>';
}

/** Drop one attribute from a self-closing element; return '' if none remain. */
function stripAttr(elXml, attr) {
  const m = /^(<w:\w+)\b([^>]*?)\s*\/?>$/.exec(elXml);
  const attrs = m[2].replace(new RegExp('\\s*' + attr + '="[^"]*"'), '');
  return /\S/.test(attrs) ? m[1] + attrs + '/>' : '';
}

/** Add or replace one attribute on a self-closing element. */
function setAttr(elXml, attr, value) {
  const m = /^(<w:\w+)\b([^>]*?)\s*\/?>$/.exec(elXml);
  const attrRe = new RegExp('\\b' + attr + '="[^"]*"');
  const attrs = attrRe.test(m[2])
    ? m[2].replace(attrRe, attr + '="' + value + '"')
    : m[2] + ' ' + attr + '="' + value + '"';
  return m[1] + attrs + '/>';
}

/**
 * Set or clear attributes of `<w:spacing>` in a pPr, keeping any it already
 * carried — a file's `w:beforeAutospacing`, say, rides through a line-spacing
 * change untouched. `changes` maps attribute name to a value or null; an
 * element left with no attributes is removed, so a paragraph that had no
 * direct spacing returns to having none.
 */
function withSpacing(pPr, changes) {
  const { open, inner, close } = splitPPr(pPr);
  const existing = pPrChildren(inner).find((c) => c.tag === 'spacing');
  let el = existing ? inner.slice(existing.start, existing.end) : '<w:spacing/>';
  for (const [attr, value] of Object.entries(changes)) {
    el = value === null ? (stripAttr(el, attr) || '<w:spacing/>') : setAttr(el, attr, value);
  }
  if (!/\bw:[a-zA-Z]+="/.test(el)) el = '';
  if (existing) return joinPPr(open, inner.slice(0, existing.start) + el + inner.slice(existing.end), close);
  if (el === '') return joinPPr(open, inner, close);
  return joinPPr(open, insertOrdered(inner, 'spacing', el), close);
}

/**
 * Set or clear `<w:pageBreakBefore/>` in a pPr. A boolean paragraph property,
 * so clearing removes the element rather than writing `w:val="0"` — a
 * paragraph that never asked for a break should not carry the vocabulary of
 * one.
 */
function withPageBreakBefore(pPr, on) {
  const { open, inner, close } = splitPPr(pPr);
  const existing = pPrChildren(inner).find((c) => c.tag === 'pageBreakBefore');
  const without = existing ? inner.slice(0, existing.start) + inner.slice(existing.end) : inner;
  if (!on) return joinPPr(open, without, close);
  return joinPPr(open, insertOrdered(without, 'pageBreakBefore', '<w:pageBreakBefore/>'), close);
}

/** Set or clear one paragraph property. Neutral in, docx out. */
function withParagraphProp(pPr, prop, value) {
  if (prop === 'align') {
    if (value == null) return withJc(pPr, null);
    const docx = TO_DOCX_ALIGN[value];
    if (!docx) throw new Error('unknown alignment: ' + value);
    return withJc(pPr, docx);
  }
  if (prop === 'indentTwips') {
    return withInd(pPr, value == null ? null : Math.max(0, Math.round(Number(value))));
  }
  if (prop === 'style') {
    return withPStyle(pPr, value == null ? null : value);
  }
  if (prop === 'lineSpacing') {
    // A multiplier: 240 twentieths of a line is single spacing. Setting one
    // pins the rule to `auto`; clearing removes both attributes together, so
    // no orphaned lineRule survives to change the meaning of a later line.
    if (value == null) return withSpacing(pPr, { 'w:line': null, 'w:lineRule': null });
    const factor = Number(value);
    if (!Number.isFinite(factor) || factor <= 0) throw new Error('line spacing must be a positive multiplier: ' + value);
    return withSpacing(pPr, { 'w:line': String(Math.round(factor * 240)), 'w:lineRule': 'auto' });
  }
  if (prop === 'spaceBeforePts') {
    if (value == null) return withSpacing(pPr, { 'w:before': null });
    return withSpacing(pPr, { 'w:before': String(Math.max(0, Math.round(Number(value) * 20))) });
  }
  if (prop === 'spaceAfterPts') {
    if (value == null) return withSpacing(pPr, { 'w:after': null });
    return withSpacing(pPr, { 'w:after': String(Math.max(0, Math.round(Number(value) * 20))) });
  }
  if (prop === 'pageBreakBefore') {
    return withPageBreakBefore(pPr, Boolean(value));
  }
  throw new Error('unknown paragraph property: ' + prop);
}

// ---- charts and shapes: raw xml in, paintable images out --------------------
//
// The format layer hands over a chart part's xml or a `wps:wsp` verbatim; the
// drawing COMMONS — the same package the sheet grid and Studio draw with —
// turns them into an SVG data URI that rides the ordinary images pipeline.
// One rendering stack for the whole estate, and the painter never learns a
// new node type. A drawing we cannot interpret paints nothing and survives
// the round trip untouched, exactly as before.

function drawingToImage(d) {
  try {
    if (d.kind === 'chart') {
      const width = d.widthPx || 480;
      const height = d.heightPx || 300;
      const spec = parseChartXml(d.chartXml, { width, height });
      if (!spec) return null;
      return { name: d.name ?? spec.title ?? 'Chart', widthPx: width, heightPx: height, href: svgDataUri(buildChart(spec)) };
    }
    if (d.kind === 'shape') {
      const width = d.widthPx || 120;
      const height = d.heightPx || 80;
      const descriptor = parseShapeXml(d.shapeXml);
      if (!descriptor) return null;
      const built = buildShape(descriptor, { x: 0, y: 0, width, height });
      return {
        name: d.name ?? descriptor.geometry,
        widthPx: width,
        heightPx: height,
        href: svgDataUri(scene({ width, height, background: 'none', children: [built] })),
      };
    }
  } catch { /* a drawing that will not paint is an empty margin, not a crash */ }
  return null;
}

function withDrawingsPainted(p) {
  if (!p?.drawings?.length) return p;
  const painted = p.drawings.map(drawingToImage).filter(Boolean);
  if (!painted.length) return p;
  return { ...p, images: [...(p.images ?? []), ...painted] };
}

/** Open a .docx for editing. The Workspace entry point. */
export const openDocx = (buf) => new DocView(OoxmlBackend.open(buf));
