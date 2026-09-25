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
import { ENVELOPE_SIZES, ENVELOPE_STYLES_XML, envelopeXml, envelopeDocumentXml, labelSheetXml, labelParagraph, nextRecordParagraph, labelProduct, NEXT_FIELD_RUNS } from './labels.js';
import { readParagraphStyles, readCharacterStyles, readNumberingDefs, readThemeFonts, readThemeColours, STANDARD_STYLES_XML } from './docstyles.js';
import {
  findDrawings, readDrawing, readBodyPr, readShapeLook, DRAWING_NS, Z_BASE, Z_STEP, textBoxRun, fallbackFor,
  withAnchorAttrs, withDocPr, withPosition, withWrap, withExtent, withTransform, withShapeFill, withShapeLine, withBodyPr,
  toAnchor, toInline, memberXml, groupMembers, groupGraphic, memberToDrawing, anchorXml, EMU_PER_PX,
} from './drawings.js';

/** What the Arrange commands need of a drawing, off its XML: its id, what it is, its z-order and its turn. */
function arrangeOf(drawingXml) {
  const d = readDrawing(drawingXml);
  return {
    id: d.id, kind: d.kind, relativeHeight: d.relativeHeight,
    ...(d.rot ? { rot: d.rot } : {}), ...(d.flipH ? { flipH: true } : {}), ...(d.flipV ? { flipV: true } : {}),
    ...(d.hidden ? { hidden: true } : {}),
  };
}

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
  textOf, parseRuns, renderRuns, renderRun, firstRunProps, RPR_RE, mapComplexFieldResults, mergeFieldsOnly,
} from './runs.js';

/**
 * Paper sizes Windows numbers as envelopes (DMPAPER_ENV_*): No. 9–14, DL,
 * C3–C6, C65, B4–B6, Italy, Monarch, 6¾. Word writes the number as
 * `w:pgSz w:code` on the section an envelope is added as.
 */
const ENVELOPE_PAPER = new Set([19, 20, 21, 22, 23, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38]);
const isEnvelopeSection = (sectPrXml) => {
  const code = /<w:pgSz\b[^>]*\bw:code="(\d+)"/.exec(String(sectPrXml || ''));
  return code ? ENVELOPE_PAPER.has(Number(code[1])) : false;
};

/**
 * The section whose page the editor draws and Layout sets up: the first one,
 * as it always was — unless that is an envelope Mailings → Envelopes added
 * at the front of a letter, in which case the letter's own, after it.
 * Answers a regex match (`index`, `[0]`) — only the opening tag with
 * `openOnly` — or null.
 */
function mainSectPr(body, { openOnly = false } = {}) {
  let first = null;
  let main = null;
  for (const m of String(body).matchAll(/<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)) {
    if (!first) first = m;
    if (!isEnvelopeSection(m[0])) { main = m; break; }
  }
  const hit = main || first;
  if (!hit || !openOnly) return hit;
  const open = /^<w:sectPr\b[^>]*?\/?>/.exec(hit[0]);
  return Object.assign([open[0]], { index: hit.index });
}

/**
 * A mail merge source's path from the `file:///…` target Word writes —
 * `file:///C:\Users\…\list.csv` on Windows, `file:///home/…` elsewhere.
 */
function sourcePathOf(target) {
  let s = String(target);
  try { s = decodeURI(s); } catch { /* a stray % — keep the text as it is */ }
  s = s.replace(/^file:\/*/i, '');
  if (!/^[A-Za-z]:[\\/]/.test(s) && !s.startsWith('\\\\')) s = '/' + s;
  return s;
}

/** Whether a section is an envelope Word added, and its paper number, from its `w:sectPr`. */
function sectionKind(sectPrXml) {
  const code = /<w:pgSz\b[^>]*\bw:code="(\d+)"/.exec(String(sectPrXml || ''));
  return { envelope: isEnvelopeSection(sectPrXml), code: code ? Number(code[1]) : null };
}

/** One section's page: its size, orientation and margins, from its `w:sectPr`. */
function pageOf(sectPrXml) {
  const s = parseSection(sectPrXml || '');
  return { widthPx: s.widthPx, heightPx: s.heightPx, orientation: s.orientation, margins: s.margins };
}

/** A paragraph's own `w:sectPr` — the end of a section — and the break it asks for, or null. */
function paragraphSectionBreak(pPrXml) {
  if (!pPrXml || !pPrXml.includes('<w:sectPr')) return null;
  const sect = /<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(pPrXml);
  if (!sect) return null;
  const type = /<w:type\b[^>]*\bw:val="([^"]*)"/.exec(sect[0]);
  return { type: type ? type[1] : 'nextPage', sectPr: sect[0] };
}

/**
 * An explicit page break. `w:pageBreakBefore` is a paragraph property; a
 `w:br w:type="page"` is a run-level break sitting inside the paragraph. Word
 * honours both, so both have to be seen or a report breaks in the wrong places.
 */
const PAGE_BREAK_BEFORE = /<w:pageBreakBefore\b(?![^>]*w:val="(?:0|false)")/;
const EXPLICIT_BREAK = /<w:br\b[^>]*w:type="page"/;

/**
 * Where an anchored drawing sits and how text treats it, from `wp:anchor`:
 * the wrap (square, tight, through, topAndBottom, none), which side the text
 * runs on, the horizontal and vertical positions (an alignment or an offset,
 * and what each is relative to), the distances text keeps from it, and
 * whether it is behind the text. An inline drawing has none of this — it is
 * a character in the line — and answers `anchored: false`.
 */
export function anchorLayout(inner) {
  const anchor = /<wp:anchor\b([^>]*)>/.exec(inner);
  if (!anchor) return { anchored: false };
  const a = attrs(anchor[1]);
  const emu = (v) => (v == null || v === '' ? 0 : Number(v) / 9525);
  const wrapEl = /<wp:(wrapSquare|wrapTight|wrapThrough|wrapTopAndBottom|wrapNone)\b([^>]*)\/?>/.exec(inner);
  const wrap = wrapEl
    ? { wrapSquare: 'square', wrapTight: 'tight', wrapThrough: 'through', wrapTopAndBottom: 'topAndBottom', wrapNone: 'none' }[wrapEl[1]]
    : 'none';
  const wrapSide = wrapEl ? attrs(wrapEl[2] || '').wrapText ?? 'bothSides' : null;
  const position = (axis) => {
    const el = new RegExp('<wp:position' + axis + '\\b([^>]*)>([\\s\\S]*?)<\\/wp:position' + axis + '>').exec(inner);
    if (!el) return { rel: null, align: null, offsetPx: null };
    const align = /<wp:align>([^<]*)<\/wp:align>/.exec(el[2])?.[1] ?? null;
    const off = /<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(el[2]);
    return { rel: attrs(el[1]).relativeFrom ?? null, align, offsetPx: off ? Number(off[1]) / 9525 : null };
  };
  const h = position('H');
  const v = position('V');
  return {
    anchored: true,
    wrap,
    wrapSide,
    hAlign: h.align, hOffsetPx: h.offsetPx, hRel: h.rel,
    vAlign: v.align, vOffsetPx: v.offsetPx, vRel: v.rel,
    dist: { l: emu(a.distL), r: emu(a.distR), t: emu(a.distT), b: emu(a.distB) },
    behind: a.behindDoc === '1' || a.behindDoc === 'true',
  };
}
/**
 * The two "keep" properties a paginator honours: a heading that stays with
 * the paragraph after it, and a paragraph whose lines are not split across
 * pages. Both are toggles that may be written as w:val="0" to turn off.
 */
const KEEP_NEXT = /<w:keepNext\b(?![^>]*w:val="(?:0|false)")/;
const KEEP_LINES = /<w:keepLines\b(?![^>]*w:val="(?:0|false)")/;

/**
 * A drop cap, off `<w:framePr>`. Word makes one by splitting the initial
 * letter into its own paragraph and framing it — `w:dropCap="drop"` (the
 * letter sits in the column) or `"margin"` (it hangs in the margin) — with
 * `w:lines` saying how many lines deep it stands. Any other framePr (a pull
 * quote frame, say) reads back as no drop cap; this file only ever WRITES
 * the drop-cap shape, so that is the only shape it needs to recognise.
 */
const FRAME_PR = /<w:framePr\b([^>]*)\/?>/;
function readDropCap(pPr) {
  const m = FRAME_PR.exec(pPr || '');
  if (!m) return null;
  const kind = /\bw:dropCap="([^"]*)"/.exec(m[1])?.[1];
  if (kind !== 'drop' && kind !== 'margin') return null;
  const lines = /\bw:lines="(\d+)"/.exec(m[1]);
  return { kind, lines: lines ? Number(lines[1]) : 3 };
}

/**
 * A paragraph in a frame placed on the page — Word's Envelope Address, a
 * text frame from an older document: `w:framePr` anchored to the page with
 * its left and top given, in px from the page's own top-left corner, and
 * the frame's size. A drop cap is a frame too, and is read above instead;
 * a frame placed by alignment alone is left in the flow, as before.
 */
function readFrame(pPr) {
  const m = FRAME_PR.exec(pPr || '');
  if (!m || /\bw:dropCap="(?:drop|margin)"/.test(m[1])) return null;
  const a = attrs(m[1]);
  if (a['w:hAnchor'] !== 'page' || a['w:vAnchor'] !== 'page' || a['w:x'] === undefined || a['w:y'] === undefined) return null;
  return {
    xPx: twipsToPx(a['w:x']), yPx: twipsToPx(a['w:y']),
    widthPx: a['w:w'] ? twipsToPx(a['w:w']) : null, heightPx: a['w:h'] ? twipsToPx(a['w:h']) : null,
    exact: a['w:hRule'] === 'exact',
  };
}

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
    let cleared = false;
    for (const m of tabs[1].matchAll(/<w:tab\b([^>]*)\/>/g)) {
      const a = attrs(m[1]);
      if (a['w:val'] === 'clear') { cleared = true; continue; }
      if (a['w:pos'] === undefined) continue;
      stops.push({ align: a['w:val'] || 'left', posPx: twipsToPx(a['w:pos']), leader: a['w:leader'] && a['w:leader'] !== 'none' ? a['w:leader'] : null });
    }
    // A paragraph's own stops replace its style's; one that only clears
    // stops (what the ruler writes when the last one is dragged away) has
    // none, and an empty list says so where a missing one would not.
    if (stops.length || cleared) out.tabs = stops.sort((x, y) => x.posPx - y.posPx);
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

/**
 * What a table says about itself before its first row: the grid's column
 * widths and the table's own width — a share of the text width (`pct`, in
 * fiftieths of a percent), a fixed one (`dxa`), or nothing. The page draws
 * the columns from these; the ruler and the grips on the page move them.
 */
function tableHead(body, at) {
  const firstRow = body.indexOf('<w:tr', at);
  const head = body.slice(at, firstRow === -1 ? at + 4000 : firstRow);
  const grid = /<w:tblGrid\b[^>]*>([\s\S]*?)<\/w:tblGrid>/.exec(head);
  const gridPx = grid ? [...grid[1].matchAll(/<w:gridCol\b[^>]*\bw:w="(\d+)"/g)].map((g) => twipsToPx(Number(g[1]))) : [];
  const tblW = /<w:tblW\b([^>]*)\/>/.exec(head);
  const w = tblW ? attrs(tblW[1]) : {};
  const tableWidth = w['w:type'] === 'pct'
    ? { type: 'pct', value: /%\s*$/.test(String(w['w:w'])) ? parseFloat(w['w:w']) * 50 : Number(w['w:w']) || 5000 }
    : w['w:type'] === 'dxa' && Number(w['w:w']) > 0 ? { type: 'dxa', value: Number(w['w:w']) }
    : null;
  // A table with every border explicitly off, laid out to fixed widths — a
  // sheet of labels — is drawn as Word draws one: no lines (the dashed
  // gridlines only), its own cell margins, the widths as given.
  const borders = /<w:tblBorders\b[^>]*>([\s\S]*?)<\/w:tblBorders>/.exec(head);
  const sides = borders ? [...borders[1].matchAll(/<w:(top|left|bottom|right|insideH|insideV|start|end)\b[^>]*\bw:val="([^"]*)"/g)] : [];
  const bare = sides.length >= 4 && sides.every((s) => s[2] === 'nil' || s[2] === 'none');
  const fixed = /<w:tblLayout\b[^>]*\bw:type="fixed"/.test(head);
  const mar = /<w:tblCellMar\b[^>]*>([\s\S]*?)<\/w:tblCellMar>/.exec(head);
  const side = (name) => { const m = mar ? new RegExp('<w:' + name + '\\b[^>]*\\bw:w="(\\d+)"').exec(mar[1]) : null; return m ? twipsToPx(Number(m[1])) : null; };
  const look = bare || fixed ? { bare, fixed, ...(mar ? { cellMarginPx: { left: side('left') ?? side('start') ?? 0, right: side('right') ?? side('end') ?? 0, top: side('top') ?? 0, bottom: side('bottom') ?? 0 } } : {}) } : null;
  return { gridPx: gridPx.length ? gridPx : null, tableWidth, ...(look ? { look } : {}) };
}

/** A row's own height, if the file sets one, and whether it is exact or a floor. */
function rowHead(body, at) {
  const firstCell = body.indexOf('<w:tc', at);
  const head = body.slice(at, firstCell === -1 ? at + 1000 : firstCell);
  const h = /<w:trHeight\b([^>]*)\/>/.exec(head);
  if (!h) return {};
  const a = attrs(h[1]);
  const val = Number(a['w:val']);
  return val > 0 ? { heightPx: twipsToPx(val), rule: a['w:hRule'] || 'atLeast' } : {};
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
/** The watermark's paragraph: the one holding a VML text path. */
const WATERMARK_P = /<v:textpath\b/;
const VML_NS = 'urn:schemas-microsoft-com:vml';
const VML_OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
/** WordArt's plain-text shape, as Word declares it before every watermark. */
const SHAPETYPE_136 = '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">'
  + '<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>'
  + '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>'
  + '<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>'
  + '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>'
  + '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>'
  + '<o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>';

/**
 * The watermark's paragraph as Word writes one: the shape type, then the
 * shape — centred on the margins, rotated, filled in the colour at half
 * opacity, its words a text path — in a run the spell checker leaves alone.
 */
function watermarkParagraph(text, colour, rotation) {
  const attr = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const fill = /^#?[0-9a-fA-F]{6}$/.test(String(colour)) ? '#' + String(colour).replace('#', '').toLowerCase() : String(colour || 'silver');
  const deg = Math.round(Number(rotation) || 0);
  return '<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:pict>' + SHAPETYPE_136
    + '<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t136"'
    + ' style="position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:' + deg
    + ';z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"'
    + ' o:allowincell="f" fillcolor="' + attr(fill) + '" stroked="f">'
    + '<v:fill opacity=".5"/>'
    + '<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="' + attr(text) + '"/>'
    + '</v:shape></w:pict></w:r></w:p>';
}
const COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const COMMENTS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const CHART_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const EMPTY_COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>';
/** The image formats the insert accepts — what every Word since 2007 renders. */
const IMAGE_EXTENSIONS = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp' };
/** CSS pixels to EMUs, the unit DrawingML measures in. 1px at 96dpi = 9525. */
const PX_TO_EMU = 9525;

/** One level of a list definition: its number format and text, half an inch further in per level. */
const lvlXml = (i, fmt, text, rPr = '') =>
  '<w:lvl w:ilvl="' + i + '"><w:start w:val="1"/><w:numFmt w:val="' + fmt + '"/>' +
  '<w:lvlText w:val="' + text + '"/><w:lvlJc w:val="left"/>' +
  '<w:pPr><w:ind w:left="' + (720 * (i + 1)) + '" w:hanging="360"/></w:pPr>' + rPr +
  '</w:lvl>';

/** Word's bullet list: a round bullet, then a hollow one, then a square, and round again below. */
const bulletAbstract = (id) => {
  const BULLETS = [
    ['&#61623;', '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'],
    ['o', '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/></w:rPr>'],
    ['&#61607;', '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr>'],
  ];
  let xml = '<w:abstractNum w:abstractNumId="' + id + '"><w:multiLevelType w:val="hybridMultilevel"/>';
  for (let i = 0; i < 9; i++) xml += lvlXml(i, 'bullet', BULLETS[i % 3][0], BULLETS[i % 3][1]);
  return xml + '</w:abstractNum>';
};

/** Word's numbered list: 1. a. i. and round again — nine levels, so Increase indent has somewhere to go. */
const numberAbstract = (id) => {
  const FMT = ['decimal', 'lowerLetter', 'lowerRoman'];
  let xml = '<w:abstractNum w:abstractNumId="' + id + '"><w:multiLevelType w:val="hybridMultilevel"/>';
  for (let i = 0; i < 9; i++) xml += lvlXml(i, FMT[i % 3], '%' + (i + 1) + '.');
  return xml + '</w:abstractNum>';
};

/** The multilevel list — 1. 1.1. 1.1.1. — each level's number carrying the ones above it. */
const outlineAbstract = (id) => {
  let xml = '<w:abstractNum w:abstractNumId="' + id + '"><w:multiLevelType w:val="multilevel"/>';
  for (let i = 0; i < 9; i++) {
    const text = Array.from({ length: i + 1 }, (_, k) => '%' + (k + 1)).join('.') + '.';
    xml += lvlXml(i, 'decimal', text);
  }
  return xml + '</w:abstractNum>';
};

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
    // `_paragraphTextBoxes` reads them from there. Listing them IN THE FLOW
    // made a cover page's text box lines into body paragraphs and LOST the
    // paragraph that carried the box.
    //
    // They are listed now, but as a story of their own, the way Word keeps a
    // text box's words: after every body paragraph, each carrying the
    // container `x<offset of its w:txbxContent>`, so a caret can live in a box
    // and type there while the body's indices — and every adjacency the
    // editor's guards rely on — are exactly what they were. Only a box whose
    // anchor is a plain body paragraph (not a cell's, not a content
    // control's) and whose content is paragraphs alone is listed; any other
    // box stays opaque and is drawn from its anchor, read-only, as before.
    // A box's VML twin in `mc:Fallback` is never listed: it repeats the words.
    let boxDepth = 0;
    const fallbacks = [...body.matchAll(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g)].map((f) => [f.index, f.index + f[0].length]);
    const inFallback = (at) => fallbacks.some(([a, b]) => at >= a && at < b);
    let openBox = null; // the addressable box being read: { content, pStart }
    const pendingBoxes = []; // its paragraphs, waiting for the anchor to close
    const boxParas = []; // every box's paragraphs, listed after the body
    const boxIndex = new Map(); // txbxContent offset -> its paragraphs
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
    // What the page needs to draw a top-level table's cell as the file draws
    // it: the grid and width (once per table, carried by each paragraph in
    // it), a cell's span, a row's height. A nested table draws flat.
    const tableMeta = () => {
      if (stack.length !== 3 || stack[0].tag !== 'tbl') return {};
      const [tbl, tr, tc] = stack;
      return {
        ...(tbl.gridPx ? { gridPx: tbl.gridPx, tableWidth: tbl.tableWidth } : {}),
        ...(tbl.look ? { tableLook: tbl.look } : {}),
        ...(tc.vAlign ? { cellVAlign: tc.vAlign } : {}),
        ...(tc.span > 1 ? { cellSpan: tc.span } : {}),
        ...(tr.heightPx ? { rowHeightPx: tr.heightPx, rowRule: tr.rule } : {}),
      };
    };

    let m;
    while ((m = re.exec(body))) {
      const tag = m[0];
      if (m[1] || m[3]) { // open/self-close or close of tbl|tr|tc|sdt|txbxContent
        const name = m[1] ?? m[3];
        if (name === 'txbxContent') {
          if (m[1] && m[2] !== '/') {
            if (boxDepth === 0 && !openBox && pStart >= 0 && stack.length === 0 && sdtDepth === 0 && !inFallback(m.index)) {
              const close = body.indexOf('</w:txbxContent>', m.index);
              const inner = close >= 0 ? body.slice(m.index + tag.length, close) : '';
              if (close >= 0 && !/<w:(?:tbl|sdt|txbxContent)\b/.test(inner)) {
                openBox = { content: m.index, pStart: -1, paras: [] };
                continue;
              }
            }
            boxDepth += 1;
          } else if (m[3]) {
            if (openBox && boxDepth === 0) { pendingBoxes.push(openBox); openBox = null; continue; }
            boxDepth = Math.max(0, boxDepth - 1);
          }
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
          if (name === 'tbl') stack.push({ tag: 'tbl', id: m.index, nextRow: 0, ...tableHead(body, m.index) });
          else if (name === 'tr') {
            const top = stack[stack.length - 1];
            stack.push({ tag: 'tr', index: top?.tag === 'tbl' ? top.nextRow++ : 0, nextCell: 0, ...rowHead(body, m.index) });
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
            const span = /<w:gridSpan\b[^>]*\bw:val="(\d+)"/.exec(head);
            const vAlign = /<w:vAlign\b[^>]*\bw:val="(center|bottom)"/.exec(head);
            stack.push({ tag: 'tc', index: top?.tag === 'tr' ? top.nextCell++ : 0, hidden, span: span ? Number(span[1]) : 1, ...(vAlign ? { vAlign: vAlign[1] } : {}) });
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
      if (openBox) {
        if (/^<w:p\b/.test(tag)) {
          if (tag.endsWith('/>')) openBox.paras.push({ xml: tag, start: m.index, end: m.index + tag.length, text: '' });
          else openBox.pStart = m.index;
        } else if (tag === '</w:p>' && openBox.pStart >= 0) {
          const xml = body.slice(openBox.pStart, m.index + tag.length);
          openBox.paras.push({ xml, start: openBox.pStart, end: m.index + tag.length, text: textOf(xml) });
          openBox.pStart = -1;
        }
        continue;
      }
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
          out.push({ index: out.length, xml: tag, start: m.index, end: m.index + tag.length, text: '', container: key, ...tableMeta(), ...(hiddenCell ? { hiddenCell } : {}), ...(inSdt ? { inSdt } : {}) });
        } else {
          pStart = m.index;
        }
        continue;
      }
      if (tag === '</w:p>' && pStart >= 0) {
        const xml = body.slice(pStart, m.index + tag.length);
        const own = xml.includes('<w:txbxContent') ? stripTextBoxes(xml) : xml;
        out.push({ index: out.length, xml, start: pStart, end: m.index + tag.length, text: textOf(own), container: key, ...tableMeta(), ...(hiddenCell ? { hiddenCell } : {}), ...(inSdt ? { inSdt } : {}) });
        for (const b of pendingBoxes) boxParas.push({ ...b, anchor: out.length - 1 });
        pendingBoxes.length = 0;
        pStart = -1;
      }
    }

    // The text boxes' story, after the body: each box's paragraphs in order,
    // each knowing its box (the offset of its content) and its anchor.
    for (const b of boxParas) {
      const indices = [];
      for (const p of b.paras) {
        indices.push(out.length);
        out.push({ index: out.length, ...p, container: 'x' + b.content, box: { content: b.content, anchor: b.anchor } });
      }
      boxIndex.set(b.content, indices);
    }

    this._editParagraphsFor = this.xml;
    this._editParagraphs = out;
    this._boxIndex = boxIndex;
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

  /** Page size and margins from `w:sectPr`, in CSS pixels — and the page colour. */
  section() {
    const { body } = this._body();
    // The first section, exactly as it always read — unless an envelope was
    // added in front of the letter (see `mainSectPr`): then the letter's.
    const main = mainSectPr(body);
    const first = /<w:sectPr\b/.exec(body);
    if (!main || !first || main.index === first.index) return { ...parseSection(body), background: this.pageColour() };
    return { ...parseSection(main[0]), evenAndOdd: /<w:evenAndOddHeaders\b[^>]*\/?>/.test(body), background: this.pageColour() };
  }

  /**
   * Every section, in order: where it ends (the edit-space index of its last
   * paragraph — the one carrying its `w:sectPr` — or null for the last
   * section, which the body's own `w:sectPr` closes), the break that starts
   * the one after it, and its page. A merged letter has one per record; a
   * letter with an envelope added has the envelope's first.
   */
  sections() {
    const out = [];
    for (const p of this.editParagraphs()) {
      if (p.container !== null) continue;
      const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/.exec(p.xml);
      const brk = paragraphSectionBreak(pPr ? pPr[0] : null);
      if (brk) out.push({ endsAt: p.index, type: brk.type, ...pageOf(brk.sectPr), ...sectionKind(brk.sectPr), sectPrXml: brk.sectPr });
    }
    // The last section is closed by the body's own `w:sectPr`, last in it.
    const { body } = this._body();
    const all = [...body.matchAll(/<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)];
    const last = all.length ? all[all.length - 1][0] : '';
    out.push({ endsAt: null, type: null, ...pageOf(last), ...sectionKind(last), sectPrXml: last });
    return out;
  }

  // ---- mailings: envelopes and labels ---------------------------------------

  /** Word's Envelope Address and Envelope Return styles, written once where the file has neither. */
  _ensureEnvelopeStyles() {
    this.ensureParagraphStyles();
    const part = 'word/styles.xml';
    const xml = this.pkg.text(part);
    if (/<w:style\b[^>]*\bw:styleId="EnvelopeAddress"/.test(xml)) return;
    this.pkg.write_(part, xml.replace('</w:styles>', ENVELOPE_STYLES_XML + '</w:styles>'));
  }

  /**
   * The envelope Mailings → Envelopes added in front of the letter, read
   * back — its size, the delivery and return addresses as lines, and the
   * last paragraph of its section — or null when there is none.
   */
  envelope() {
    const sections = this.sections();
    if (sections.length < 2 || !sections[0].envelope) return null;
    const first = sections[0];
    const paras = this.editParagraphs().filter((p) => p.container === null && p.index <= first.endsAt);
    const linesOf = (style) => paras.filter((p) => new RegExp('<w:pStyle\\b[^>]*w:val="' + style + '"').test(p.xml)).flatMap((p) => textOf(p.xml).split('\n')).filter((l) => l.trim());
    return {
      size: ENVELOPE_SIZES.find((s) => s.code === first.code)?.id ?? null,
      widthPx: first.widthPx, heightPx: first.heightPx,
      delivery: linesOf('EnvelopeAddress'),
      returnAddress: linesOf('EnvelopeReturn'),
      endsAt: first.endsAt,
    };
  }

  /**
   * Envelopes → Add to Document (Change Document when there is one): the
   * envelope as the document's first section — see labels.js
   * `envelopeXml` — in place of the one already there. The letter after it
   * keeps its own page, and its page numbers start at 1 again.
   */
  addEnvelope(spec = {}) {
    const size = ENVELOPE_SIZES.find((s) => s.id === spec.size) || ENVELOPE_SIZES[0];
    this._ensureEnvelopeStyles();
    const xml = envelopeXml({ ...spec, size });
    const had = this.envelope();
    const end = had ? this.editParagraphs()[had.endsAt].end : 0;
    this._spliceBody(0, end, xml);
    return this;
  }

  /** Start Mail Merge → Envelopes: the whole document one envelope — see labels.js `envelopeDocumentXml`. */
  setEnvelopeDocument(spec = {}) {
    const size = ENVELOPE_SIZES.find((s) => s.id === spec.size) || ENVELOPE_SIZES[0];
    this._ensureEnvelopeStyles();
    const { body, sectPr } = envelopeDocumentXml({ ...spec, size });
    const { prefix, suffix } = this._body();
    this.xml = prefix + body + sectPr + suffix;
    this.dirty = true;
    return this;
  }

  /**
   * Labels → New Document, and Start Mail Merge → Labels: the whole body
   * becomes a sheet of labels (labels.js `labelSheetXml`) on the product's
   * own paper. `mode`: 'full' — every label the same `lines`; 'single' —
   * only the label at `row`, `col` (1-based); 'merge' — the first label
   * empty for the fields, every other one starting with «Next Record».
   */
  setLabelSheet({ product, lines = [], mode = 'full', row = 1, col = 1, font = null } = {}) {
    const p = labelProduct(product);
    const cell = (r, c) => {
      if (mode === 'merge') return r === 0 && c === 0 ? labelParagraph([], font) : nextRecordParagraph();
      if (mode === 'single') return r === row - 1 && c === col - 1 ? labelParagraph(lines, font) : null;
      return labelParagraph(lines, font);
    };
    const { body, sectPr } = labelSheetXml(p, cell);
    const { prefix, suffix } = this._body();
    this.xml = prefix + body + sectPr + suffix;
    this.dirty = true;
    return this;
  }

  /**
   * Mailings → Update Labels: what the first label on the sheet holds — its
   * fields, typically an Address Block — copied into every other label,
   * after that label's «Next Record», so each takes the next recipient.
   * The narrow cells between labels are left alone. Answers how many labels
   * it filled.
   */
  updateLabels() {
    const { body } = this._body();
    const start = body.indexOf('<w:tbl>') >= 0 ? body.indexOf('<w:tbl>') : body.search(/<w:tbl\b/);
    if (start < 0) throw new Error('There is no sheet of labels in this document. Start Mail Merge → Labels makes one.');
    const end = body.indexOf('</w:tbl>', start) + '</w:tbl>'.length;
    const table = body.slice(start, end);
    const cells = [...table.matchAll(/<w:tc>(<w:tcPr>[\s\S]*?<\/w:tcPr>)([\s\S]*?)<\/w:tc>/g)];
    const width = (tcPr) => Number(/<w:tcW\b[^>]*\bw:w="(\d+)"/.exec(tcPr)?.[1] || 0);
    const widest = Math.max(...cells.map((c) => width(c[1])));
    const labels = cells.filter((c) => width(c[1]) >= widest * 0.6);
    if (labels.length < 2) return 0;
    // The first label's content, without a «Next Record» of its own, then
    // with one in front of it for every other label.
    const content = labels[0][2].split(NEXT_FIELD_RUNS).join('');
    const withNext = content.replace(/^(\s*<w:p\b[^>]*>(?:<w:pPr>[\s\S]*?<\/w:pPr>)?)/, (m) => m + NEXT_FIELD_RUNS);
    let out = '';
    let at = 0;
    let filled = 0;
    for (const c of labels.slice(1)) {
      out += table.slice(at, c.index) + '<w:tc>' + c[1] + withNext + '</w:tc>';
      at = c.index + c[0].length;
      filled += 1;
    }
    out += table.slice(at);
    this._spliceBody(start, end, out);
    return filled;
  }

  /**
   * Start offsets (in the body) of the paragraphs that open a new page
   * because the paragraph before them ends a section — Word's Next Page,
   * Odd Page and Even Page breaks; a Continuous one starts no page.
   */
  _sectionStarts() {
    if (this._sectionStartsFor === this.xml) return this._sectionStartsSet;
    const set = new Set();
    const list = this.editParagraphs().filter((p) => p.container === null);
    for (let i = 0; i + 1 < list.length; i++) {
      const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/.exec(list[i].xml);
      const brk = paragraphSectionBreak(pPr ? pPr[0] : null);
      if (brk && brk.type !== 'continuous') set.add(list[i + 1].start);
    }
    this._sectionStartsFor = this.xml;
    this._sectionStartsSet = set;
    return set;
  }

  /** The page colour, `<w:background w:color="RRGGBB"/>` before the body, as '#RRGGBB' or null. */
  pageColour() {
    const head = this.xml.slice(0, Math.max(0, this.xml.indexOf('<w:body')));
    const m = /<w:background\b[^>]*\bw:color="([0-9A-Fa-f]{6})"/.exec(head);
    return m ? '#' + m[1].toUpperCase() : null;
  }

  /**
   * Set or clear the page colour: `<w:background>` as the document's first
   * child, and `<w:displayBackgroundShape/>` in the settings so Word shows
   * it (without that, Word keeps the colour and draws a white page).
   */
  /**
   * Declare Office Math's namespace on `<w:document>`, as Word does on every
   * document, before an equation is written into one that lacks it — an
   * `m:` element with no `xmlns:m` in scope is a file Word refuses.
   */
  ensureMathNamespace() {
    const open = /<w:document\b[^>]*>/.exec(this.xml);
    if (!open || /\bxmlns:m=/.test(open[0])) return false;
    const tag = open[0].replace(/^<w:document\b/, '<w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    this.xml = this.xml.slice(0, open.index) + tag + this.xml.slice(open.index + open[0].length);
    this.dirty = true;
    return true;
  }

  setPageColour(colour) {
    const open = /<w:document\b[^>]*>/.exec(this.xml);
    if (!open) throw new Error('document has no <w:document>');
    const withoutOld = this.xml.replace(/<w:background\b[^>]*(?:\/>|>[\s\S]*?<\/w:background>)/, '');
    if (colour == null) {
      this.xml = withoutOld;
      this.dirty = true;
      return this;
    }
    const hex = String(colour).replace('#', '').toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(hex)) throw new Error('a page colour is six hex digits: ' + colour);
    const at = /<w:document\b[^>]*>/.exec(withoutOld);
    this.xml = withoutOld.slice(0, at.index + at[0].length) + '<w:background w:color="' + hex + '"/>' + withoutOld.slice(at.index + at[0].length);
    this.dirty = true;

    const settingsPart = 'word/settings.xml';
    if (this.pkg.has(settingsPart)) {
      const settings = this.pkg.text(settingsPart);
      if (!/<w:displayBackgroundShape\b/.test(settings)) {
        this.pkg.write_(settingsPart, Buffer.from(settings.replace(/(<w:settings\b[^>]*>)/, '$1<w:displayBackgroundShape/>'), 'utf8'));
      }
    } else {
      this.pkg.addPart(settingsPart, Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:displayBackgroundShape/></w:settings>', 'utf8'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
      this.pkg.addRelationshipTo(this.mainPart, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
    }
    return this;
  }

  /**
   * Mailings: the mail merge this document is the main document of, read
   * from `w:mailMerge` in settings.xml — the kind of document it is
   * (`formLetters`, `email`, `envelopes`, `mailingLabels`, `catalog`), the
   * data source it is attached to and how Word reaches it, the record last
   * previewed. Null for an ordinary document.
   *
   * `path` is the source file, from the `mailMergeSource` relationship
   * Word writes (`file:///C:\…\list.csv`), else from the query or the
   * connection string; `sheet` is the worksheet an Excel source names in
   * its query (`SELECT * FROM \`Sheet1$\``).
   */
  mailMerge() {
    const part = 'word/settings.xml';
    if (!this.pkg.has(part)) return null;
    const m = /<w:mailMerge\b[^>]*>([\s\S]*?)<\/w:mailMerge>/.exec(this.pkg.text(part));
    if (!m) return null;
    const inner = m[1];
    const val = (name) => {
      const r = new RegExp('<w:' + name + '\\b[^>]*\\bw:val="([^"]*)"').exec(inner);
      return r ? unesc(r[1]) : null;
    };
    const rid = /<w:dataSource\b[^>]*\br:id="([^"]+)"/.exec(inner);
    const rel = rid ? this.pkg.rels(part).find((r) => r.Id === rid[1]) : null;
    const query = val('query');
    const connect = val('connectString');
    const fromTarget = rel?.Target ? sourcePathOf(unesc(rel.Target)) : null;
    const fromQuery = query ? (/\bFROM\s+(?!`)(.+?)\s*$/i.exec(query)?.[1] ?? null) : null;
    const fromConnect = connect ? (/Data Source=([^;]+)/i.exec(connect)?.[1] ?? null) : null;
    const sheet = query ? (/`([^`]+?)\$?`/.exec(query)?.[1] ?? null) : null;
    return {
      type: val('mainDocumentType') || 'formLetters',
      dataType: val('dataType'),
      connectString: connect,
      query,
      path: fromTarget || fromConnect || fromQuery || null,
      sheet: sheet && !/^Rutba Contacts$/i.test(sheet) ? sheet : null,
      contacts: Boolean(query && /`Rutba Contacts`/i.test(query)),
      destination: val('destination'),
      addressField: val('addressFieldName'),
      subject: val('mailSubject'),
      viewMergedData: /<w:viewMergedData\b(?![^>]*w:val="(?:0|false)")/.test(inner),
      activeRecord: Number(val('activeRecord')) || null,
    };
  }

  /**
   * Write `w:mailMerge` as Word writes it — or, with null (Start Mail Merge
   * → Normal Word Document), take it away with its data source link.
   *
   * `spec`: `{ type, source: { kind: 'csv'|'tsv'|'xlsx'|'contacts', path,
   * sheet }, destination, addressField, subject, viewMergedData,
   * activeRecord }`. A text file is `textFile` queried by its path; a
   * workbook is reached the way Word reaches one, through the ACE OLE DB
   * provider, with the sheet in the query; the address book has no file,
   * and is named in the query for this suite to find again. The source file
   * is also a `mailMergeSource` relationship from settings.xml, external,
   * which is what Word follows to reattach the list when the letter opens.
   * Only the `w:mailMerge` element and that one relationship are touched:
   * every other setting rides through as it was.
   */
  setMailMerge(spec) {
    const part = 'word/settings.xml';
    const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource';
    const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    if (!this.pkg.has(part)) {
      if (!spec) return this;
      this.pkg.addPart(part, Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="' + WORD_NS + '" xmlns:r="' + R_NS + '"></w:settings>', 'utf8'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
      this.pkg.addRelationshipTo(this.mainPart, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
    }
    let settings = this.pkg.text(part);
    // The old element and its source link go first, whatever replaces them.
    const old = /<w:mailMerge\b[^>]*>[\s\S]*?<\/w:mailMerge>|<w:mailMerge\b[^>]*\/>/.exec(settings);
    if (old) {
      const rid = /<w:dataSource\b[^>]*\br:id="([^"]+)"/.exec(old[0]);
      settings = settings.slice(0, old.index) + settings.slice(old.index + old[0].length);
      if (rid) this._removeRel(part, rid[1]);
    }
    if (spec) {
      const v = (name, value) => (value == null || value === '' ? '' : '<w:' + name + ' w:val="' + esc(String(value)) + '"/>');
      const src = spec.source || null;
      let dataType = '';
      let connect = '';
      let query = '';
      let dataSource = '';
      if (src?.kind === 'contacts') {
        dataType = 'native';
        query = 'SELECT * FROM `Rutba Contacts` ';
      } else if (src?.path) {
        const winPath = String(src.path);
        if (src.kind === 'xlsx') {
          dataType = 'native';
          connect = 'Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=' + winPath + ';Mode=Read;Extended Properties="HDR=YES;IMEX=1;";';
          query = 'SELECT * FROM `' + (src.sheet || 'Sheet1') + '$` ';
        } else {
          dataType = 'textFile';
          query = 'SELECT * FROM ' + winPath;
        }
        const target = 'file:///' + winPath.replace(/^\/+/, '');
        const rId = this.pkg.addRelationshipTo(part, REL, target, { external: true });
        dataSource = '<w:dataSource r:id="' + rId + '"/>';
      }
      const element = '<w:mailMerge>' +
        v('mainDocumentType', spec.type || 'formLetters') +
        (src ? '<w:linkToQuery/>' : '') +
        v('dataType', dataType) +
        (src ? '<w:connectString w:val="' + esc(connect) + '"/>' : '') +
        v('query', query) +
        dataSource +
        v('destination', spec.destination) +
        v('addressFieldName', spec.addressField) +
        v('mailSubject', spec.subject) +
        (spec.viewMergedData ? '<w:viewMergedData/>' : '') +
        v('activeRecord', spec.activeRecord) +
        '</w:mailMerge>';
      // Where the schema puts it: after the document-wide switches Word
      // writes first, before revision tracking and the default tab stop.
      const later = /<w:(?:revisionView|trackRevisions|doNotTrackMoves|doNotTrackFormatting|documentProtection|autoFormatOverride|styleLockTheme|styleLockQFSet|defaultTabStop|autoHyphenation|consecutiveHyphenLimit|hyphenationZone|doNotHyphenateCaps|showEnvelope|summaryLength|clickAndTypeStyle|defaultTableStyle|evenAndOddHeaders|bookFoldRevPrinting|bookFoldPrinting|bookFoldPrintingSheets|drawingGrid\w*|displayHorizontalDrawingGridEvery|displayVerticalDrawingGridEvery|doNotUseMarginsForDrawingGridOrigin|doNotShadeFormData|noPunctuationKerning|characterSpacingControl|printTwoOnOne|strictFirstAndLastChars|noLineBreaksAfter|noLineBreaksBefore|savePreviewPicture|doNotValidateAgainstSchema|saveInvalidXml|ignoreMixedContent|alwaysShowPlaceholderText|doNotDemarcateInvalidXml|saveXmlDataOnly|useXSLTWhenSaving|saveThroughXslt|showXMLTags|alwaysMergeEmptyNamespace|updateFields|hdrShapeDefaults|footnotePr|endnotePr|compat|docVars|rsids|mathPr|attachedSchema|themeFontLang|clrSchemeMapping|doNotIncludeSubdocsInStats|doNotAutoCompressPictures|forceUpgrade|captions|readModeInkLockDown|smartTagType|schemaLibrary|shapeDefaults|doNotEmbedSmartTags|decimalSymbol|listSeparator)\b|<w1[45]:|<\/w:settings>/.exec(settings);
      settings = settings.slice(0, later.index) + element + settings.slice(later.index);
      // The r: prefix must be in scope for the data source reference.
      if (dataSource && !/<w:settings\b[^>]*\bxmlns:r=/.test(settings)) {
        settings = settings.replace(/<w:settings\b/, '<w:settings xmlns:r="' + R_NS + '"');
      }
    }
    this.pkg.write_(part, settings);
    this.dirty = true;
    return this;
  }

  /** Take one relationship off a part's .rels, by its id. */
  _removeRel(fromPart, id) {
    const relsPath = OoxmlPackage.relsPathFor(fromPart);
    if (!this.pkg.has(relsPath)) return;
    const xml = this.pkg.text(relsPath);
    const next = xml.replace(new RegExp('<Relationship\\b[^>]*\\bId="' + id + '"[^>]*/>'), '');
    if (next !== xml) this.pkg.write_(relsPath, next);
  }

  /** A fresh `w:id` for the next `w:ins`/`w:del` — one past the highest either kind already carries. */
  nextTrackChangeId() {
    let maxId = -1;
    for (const m of this.xml.matchAll(/<w:(?:ins|del)\b[^>]*\bw:id="(\d+)"/g)) maxId = Math.max(maxId, Number(m[1]));
    return maxId + 1;
  }

  /** Review → Track Changes: is this document recording, right now — `<w:trackRevisions/>` in settings.xml. */
  trackRevisions() {
    const part = 'word/settings.xml';
    if (!this.pkg.has(part)) return false;
    return /<w:trackRevisions\b/.test(this.pkg.text(part));
  }

  /**
   * Turn recording on or off. Same append-or-create shape as `setPageColour`'s
   * `w:displayBackgroundShape`: an existing settings part is only ever added
   * to or trimmed, never rewritten whole, so a template's other settings
   * (compatibility flags, the default tab stop) ride through untouched.
   */
  setTrackRevisions(on) {
    const part = 'word/settings.xml';
    if (!this.pkg.has(part)) {
      if (!on) return this;
      this.pkg.addPart(part, Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:trackRevisions/></w:settings>', 'utf8'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
      this.pkg.addRelationshipTo(this.mainPart, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
      this.dirty = true;
      return this;
    }
    const settings = this.pkg.text(part);
    const has = /<w:trackRevisions\b[^>]*\/>/.test(settings);
    if (on && !has) {
      this.pkg.write_(part, settings.replace(/(<w:settings\b[^>]*>)/, '$1<w:trackRevisions/>'));
      this.dirty = true;
    } else if (!on && has) {
      this.pkg.write_(part, settings.replace(/<w:trackRevisions\b[^>]*\/>/, ''));
      this.dirty = true;
    }
    return this;
  }

  /**
   * Pictures embedded in one paragraph, as data URIs — with, for a picture
   * in a `wp:anchor`, where it floats and how the text treats it (see
   * `anchorLayout`), which is what lets the page wrap words round it.
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
      // A group's pictures are the group's — `_paragraphGroups` draws them.
      if (/<a:graphicData\b[^>]*wordprocessingGroup/.test(inner)) continue;
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
        ...anchorLayout(inner),
        ...arrangeOf(m[0]),
      });
    }
    return out;
  }

  /**
   * A picture's size: the drawing's extent and the picture's own transform,
   * both in EMU, both from the pixels the page measured. Everything else
   * about the drawing — where it sits, how the words treat it — is untouched.
   */
  setImageSize(index, imageIndex, { widthPx, heightPx }) {
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    const w = Math.round(Number(widthPx));
    const h = Math.round(Number(heightPx));
    if (!(w > 0) || !(h > 0)) throw new Error('a picture needs a positive width and height');
    const drawings = [...p.xml.matchAll(/<w:drawing\b[^>]*>[\s\S]*?<\/w:drawing>/g)].filter((m) => /<a:blip\b/.test(m[0]));
    const d = drawings[imageIndex];
    if (!d) throw new Error('no picture ' + imageIndex + ' in paragraph ' + index);
    const cx = w * PX_TO_EMU;
    const cy = h * PX_TO_EMU;
    let drawing = d[0].replace(/<wp:extent\b[^>]*\/>/, '<wp:extent cx="' + cx + '" cy="' + cy + '"/>');
    // The picture's own transform: the first a:ext inside pic:spPr.
    drawing = drawing.replace(/(<pic:spPr\b[\s\S]*?<a:ext\b)[^>]*(\/>)/, '$1 cx="' + cx + '" cy="' + cy + '"$2');
    const xml = p.xml.slice(0, d.index) + drawing + p.xml.slice(d.index + d[0].length);
    this._spliceBody(p.start, p.end, xml);
    return this;
  }

  /**
   * How a picture sits in its paragraph: in the line, or floating with the
   * text wrapping round it, or behind or in front of it. Rewrites the
   * picture's drawing between `wp:inline` and `wp:anchor` in place; the
   * picture itself — the blip, the size, the name — is untouched, so the
   * file round-trips whichever way it went. `wrap` is one of inline,
   * square, tight, topAndBottom, behind, front; `hAlign` left, center or
   * right, relative to the margins.
   */
  setImageLayout(index, imageIndex, { wrap = 'inline', hAlign = 'left' } = {}) {
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    const drawings = [...p.xml.matchAll(/<w:drawing\b[^>]*>[\s\S]*?<\/w:drawing>/g)].filter((m) => /<a:blip\b/.test(m[0]));
    const d = drawings[imageIndex];
    if (!d) throw new Error('no picture ' + imageIndex + ' in paragraph ' + index);
    const inner = d[0];
    const extent = /<wp:extent\b[^>]*\/>/.exec(inner)?.[0] ?? '<wp:extent cx="914400" cy="914400"/>';
    const docPr = /<wp:docPr\b[^>]*\/>|<wp:docPr\b[^>]*>[\s\S]*?<\/wp:docPr>/.exec(inner)?.[0] ?? '<wp:docPr id="1" name="Picture"/>';
    const graphic = /<a:graphic\b[\s\S]*<\/a:graphic>/.exec(inner)?.[0];
    if (!graphic) throw new Error('the picture has no graphic');
    const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
    let drawing;
    if (wrap === 'inline') {
      drawing = '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" ' + WP + '>' + extent + docPr + '<wp:cNvGraphicFramePr/>' + graphic + '</wp:inline></w:drawing>';
    } else {
      const align = { left: 'left', center: 'center', centre: 'center', right: 'right' }[hAlign] || 'left';
      // Word wants a wrap polygon with wrapTight; a square is what it draws
      // for a rectangular picture anyway.
      const wrapXml = wrap === 'square' || wrap === 'tight' ? '<wp:wrapSquare wrapText="bothSides"/>'
        : wrap === 'topAndBottom' ? '<wp:wrapTopAndBottom/>'
          : '<wp:wrapNone/>';
      const behind = wrap === 'behind' ? '1' : '0';
      drawing = '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="' + behind + '" locked="0" layoutInCell="1" allowOverlap="1" ' + WP + '>'
        + '<wp:simplePos x="0" y="0"/>'
        + '<wp:positionH relativeFrom="margin"><wp:align>' + align + '</wp:align></wp:positionH>'
        + '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        + extent + '<wp:effectExtent l="0" t="0" r="0" b="0"/>' + wrapXml + docPr + '<wp:cNvGraphicFramePr/>' + graphic
        + '</wp:anchor></w:drawing>';
    }
    const xml = p.xml.slice(0, d.index) + drawing + p.xml.slice(d.index + inner.length);
    this._spliceBody(p.start, p.end, xml);
    return this;
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
   * @returns {{bullet: string, number: string, outline: string}} the numIds to write into numPr
   */
  ensureListNumbering() {
    const part = 'word/numbering.xml';
    const has = this.pkg.has(part);
    const xml = has ? this.pkg.text(part) : null;
    const found = this._classifyNumbering(xml);

    let bulletId = found.bullet;
    let numberId = found.number;
    let outlineId = found.outline;
    if (bulletId && numberId && outlineId) return { bullet: bulletId, number: numberId, outline: outlineId };

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
    if (!outlineId) {
      const aId = nextAbstract++;
      const nId = nextNum++;
      abstracts.push(outlineAbstract(aId));
      nums.push(numDef(nId, aId));
      outlineId = String(nId);
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
    return { bullet: bulletId, number: numberId, outline: outlineId };
  }

  /**
   * Which numIds already carry a bullet or a decimal level-0 definition, plus the
   * highest ids in use — everything `ensureListNumbering` needs to decide whether
   * to reuse or to mint fresh ids that cannot collide.
   */
  _classifyNumbering(xml) {
    const out = { bullet: null, number: null, outline: null, maxNum: 0, maxAbstract: -1 };
    if (!xml) return out;

    // A `<w:num>` points at a `<w:abstractNum>`; the abstract's level-0 numFmt
    // (its first, since level 0 is written first) is what says bullet vs number,
    // and a second level whose text carries the first's number — "%1.%2." —
    // is what says multilevel.
    const absFmt = new Map();
    const absOutline = new Set();
    for (const m of xml.matchAll(/<w:abstractNum\b([^>]*)>([\s\S]*?)<\/w:abstractNum>/g)) {
      const id = attrs(m[1])['w:abstractNumId'];
      if (id !== undefined) out.maxAbstract = Math.max(out.maxAbstract, Number(id));
      const fmt = /<w:numFmt\b[^>]*\bw:val="([^"]*)"/.exec(m[2]);
      absFmt.set(String(id), fmt ? fmt[1] : 'decimal');
      const lvl1 = /<w:lvl w:ilvl="1"[^>]*>([\s\S]*?)<\/w:lvl>/.exec(m[2]);
      if (lvl1 && /<w:lvlText\b[^>]*\bw:val="%1\.%2\.?"/.test(lvl1[1])) absOutline.add(String(id));
    }
    for (const m of xml.matchAll(/<w:num\b([^>]*)>([\s\S]*?)<\/w:num>/g)) {
      const numId = attrs(m[1])['w:numId'];
      if (numId === undefined) continue;
      out.maxNum = Math.max(out.maxNum, Number(numId));
      const absId = /<w:abstractNumId\b[^>]*\bw:val="([^"]*)"/.exec(m[2]);
      const fmt = absId ? absFmt.get(absId[1]) : undefined;
      if (fmt === 'bullet') { if (!out.bullet) out.bullet = String(numId); }
      else if (fmt === 'decimal' && absId && absOutline.has(absId[1])) { if (!out.outline) out.outline = String(numId); }
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
   * The "Caption" style, written once — Word's own on a document that never
   * had one: italic, 9pt, a little space after so two captions in a row do
   * not run together. `ensureParagraphStyles` above only ever writes the
   * WHOLE part fresh, for a document that has none at all; this instead adds
   * one STYLE to an existing styles.xml, the way `ensureListNumbering` adds
   * one numbering definition to an existing numbering.xml rather than
   * replacing the file — an existing catalogue is a template author's own
   * and is never rewritten, only ever added to.
   */
  _ensureCaptionStyle() {
    this.ensureParagraphStyles();
    const part = 'word/styles.xml';
    const xml = this.pkg.text(part);
    if (/<w:style\b[^>]*\bw:styleId="Caption"/.test(xml)) return;
    const style = '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/>'
      + '<w:pPr><w:spacing w:after="200"/></w:pPr><w:rPr><w:i/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>';
    this.pkg.write_(part, xml.replace('</w:styles>', style + '</w:styles>'));
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
  /**
   * Set or clear the page borders: `<w:pgBorders>` in the body's
   * `w:sectPr`, where the schema puts it — after the margins, before the
   * line numbering, page numbering and columns. Sides come as
   * { top, left, bottom, right } of { style, widthPx, colour, spacePt } and
   * `offsetFrom` 'page' (the default, Word's own: 24 pt in from the edge)
   * or 'text' (out from the text by the space). null, or no sides, takes
   * the element off.
   */
  setPageBorders(borders) {
    const { prefix, body, suffix } = this._body();
    const at = mainSectPr(body);
    if (!at) throw new Error('this document has no section properties to border');
    let sectPr = at[0];
    if (/^<w:sectPr\b[^>]*\/>$/.test(sectPr)) sectPr = sectPr.replace(/\/>$/, '>') + '</w:sectPr>';
    sectPr = sectPr.replace(/<w:pgBorders\b[^>]*\/>|<w:pgBorders\b[^>]*>[\s\S]*?<\/w:pgBorders>/, '');
    const sides = ['top', 'left', 'bottom', 'right'].filter((s) => borders && borders[s]);
    if (sides.length) {
      const offsetFrom = borders.offsetFrom === 'text' ? 'text' : 'page';
      const xml = '<w:pgBorders w:offsetFrom="' + offsetFrom + '">' + sides.map((s) => {
        const b = borders[s];
        const sz = Math.max(2, Math.round((Number(b.widthPx) || 1) * 6));
        const colour = b.colour ? String(b.colour).replace('#', '').toUpperCase() : 'auto';
        const space = b.spacePt === undefined ? (offsetFrom === 'page' ? 24 : 4) : Math.max(0, Math.round(Number(b.spacePt) || 0));
        return '<w:' + s + ' w:val="' + (b.style || 'single') + '" w:sz="' + sz + '" w:space="' + space + '" w:color="' + colour + '"/>';
      }).join('') + '</w:pgBorders>';
      // After the last child the schema puts before it; failing one, before
      // the first it puts after; failing both, at the end.
      let before = -1;
      for (const m of sectPr.matchAll(/<w:(?:headerReference|footerReference|footnotePr|endnotePr|type|pgSz|pgMar|paperSrc)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:(?:footnotePr|endnotePr)>)/g)) {
        before = m.index + m[0].length;
      }
      const later = /<w:(?:lnNumType|pgNumType|cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\b/.exec(sectPr);
      sectPr = before >= 0
        ? sectPr.slice(0, before) + xml + sectPr.slice(before)
        : later
          ? sectPr.slice(0, later.index) + xml + sectPr.slice(later.index)
          : sectPr.replace('</w:sectPr>', xml + '</w:sectPr>');
    }
    this.xml = prefix + body.slice(0, at.index) + sectPr + body.slice(at.index + at[0].length) + suffix;
    this.dirty = true;
    return this;
  }

  /**
   * Line numbering — Layout → Line Numbers: `<w:lnNumType>` in the section,
   * after the page borders where the schema puts it. { countBy, restart
   * ('continuous' | 'newPage' | 'newSection'), start, distancePx }; null
   * takes it off.
   */
  setLineNumbers(spec) {
    const { prefix, body, suffix } = this._body();
    const at = mainSectPr(body);
    if (!at) throw new Error('this document has no section properties to number');
    let sectPr = at[0];
    if (/^<w:sectPr\b[^>]*\/>$/.test(sectPr)) sectPr = sectPr.replace(/\/>$/, '>') + '</w:sectPr>';
    sectPr = sectPr.replace(/<w:lnNumType\b[^>]*\/>|<w:lnNumType\b[^>]*>[\s\S]*?<\/w:lnNumType>/, '');
    if (spec) {
      const countBy = Math.max(1, Math.round(Number(spec.countBy) || 1));
      const restart = ['continuous', 'newPage', 'newSection'].includes(spec.restart) ? spec.restart : 'continuous';
      const start = Math.max(1, Math.round(Number(spec.start) || 1));
      const distance = spec.distancePx != null ? ' w:distance="' + Math.max(0, Math.round(Number(spec.distancePx) * 15)) + '"' : '';
      const xml = '<w:lnNumType w:countBy="' + countBy + '"' + (start !== 1 ? ' w:start="' + start + '"' : '') + distance + ' w:restart="' + restart + '"/>';
      let before = -1;
      for (const m of sectPr.matchAll(/<w:(?:headerReference|footerReference|footnotePr|endnotePr|type|pgSz|pgMar|paperSrc|pgBorders)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:(?:footnotePr|endnotePr|pgBorders)>)/g)) {
        before = m.index + m[0].length;
      }
      const later = /<w:(?:pgNumType|cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\b/.exec(sectPr);
      sectPr = before >= 0
        ? sectPr.slice(0, before) + xml + sectPr.slice(before)
        : later
          ? sectPr.slice(0, later.index) + xml + sectPr.slice(later.index)
          : sectPr.replace('</w:sectPr>', xml + '</w:sectPr>');
    }
    this.xml = prefix + body.slice(0, at.index) + sectPr + body.slice(at.index + at[0].length) + suffix;
    this.dirty = true;
    return this;
  }

  /**
   * Layout → Columns: `<w:cols>` in the section, after the page numbering
   * where the schema puts it (see the `later` regex below, which already
   * knew that before this wrote anything there). `{ count, spaceTwips =
   * 720, separator = false, widths }`; `widths` — one per column, in
   * twips — is how Word's own Left and Right presets write two columns of
   * different widths (`w:equalWidth="0"` plus a `w:col` each); leaving it
   * out shares the content width equally, the ordinary case. `count` 1, or
   * null, takes the element off — Word never writes `w:cols` for one column.
   */
  setPageSetup({ orientation, size, margins, columns } = {}) {
    const PAPER = { A4: [11906, 16838], Letter: [12240, 15840], Legal: [12240, 20160] };
    const MARGIN_PRESETS = {
      normal: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      narrow: { top: 720, right: 720, bottom: 720, left: 720 },
      wide: { top: 1440, right: 2880, bottom: 1440, left: 2880 },
    };

    const { prefix, body, suffix } = this._body();
    const at = mainSectPr(body);
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
      const current = parseSection(at[0]);
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

    if (columns !== undefined) {
      sectPr = sectPr.replace(/<w:cols\b[^>]*\/>|<w:cols\b[^>]*>[\s\S]*?<\/w:cols>/, '');
      const count = columns ? Math.max(1, Math.min(12, Math.round(Number(columns.count) || 1))) : 1;
      if (columns && count > 1) {
        const spaceTwips = Math.max(0, Math.round(Number(columns.spaceTwips ?? 720)));
        const sep = columns.separator ? ' w:sep="1"' : '';
        const widths = Array.isArray(columns.widths) && columns.widths.length === count ? columns.widths : null;
        const xml = widths
          ? '<w:cols w:num="' + count + '" w:space="' + spaceTwips + '"' + sep + ' w:equalWidth="0">' +
            widths.map((w, i) => '<w:col w:w="' + Math.max(1, Math.round(Number(w) || 0)) + '"' + (i < widths.length - 1 ? ' w:space="' + spaceTwips + '"' : '') + '/>').join('') +
            '</w:cols>'
          : '<w:cols w:num="' + count + '" w:space="' + spaceTwips + '"' + sep + '/>';
        // Before the last child the schema puts before `w:cols` (through
        // `w:pgNumType`); failing that, before the first it puts after.
        let before = -1;
        for (const m of sectPr.matchAll(/<w:(?:headerReference|footerReference|footnotePr|endnotePr|type|pgSz|pgMar|paperSrc|pgBorders|lnNumType|pgNumType)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:(?:footnotePr|endnotePr|pgBorders)>)/g)) {
          before = m.index + m[0].length;
        }
        const later = /<w:(?:formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\b/.exec(sectPr);
        sectPr = before >= 0
          ? sectPr.slice(0, before) + xml + sectPr.slice(before)
          : later
            ? sectPr.slice(0, later.index) + xml + sectPr.slice(later.index)
            : sectPr.replace('</w:sectPr>', xml + '</w:sectPr>');
      }
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
      if (/<a:graphicData\b[^>]*wordprocessingGroup/.test(inner)) continue; // a group — _paragraphGroups
      const place = { ...anchorLayout(inner), ...arrangeOf(m[0]) };
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
          out.push({ ...place, kind: 'chart', name, widthPx, heightPx, chartXml: this.pkg.text(part) });
        }
        continue;
      }
      const wsp = /<wps:wsp\b[\s\S]*?<\/wps:wsp>/.exec(inner);
      // A shape with a text box is a TEXT box — `_paragraphTextBoxes` draws
      // it with its words; painting the frame here too would double it.
      if (wsp && !/<wps:txbx\b/.test(wsp[0])) out.push({ ...place, kind: 'shape', name, widthPx, heightPx, shapeXml: wsp[0] });
    }
    return out;
  }

  /**
   * The text boxes anchored in a paragraph: each one's frame — size, fill,
   * outline, position, wrap, its insets and where its words sit — and its
   * paragraphs, decorated like the body's (style, alignment, indents,
   * shading, runs, pictures).
   *
   * A box the edit space lists (see `editParagraphs`) names its paragraphs'
   * indices in `blocks`: the page draws those, and the caret types in them.
   * `paragraphs` is still read for everyone who draws without an address —
   * the printout, a box the edit space leaves opaque.
   *
   * Only the `mc:Choice` of an AlternateContent is read; the VML fallback
   * repeats the same words for older Words and would show them twice.
   * `base` is the paragraph's offset in the body, which is how a box is
   * matched to its paragraphs in the edit space.
   */
  _paragraphTextBoxes(paragraphXml, base = null) {
    const xml = String(paragraphXml);
    if (!xml.includes('<w:txbxContent')) return [];
    const colours = this.themeColours();
    const fallbacks = [...xml.matchAll(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g)].map((f) => [f.index, f.index + f[0].length]);
    const inFallback = (at) => fallbacks.some(([a, b]) => at >= a && at < b);
    const units = findDrawings(xml).map((d) => ({ start: d.start, xml: d.xml, drawing: true }));
    for (const m of xml.matchAll(/<w:pict\b[^>]*>[\s\S]*?<\/w:pict>/g)) {
      if (!inFallback(m.index) && m[0].includes('<w:txbxContent')) units.push({ start: m.index, xml: m[0], drawing: false });
    }
    units.sort((a, b) => a.start - b.start);
    const out = [];
    for (const unit of units) {
      const inner = unit.xml;
      const open = /<w:txbxContent\b[^>]*>/.exec(inner);
      if (!open) continue;
      const d = unit.drawing ? readDrawing(inner) : null;
      if (d && d.kind === 'group') continue; // `_paragraphGroups` draws a group's boxes
      const close = inner.indexOf('</w:txbxContent>', open.index);
      const content = inner.slice(open.index + open[0].length, close < 0 ? inner.length : close);
      const extent = /<wp:extent\b([^>]*)\/>/.exec(inner);
      const ext = extent ? attrs(extent[1]) : {};
      let widthPx = ext.cx ? Math.round(Number(ext.cx) / 9525) : null;
      let heightPx = ext.cy ? Math.round(Number(ext.cy) / 9525) : null;
      if (widthPx === null) {
        // VML: <v:shape style="width:451.3pt;height:38.4pt">
        const style = /<v:(?:shape|rect)\b[^>]*\bstyle="([^"]*)"/.exec(inner);
        const dim = (name) => { const dm = new RegExp('(?:^|;)\\s*' + name + ':\\s*([\\d.]+)(pt|px|in|cm)?').exec(style?.[1] ?? ''); return dm ? Math.round(Number(dm[1]) * ({ pt: 96 / 72, px: 1, in: 96, cm: 96 / 2.54 })[dm[2] || 'pt']) : null; };
        widthPx = dim('width');
        heightPx = dim('height');
      }
      const namePr = /<wp:docPr\b([^>]*)\/?>/.exec(inner);
      const layout = anchorLayout(inner);
      const hAlign = layout.hAlign ?? null;
      const spPr = /<wps:spPr\b[^>]*>([\s\S]*?)<\/wps:spPr>/.exec(inner)?.[1] ?? '';
      const line = /<a:ln\b[^>]*>([\s\S]*?)<\/a:ln>/.exec(spPr);
      const fill = colourOf(spPr.replace(/<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/, ''), colours);
      const vmlFill = /<v:(?:shape|rect)\b[^>]*\bfillcolor="([^"]*)"/.exec(inner);
      const paragraphs = this._liteParagraphs(content);
      const body = /<wps:bodyPr\b/.test(inner) ? readBodyPr(inner) : null;
      const look = readShapeLook(inner);
      const at = base != null ? base + unit.start + open.index : null;
      const blocks = at != null ? this._boxIndex?.get(at) ?? null : null;
      out.push({
        ...(d ? { id: d.id, relativeHeight: d.relativeHeight, rot: d.rot, flipH: d.flipH, flipV: d.flipV, hidden: d.hidden } : {}),
        kind: 'textbox',
        name: namePr ? (attrs(namePr[1])['name'] ?? null) : null,
        widthPx, heightPx, hAlign,
        anchored: layout.anchored, wrap: layout.wrap ?? null, wrapSide: layout.wrapSide ?? null,
        dist: layout.dist ?? null, behind: Boolean(layout.behind),
        vRel: layout.vRel ?? null, vOffsetPx: layout.vOffsetPx ?? null, vAlign: layout.vAlign ?? null,
        hRel: layout.hRel ?? null, hOffsetPx: layout.hOffsetPx ?? null,
        fill: fill ?? (vmlFill ? vmlFill[1] : null),
        line: line ? colourOf(line[1], colours) : null,
        ...(line ? { lineWidthPx: Math.round((look.lineWidthPx ?? 1) * 100) / 100 } : {}),
        ...(body ? { insets: body.insets, vAnchor: body.anchor, vert: body.vert, autoFit: body.autoFit } : {}),
        ...(blocks && blocks.length ? { blocks } : {}),
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
    for (const m of this.xml.matchAll(/<(?:wp:docPr|wps:cNvPr|pic:cNvPr|wpg:cNvPr)\b[^>]*\bid="(\d+)"/g)) {
      max = Math.max(max, Number(m[1]));
    }
    return max + 1;
  }

  /* ── floating drawings: text boxes, Arrange, groups ─────────────────────── */

  /**
   * Declare on `<w:document>` the namespaces a floating drawing writes —
   * markup compatibility, the drawing canvases, DrawingML, VML — each only
   * if it is missing, the way Word declares all of them on every document.
   */
  ensureDrawingNamespaces() {
    const open = /<w:document\b[^>]*>/.exec(this.xml);
    if (!open) return false;
    let tag = open[0];
    for (const [prefix, uri] of Object.entries(DRAWING_NS)) {
      if (!new RegExp('\\sxmlns:' + prefix + '=').test(tag)) tag = tag.replace(/^<w:document\b/, '<w:document xmlns:' + prefix + '="' + uri + '"');
    }
    if (tag === open[0]) return false;
    this.xml = this.xml.slice(0, open.index) + tag + this.xml.slice(open.index + open[0].length);
    this.dirty = true;
    return true;
  }

  /**
   * Every drawing on the page, in document order — what the Selection Pane
   * lists and the Arrange commands act on: its id and name, what it is, the
   * paragraph that holds it, whether it floats and how, its place in the
   * z-order, its size and its turn, and a text box's first words.
   */
  drawings() {
    const { body } = this._body();
    if (!body.includes('<w:drawing')) return [];
    const paras = this.editParagraphs().filter((p) => !p.box);
    const out = [];
    let k = 0;
    for (const d of findDrawings(body)) {
      while (k < paras.length && paras[k].end <= d.start) k += 1;
      // The innermost paragraph holding it: body paragraphs and cell
      // paragraphs are listed in document order and never overlap.
      let block = -1;
      for (let i = k; i < paras.length && paras[i].start <= d.start; i++) {
        if (d.end <= paras[i].end) { block = paras[i].index; break; }
      }
      const info = readDrawing(d.xml);
      const layout = anchorLayout(d.xml);
      const words = info.kind === 'textbox' ? textOf((/<w:txbxContent\b[^>]*>([\s\S]*?)<\/w:txbxContent>/.exec(d.xml) || [])[1] || '') : null;
      out.push({
        ...info, ...layout, block,
        ...(info.kind === 'group' ? { members: groupMembers(d.xml).members.map((m) => ({ id: m.id, name: m.name, kind: m.kind })) } : {}),
        ...(words != null ? { text: words.slice(0, 120) } : {}),
      });
    }
    return out;
  }

  /** One drawing's span in the body, by its id — or a sentence saying there is none. */
  _drawingById(id) {
    const { body } = this._body();
    const hit = findDrawings(body).find((d) => d.id === Number(id));
    if (!hit) throw new Error('there is no drawing ' + id + ' in this document');
    return hit;
  }

  /** The highest place in the z-order a drawing holds now. */
  _topZ() {
    let top = Z_BASE;
    for (const m of this.xml.matchAll(/<wp:anchor\b[^>]*\brelativeHeight="(\d+)"/g)) top = Math.max(top, Number(m[1]));
    return top;
  }

  /**
   * Rewrite one drawing through `fn(xml, info)`; a text box's VML twin is
   * written again from what the box has become.
   */
  _rewriteDrawing(id, fn) {
    const d = this._drawingById(id);
    const next = fn(d.xml, readDrawing(d.xml));
    if (next === d.xml) return this;
    const { prefix, body, suffix } = this._body();
    let out = body.slice(0, d.start) + next + body.slice(d.end);
    if (d.alternate) {
      const unitEnd = d.unitEnd + (next.length - d.xml.length);
      const unit = out.slice(d.unitStart, unitEnd);
      const fb = /<mc:Fallback\b[^>]*>[\s\S]*<\/mc:Fallback>/.exec(unit);
      if (fb && fb[0].includes('<w:txbxContent') && /<wps:txbx\b/.test(next)) {
        let twin = null;
        try { twin = fallbackFor(next, this.themeColours()); } catch { twin = null; }
        if (twin) out = out.slice(0, d.unitStart) + unit.slice(0, fb.index) + '<mc:Fallback>' + twin + '</mc:Fallback>' + unit.slice(fb.index + fb[0].length) + out.slice(unitEnd);
      }
    }
    this.xml = prefix + out + suffix;
    this.dirty = true;
    return this;
  }

  /**
   * Change one drawing the way Word's Arrange and Format commands do, each
   * only if it is given:
   *   wrap        'inline' | 'square' | 'tight' | 'through' | 'topAndBottom' |
   *               'behind' | 'front' — In Line with Text through In Front of Text
   *   side        which side the words take: bothSides, left, right, largest
   *   h, v        { rel, align } or { rel, offsetPx } — Position and Align
   *   widthPx, heightPx   the size
   *   relativeHeight      a place in the z-order
   *   rot, flipH, flipV   Rotate
   *   name, hidden        the Selection Pane's name and eye
   *   fill, line          a shape's fill ('#RRGGBB' or null) and outline
   *                       ({ colour, widthPx } or null)
   *   insets, vAnchor, vert, autoFit   a text box's margins, vertical
   *                       alignment, text direction and "resize to fit text"
   */
  updateDrawing(id, patch = {}) {
    return this._rewriteDrawing(id, (xml, info) => {
      let out = xml;
      if (patch.wrap !== undefined) {
        if (patch.wrap === 'inline') {
          if (info.anchored) out = toInline(out);
        } else {
          const behind = patch.wrap === 'behind';
          const wrap = patch.wrap === 'behind' || patch.wrap === 'front' ? 'none' : patch.wrap;
          if (!['square', 'tight', 'through', 'topAndBottom', 'none'].includes(wrap)) throw new Error('unknown wrap: ' + patch.wrap);
          const side = patch.side || (/<wp:wrap(?:Square|Tight|Through)\b[^>]*\bwrapText="([^"]*)"/.exec(out)?.[1]) || 'bothSides';
          if (!info.anchored) {
            out = toAnchor(out, {
              wrap, side, behind, relativeHeight: this._topZ() + Z_STEP,
              h: patch.h || { rel: 'column', offsetPx: 0 }, v: patch.v || { rel: 'paragraph', offsetPx: 0 },
            });
          } else {
            out = withWrap(out, wrap, side);
            out = withAnchorAttrs(out, { behindDoc: behind ? '1' : '0' });
          }
        }
      }
      const anchored = /<wp:anchor\b/.test(out);
      if (anchored && patch.h) out = withPosition(out, 'H', patch.h);
      if (anchored && patch.v) out = withPosition(out, 'V', patch.v);
      if (anchored && patch.relativeHeight != null) out = withAnchorAttrs(out, { relativeHeight: String(Math.max(0, Math.round(patch.relativeHeight))) });
      if (patch.widthPx != null || patch.heightPx != null) {
        const w = patch.widthPx ?? info.widthPx ?? 96;
        const h = patch.heightPx ?? info.heightPx ?? 96;
        if (!(w > 0) || !(h > 0)) throw new Error('a drawing needs a positive width and height');
        out = withExtent(out, w, h);
      }
      if (patch.rot !== undefined || patch.flipH !== undefined || patch.flipV !== undefined) {
        out = withTransform(out, { rot: patch.rot, flipH: patch.flipH, flipV: patch.flipV });
      }
      if (patch.name !== undefined || patch.hidden !== undefined) out = withDocPr(out, { name: patch.name, hidden: patch.hidden });
      if (patch.fill !== undefined) out = withShapeFill(out, patch.fill == null ? null : patch.fill);
      if (patch.line !== undefined) out = withShapeLine(out, patch.line == null ? { none: true } : patch.line);
      if (patch.insets !== undefined || patch.vAnchor !== undefined || patch.vert !== undefined || patch.autoFit !== undefined) {
        out = withBodyPr(out, { insets: patch.insets, anchor: patch.vAnchor, vert: patch.vert, autoFit: patch.autoFit });
      }
      return out;
    });
  }

  /**
   * Bring Forward, Send Backward, Bring to Front, Send to Back — the
   * floating drawings' order, which Word keeps in each anchor's
   * `relativeHeight`: the chosen ones move, the others keep their order,
   * and only the drawings whose place changed are rewritten.
   */
  orderDrawings(ids, how) {
    const chosen = new Set((Array.isArray(ids) ? ids : [ids]).map(Number));
    const floating = this.drawings().filter((d) => d.anchored);
    if (!floating.some((d) => chosen.has(d.id))) throw new Error('Bring Forward and Send Backward move floating drawings — choose a wrap other than In Line with Text first.');
    const order = floating
      .map((d, i) => ({ id: d.id, z: d.relativeHeight ?? Z_BASE, i }))
      .sort((a, b) => a.z - b.z || a.i - b.i);
    const list = order.map((d) => d.id);
    let next;
    if (how === 'front') next = [...list.filter((id) => !chosen.has(id)), ...list.filter((id) => chosen.has(id))];
    else if (how === 'back') next = [...list.filter((id) => chosen.has(id)), ...list.filter((id) => !chosen.has(id))];
    else if (how === 'forward' || how === 'backward') {
      next = list.slice();
      const up = how === 'forward';
      const idx = next.map((id, i) => (chosen.has(id) ? i : -1)).filter((i) => i >= 0);
      for (const i of up ? idx.reverse() : idx) {
        let j = i;
        // Past the next one that is not chosen, as Word steps one place.
        const k = up ? j + 1 : j - 1;
        if (k < 0 || k >= next.length || chosen.has(next[k])) continue;
        [next[j], next[k]] = [next[k], next[j]];
        j = k;
      }
    } else throw new Error('unknown order: ' + how);
    // Heights as Word spaces them, rewritten only where they changed.
    const current = new Map(order.map((d) => [d.id, d.z]));
    next.forEach((id, i) => {
      const z = Z_BASE + (i + 1) * Z_STEP;
      if (current.get(id) !== z) this.updateDrawing(id, { relativeHeight: z });
    });
    return this;
  }

  /**
   * A drawing's run moved to another paragraph — its anchor, which Word
   * moves with a floating drawing dragged down the page. The drawing lands
   * at the paragraph's start; a run that held nothing else goes with it.
   */
  moveDrawing(id, toIndex) {
    const d = this._drawingById(id);
    const target = this.editParagraph(toIndex);
    if (!target) throw new Error('no paragraph at index ' + toIndex);
    if (target.box) throw new Error('a drawing cannot be anchored inside a text box');
    if (target.start <= d.start && d.end <= target.end) return this;
    const { prefix, body, suffix } = this._body();
    const unit = body.slice(d.unitStart, d.unitEnd);
    const run = d.runStart >= 0 ? body.slice(d.runStart, d.runEnd) : null;
    const rest = run ? run.slice(0, d.unitStart - d.runStart) + run.slice(d.unitEnd - d.runStart) : '';
    const bare = !run || /^<w:r\b[^>]*>\s*(?:<w:rPr\b[^>]*\/>|<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>)?\s*<\/w:r>$/.test(rest);
    const [cutFrom, cutTo] = run && bare ? [d.runStart, d.runEnd] : [d.unitStart, d.unitEnd];
    const moved = '<w:r>' + unit + '</w:r>';
    const openTag = /<w:p\b[^>]*?>/.exec(target.xml)[0];
    const pPr = /^<w:p\b[^>]*?>\s*(<w:pPr\b[^>]*\/>|<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>)?/.exec(target.xml);
    let at = target.start + (pPr ? pPr[0].length : openTag.length);
    let out;
    if (target.xml.endsWith('/>') && /^<w:p\b[^>]*\/>$/.test(target.xml)) {
      // An empty paragraph written short: open it to take the run.
      const opened = target.xml.replace(/\/>$/, '>') + moved + '</w:p>';
      out = cutFrom < target.start
        ? body.slice(0, cutFrom) + body.slice(cutTo, target.start) + opened + body.slice(target.end)
        : body.slice(0, target.start) + opened + body.slice(target.end, cutFrom) + body.slice(cutTo);
    } else if (at <= cutFrom) {
      out = body.slice(0, at) + moved + body.slice(at, cutFrom) + body.slice(cutTo);
    } else {
      out = body.slice(0, cutFrom) + body.slice(cutTo, at) + moved + body.slice(at);
    }
    this.xml = prefix + out + suffix;
    this.dirty = true;
    return this;
  }

  /** A drawing taken out of the document — its run, when the run held nothing else. The paragraph stays. */
  removeDrawing(id) {
    const d = this._drawingById(id);
    const { prefix, body, suffix } = this._body();
    const run = d.runStart >= 0 ? body.slice(d.runStart, d.runEnd) : null;
    const rest = run ? run.slice(0, d.unitStart - d.runStart) + run.slice(d.unitEnd - d.runStart) : '';
    const bare = !run || /^<w:r\b[^>]*>\s*(?:<w:rPr\b[^>]*\/>|<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>)?\s*<\/w:r>$/.test(rest);
    const [from, to] = run && bare ? [d.runStart, d.runEnd] : [d.unitStart, d.unitEnd];
    this.xml = prefix + body.slice(0, from) + body.slice(to) + suffix;
    this.dirty = true;
    return this;
  }

  /**
   * Insert → Text Box: a floating text box, written as Word 2010 and later
   * write one (see drawings.js `textBoxRun`), anchored at the start of
   * paragraph `index`. `paragraphs` are its words: `{ text, bold, italic,
   * sizePt, colour, align, font }` each. Answers the new box's id.
   */
  insertTextBox(index, spec = {}) {
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    if (p.box) throw new Error('a text box cannot go inside another text box');
    if (p.container) throw new Error('A text box goes beside the words, not in a table cell — move the caret out of the table first.');
    if (p.structural) throw new Error('A text box cannot be anchored in this paragraph — it is part of a field or a content control.');
    this.ensureDrawingNamespaces();
    const id = this._nextDrawingId();
    const words = (spec.paragraphs && spec.paragraphs.length ? spec.paragraphs : [{ text: '' }]).map((w) => {
      const rPr = [
        w.font ? '<w:rFonts w:ascii="' + esc(w.font) + '" w:hAnsi="' + esc(w.font) + '"/>' : '',
        w.bold ? '<w:b/>' : '', w.italic ? '<w:i/>' : '',
        w.colour ? '<w:color w:val="' + String(w.colour).replace('#', '').toUpperCase() + '"/>' : '',
        w.sizePt ? '<w:sz w:val="' + Math.round(w.sizePt * 2) + '"/><w:szCs w:val="' + Math.round(w.sizePt * 2) + '"/>' : '',
      ].join('');
      const pPr = [
        w.style ? '<w:pStyle w:val="' + esc(w.style) + '"/>' : '',
        '<w:spacing w:after="' + (w.afterTwips ?? 0) + '" w:line="' + (w.lineTwips ?? 240) + '" w:lineRule="auto"/>',
        w.align ? '<w:jc w:val="' + (w.align === 'justify' ? 'both' : w.align) + '"/>' : '',
      ].join('');
      const run = w.text ? '<w:r>' + (rPr ? '<w:rPr>' + rPr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + esc(w.text) + '</w:t></w:r>' : '';
      return '<w:p><w:pPr>' + pPr + (rPr ? '<w:rPr>' + rPr + '</w:rPr>' : '') + '</w:pPr>' + run + '</w:p>';
    }).join('');
    const run = textBoxRun({
      id,
      name: spec.name || 'Text Box ' + id,
      widthPx: spec.widthPx ?? 240, heightPx: spec.heightPx ?? 96,
      h: spec.h || { rel: 'column', offsetPx: 0 }, v: spec.v || { rel: 'paragraph', offsetPx: 0 },
      wrap: spec.wrap || 'square', behind: Boolean(spec.behind),
      relativeHeight: this._topZ() + Z_STEP,
      fill: spec.fill === undefined ? 'FFFFFF' : spec.fill == null ? null : String(spec.fill).replace('#', ''),
      line: spec.line === undefined ? '000000' : spec.line == null ? null : String(spec.line).replace('#', ''),
      lineWidthPx: spec.lineWidthPx ?? 1,
      insets: spec.insets || { l: 9.6, t: 4.8, r: 9.6, b: 4.8 },
      anchor: spec.vAnchor || 'top', vert: spec.vert || 'horz', autoFit: Boolean(spec.autoFit),
      paragraphs: words,
    });
    const fresh = this.editParagraph(index);
    const openTag = /<w:p\b[^>]*?>/.exec(fresh.xml)[0];
    if (/^<w:p\b[^>]*\/>$/.test(fresh.xml)) {
      this._spliceBody(fresh.start, fresh.end, fresh.xml.replace(/\/>$/, '>') + run + '</w:p>');
    } else {
      const lead = /^<w:p\b[^>]*?>\s*(<w:pPr\b[^>]*\/>|<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>)?/.exec(fresh.xml);
      const at = fresh.start + (lead ? lead[0].length : openTag.length);
      this._spliceBody(at, at, run);
    }
    return id;
  }

  /** The edit-space paragraphs of a text box, by the box's id — where the caret goes in it. */
  textBoxBlocks(id) {
    const d = this._drawingById(id);
    return this.editParagraphs().filter((p) => p.box && p.box.content > d.start && p.box.content < d.end).map((p) => p.index);
  }

  /**
   * Group: the floating drawings named, gathered into one `wpg:wgp` group
   * the way Word writes it, anchored in the paragraph of the first. `rects`
   * says where each one stands on the page, in px (`{ id, x, y, w, h }`), the
   * group's own box is theirs together, and `place` is where that box stands
   * from the anchor — `{ h: { rel, offsetPx }, v: { rel, offsetPx } }` — which
   * only the page can know, since the paragraphs' places are its layout.
   */
  groupDrawings(ids, { rects = [], place = null } = {}) {
    const list = (ids || []).map(Number);
    if (list.length < 2) throw new Error('Select two or more drawings to group them.');
    const all = this.drawings();
    const chosen = list.map((id) => all.find((d) => d.id === id));
    if (chosen.some((d) => !d)) throw new Error('one of those drawings is not in this document');
    if (chosen.some((d) => !d.anchored)) throw new Error('A drawing in line with the text cannot be grouped — give each one a wrap first (Layout → Wrap Text).');
    if (chosen.some((d) => d.kind === 'chart' || d.kind === 'canvas' || d.kind === 'other')) throw new Error('Only pictures, shapes, text boxes and groups can be grouped.');
    const byId = new Map(rects.map((r) => [Number(r.id), r]));
    const boxes = chosen.map((d) => byId.get(d.id) || { x: d.hOffsetPx || 0, y: d.vOffsetPx || 0, w: d.widthPx || 96, h: d.heightPx || 96 });
    const x0 = Math.min(...boxes.map((b) => b.x));
    const y0 = Math.min(...boxes.map((b) => b.y));
    const x1 = Math.max(...boxes.map((b) => b.x + b.w));
    const y1 = Math.max(...boxes.map((b) => b.y + b.h));
    const cx = Math.max(1, Math.round((x1 - x0) * EMU_PER_PX));
    const cy = Math.max(1, Math.round((y1 - y0) * EMU_PER_PX));
    const { body } = this._body();
    const spans = findDrawings(body);
    // Members in their order on the page, the lowest first: in a group the
    // later member is drawn over the earlier, as a higher relativeHeight was.
    const stacked = chosen.map((d, i) => ({ d, b: boxes[i] })).sort((p, q) => (p.d.relativeHeight ?? 0) - (q.d.relativeHeight ?? 0));
    const members = stacked.map(({ d, b }) => {
      const span = spans.find((s) => s.id === d.id);
      return memberXml(span.xml, { x: (b.x - x0) * EMU_PER_PX, y: (b.y - y0) * EMU_PER_PX, cx: b.w * EMU_PER_PX, cy: b.h * EMU_PER_PX });
    });
    const first = chosen[0];
    const id = this._nextDrawingId();
    const wrap = first.behind ? 'none' : first.wrap || 'square';
    const graphic = groupGraphic(members, { cx, cy });
    const drawing = anchorXml({
      open: '<w:drawing>',
      extent: '<wp:extent cx="' + cx + '" cy="' + cy + '"/>',
      effectExtent: '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
      docPr: '<wp:docPr id="' + id + '" name="Group ' + id + '"/>',
      frame: '<wp:cNvGraphicFramePr/>',
      graphic,
    }, {
      wrap, side: first.wrapSide || 'bothSides', behind: first.behind,
      relativeHeight: Math.max(...chosen.map((d) => d.relativeHeight || Z_BASE)),
      h: place?.h || { rel: first.hRel || 'column', offsetPx: x0 },
      v: place?.v || { rel: first.vRel || 'paragraph', offsetPx: y0 },
    });
    this.ensureDrawingNamespaces();
    // Take the members out, last first so the earlier offsets hold, then
    // put the group where the first one was.
    const anchorBlock = first.block;
    for (const d of [...chosen].sort((a, b) => spans.find((s) => s.id === b.id).start - spans.find((s) => s.id === a.id).start)) this.removeDrawing(d.id);
    const target = this.editParagraph(anchorBlock);
    const lead = /^<w:p\b[^>]*?>\s*(<w:pPr\b[^>]*\/>|<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>)?/.exec(target.xml);
    const run = '<w:r>' + drawing + '</w:r>';
    if (/^<w:p\b[^>]*\/>$/.test(target.xml)) this._spliceBody(target.start, target.end, target.xml.replace(/\/>$/, '>') + run + '</w:p>');
    else this._spliceBody(target.start + lead[0].length, target.start + lead[0].length, run);
    return id;
  }

  /**
   * Ungroup: each member of a group a floating drawing of its own again,
   * where it stood in the group, anchored where the group was. `place` is
   * where the group's box stands from its anchor, in px — the page's to say
   * when the group is aligned rather than offset.
   */
  ungroupDrawing(id, { place = null } = {}) {
    const d = this._drawingById(id);
    const info = readDrawing(d.xml);
    if (info.kind !== 'group') throw new Error('That drawing is not a group.');
    const layout = anchorLayout(d.xml);
    const g = groupMembers(d.xml);
    const sx = (g.ext.cx || 1) / (g.chExt.cx || 1);
    const sy = (g.ext.cy || 1) / (g.chExt.cy || 1);
    const h0 = place?.h ?? { rel: layout.hRel || 'column', offsetPx: layout.hOffsetPx || 0 };
    const v0 = place?.v ?? { rel: layout.vRel || 'paragraph', offsetPx: layout.vOffsetPx || 0 };
    const taken = new Set([...this.xml.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1])));
    let fresh = this._nextDrawingId();
    const runs = g.members.map((m, i) => {
      let mid = m.id;
      if (mid == null || taken.has(mid)) mid = fresh++;
      taken.add(mid);
      const widthPx = (m.cx * sx) / EMU_PER_PX;
      const heightPx = (m.cy * sy) / EMU_PER_PX;
      const xPx = ((m.x - g.chOff.x) * sx) / EMU_PER_PX;
      const yPx = ((m.y - g.chOff.y) * sy) / EMU_PER_PX;
      const drawing = memberToDrawing(m, {
        id: mid, widthPx, heightPx,
        h: { rel: h0.rel, offsetPx: (h0.offsetPx || 0) + xPx },
        v: { rel: v0.rel, offsetPx: (v0.offsetPx || 0) + yPx },
        wrap: layout.wrap || 'square', side: layout.wrapSide || 'bothSides', behind: layout.behind,
        relativeHeight: (info.relativeHeight || Z_BASE) + i,
      });
      return '<w:r>' + drawing + '</w:r>';
    }).join('');
    const { prefix, body, suffix } = this._body();
    const run = d.runStart >= 0 ? body.slice(d.runStart, d.runEnd) : null;
    const rest = run ? run.slice(0, d.unitStart - d.runStart) + run.slice(d.unitEnd - d.runStart) : '';
    const bare = !run || /^<w:r\b[^>]*>\s*(?:<w:rPr\b[^>]*\/>|<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>)?\s*<\/w:r>$/.test(rest);
    const [from, to] = run && bare ? [d.runStart, d.runEnd] : [d.unitStart, d.unitEnd];
    // Runs where the group's run was — or, when the group shared its run, after it.
    const insert = run && bare ? runs : '</w:r>' + runs + '<w:r>';
    this.xml = prefix + body.slice(0, from) + insert + body.slice(to) + suffix;
    this.dirty = true;
    return g.members.length;
  }

  /**
   * A group's members, to draw: each one's box inside the group (px, in the
   * group's own drawn size), what it is, and what it shows — a picture's
   * bytes, a shape's XML, a text box's look and paragraphs (and, when the
   * edit space lists them, their indices).
   */
  _paragraphGroups(paragraphXml, base = null) {
    const out = [];
    const rels = new Map(this.pkg.rels(this.mainPart).map((r) => [r.Id, r.Target]));
    const colours = this.themeColours();
    for (const d of findDrawings(String(paragraphXml))) {
      const info = readDrawing(d.xml);
      if (info.kind !== 'group') continue;
      const g = groupMembers(d.xml);
      const W = info.widthPx || 1;
      const H = info.heightPx || 1;
      const sx = W / Math.max(1, g.chExt.cx);
      const sy = H / Math.max(1, g.chExt.cy);
      const members = g.members.map((m) => {
        const box = { xPx: (m.x - g.chOff.x) * sx, yPx: (m.y - g.chOff.y) * sy, widthPx: m.cx * sx, heightPx: m.cy * sy, rot: m.rot, flipH: m.flipH, flipV: m.flipV };
        if (m.kind === 'picture') {
          const relId = /<a:blip\b[^>]*\br:embed="([^"]*)"/.exec(m.xml)?.[1];
          const target = relId ? rels.get(relId) : null;
          const part = target ? OoxmlPackage.resolveTarget(this.mainPart, target) : null;
          const bytes = part && this.pkg.has(part) ? this.pkg.read(part) : null;
          return { ...box, kind: 'picture', id: m.id, name: m.name, href: bytes ? toDataUri(bytes, part) : null };
        }
        if (m.kind === 'textbox') {
          const open = /<w:txbxContent\b[^>]*>/.exec(m.xml);
          const content = open ? m.xml.slice(open.index + open[0].length, m.xml.indexOf('</w:txbxContent>', open.index)) : '';
          const at = base != null && open ? base + d.start + d.xml.indexOf(m.xml) + open.index : null;
          const blocks = at != null ? this._boxIndex?.get(at) ?? null : null;
          const spPr = /<wps:spPr\b[^>]*>([\s\S]*?)<\/wps:spPr>/.exec(m.xml)?.[1] ?? '';
          const line = /<a:ln\b[^>]*>([\s\S]*?)<\/a:ln>/.exec(spPr);
          const bodyPr = readBodyPr(m.xml);
          return {
            ...box, kind: 'textbox', id: m.id, name: m.name,
            fill: colourOf(spPr.replace(/<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/, ''), colours), line: line ? colourOf(line[1], colours) : null,
            insets: bodyPr.insets, vAnchor: bodyPr.anchor, vert: bodyPr.vert,
            paragraphs: this._liteParagraphs(content), ...(blocks && blocks.length ? { blocks } : {}),
          };
        }
        return { ...box, kind: m.kind, id: m.id, name: m.name, shapeXml: m.kind === 'shape' ? m.xml : null };
      });
      out.push({ ...anchorLayout(d.xml), ...arrangeOf(d.xml), kind: 'group', name: info.name, widthPx: W, heightPx: H, members });
    }
    return out;
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
      throw new Error('unsupported image type: ' + contentType + ' (png, jpeg, gif, bmp or webp)');
    }
    // Bytes come as a Buffer, as the Uint8Array a file dialog's read hands
    // over the IPC, or as base64 text — never as the digits of an array.
    const bytes = Buffer.isBuffer(data) ? data : data instanceof Uint8Array ? Buffer.from(data) : Buffer.from(String(data), 'base64');
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
   * The nth picture out of a paragraph — its run, so a paragraph of words
   * keeps them; a paragraph that held only the picture goes with it, unless
   * it is the body's last, which is left empty. The media part stays in the
   * package, unreferenced and harmless, as an undone insert leaves it.
   */
  removeImage(index, imageIndex = 0) {
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    const xml = p.xml;
    let seen = -1;
    let hit = null;
    for (const m of xml.matchAll(/<w:drawing\b[^>]*>[\s\S]*?<\/w:drawing>/g)) {
      if (!/<a:blip\b/.test(m[0])) continue;
      seen += 1;
      if (seen === Number(imageIndex)) { hit = m; break; }
    }
    if (!hit) throw new Error('no picture ' + imageIndex + ' in paragraph ' + index);
    const runOpen = Math.max(xml.lastIndexOf('<w:r>', hit.index), xml.lastIndexOf('<w:r ', hit.index));
    const runClose = xml.indexOf('</w:r>', hit.index + hit[0].length);
    if (runOpen < 0 || runClose < 0) throw new Error('the picture is not in a run');
    const without = xml.slice(0, runOpen) + xml.slice(runClose + '</w:r>'.length);
    const empty = !/<w:r\b/.test(without.replace(/<w:pPr\b[\s\S]*?<\/w:pPr>/, ''));
    if (empty && this.editParagraphCount() > 1 && !p.container) {
      this._spliceBody(p.start, p.end, '');
    } else {
      this._spliceBody(p.start, p.end, empty ? p.open + (p.pPr || '') + '</w:p>' : without);
    }
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
   * The watermark: faint words across every page, as Word keeps them — a
   * WordArt shape (a VML text path) in the default header's first
   * paragraph, rotated, in a fill colour, behind the body. Text sets or
   * replaces it; null takes it out and leaves the header's words alone. A
   * document with no header gets one for it.
   */
  setWatermark(text, { colour = 'silver', rotation = 315 } = {}) {
    const words = text == null ? '' : String(text).trim();
    let band = this.headerFooters().headers.default ?? null;
    if (!band) {
      if (!words) return this;
      this._addBand('header', '<w:p/>');
      band = this.headerFooters().headers.default;
    }
    let xml = this.pkg.text(band.part);
    // Whatever paragraph was the watermark's goes; the words stay.
    xml = xml.replace(/<w:p\b[^>]*?(?:\/>|>[\s\S]*?<\/w:p>)/g, (p) => (WATERMARK_P.test(p) ? '' : p));
    if (words) {
      const root = /<w:hdr\b[^>]*>/.exec(xml);
      if (!root) throw new Error('unrecognised header part: ' + band.part);
      let open = root[0];
      if (!/\bxmlns:v=/.test(open)) open = open.replace(/>$/, ' xmlns:v="' + VML_NS + '">');
      if (!/\bxmlns:o=/.test(open)) open = open.replace(/>$/, ' xmlns:o="' + VML_OFFICE_NS + '">');
      xml = xml.slice(0, root.index) + open + watermarkParagraph(words, colour, rotation) + xml.slice(root.index + root[0].length);
    }
    if (!/<w:p\b/.test(xml)) xml = xml.replace(/<\/w:hdr>/, '<w:p/></w:hdr>');
    this.pkg.write_(band.part, xml);
    this._undoParts.add(band.part);
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
      // The watermark's paragraph is the band's shape, not its words: it
      // stays first, and the lines follow it.
      const all = [...xml.matchAll(/<w:p\b[^>]*?(?:\/>|>[\s\S]*?<\/w:p>)/g)].map((m) => m[0]);
      const watermark = all.find((p) => WATERMARK_P.test(p)) || '';
      const olds = all.filter((p) => !WATERMARK_P.test(p));
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
      this.pkg.write_(existing.part, root[1] + watermark + body + root[2]);
      this._undoParts.add(existing.part);
      return this;
    }

    // No band yet: a part, a content type, a relationship, a reference.
    return this._addBand(which, texts.map((t) => '<w:p>' + (t === '' ? '' : renderRun(null, t)) + '</w:p>').join('') || '<w:p/>');
  }

  /** A new default band around the paragraphs given: its part, content type, relationship and `sectPr` reference. */
  _addBand(which, body) {
    const tag = which === 'header' ? 'hdr' : 'ftr';
    let n = 1;
    while (this.pkg.has('word/' + which + n + '.xml')) n += 1;
    const partName = 'word/' + which + n + '.xml';
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
    const at = mainSectPr(docBody, { openOnly: true });
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

  /**
   * Bookmark spans — Word's older anchor mechanism, a NAME on a range rather
   * than a content control's tag on a spot. `from`/`to` are paragraph indexes
   * in `paragraphs()` order: where the `w:bookmarkStart` sits, and where the
   * matching `w:bookmarkEnd` (same `w:id`) does. An end that lands at body
   * level between two paragraphs belongs to the one before it — a mark
   * cannot reach into a paragraph it never touches; an end that cannot be
   * found at all collapses the span onto its start. Word's own bookmarks
   * (`_GoBack` and friends) are left out, as `_`-prefixed names always are.
   */
  bookmarks() {
    const { body } = this._body();
    const paragraphs = this.paragraphs();
    const paragraphAt = (offset) => {
      for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        if (offset >= p.start && offset <= p.end) return i;
        if (offset < p.start) return Math.max(0, i - 1);
      }
      return paragraphs.length - 1;
    };
    const starts = new Map();
    for (const m of body.matchAll(/<w:bookmarkStart\b([^>]*)\/>/g)) {
      const a = attrs(m[1]);
      if (!a['w:name'] || a['w:name'].startsWith('_')) continue;
      starts.set(Number(a['w:id']), { name: a['w:name'], offset: m.index });
    }
    const ends = new Map();
    for (const m of body.matchAll(/<w:bookmarkEnd\b([^>]*)\/>/g)) {
      ends.set(Number(attrs(m[1])['w:id']), m.index);
    }
    const out = [];
    for (const [id, { name, offset }] of starts) {
      const from = paragraphAt(offset);
      const endOffset = ends.get(id);
      const to = endOffset === undefined ? from : Math.max(from, paragraphAt(endOffset));
      out.push({ id, name, from, to });
    }
    return out;
  }

  /**
   * Mint a bookmark over paragraphs `from`..`to` (inclusive, `paragraphs()`
   * order) — Insert > Bookmark. Word's own name rule: letters, digits and
   * underscores, starting with a letter, forty characters at most; a name
   * outside that is one Word itself would refuse, so it is an error here too.
   *
   * Word's Add REPLACES a bookmark of the same name rather than stacking a
   * second one on it, matched here so a document that has been through this
   * twice does not surprise anyone who knows Word's own dialog. The new id is
   * one past the highest `w:bookmarkStart` id already in the body.
   *
   * The end is spliced in first: it sits at or after the start, so writing it
   * first leaves the start paragraph's offset untouched for the second splice.
   */
  addBookmark(name, from, to = from) {
    if (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(String(name))) {
      throw new Error('a bookmark name is letters, digits and underscores, starting with a letter, up to 40 characters');
    }
    this.removeBookmark(name);
    return this._mintBookmark(name, from, to);
  }

  /**
   * The unchecked core of `addBookmark` — no name validation, no replace of
   * an existing one of the same name. Split out for `_Toc` bookmarks (a
   * table of contents entry's anchor on its heading): Word's own name rule
   * above refuses a leading underscore, the very thing that keeps a hidden
   * bookmark out of `bookmarks()` and the Bookmark dialog's list.
   */
  _mintBookmark(name, from, to = from) {
    let maxId = -1;
    for (const m of this.xml.matchAll(/<w:bookmarkStart\b[^>]*\bw:id="(\d+)"/g)) {
      maxId = Math.max(maxId, Number(m[1]));
    }
    const id = maxId + 1;
    const startXml = '<w:bookmarkStart w:id="' + id + '" w:name="' + esc(String(name)) + '"/>';
    const endXml = '<w:bookmarkEnd w:id="' + id + '"/>';

    // A self-closing `<w:p/>` has nowhere to hang a mark — opened up first,
    // the way the model's own `open` normalises it.
    const openUp = (p) => (/^<w:p\b[^>]*\/>$/.test(p.xml) ? p.open + '</w:p>' : p.xml);

    const pTo = this.paragraph(to);
    if (!pTo) throw new Error('no paragraph at index ' + to);
    const toXml = openUp(pTo);
    const withEnd = toXml.slice(0, toXml.length - '</w:p>'.length) + endXml + '</w:p>';
    this._spliceBody(pTo.start, pTo.end, withEnd);

    const pFrom = this.paragraph(from);
    if (!pFrom) throw new Error('no paragraph at index ' + from);
    const fromXml = openUp(pFrom);
    const headLen = pFrom.open.length + (pFrom.pPr ? pFrom.pPr.length : 0);
    const withStart = fromXml.slice(0, headLen) + startXml + fromXml.slice(headLen);
    this._spliceBody(pFrom.start, pFrom.end, withStart);

    this.dirty = true;
    return id;
  }

  /** Take a bookmark's `w:bookmarkStart`/`w:bookmarkEnd` out. True if one was there. */
  removeBookmark(name) {
    const { prefix, body, suffix } = this._body();
    const ids = [];
    for (const m of body.matchAll(/<w:bookmarkStart\b([^>]*)\/>/g)) {
      const a = attrs(m[1]);
      if (a['w:name'] === String(name)) ids.push(a['w:id']);
    }
    if (!ids.length) return false;
    let next = body;
    for (const id of ids) {
      const safe = String(id).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      next = next.replace(new RegExp('<w:bookmarkStart\\b[^>]*w:id="' + safe + '"[^>]*/>'), '');
      next = next.replace(new RegExp('<w:bookmarkEnd\\b[^>]*w:id="' + safe + '"[^>]*/>'), '');
    }
    this.xml = prefix + next + suffix;
    this.dirty = true;
    return true;
  }

  /**
   * Every field in the body, in paragraph order — a `<w:fldSimple>` and a
   * COMPLETE complex field alike, since `parseRuns` folds the latter to the
   * same one-run shape before either reaches here: the code Word wrote
   * (`instr`), what kind of field it is, a REF's bookmark or a SEQ's label
   * (`name`), and the cached words currently shown. `index` is the field
   * run's position among that paragraph's runs, so a caller can find the
   * same field again after an edit moves nothing but its neighbours.
   */
  fields() {
    const out = [];
    for (const entry of this.paragraphs()) {
      const p = this.paragraph(entry.index);
      p.runs.forEach((r, index) => {
        if (!r.field) return;
        out.push({ paragraph: p.index, index, instr: r.field.instr, kind: r.field.kind, name: r.field.name, text: r.text });
      });
    }
    return out;
  }

  /**
   * References → Update Fields (or F9): every REF field's words refreshed
   * from its bookmark — the text of paragraphs `from`..`to`, joined with a
   * space, the way a multi-paragraph bookmark collapses onto a field's one
   * line. A REF whose bookmark has since been removed gets the words Word
   * itself puts there rather than leave the old, now-wrong ones standing.
   *
   * Every SEQ field is renumbered in the same pass — see
   * `_renumberSeqFields` — because F9 in Word refreshes every field kind at
   * once, not REF alone; a caption inserted mid-document, or one deleted,
   * leaves the rest wrong until this runs.
   *
   * Returns how many fields actually changed, for the ribbon's toast.
   */
  refreshRefFields() {
    const NOT_FOUND = 'Error! Reference source not found.';
    const marks = new Map(this.bookmarks().map((b) => [b.name, b]));
    const paragraphs = this.paragraphs();
    let changed = 0;
    for (const entry of paragraphs) {
      if (!/<w:fldSimple\b/.test(entry.xml)) continue;
      const p = this.paragraph(entry.index);
      let touched = false;
      const next = p.runs.map((r) => {
        if (!r.field || r.field.kind !== 'ref') return r;
        const mark = marks.get(r.field.name);
        const words = mark
          ? paragraphs.slice(mark.from, mark.to + 1).map((q) => q.text).join(' ')
          : NOT_FOUND;
        if (words === r.text) return r;
        touched = true;
        changed += 1;
        return { ...r, text: words };
      });
      if (touched) this.setParagraphRuns(p.index, next);
    }
    changed += this._renumberSeqFields();
    // A table of contents is a field too — F9 rebuilds its entries from the
    // current headings and bookmarks exactly as References → Update Table
    // does, keeping whatever page numbers it last cached (see
    // `updateTableOfContents`). Not counted into `changed`: it is a
    // structural rebuild, not a word-for-word refresh the REF/SEQ count means.
    if (this.hasTableOfContents()) this.updateTableOfContents();
    if (changed) this.dirty = true;
    return changed;
  }

  /**
   * Every SEQ field's cached number, right now: the 1-based count of that
   * label's SEQ fields up to and including this one, in document order —
   * `Figure 1`, `Figure 2`, `Table 1`, each label counted on its own, the
   * way Word's own captions are. A complex field stays structural (see
   * `_decorate`), so this rewrites the paragraph's raw XML directly rather
   * than going through `setParagraphRuns`, touching only the result run's
   * `<w:t>` and leaving the begin/instrText/separate/end markers exactly as
   * they were. Returns how many results actually changed.
   */
  _renumberSeqFields() {
    const counts = new Map();
    let changed = 0;
    const total = this.paragraphCount();
    for (let i = 0; i < total; i += 1) {
      const p = this.paragraph(i);
      if (!/<w:fldChar\b[^>]*\bw:fldCharType="begin"/.test(p.xml)) continue;
      const next = mapComplexFieldResults(p.xml, (instr, text) => {
        const m = /^\s*SEQ\s+(\S+)/i.exec(instr);
        if (!m) return undefined;
        const label = m[1];
        const n = (counts.get(label) ?? 0) + 1;
        counts.set(label, n);
        return String(n) === text ? undefined : String(n);
      });
      if (next !== p.xml) {
        this._spliceBody(p.start, p.end, next);
        changed += 1;
      }
    }
    return changed;
  }

  /**
   * Insert → Captions → Insert Caption: a new paragraph after `at`
   * (`paragraphs()` order — the block the caret sits in, or a picture's or
   * table's own paragraph when the caller resolved the caret there first),
   * styled "Caption" (written once, see `_ensureCaptionStyle`), reading
   * `Figure 3: <text>` the way Word's own dialog writes one.
   *
   * The running number is a SEQ field — `SEQ Figure \* ARABIC`, a COMPLEX
   * field (`w:fldChar` begin/separate/end around an `<w:instrText>`), the
   * shape Word itself writes for a caption; a REF's `w:fldSimple` is a
   * different, simpler thing. Its result is written as a placeholder and
   * then corrected by `_renumberSeqFields` in the same call — inserting a
   * caption ahead of others with the same label must renumber THEM too, not
   * just count what already existed, and that is the one place this engine
   * already knows how to do it.
   */
  addCaption({ label, text = '', at }) {
    if (!['Figure', 'Table', 'Equation'].includes(label)) {
      throw new Error('a caption\'s label is Figure, Table or Equation');
    }
    const p = this.paragraph(at);
    if (!p) throw new Error('no paragraph at index ' + at);
    this._ensureCaptionStyle();

    const instr = ' SEQ ' + label + ' \\* ARABIC ';
    const seqXml =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve">' + esc(instr) + '</w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t xml:space="preserve">1</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

    const xml =
      '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>' +
      renderRun(null, label + ' ') +
      seqXml +
      renderRun(null, ': ' + String(text)) +
      '</w:p>';

    this._spliceBody(p.end, p.end, xml);
    this._renumberSeqFields();
    return this;
  }

  // ---- table of contents ---------------------------------------------------

  /**
   * The "TOC1".."TOCn" styles, written once — Word's own defaults on a
   * document that never had a table of contents: no indent for level 1,
   * 220 twips more per level after it, a little space after each entry so
   * they do not run together. Same append-never-rewrite rule as
   * `_ensureCaptionStyle`: an existing catalogue is a template author's own.
   */
  _ensureTocStyles(levels = 3) {
    this.ensureParagraphStyles();
    const part = 'word/styles.xml';
    let xml = this.pkg.text(part);
    let changed = false;
    for (let level = 1; level <= levels; level++) {
      const id = 'TOC' + level;
      if (new RegExp('<w:style\\b[^>]*\\bw:styleId="' + id + '"').test(xml)) continue;
      const indent = (level - 1) * 220;
      const style = '<w:style w:type="paragraph" w:styleId="' + id + '"><w:name w:val="toc ' + level + '"/><w:basedOn w:val="Normal"/>'
        + '<w:pPr>' + (indent ? '<w:ind w:left="' + indent + '"/>' : '') + '<w:spacing w:after="100"/></w:pPr></w:style>';
      xml = xml.replace('</w:styles>', style + '</w:styles>');
      changed = true;
    }
    if (changed) this.pkg.write_(part, xml);
  }

  /**
   * The print area's width, in twips, from the section's own page size and
   * margins — where Word right-aligns a TOC entry's page number, dot-leader
   * tab and all, regardless of the entry's own indent.
   */
  _tocTabPositionTwips() {
    const { body } = this._body();
    const sectPr = mainSectPr(body);
    const sz = sectPr ? /<w:pgSz\b([^>]*)\/>/.exec(sectPr[0]) : null;
    const mar = sectPr ? /<w:pgMar\b([^>]*)\/>/.exec(sectPr[0]) : null;
    const szAttrs = sz ? attrs(sz[1]) : {};
    const marAttrs = mar ? attrs(mar[1]) : {};
    const width = Number(szAttrs['w:w']) || 11906;
    const left = Number(marAttrs['w:left']) || 1440;
    const right = Number(marAttrs['w:right']) || 1440;
    return Math.max(720, width - left - right);
  }

  /**
   * Every heading a table of contents can list, in document order: a
   * top-level paragraph styled Heading1..Heading<levels> with words in it.
   * A heading inside the TOC's own `w:sdt` never reaches this — `paragraphs()`
   * treats a body-level content control as opaque, which is exactly why
   * rebuilding a table of contents never lists its own old entries.
   */
  _headingsForToc(levels) {
    const out = [];
    for (const entry of this.paragraphs()) {
      const m = /^Heading([1-9])$/.exec(this.paragraph(entry.index).style || '');
      if (!m) continue;
      const level = Number(m[1]);
      if (level > levels) continue;
      const text = entry.text.trim();
      if (!text) continue;
      out.push({ index: entry.index, level, text });
    }
    return out;
  }

  /** Every `_Toc<digits>` bookmark in the body, gone — see `_buildToc`. */
  _stripTocBookmarks() {
    const names = new Set();
    for (const m of this.xml.matchAll(/<w:bookmarkStart\b[^>]*\bw:name="(_Toc\d+)"/g)) names.add(m[1]);
    for (const name of names) this.removeBookmark(name);
  }

  /**
   * One TOC entry's content: the heading's own words, a right tab, and a
   * PAGEREF field for its page — all inside a hyperlink to the heading's
   * `_Toc` bookmark, exactly the shape Word writes (Ctrl+click in the
   * window follows `w:anchor` straight to `gotoBookmark`). `page` undefined
   * or null leaves the field's cached result blank rather than lying with a
   * stale number — correct once Update Table supplies one.
   */
  _tocEntryXml({ text, bookmarkName, page }) {
    const pageText = page === undefined || page === null || page === '' ? '' : String(page);
    const pageField =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGEREF ' + bookmarkName + ' \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      renderRun(null, pageText) +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    return '<w:hyperlink w:anchor="' + esc(bookmarkName) + '" w:history="1">' +
      renderRun(null, text) +
      '<w:r><w:tab/></w:r>' +
      pageField +
      '</w:hyperlink>';
  }

  /**
   * The whole field, begin to end, wrapped the way Word 365 writes a table
   * of contents building block: `w:sdt`/`w:docPartObj`/`w:docPartGallery`
   * "Table of Contents" round a run of TOC1..TOCn paragraphs. A document
   * with no headings gets the one line Word itself shows, "No table of
   * contents entries found.", inside the field so Update Table still finds
   * it.
   *
   * Word puts the outer `TOC` field's begin/separate in the FIRST entry
   * paragraph, ahead of its hyperlink, and the end in the LAST, after —
   * this engine cannot follow it there. Its complex-field reader
   * (`foldComplexFields`, shared with SEQ and REF) matches one begin to the
   * NEAREST end within a paragraph; with the outer begin sharing a
   * paragraph with an entry's own complete PAGEREF field, the nearest end
   * is the PAGEREF's, not the TOC field's own — the match pairs wrong and
   * swallows the entry's hyperlink whole. Giving the begin/separate and the
   * end a bare paragraph of their own, front and back, keeps every
   * paragraph's own field self-contained: still one field, still found and
   * refreshed as one (`_tocSpan`/`updateTableOfContents` read the whole
   * `w:sdt` span, not paragraph by paragraph), just one paragraph wider at
   * each end than Word's own.
   */
  _tocFieldXml(headings, levels, pages) {
    const tabTwips = this._tocTabPositionTwips();
    const instr = ' TOC \\o "1-' + levels + '" \\h \\z \\u ';
    const pPrFor = (level) =>
      '<w:pPr><w:pStyle w:val="TOC' + level + '"/>' +
      '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="' + tabTwips + '"/></w:tabs>' +
      '<w:rPr><w:noProof/></w:rPr></w:pPr>';
    // The instruction is written raw, not through `esc`: it is a literal this
    // method fully controls (digits, spaces, backslashes and quotes only,
    // never `&`/`<`/`>`), and Word itself writes a field's `"1-3"` unescaped —
    // `w:instrText` is text content, not an attribute value, so quoting it is
    // unnecessary. `_tocSpan`/`tableOfContents` read the level back the same
    // literal way.
    const fieldBegin =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve">' + instr + '</w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const fieldEnd = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

    let content;
    if (!headings.length) {
      content = '<w:p>' + pPrFor(1) + fieldBegin + renderRun(null, 'No table of contents entries found.') + fieldEnd + '</w:p>';
    } else {
      const beginPara = '<w:p>' + pPrFor(headings[0].level) + fieldBegin + '</w:p>';
      const endPara = '<w:p>' + pPrFor(headings[headings.length - 1].level) + fieldEnd + '</w:p>';
      const entries = headings.map((h, i) => {
        const entry = this._tocEntryXml({ text: h.text, bookmarkName: h.bookmark, page: pages ? pages[i] : undefined });
        return '<w:p>' + pPrFor(h.level) + entry + '</w:p>';
      }).join('');
      content = beginPara + entries + endPara;
    }
    return '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr>' +
      '<w:sdtContent>' + content + '</w:sdtContent></w:sdt>';
  }

  /**
   * Current headings, freshly bookmarked, as the field's XML — the one place
   * insert and update both build from, so the two can never draw the entries
   * differently. Every existing `_Toc` bookmark is stripped and every current
   * heading remarked in document order: simpler and no less correct than
   * matching old bookmarks to headings that may have moved, been renamed or
   * gone, and it is what makes "update after adding/removing/renaming
   * headings" trustworthy.
   */
  _buildToc(levels, pages) {
    const headings = this._headingsForToc(levels);
    this._stripTocBookmarks();
    let nextId = 100000001;
    const withBookmarks = headings.map((h) => {
      const name = '_Toc' + String(nextId++);
      this._mintBookmark(name, h.index, h.index);
      return { ...h, bookmark: name };
    });
    return { xml: this._tocFieldXml(withBookmarks, levels, pages), count: headings.length };
  }

  /**
   * Where an existing table of contents sits in the body, start to end —
   * the `w:sdt` building block this engine writes, or (a file from an older
   * Word, or one built by hand without the content control) a bare `TOC`
   * complex field found by a balanced scan of `w:fldChar`s, since a PAGEREF
   * field nested in every entry means the first `fldChar` "end" met is an
   * entry's own, not the table's. Null when there is none.
   */
  _tocSpan() {
    const { body } = this._body();
    const sdtRe = /<w:sdt\b[^>]*>(?:(?!<w:sdt\b)[\s\S])*?<w:docPartGallery\b[^>]*\bw:val="Table of Contents"[\s\S]*?<\/w:sdt>/;
    const m = sdtRe.exec(body);
    if (m) return { start: m.index, end: m.index + m[0].length };

    const fldRe = /<w:fldChar\b[^>]*\bw:fldCharType="(begin|end)"[^>]*\/>/g;
    let mm;
    let startIdx = -1;
    while ((mm = fldRe.exec(body))) {
      if (mm[1] !== 'begin') continue;
      const after = mm.index + mm[0].length;
      const nextFld = /<w:fldChar\b/.exec(body.slice(after));
      const window = body.slice(after, nextFld ? after + nextFld.index : after + 300);
      const instr = [...window.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)].map((x) => unesc(x[1])).join('');
      if (/^\s*TOC\b/i.test(instr)) { startIdx = mm.index; break; }
    }
    if (startIdx === -1) return null;
    fldRe.lastIndex = startIdx;
    let depth = 0;
    let endIdx = -1;
    while ((mm = fldRe.exec(body))) {
      if (mm[1] === 'begin') depth += 1;
      else {
        depth -= 1;
        if (depth === 0) { endIdx = mm.index + mm[0].length; break; }
      }
    }
    if (endIdx === -1) return null;
    const paraStart = body.lastIndexOf('<w:p', startIdx);
    const paraEndClose = body.indexOf('</w:p>', endIdx);
    return {
      start: paraStart === -1 ? startIdx : paraStart,
      end: paraEndClose === -1 ? endIdx : paraEndClose + '</w:p>'.length,
    };
  }

  /** Every PAGEREF's cached result, in the field's own order — what a plain Update Table (no fresh page map) keeps rather than blanks. */
  _tocCachedPages(spanXml) {
    const out = [];
    for (const m of spanXml.matchAll(/PAGEREF\s+_Toc\d+[^<]*<\/w:instrText>[\s\S]*?<w:fldChar\b[^>]*"separate"[^>]*\/>[\s\S]*?<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)) {
      out.push(unesc(m[1]) || undefined);
    }
    return out;
  }

  /** True when the body carries a table of contents field — for the ribbon's Update Table / Remove Table of Contents buttons. */
  hasTableOfContents() { return this._tocSpan() !== null; }

  /**
   * The table of contents read back — every entry's words, its bookmark and
   * its cached page — for the window and the tests, since the entries live
   * inside an opaque `w:sdt` that `fields()` never reaches.
   */
  tableOfContents() {
    const span = this._tocSpan();
    if (!span) return null;
    const { body } = this._body();
    const xml = body.slice(span.start, span.end);
    const oLevel = /\\o\s+"1-(\d)"/.exec(xml);
    const entries = [];
    for (const m of xml.matchAll(/<w:hyperlink\b[^>]*\bw:anchor="([^"]+)"[^>]*>([\s\S]*?)<\/w:hyperlink>/g)) {
      const inner = m[2];
      const tabIdx = inner.indexOf('<w:tab/>');
      const text = textOf(tabIdx === -1 ? inner : inner.slice(0, tabIdx));
      const pageMatch = /<w:fldChar\b[^>]*"separate"[^>]*\/>[\s\S]*?<w:t\b[^>]*>([\s\S]*?)<\/w:t>/.exec(inner);
      entries.push({ anchor: unesc(m[1]), text, page: pageMatch ? (unesc(pageMatch[1]) || null) : null });
    }
    return { levels: oLevel ? Number(oLevel[1]) : 3, entries };
  }

  /**
   * Insert → Table of Contents, at paragraph `at`: a field, not text — see
   * `_tocFieldXml`. `pages`, aligned with the headings this finds (document
   * order), is the on-screen page of each; omitted, every entry's page is
   * blank until Update Table supplies one.
   */
  insertTableOfContents({ at, levels = 3, pages } = {}) {
    const p = this.paragraph(at);
    if (!p) throw new Error('no paragraph at index ' + at);
    this._ensureTocStyles(levels);
    const { xml } = this._buildToc(levels, pages);
    this._spliceBody(p.end, p.end, xml);
    this.dirty = true;
    return this;
  }

  /**
   * References → Update Table (and F9, alongside REF/SEQ — see
   * `refreshRefFields`): the entries rebuilt from the CURRENT headings and
   * their current bookmarks, keeping the field where it was. `pages` fresh
   * from the caller's on-screen pagination wins; without one, each entry
   * keeps the page number it last cached, sliced or padded to the new count
   * — plain F9 must not blank out numbers nobody asked it to forget.
   */
  updateTableOfContents({ pages } = {}) {
    const span = this._tocSpan();
    if (!span) throw new Error('this document has no table of contents to update');
    const { body } = this._body();
    const existing = body.slice(span.start, span.end);
    const oLevel = /\\o\s+"1-(\d)"/.exec(existing);
    const levels = oLevel ? Number(oLevel[1]) : 3;
    const effectivePages = pages || this._tocCachedPages(existing);
    this._ensureTocStyles(levels);
    const { xml } = this._buildToc(levels, effectivePages);
    this._spliceBody(span.start, span.end, xml);
    this.dirty = true;
    return this;
  }

  /** References → Remove Table of Contents. True when one was there. */
  removeTableOfContents() {
    const span = this._tocSpan();
    if (!span) return false;
    this._spliceBody(span.start, span.end, '');
    this.dirty = true;
    return true;
  }

  // ---- tracked changes -------------------------------------------------------

  /**
   * Accept every tracked change in one paragraph — `editParagraphs()`
   * order, a table cell's own paragraphs addressed exactly as the editor
   * addresses them, not `paragraphs()`'s top-level-only space. An insertion
   * loses its `w:ins` wrapper, keeping the words; a deletion and its
   * `w:delText` are gone outright. True when the paragraph carried a
   * change; false leaves it untouched rather than rewriting for nothing.
   *
   * Regex, not `setEditParagraphRuns`: an accept or reject is a pure
   * strip-the-wrapper edit that never needs the run model at all, and
   * staying off it means this works whether or not a paragraph carrying
   * `w:ins`/`w:del` is one this engine will rebuild from runs (see
   * `_decorate`) — the two are free to evolve apart.
   */
  acceptParagraphChanges(index) {
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    if (!/<w:(?:ins|del)\b/.test(p.xml)) return false;
    const next = p.xml
      .replace(/<w:ins\b[^>]*>([\s\S]*?)<\/w:ins>/g, '$1')
      .replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g, '');
    this._spliceBody(p.start, p.end, next);
    this.dirty = true;
    return true;
  }

  /**
   * Reject every tracked change in one paragraph: an insertion and its
   * words are gone outright; a deletion's `w:del` wrapper comes off and its
   * `w:delText` runs become ordinary `w:t` again — the paragraph as it read
   * before either change.
   */
  rejectParagraphChanges(index) {
    const p = this.editParagraph(index);
    if (!p) throw new Error('no paragraph at index ' + index);
    if (!/<w:(?:ins|del)\b/.test(p.xml)) return false;
    const next = p.xml
      .replace(/<w:ins\b[^>]*>[\s\S]*?<\/w:ins>/g, '')
      .replace(/<w:del\b[^>]*>([\s\S]*?)<\/w:del>/g, (whole, inner) => inner
        .replace(/<w:delText\b([^>]*)\/>/g, '<w:t$1/>')
        .replace(/<w:delText\b([^>]*)>/g, '<w:t$1>')
        .replace(/<\/w:delText>/g, '</w:t>'));
    this._spliceBody(p.start, p.end, next);
    this.dirty = true;
    return true;
  }

  /** Review → Accept All / Reject All. True when anything in the body changed. */
  acceptAllChanges() {
    let changed = false;
    for (let i = 0; i < this.editParagraphCount(); i++) if (this.acceptParagraphChanges(i)) changed = true;
    return changed;
  }

  rejectAllChanges() {
    let changed = false;
    for (let i = 0; i < this.editParagraphCount(); i++) if (this.rejectParagraphChanges(i)) changed = true;
    return changed;
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
    // w:txbxContent left the list the day a text box became editable: a
    // rebuild keeps the run that holds a box whole (`_keptFragments`), so
    // the paragraph that anchors one is typed in like any other. What the
    // tests below look for is the paragraph's OWN XML — a field in a box's
    // words is the box's, not the anchor's.
    // A note reference is NOT on the list: it is a run of its own with one
    // character of text, and the rebuild writes the element back from it.
    // w:bookmarkStart left the list 2026-09-24: the rebuilders now carry a
    // paragraph's bookmark marks through via `_leadFragments`/`_keptFragments`
    // (see those), so a bookmarked paragraph no longer has to go read-only.
    // w:fldSimple left the list the same day a REF field needed one: parseRuns
    // now reads a simple field as ONE run carrying `field`, and renderRuns
    // writes the `<w:fldSimple>` wrapper back from it — see runs.js. A
    // COMPLEX field (`w:fldChar` begin/separate/end, `w:instrText`) is not
    // this lucky: its result is split across several runs with no single one
    // to own the wrapper, so it stays structural.
    // w:ins and w:del left the list the day recording tracked changes was
    // built: parseRuns/renderRuns now carry an insertion or a deletion as a
    // run of their own (runs.js's `flatInsRuns`/`flatDelRuns`, the ins/del
    // counterpart of a hyperlink's group wrapper), so a paragraph mid-review
    // rebuilds like any other — typing beside somebody's tracked change no
    // longer has to wait for it to be resolved first. Its tracked state is
    // still summarised below for the margin and the Reviewing Pane.
    const ownXml = p.xml.includes('<w:txbxContent') ? stripTextBoxes(p.xml) : p.xml;
    const structural = ['w:fldChar', 'w:commentRangeStart', 'w:sdt']
      .filter((tag) => new RegExp('<' + tag + '\\b').test(ownXml))
      // Mail merge fields are complex fields too, and a letter is made of
      // them — "Dear «FirstName»," must stay a line a person can type in.
      // When every complex field in the paragraph is a whole merge field at
      // its top level, each is one run the rebuild writes back verbatim
      // (runs.js `foldMergeFields`), so the paragraph is not locked.
      .filter((tag) => tag !== 'w:fldChar' || !mergeFieldsOnly(p.xml));
    // A paragraph INSIDE a body-level content control carries no sdt tag of
    // its own; it is read-only for the same reason one that does is.
    if (p.inSdt && !structural.includes('w:sdt')) structural.push('w:sdt');
    // The paragraph's OWN words: a text box's are the box's (`textBoxes`
    // below), not the anchor's — counted once, drawn once.
    // Its properties are its own too: an anchor with no `w:pPr` must not
    // take the first one inside its box for its own.
    const own = ownXml;
    const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>|<w:pPr\b[^>]*\/>/.exec(own);
    const style = /<w:pStyle\b[^>]*w:val="([^"]*)"/.exec(pPr ? pPr[0] : '');
    // `open` is normalised to a real opening tag: a self-closing `<w:p/>` (an
    // empty cell, a blank line) must not be used as a prefix and then closed
    // AGAIN by the rebuilders — `<w:p/>…</w:p>` is how typing into a fresh
    // table cell once produced XML no scan could read.
    const rawOpen = /<w:p\b[^>]*?>/.exec(p.xml)?.[0] ?? '<w:p>';
    return {
      index: p.index,
      container: p.container ?? null,
      ...(p.hiddenCell ? { hiddenCell: true } : {}),
      ...(p.gridPx ? { gridPx: p.gridPx, tableWidth: p.tableWidth ?? null } : {}),
      ...(p.cellSpan > 1 ? { cellSpan: p.cellSpan } : {}),
      ...(p.rowHeightPx ? { rowHeightPx: p.rowHeightPx, rowRule: p.rowRule } : {}),
      ...(p.tableLook ? { tableLook: p.tableLook } : {}),
      ...(p.cellVAlign ? { cellVAlign: p.cellVAlign } : {}),
      xml: p.xml,
      start: p.start,
      end: p.end,
      text: own === p.xml ? p.text : textOf(own),
      open: rawOpen.endsWith('/>') ? rawOpen.replace(/\/>$/, '>') : rawOpen,
      pPr: pPr ? pPr[0] : null,
      style: style ? style[1] : null,
      // An explicit page break is the author's instruction, not a suggestion —
      // the paginator must not decide it knows better.
      // A section break before it — Next Page, Odd or Even — is one too: a
      // merged letter's records each start a page that way.
      pageBreakBefore: PAGE_BREAK_BEFORE.test(pPr ? pPr[0] : '') || EXPLICIT_BREAK.test(p.xml) || (p.container == null && this._sectionStarts().has(p.start)),
      ...(p.container == null && paragraphSectionBreak(pPr ? pPr[0] : null) ? { sectionBreak: paragraphSectionBreak(pPr[0]).type } : {}),
      keepNext: KEEP_NEXT.test(pPr ? pPr[0] : ''),
      keepLines: KEEP_LINES.test(pPr ? pPr[0] : ''),
      // A drop cap rides the same way — a paragraph property the paginator
      // and the painter both read straight off the block, not something
      // they have to ask the format layer for.
      dropCap: readDropCap(pPr ? pPr[0] : ''),
      // A frame placed on the page: drawn there, out of the flow.
      ...(p.container == null && readFrame(pPr ? pPr[0] : '') ? { frame: readFrame(pPr[0]) } : {}),
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
        const drawings = this._paragraphRichDrawings(own);
        return drawings.length ? { drawings } : {};
      })()),
      // Text boxes anchored in this paragraph, their paragraphs read the way
      // the body's are — a cover page is nothing but these.
      ...((() => {
        const textBoxes = this._paragraphTextBoxes(p.xml, p.start);
        return textBoxes.length ? { textBoxes } : {};
      })()),
      // Groups anchored here: each one's members, placed in its box.
      ...((() => {
        const groups = p.xml.includes('wordprocessingGroup') ? this._paragraphGroups(p.xml, p.start) : [];
        return groups.length ? { groups } : {};
      })()),
      // A paragraph in a text box knows its box and the paragraph anchoring it.
      ...(p.box ? { box: p.box } : {}),
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
    const next = body.slice(0, start) + replacement + body.slice(end);
    this.xml = prefix + this._syncTextBoxFallback(next, start) + suffix;
    this.dirty = true;
  }

  /**
   * An edit inside a text box's words, in the DrawingML Word 2010 reads,
   * leaves the VML twin an older Word reads showing the old words; the twin
   * is written again from the box whenever that happens, so the two agree.
   * An edit anywhere else costs one backwards search.
   */
  _syncTextBoxFallback(body, at) {
    const open = body.lastIndexOf('<w:txbxContent', at);
    if (open < 0 || body.lastIndexOf('</w:txbxContent>', at) > open) return body;
    const ac = body.lastIndexOf('<mc:AlternateContent', open);
    if (ac < 0 || body.lastIndexOf('</mc:AlternateContent>', open) > ac) return body;
    const acEnd = body.indexOf('</mc:AlternateContent>', open);
    if (acEnd < 0) return body;
    const unit = body.slice(ac, acEnd);
    if (unit.indexOf('<mc:AlternateContent', 1) >= 0) return body; // a box holding another: left alone
    const choice = /<mc:Choice\b[^>]*>([\s\S]*?)<\/mc:Choice>/.exec(unit);
    const fb = /<mc:Fallback\b[^>]*>[\s\S]*<\/mc:Fallback>/.exec(unit);
    if (!choice || !fb || fb.index < choice.index || !fb[0].includes('<w:txbxContent')) return body;
    if (ac + choice.index + choice[0].length < open) return body; // the edit was in the twin itself
    let twin;
    try { twin = fallbackFor(choice[1], this.themeColours()); } catch { return body; }
    return body.slice(0, ac) + unit.slice(0, fb.index) + '<mc:Fallback>' + twin + '</mc:Fallback>' + unit.slice(fb.index + fb[0].length) + body.slice(acEnd);
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
   *   - any other top-level child (a proofErr range) rides along verbatim —
   *     an equation no longer does: it is a run of the model's own, written
   *     in place. Structural paragraphs never reach a rebuild, so bookmarks
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
      // A bookmarkStart is pulled out to `_leadFragments` instead, and pinned
      // at the paragraph's own start rather than wherever it happened to
      // sit — see that method.
      if (childTag === 'w:bookmarkStart') return;
      // An equation is a run the model OWNS since equations were read (a
      // character of the text, written back verbatim by renderRuns) — kept
      // here as well, it would be written twice.
      if (childTag === 'm:oMath' || childTag === 'm:oMathPara') return;
      // A mail merge field's runs — begin, instruction, separate, end — are a
      // field the model owns as one run and writes back whole (runs.js
      // `foldMergeFields`); kept here too, each field would be written twice.
      // A run holding a text box — its `<w:t>`s are the BOX's words, not the
      // paragraph's (parseRuns reads the paragraph with them stripped) — is
      // kept whole: the drawing, its VML twin and the box's paragraphs.
      if (/<w:txbxContent\b/.test(chunk)) { kept.push(chunk); return; }
      if (/<w:(?:fldChar|instrText)\b/.test(chunk)) return;
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

  /**
   * A paragraph's own `w:bookmarkStart` marks — top-level ones, siblings of
   * its runs rather than something a run wraps. Word can plant one mid-
   * paragraph; a rebuild from runs has no run to hang it inside of, so every
   * one of them is pinned right after `w:pPr` instead, ahead of the text.
   * That widens a mid-paragraph mark to the paragraph's own start the first
   * time the paragraph is edited — the honest cost of a bookmarked paragraph
   * staying editable rather than going read-only forever.
   */
  _leadFragments(p) {
    if (/^<w:p\b[^>]*\/>$/.test(p.xml)) return '';
    const openRaw = /<w:p\b[^>]*?>/.exec(p.xml)[0];
    const inner = p.xml.slice(openRaw.length, p.xml.length - '</w:p>'.length);
    const kept = [];
    const re = /<([A-Za-z0-9]+:[A-Za-z0-9]+)\b[^>]*?(\/?)>|<\/([A-Za-z0-9]+:[A-Za-z0-9]+)\s*>/g;
    let depth = 0;
    let m;
    while ((m = re.exec(inner))) {
      if (m[1]) {
        if (m[2] === '/') {
          if (depth === 0 && m[1] === 'w:bookmarkStart') kept.push(m[0]);
        } else {
          depth += 1;
        }
      } else if (m[3] && depth > 0) {
        depth -= 1;
      }
    }
    return kept.join('');
  }

  _setRuns(p, runs) {
    const rebuilt = p.open + (p.pPr ?? '') + this._leadFragments(p) + renderRuns(runs) + this._keptFragments(p) + '</w:p>';
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
    // The kept fragments — and the paragraph's own bookmark marks — stay with
    // the FIRST half: a split is Enter at the caret, and neither the image
    // nor the bookmark the paragraph carried follows the caret onto the new
    // line.
    const first = p.open + (p.pPr ?? '') + this._leadFragments(p) + renderRuns(before) + this._keptFragments(p) + '</w:p>';
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
      + this._leadFragments(p) + this._leadFragments(next)
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

  /**
   * A run of body paragraphs in the order of their words — Home → Sort:
   * A to Z or Z to A, numbers in their order and case set aside, as Word's
   * sort does for paragraphs of text. Each paragraph moves whole, its look
   * with it. The run must be body paragraphs with nothing else between
   * them: not a table's cells, which the table sorts by its own rows.
   */
  sortParagraphs(from, to, { descending = false } = {}) {
    const ps = [];
    for (let i = from; i <= to; i++) {
      const p = this.editParagraph(i);
      if (!p) throw new Error('no paragraph at index ' + i);
      if (p.container) throw new Error("Sort works on the body's paragraphs; a table sorts by its own rows.");
      this._assertEditable(p, 'sorting it');
      ps.push(p);
    }
    if (ps.length < 2) return this;
    const { body } = this._body();
    if (body.slice(ps[0].start, ps[ps.length - 1].end) !== ps.map((p) => p.xml).join('')) {
      throw new Error('Something other than paragraphs sits inside the selection; select a plain run of paragraphs to sort.');
    }
    const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
    const order = ps.map((_, i) => i).sort((a, b) => (collator.compare(ps[a].text, ps[b].text) * (descending ? -1 : 1)) || a - b);
    this._spliceBody(ps[0].start, ps[ps.length - 1].end, order.map((i) => ps[i].xml).join(''));
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
    const sectPr = mainSectPr(body)?.[0] ?? null;
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

  /**
   * Several columns' widths at once, in twips by column index — what a
   * border dragged on the page asks for, since the column on each side of
   * it changes — and, when the table's own width is a fixed one, that width
   * brought up to date with the grid, so Word draws the table at the size
   * it was dragged to rather than scaling the columns back into the old one.
   */
  setTableColumnWidths(tableStart, widths) {
    const entries = Object.entries(widths || {}).map(([i, w]) => [Number(i), w]).filter(([i]) => Number.isInteger(i) && i >= 0);
    if (!entries.length) throw new Error('no column widths to set');
    for (const [i, w] of entries) this.setTableColumnWidth(tableStart, i, w);
    const parts = this._tableParts(tableStart);
    if (!parts.grid) return this;
    const { body } = this._body();
    const sum = [...body.slice(parts.grid.start, parts.grid.end).matchAll(/<w:gridCol\b[^>]*\bw:w="(\d+)"/g)].reduce((a, g) => a + Number(g[1]), 0);
    const head = body.slice(parts.innerStart, parts.grid.start);
    const tblW = /<w:tblW\b([^>]*)\/>/.exec(head);
    if (tblW && /\bw:type="dxa"/.test(tblW[1]) && sum > 0) {
      const at = parts.innerStart + tblW.index;
      this._spliceBody(at, at + tblW[0].length, '<w:tblW w:w="' + sum + '" w:type="dxa"/>');
    }
    return this;
  }

  /**
   * One row's height in twips, as a floor ("at least" — what Word writes when
   * a row's bottom border is dragged, so a cell that needs more still gets
   * it); `null` lets the row take its content's height again.
   */
  setTableRowHeight(tableStart, rowIndex, twips) {
    const parts = this._tableParts(tableStart);
    const row = parts.rows[rowIndex];
    if (!row) throw new Error('no row ' + rowIndex + ' in this table');
    const { body } = this._body();
    const rowXml = body.slice(row.start, row.end);
    const open = /^<w:tr\b[^>]*>/.exec(rowXml)[0];
    // trPr follows tblPrEx, when a row has one, and precedes the cells.
    const lead = /^<w:tblPrEx\b[^>]*>[\s\S]*?<\/w:tblPrEx>|^<w:tblPrEx\b[^>]*\/>/.exec(rowXml.slice(open.length));
    const pre = open + (lead ? lead[0] : '');
    const trPr = /^<w:trPr\b[^>]*>[\s\S]*?<\/w:trPr>|^<w:trPr\b[^>]*\/>/.exec(rowXml.slice(pre.length));
    let inner = trPr && !trPr[0].endsWith('/>') ? trPr[0].replace(/^<w:trPr\b[^>]*>/, '').replace(/<\/w:trPr>$/, '') : '';
    inner = inner.replace(/<w:trHeight\b[^>]*\/>/g, '');
    if (twips != null) {
      const height = Math.round(Number(twips));
      if (!(height >= 20 && height <= 31680)) throw new Error('a row height must be between 1pt and 22 inches');
      const el = '<w:trHeight w:val="' + height + '" w:hRule="atLeast"/>';
      // In the schema's order trHeight precedes tblHeader, tblCellSpacing, jc and hidden.
      const after = /<w:(tblHeader|tblCellSpacing|jc|hidden)\b/.exec(inner);
      inner = after ? inner.slice(0, after.index) + el + inner.slice(after.index) : inner + el;
    }
    const rest = rowXml.slice(pre.length + (trPr ? trPr[0].length : 0));
    this._spliceBody(row.start, row.end, pre + (inner ? '<w:trPr>' + inner + '</w:trPr>' : '') + rest);
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
