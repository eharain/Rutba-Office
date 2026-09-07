/**
 * Tables, and the page they sit on.
 *
 * A table is the last common thing in a business document that rendered as
 * nothing at all. An invoice, a quotation, a bank confirmation — the numbers in
 * every one of them are in a table, and until now the editor showed the prose
 * around them and silently omitted the part the reader cares about. Preserving
 * it on save was never the problem; showing it was.
 *
 * READ-ONLY, DELIBERATELY. Editing inside a cell needs the whole caret model to
 * address positions inside a nested paragraph list, and a half-working table
 * editor that renumbers paragraph indices is worse than an honest read-only one.
 * `Document.paragraphs()` still returns TOP-LEVEL paragraphs only, so nothing
 * about the existing edit path changes; `blocks()` is a separate read that puts
 * tables back in document order for the renderer.
 *
 * Units are twips throughout OOXML's table model — 1/20 of a point, 1/1440 of an
 * inch. Converted to CSS pixels at the edge, at 96 per inch, because that is the
 * only place a pixel means anything.
 */
import { attrs } from './package.js';
import { textOf, parseRuns } from './runs.js';

/** Twips -> CSS pixels. 1440 twips per inch, 96 pixels per inch. */
export const TWIPS_PER_INCH = 1440;
export const twipsToPx = (twips) => (Number(twips) || 0) * (96 / TWIPS_PER_INCH);
/** Eighths of a point -> pixels; OOXML measures border width that way. */
export const eighthPointsToPx = (eighths) => (Number(eighths) || 0) / 8 * (96 / 72);

/**
 * Immediate children of `xml` named `name`, skipping any nested inside `guard`.
 *
 * A nested table's rows must not be read as the outer table's rows, so finding
 * `w:tr` inside a `w:tbl` guards on `w:tbl`. Written as a tag scan rather than a
 * regex with backreferences because same-name nesting is exactly what a regex
 * cannot express.
 */
function childElements(xml, name, guard) {
  const source = String(xml);
  const out = [];
  const parts = [name, guard].filter(Boolean).join('|');
  const re = new RegExp('<(' + parts + ')\\b[^>]*?(/?)>|</(' + parts + ')>', 'g');

  let guardDepth = 0;
  let start = -1;
  let depth = 0;
  let m;
  while ((m = re.exec(source))) {
    const tag = m[0];
    const opening = !tag.startsWith('</');
    const selfClosing = opening && m[2] === '/';
    const tagName = opening ? m[1] : m[3];

    if (guard && tagName === guard) {
      if (opening && !selfClosing) guardDepth += 1;
      else if (!opening) guardDepth -= 1;
      continue;
    }
    if (tagName !== name) continue;
    if (guardDepth > 0) continue;

    if (selfClosing) {
      if (depth === 0) out.push(tag);
      continue;
    }
    if (opening) {
      if (depth === 0) start = m.index;
      depth += 1;
      continue;
    }
    depth -= 1;
    if (depth === 0 && start >= 0) {
      out.push(source.slice(start, m.index + tag.length));
      start = -1;
    }
  }
  return out;
}

/** The first `<name …>…</name>` or `<name …/>` in a fragment. */
function firstElement(xml, name) {
  const re = new RegExp('<' + name + '\\b[^>]*?(/>|>[\\s\\S]*?</' + name + '>)');
  const m = re.exec(String(xml));
  return m ? m[0] : null;
}

/** `<w:val>`-style attribute of a child element, or null. */
function val(xml, name, attribute = 'w:val') {
  const el = firstElement(xml, name);
  if (!el) return null;
  const a = attrs(el);
  return a[attribute] ?? null;
}

// insideH/insideV are what actually draw a table's grid — a `TableGrid` table
// usually says nothing per-cell and relies entirely on these.
const BORDER_SIDES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];

/**
 * The part of an element that precedes its first child of `stop`.
 *
 * Properties elements (`w:tblPr`, `w:trPr`, `w:tcPr`) are always the first child,
 * so reading only the head keeps a NESTED table's properties from being picked
 * up by the table that contains it.
 */
function headBefore(xml, stop) {
  const i = String(xml).search(stop);
  return i >= 0 ? String(xml).slice(0, i) : String(xml);
}

/**
 * A border set. `w:none` and `w:nil` are real values meaning "explicitly no
 * border", which is different from "unspecified" — the first must beat an
 * inherited border, so it survives as `style: 'none'` rather than as null.
 */
function readBorders(container) {
  if (!container) return null;
  const out = {};
  let any = false;
  for (const side of BORDER_SIDES) {
    const el = firstElement(container, 'w:' + side);
    if (!el) continue;
    const a = attrs(el);
    const style = a['w:val'] ?? 'single';
    out[side] = {
      style,
      widthPx: style === 'none' || style === 'nil' ? 0 : eighthPointsToPx(a['w:sz'] ?? 4),
      colour: a['w:color'] && a['w:color'] !== 'auto' ? '#' + a['w:color'].toLowerCase() : null,
    };
    any = true;
  }
  return any ? out : null;
}

/** `<w:shd w:fill="D9E2F3"/>` -> a CSS colour, or null for none/auto. */
function readShading(container) {
  const shd = container ? firstElement(container, 'w:shd') : null;
  if (!shd) return null;
  const fill = attrs(shd)['w:fill'];
  if (!fill || fill === 'auto' || /^none$/i.test(fill)) return null;
  return '#' + fill.toLowerCase();
}

/**
 * A cell paragraph carrying any of these cannot be rebuilt from its runs — the
 * same list `Document` uses for body paragraphs, duplicated here because this
 * file must not import document.js (document.js imports this one).
 */
const STRUCTURAL_TAGS = [
  // w:hyperlink left this list 2026-08-27: links are modelled runs now, and a
  // rebuild puts the wrapper back — see parseRuns/renderRuns. w:ins and
  // w:del JOINED it the same day: a rebuild cannot re-attribute a tracked
  // change, so mid-review paragraphs are read-only until resolved in Word.
  'w:fldSimple', 'w:fldChar', 'w:bookmarkStart', 'w:commentRangeStart', 'w:sdt', 'w:ins', 'w:del',
];

/**
 * Paragraphs inside a cell, as render-ready blocks. Nested tables recurse.
 *
 * `inSdt` marks content wrapped in a block-level content control: it still
 * renders, but it gets no edit address — `Document.editParagraphs()` treats an
 * `w:sdt` subtree as opaque, and the two scans must agree about what is
 * addressable or the index assignment in `Document.blocks()` would mislabel a
 * paragraph. `structural` is the cell counterpart of a body paragraph's flag:
 * addressable, renderable, but refused for typing.
 */
function cellBlocks(cellXml) {
  const out = [];
  // Document order matters — a paragraph after a nested table must stay after it.
  const re = /<w:(tbl|sdt)\b[^>]*?(\/?)>|<\/w:(tbl|sdt)>|<w:p\b[^>]*?(\/?)>|<\/w:p>/g;
  const source = String(cellXml);
  let depth = 0;
  let sdtDepth = 0;
  let tblStart = -1;
  let pStart = -1;
  let m;
  while ((m = re.exec(source))) {
    const tag = m[0];
    if (m[1] === 'sdt' || m[3] === 'sdt') {
      if (m[1] === 'sdt' && m[2] !== '/') sdtDepth += 1;
      else if (m[3] === 'sdt') sdtDepth -= 1;
      continue;
    }
    if (m[1] === 'tbl' && m[2] !== '/') {
      if (depth === 0) tblStart = m.index;
      depth += 1;
      continue;
    }
    if (tag === '</w:tbl>') {
      depth -= 1;
      if (depth === 0 && tblStart >= 0) {
        const block = { kind: 'table', table: parseTable(source.slice(tblStart, m.index + tag.length)) };
        if (sdtDepth > 0) block.inSdt = true;
        out.push(block);
        tblStart = -1;
      }
      continue;
    }
    if (depth > 0) continue;
    if (/^<w:p\b/.test(tag)) {
      if (tag.endsWith('/>')) {
        out.push({
          kind: 'paragraph', text: '', runs: [], style: null,
          structural: false, ...(sdtDepth > 0 ? { inSdt: true } : {}),
        });
      } else pStart = m.index;
      continue;
    }
    if (tag === '</w:p>' && pStart >= 0) {
      const xml = source.slice(pStart, m.index + tag.length);
      const pPr = firstElement(xml, 'w:pPr') ?? '';
      out.push({
        kind: 'paragraph',
        text: textOf(xml),
        runs: parseRuns(xml).map((r) => ({
          text: r.text, bold: r.bold, italic: r.italic, underline: r.underline,
        })),
        style: val(pPr, 'w:pStyle'),
        align: val(pPr, 'w:jc'),
        structural: STRUCTURAL_TAGS.some((t) => new RegExp('<' + t + '\\b').test(xml)),
        ...(sdtDepth > 0 ? { inSdt: true } : {}),
      });
      pStart = -1;
    }
  }
  return out;
}

/**
 * Parse one `<w:tbl>` into a render-ready shape.
 *
 * Merges are reported, not resolved: a cell says `vMerge: 'restart' | 'continue'`
 * and the renderer turns a run of continuations into a rowspan. Resolving here
 * would mean inventing a grid the file does not contain.
 */
export function parseTable(tblXml) {
  const xml = String(tblXml);
  // Scan the table's CONTENTS, not the element: the guard counts `w:tbl` depth,
  // so leaving the outer tag in would make the table nested inside itself.
  const open = /^<w:tbl\b[^>]*>/.exec(xml);
  const inner = open ? xml.slice(open[0].length, xml.lastIndexOf('</w:tbl>')) : xml;

  // `tblPr` and `tblGrid` both precede the first row, so reading only the head
  // stops a nested table's properties being mistaken for this one's.
  const head = headBefore(inner, /<w:tr\b/);
  const tblPr = firstElement(head, 'w:tblPr');
  const grid = firstElement(head, 'w:tblGrid');

  const columns = grid
    ? childElements(grid, 'w:gridCol').map((c) => twipsToPx(attrs(c)['w:w']))
    : [];

  const rows = childElements(inner, 'w:tr', 'w:tbl').map((trXml, rowIndex) => {
    const trPr = firstElement(headBefore(trXml, /<w:tc\b/), 'w:trPr');
    const cells = childElements(trXml, 'w:tc', 'w:tbl').map((tcXml, cellIndex) => {
      const tcPr = firstElement(headBefore(tcXml, /<w:(p|tbl)\b/), 'w:tcPr');
      const vMergeEl = tcPr ? firstElement(tcPr, 'w:vMerge') : null;
      const widthEl = tcPr ? firstElement(tcPr, 'w:tcW') : null;
      const widthAttrs = widthEl ? attrs(widthEl) : {};
      return {
        index: cellIndex,
        gridSpan: Number(tcPr ? val(tcPr, 'w:gridSpan') : null) || 1,
        // `<w:vMerge/>` with no w:val means continue — the spec's default.
        vMerge: vMergeEl ? (attrs(vMergeEl)['w:val'] ?? 'continue') : null,
        widthPx: widthAttrs['w:type'] === 'dxa' ? twipsToPx(widthAttrs['w:w']) : null,
        widthPct: widthAttrs['w:type'] === 'pct' ? Number(widthAttrs['w:w']) / 50 : null,
        vAlign: tcPr ? val(tcPr, 'w:vAlign') : null,
        shading: readShading(tcPr),
        borders: readBorders(tcPr ? firstElement(tcPr, 'w:tcBorders') : null),
        blocks: cellBlocks(tcXml),
        text: cellBlocks(tcXml).filter((b) => b.kind === 'paragraph').map((b) => b.text).join('\n'),
      };
    });

    return {
      index: rowIndex,
      // A header row repeats across pages; it is also the row a reader expects
      // to look different, so the renderer is told about it.
      header: Boolean(trPr && firstElement(trPr, 'w:tblHeader')),
      heightPx: trPr ? twipsToPx(val(trPr, 'w:trHeight')) || null : null,
      cells,
    };
  });

  const widthEl = tblPr ? firstElement(tblPr, 'w:tblW') : null;
  const widthAttrs = widthEl ? attrs(widthEl) : {};

  return {
    kind: 'table',
    columns,
    rows,
    align: tblPr ? val(tblPr, 'w:jc') : null,
    widthPx: widthAttrs['w:type'] === 'dxa' ? twipsToPx(widthAttrs['w:w']) : null,
    // pct is in fiftieths of a percent, so 5000 means 100%.
    widthPct: widthAttrs['w:type'] === 'pct' ? Number(widthAttrs['w:w']) / 50 : null,
    borders: readBorders(tblPr ? firstElement(tblPr, 'w:tblBorders') : null),
    shading: readShading(tblPr),
    style: tblPr ? val(tblPr, 'w:tblStyle') : null,
    rowCount: rows.length,
    columnCount: rows.reduce(
      (n, r) => Math.max(n, r.cells.reduce((c, cell) => c + cell.gridSpan, 0)),
      columns.length,
    ),
  };
}

/**
 * Page geometry from `w:sectPr` — the size and margins the author set.
 *
 * Without this a document renders as a ribbon of text with no edges, which is
 * the single most obvious way an editor announces it is not a word processor.
 * Landscape is `w:orient`, but the width and height are already swapped in the
 * file when it is set, so it is reported rather than applied twice.
 */
export function parseSection(bodyXml) {
  const sectPr = firstElement(bodyXml, 'w:sectPr');
  const pgSz = sectPr ? firstElement(sectPr, 'w:pgSz') : null;
  const pgMar = sectPr ? firstElement(sectPr, 'w:pgMar') : null;
  const sz = pgSz ? attrs(pgSz) : {};
  const mar = pgMar ? attrs(pgMar) : {};

  // A4 is the default here rather than US Letter: this product's first users
  // are in Pakistan and the Gulf, and a wrong default should be wrong for the
  // fewest people.
  const widthTwips = Number(sz['w:w']) || 11906;
  const heightTwips = Number(sz['w:h']) || 16838;

  const margin = (name, fallback) => {
    const v = mar['w:' + name];
    return twipsToPx(v === undefined ? fallback : v);
  };

  return {
    widthPx: twipsToPx(widthTwips),
    heightPx: twipsToPx(heightTwips),
    orientation: sz['w:orient'] === 'landscape' || widthTwips > heightTwips ? 'landscape' : 'portrait',
    margins: {
      top: margin('top', 1440),
      right: margin('right', 1440),
      bottom: margin('bottom', 1440),
      left: margin('left', 1440),
      header: margin('header', 720),
      footer: margin('footer', 720),
      gutter: margin('gutter', 0),
    },
    // Everything the renderer actually lays text into.
    get contentWidthPx() {
      return this.widthPx - this.margins.left - this.margins.right - this.margins.gutter;
    },
    // Which header/footer references actually apply. Without titlePg a file can
    // carry a first-page header Word never shows.
    titlePage: /<w:titlePg\b[^>]*\/?>/.test(String(bodyXml)),
    evenAndOdd: /<w:evenAndOddHeaders\b[^>]*\/?>/.test(String(bodyXml)),
    sectPrXml: sectPr,
    declared: Boolean(sectPr),
  };
}

export { childElements, firstElement, headBefore, readBorders, readShading, cellBlocks };
