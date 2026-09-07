/**
 * The paginated frame, printed — through `@rutba/pdf`, the estate's own writer.
 *
 * The paginator already decided everything a printer needs: which lines fall
 * on which sheet, at what size and weight, with what space above and below,
 * where a table breaks and which rows repeat. The shell draws that frame into
 * a contenteditable; this file draws the same frame into a PDF. Two renderers,
 * one layout — measurement and paint still agree by construction, and the
 * PDF cannot break a line the screen did not.
 *
 * WHAT IS EXACT, AND WHAT IS NOT. The width table the paginator measures with
 * is the base-14 Helvetica advance table rounded to two places, and the PDF
 * draws with Helvetica itself — so a line the paginator fitted fits here, and
 * the safety margin it keeps (WRAP_SAFETY) covers the rounding. What the PDF
 * cannot do is draw a font the file named: everything is Helvetica (Courier
 * for a monospace name) until the writer learns embedding, and the plan says
 * so. A document that needs its own typeface is exported by a path that has
 * one, not by this file pretending.
 *
 * Images: PNG is embedded (the writer decodes it); anything else — a chart or
 * shape the backend rendered as SVG, a JPEG — is drawn as a labelled frame of
 * the right size, so the page keeps its geometry and the reader sees that
 * something was there. Honest, not pretty.
 *
 * This is the ONE file in the editor that imports the PDF writer. The
 * Workspace seam test allows it by name: `@rutba/pdf` is a shared package
 * with no Workspace dependency, so nothing here ties the editor to a product.
 */
import { PdfDocument, decodePng, isPng } from '@rutba/pdf';
import { layoutParagraph, paginate, rowHeight } from '../paginate.js';
import { computeListLabels } from '../lists.js';
import { bandForPage, resolveFields } from '../bands.js';
import { lineHeight as lineHeightOf } from '@rutba/drawing';

/** CSS pixels (96dpi, the paginator's unit) to points (72dpi, the page's). */
const PT = 0.75;

/** Where the baseline sits in a line box, as a fraction of the box height. */
const BASELINE = 0.76;

/** Word's highlighter palette — the same map the shell paints with. */
const HIGHLIGHTS = {
  yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF',
  blue: '#0000FF', red: '#FF0000', darkBlue: '#00008B', darkCyan: '#008B8B',
  darkGreen: '#006400', darkMagenta: '#8B008B', darkRed: '#8B0000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000',
};

/** A4 in CSS pixels, for a flow with no section of its own (an email body). */
export const DEFAULT_SECTION = Object.freeze({
  widthPx: 794, heightPx: 1123,
  margins: { top: 96, right: 96, bottom: 96, left: 96, gutter: 0, header: 48, footer: 48 },
  contentWidthPx: 602,
  titlePage: false, evenAndOdd: false,
});

const CELL_PADDING = 8;
const IMAGE_GAP = 8;
const TABLE_SPACE_AFTER = 12;
const BAND_SIZE_PX = 14;

const isMono = (name) => /courier|consolas|mono|menlo/i.test(String(name || ''));

function fontFor({ bold = false, italic = false, mono = false } = {}) {
  if (mono) return bold ? 'Courier-Bold' : 'Courier';
  if (bold) return italic ? 'Helvetica-BoldOblique' : 'Helvetica-Bold';
  return italic ? 'Helvetica-Oblique' : 'Helvetica';
}

/** A run's look as the writer wants it: font name, size in points, colour. */
function styleOfRun(run, fragment) {
  const bold = Boolean(run.bold) || fragment.weight === 'bold';
  const italic = Boolean(run.italic) || Boolean(fragment.italic);
  const size = run.fontSize ? Number(run.fontSize) : (fragment.sizePx || BAND_SIZE_PX) * PT;
  const colour = run.fontColour ? `#${String(run.fontColour).replace('#', '')}` : (fragment.colour || null);
  return {
    font: fontFor({ bold, italic, mono: isMono(run.fontName) }),
    size,
    colour: colour && /^#[0-9a-fA-F]{6}$/.test(colour) ? colour : null,
    underline: Boolean(run.underline),
    strike: Boolean(run.strike),
    highlight: run.highlight && HIGHLIGHTS[run.highlight] ? HIGHLIGHTS[run.highlight] : null,
  };
}

/**
 * The runs covering [from, to) of a paragraph, split at the boundaries — the
 * painter's slice, over the serialised runs a frame carries.
 */
function sliceRunSegments(runs, from, to) {
  const out = [];
  let seen = 0;
  for (const run of runs || []) {
    const start = seen;
    const end = seen + run.text.length;
    seen = end;
    if (end <= from || start >= to) continue;
    const text = run.text.slice(Math.max(0, from - start), Math.min(run.text.length, to - start));
    if (text !== '') out.push({ ...run, text });
  }
  return out;
}

/** Width of a segment list, in points, as the writer will draw it. */
function widthOfSegments(doc, segments, fragment) {
  let total = 0;
  for (const seg of segments) {
    const s = styleOfRun(seg, fragment);
    total += doc.widthOf(seg.text, { font: s.font, size: s.size });
  }
  return total;
}

/**
 * Draw one laid-out line. `x` and `baseline` are in points. Justification
 * spreads `extraPerSpace` points across every space in the line.
 */
function drawSegments(page, doc, segments, x, baseline, fragment, { extraPerSpace = 0 } = {}) {
  let cursor = x;
  for (const seg of segments) {
    const s = styleOfRun(seg, fragment);
    // Word by word when justified, so the spaces can grow; whole when not.
    const parts = extraPerSpace > 0 ? seg.text.split(/( )/).filter((p) => p !== '') : [seg.text];
    const startX = cursor;
    for (const part of parts) {
      const w = doc.widthOf(part, { font: s.font, size: s.size });
      if (part === ' ') { cursor += w + extraPerSpace; continue; }
      if (s.highlight) page.rect(cursor, baseline - s.size * 0.8, w, s.size * 1.05, { fill: s.highlight });
      page.text(part, cursor, baseline, { font: s.font, size: s.size, colour: s.colour });
      cursor += w;
    }
    if (s.underline) page.line(startX, baseline + s.size * 0.12, cursor, baseline + s.size * 0.12, { width: Math.max(0.4, s.size * 0.06), colour: s.colour || '#000000' });
    if (s.strike) page.line(startX, baseline - s.size * 0.3, cursor, baseline - s.size * 0.3, { width: Math.max(0.4, s.size * 0.06), colour: s.colour || '#000000' });
  }
  return cursor - x;
}

/**
 * Lines of one paragraph fragment, left-aligned, centred, right-aligned or
 * justified inside `widthPx` starting at `xPx`. Returns the height used (px).
 */
function drawParagraphLines(page, doc, { lines, fragment, runs, xPx, yPx, widthPx, listLabel = null, lastIsFinal = true }) {
  const lineHeightPx = fragment.lineHeightPx;
  const align = fragment.align || null;
  lines.forEach((line, i) => {
    const top = yPx + i * lineHeightPx;
    const baseline = (top + lineHeightPx * BASELINE) * PT;
    const segments = line.text === ''
      ? []
      : (runs && runs.length ? sliceRunSegments(runs, line.start, line.end) : [{ text: line.text }]);
    if (!segments.length) return;
    const lineWidth = widthOfSegments(doc, segments, fragment);
    const room = widthPx * PT;
    let x = xPx * PT;
    let extraPerSpace = 0;
    if (align === 'center') x += Math.max(0, (room - lineWidth) / 2);
    else if (align === 'right' || align === 'end') x += Math.max(0, room - lineWidth);
    else if ((align === 'both' || align === 'justify') && !(lastIsFinal && i === lines.length - 1)) {
      const spaces = segments.reduce((n, seg) => n + (seg.text.match(/ /g) || []).length, 0);
      if (spaces > 0 && room > lineWidth) extraPerSpace = (room - lineWidth) / spaces;
    }
    if (i === 0 && listLabel) {
      // The label hangs in the gutter, the way Word draws it: its right edge a
      // little short of the text edge.
      const s = styleOfRun({}, fragment);
      const w = doc.widthOf(listLabel, { font: s.font, size: s.size });
      page.text(listLabel, Math.max(0, xPx * PT - w - 6 * PT), baseline, { font: s.font, size: s.size, colour: s.colour });
    }
    drawSegments(page, doc, segments, x, baseline, fragment, { extraPerSpace });
  });
  return lines.length * lineHeightPx;
}

/** A header or footer band: plain paragraphs at the body size. */
function drawBand(page, doc, paragraphs, { xPx, yPx, widthPx }) {
  const lh = lineHeightOf(BAND_SIZE_PX);
  let y = yPx;
  for (const p of paragraphs || []) {
    const fragment = { sizePx: BAND_SIZE_PX, lineHeightPx: lh, align: p.align || null, weight: 'normal' };
    const text = p.text ?? (p.runs || []).map((r) => r.text).join('');
    drawParagraphLines(page, doc, {
      lines: [{ text, start: 0, end: text.length }], fragment, runs: p.runs || [], xPx, yPx: y, widthPx,
    });
    y += lh;
  }
  return y - yPx;
}

/** One image under a paragraph: embedded when PNG, framed and named otherwise. */
function drawImage(page, doc, img, xPx, yPx) {
  const w = img.widthPx * PT;
  const h = img.heightPx * PT;
  const x = xPx * PT;
  const y = yPx * PT;
  const bytes = dataUriBytes(img.href);
  if (bytes && isPng(bytes)) {
    try {
      const ref = doc.addImage(decodePng(bytes));
      page.image(ref, x, y, w, h);
      return;
    } catch { /* an unsupported PNG variant: fall through to the frame */ }
  }
  page.rect(x, y, w, h, { stroke: '#bbbbbb', width: 0.5, fill: '#f4f4f4' });
  const label = String(img.name || 'image').slice(0, 60);
  page.text(label, x + 4, y + Math.min(h - 2, 12), { font: 'Helvetica-Oblique', size: 7, colour: '#888888' });
}

function dataUriBytes(href) {
  const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(String(href || ''));
  if (!m) return null;
  try {
    return m[2] ? Buffer.from(m[3], 'base64') : Buffer.from(decodeURIComponent(m[3]), 'utf8');
  } catch {
    return null;
  }
}

/** A cell's or table's border on one side, as the shell resolves it. */
function borderOf(borders, side) {
  if (!borders) return undefined;
  return borders[side] ?? (side === 'top' || side === 'bottom' ? borders.insideH : borders.insideV);
}

/**
 * Draw the rows of a table fragment from (xPx, yPx); returns the height used
 * in px. Breaks between rows are the paginator's business — this draws what
 * landed on the sheet, header repeats included.
 */
function drawTable(page, doc, table, rows, { xPx, yPx, widthPx, labelOf, depth = 0 }) {
  const columns = table.columns && table.columns.length
    ? table.columns
    : Array(Math.max(1, table.columnCount || 1)).fill(widthPx / Math.max(1, table.columnCount || 1));
  let y = yPx;
  for (const row of rows) {
    const h = rowHeight(row, table, widthPx, null);
    let cx = xPx;
    let column = 0;
    for (const cell of row.cells) {
      const span = Math.max(1, cell.gridSpan || 1);
      const cellWidth = columns.slice(column, column + span).reduce((a, b) => a + b, 0) || widthPx;
      column += span;
      if (cell.vMerge === 'continue') { cx += cellWidth; continue; }
      if (cell.shading && /^#?[0-9a-fA-F]{6}$/.test(String(cell.shading))) {
        page.rect(cx * PT, y * PT, cellWidth * PT, h * PT, { fill: `#${String(cell.shading).replace('#', '')}` });
      }
      const borders = cell.borders ?? table.borders ?? null;
      const sides = [
        ['top', cx, y, cx + cellWidth, y], ['bottom', cx, y + h, cx + cellWidth, y + h],
        ['left', cx, y, cx, y + h], ['right', cx + cellWidth, y, cx + cellWidth, y + h],
      ];
      for (const [side, x1, y1, x2, y2] of sides) {
        const b = borderOf(borders, side);
        if (b && b.widthPx === 0) continue;
        const colour = b && b.colour && /^#?[0-9a-fA-F]{6}$/.test(String(b.colour)) ? `#${String(b.colour).replace('#', '')}` : '#9a9a9a';
        page.line(x1 * PT, y1 * PT, x2 * PT, y2 * PT, { width: b && b.widthPx ? Math.max(0.4, b.widthPx * PT) : 0.5, colour });
      }
      let cy = y + CELL_PADDING;
      for (const b of cell.blocks || []) {
        if (b.kind === 'table') {
          if (depth < 3) cy += drawTable(page, doc, b.table, b.table.rows, { xPx: cx + CELL_PADDING, yPx: cy, widthPx: cellWidth - CELL_PADDING * 2, labelOf, depth: depth + 1 });
          continue;
        }
        const laid = layoutParagraph(b, cellWidth - CELL_PADDING * 2, { cache: null });
        const label = labelOf(b.blockIndex);
        const fragment = {
          sizePx: laid.style.sizePx, lineHeightPx: laid.lineHeightPx, weight: laid.style.weight,
          italic: laid.style.italic, colour: laid.style.colour, align: b.align || laid.style.align || null,
        };
        const runs = label ? [{ text: `${label.label} ` }, ...(b.runs || [])] : (b.runs || []);
        const lines = label
          ? laid.lines.map((l, i) => (i === 0 ? { ...l, end: l.end + label.label.length + 1 } : { ...l, start: l.start + label.label.length + 1, end: l.end + label.label.length + 1 }))
          : laid.lines;
        cy += drawParagraphLines(page, doc, {
          lines, fragment, runs, xPx: cx + CELL_PADDING + (laid.indentPx || 0), yPx: cy, widthPx: cellWidth - CELL_PADDING * 2,
        });
      }
      cx += cellWidth;
    }
    y += h;
  }
  return y - yPx;
}

/**
 * Render a paginated frame — `{ blocks, section, pages, listLabels? }`, the
 * shape a document session hands the shell — to PDF bytes.
 *
 * @returns {{ buffer: Buffer, pages: number }}
 */
export function renderFramePdf(frame, { title = '', author = '', created = null } = {}) {
  if (!frame || !frame.pages || !frame.pages.pages) throw new Error('renderFramePdf needs a paginated frame (frame.pages)');
  const section = frame.section || DEFAULT_SECTION;
  const m = section.margins;
  const byIndex = new Map((frame.blocks || []).map((b) => [b.index, b]));
  const labels = frame.listLabels || null;
  const labelOf = (i) => (labels instanceof Map ? labels.get(i) : (labels ? labels[i] : null)) || null;

  const doc = new PdfDocument({ size: [section.widthPx * PT, section.heightPx * PT], title, author, created });
  const xPx = m.left + (m.gutter || 0);
  const widthPx = section.contentWidthPx;

  for (const sheet of frame.pages.pages) {
    const page = doc.addPage();
    if (sheet.header) drawBand(page, doc, sheet.header, { xPx, yPx: m.header || m.top / 2, widthPx });
    if (sheet.footer) {
      const lh = lineHeightOf(BAND_SIZE_PX);
      const rows = (sheet.footer || []).length;
      drawBand(page, doc, sheet.footer, { xPx, yPx: section.heightPx - (m.footer || m.bottom / 2) - rows * lh, widthPx });
    }

    let y = m.top;
    for (const fr of sheet.fragments) {
      if (fr.kind === 'table') {
        y += drawTable(page, doc, fr.table, fr.rows, { xPx, yPx: y, widthPx, labelOf });
        if (!fr.continues) y += TABLE_SPACE_AFTER;
        continue;
      }
      if (fr.kind === 'images') {
        for (const img of fr.images) {
          drawImage(page, doc, img, xPx, y);
          y += img.heightPx + IMAGE_GAP;
        }
        continue;
      }
      y += fr.spaceBefore || 0;
      const block = byIndex.get(fr.paragraphIndex) || null;
      y += drawParagraphLines(page, doc, {
        lines: fr.lines, fragment: fr, runs: block ? block.runs : null,
        xPx: xPx + (fr.indent || 0), yPx: y, widthPx: widthPx - (fr.indent || 0),
        listLabel: fr.listLabel || null, lastIsFinal: Boolean(fr.last),
      });
      y += fr.spaceAfter || 0;
    }
  }
  return { buffer: doc.toBuffer(), pages: frame.pages.pages.length };
}

/**
 * Render a DocView — paginating it first if its backend has no pages (Mail's
 * HTML body), on the default A4 sheet or the one given.
 */
export function renderPdf(view, { title = '', author = '', created = null, section = null } = {}) {
  const numbering = typeof view.doc.numberingDefs === 'function' ? view.doc.numberingDefs() : null;
  const listLabels = computeListLabels(view.flow, view.blocks, numbering);
  let pages = view.section ? view.pages : null;
  let sheet = view.section || null;
  if (!pages) {
    sheet = section || DEFAULT_SECTION;
    const styles = typeof view.doc.paragraphStyles === 'function' ? view.doc.paragraphStyles() : null;
    pages = paginate({ flow: view.flow, blocks: view.blocks, section: sheet, styles, listLabels });
    const bands = typeof view.doc.headerFooters === 'function' ? view.doc.headerFooters() : { headers: {}, footers: {} };
    for (const page of pages.pages) {
      const header = bandForPage(bands.headers || {}, page.number, {});
      const footer = bandForPage(bands.footers || {}, page.number, {});
      page.header = header ? resolveFields(header.paragraphs, { page: page.number, of: pages.count }) : null;
      page.footer = footer ? resolveFields(footer.paragraphs, { page: page.number, of: pages.count }) : null;
    }
  }
  return renderFramePdf({ blocks: view.blocks, section: sheet, pages, listLabels }, { title, author, created });
}
