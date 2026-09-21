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
/** A paragraph's borders: one line a side, in the file's weight and colour. */
function drawParagraphBorders(page, borders, { xPx, yPx, widthPx, heightPx }) {
  if (!borders) return;
  const x1 = xPx * PT;
  const x2 = (xPx + widthPx) * PT;
  const y1 = yPx * PT;
  const y2 = (yPx + heightPx) * PT;
  for (const side of ['top', 'bottom', 'left', 'right']) {
    const b = borders[side];
    if (!b || !b.style || b.style === 'none' || b.style === 'nil') continue;
    const width = Math.max(0.5, (Number(b.widthPx) || 1) * 0.75);
    const colour = b.colour && /^#?[0-9a-fA-F]{6}$/.test(String(b.colour)) ? `#${String(b.colour).replace('#', '')}` : '#000000';
    // A double border is two lines of the weight, a weight apart, the inner
    // one inside the box.
    const offsets = b.style === 'double' ? [0, width * 2] : [0];
    for (const o of offsets) {
      if (side === 'top') page.line(x1, y1 + o, x2, y1 + o, { width, colour });
      else if (side === 'bottom') page.line(x1, y2 - o, x2, y2 - o, { width, colour });
      else if (side === 'left') page.line(x1 + o, y1, x1 + o, y2, { width, colour });
      else page.line(x2 - o, y1, x2 - o, y2, { width, colour });
    }
  }
}

/**
 * The page borders: a box on the sheet, each side in from the page edge by
 * its space (Word's default), or out from the text by it — that is, the
 * margin less the space. Drawn with the paragraph-border pen.
 */
function drawPageBorders(page, borders, section) {
  const m = section.margins;
  const inset = (side) => {
    const b = borders[side];
    const spacePx = ((b && b.spacePt) || 0) * (96 / 72);
    if (borders.offsetFrom !== 'text') return spacePx;
    const margin = side === 'top' ? m.top : side === 'bottom' ? m.bottom : side === 'left' ? m.left + (m.gutter || 0) : m.right;
    return Math.max(0, margin - spacePx);
  };
  const x1 = inset('left');
  const y1 = inset('top');
  drawParagraphBorders(page, borders, { xPx: x1, yPx: y1, widthPx: section.widthPx - inset('right') - x1, heightPx: section.heightPx - inset('bottom') - y1 });
}

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
    // A line beside a floating picture is narrower than its column, and one
    // beside a left float starts further in; the paginator says by how much.
    const room = (line.widthPx ?? widthPx) * PT;
    let x = (xPx + (line.offsetPx || 0)) * PT;
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

  // Line numbers count every body line; the count starts again on each
  // page when the section says so.
  const numbering = section.lineNumbers || null;
  let lineNo = numbering ? numbering.start || 1 : 1;

  for (const sheet of frame.pages.pages) {
    const page = doc.addPage();
    if (numbering && numbering.restart === 'newPage') lineNo = numbering.start || 1;
    // The page colour under everything, edge to edge, as Word prints it
    // when asked to print background colours.
    if (section.background) page.rect(0, 0, section.widthPx * PT, section.heightPx * PT, { fill: section.background });
    // The page borders next, so the text and the bands draw over them.
    if (section.pageBorders) drawPageBorders(page, section.pageBorders, section);
    // The watermark first, so everything else draws over it: the header's
    // WordArt, rising across the page in the grey Word draws it in.
    if (sheet.watermark?.text) drawWatermark(page, doc, sheet.watermark, section);
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
        // Row by row: a paragraph of only pictures has several to a row,
        // each standing on the row's baseline, the row placed by the
        // paragraph's alignment; an anchored picture on its own keeps the
        // side it asked for. A fragment without rows is one picture a row.
        const rows = fr.rows || fr.images.map((img) => ({ count: 1, heightPx: img.heightPx }));
        let at = 0;
        for (const row of rows) {
          const imgs = fr.images.slice(at, at + row.count);
          at += row.count;
          const rowWidth = imgs.reduce((s, img) => s + img.widthPx, 0);
          const side = imgs.length === 1 && imgs[0].anchored ? imgs[0].hAlign : fr.align === 'center' || fr.align === 'right' ? fr.align : 'left';
          let x = xPx + Math.max(0, side === 'center' ? (widthPx - rowWidth) / 2 : side === 'right' ? widthPx - rowWidth : 0);
          for (const img of imgs) {
            drawImage(page, doc, img, x, y + Math.max(0, row.heightPx - img.heightPx));
            x += img.widthPx;
          }
          y += row.heightPx + IMAGE_GAP;
        }
        continue;
      }
      if (fr.kind === 'float') {
        // Beside the words, at the side it asked for, where the paginator
        // put it — the lines round it were laid out shorter to leave the
        // room. It advances nothing: the words carry on beside it.
        const dx = fr.side === 'right' ? widthPx - fr.widthPx : 0;
        drawImage(page, doc, fr.image, xPx + Math.max(0, dx), m.top + fr.topPx);
        continue;
      }
      if (fr.kind === 'textbox') {
        drawTextBox(page, doc, fr, { xPx, yPx: y, widthPx });
        y += fr.heightPx + IMAGE_GAP;
        continue;
      }
      if (fr.kind === 'floatbox') {
        // A text box beside the words, at the side it asked for, where the
        // paginator put it; the lines round it were laid out shorter. It
        // advances nothing, like a floating picture.
        const bx = fr.side === 'right' ? xPx + Math.max(0, widthPx - fr.widthPx) : xPx;
        drawTextBox(page, doc, { ...fr, hAlign: null }, { xPx: bx, yPx: m.top + fr.topPx, widthPx: fr.widthPx });
        continue;
      }
      if (fr.kind === 'note') {
        // An endnote: its number in the gutter, its words indented past it.
        y += fr.spaceBefore || 0;
        drawNoteNumber(page, doc, fr.note, xPx, y, fr);
        y += drawParagraphLines(page, doc, {
          lines: fr.lines, fragment: fr, runs: fr.runs, xPx: xPx + NOTE_INDENT_PX, yPx: y, widthPx: widthPx - NOTE_INDENT_PX,
        });
        y += fr.spaceAfter || 0;
        continue;
      }
      y += fr.spaceBefore || 0;
      const block = byIndex.get(fr.paragraphIndex) || null;
      // The paragraph's shading and borders, edge to edge of the column,
      // over the lines it has on this page — as the screen draws them.
      if (block && (block.shading || block.borders)) {
        const heightPx = fr.lines.length * fr.lineHeightPx;
        if (block.shading && /^#?[0-9a-fA-F]{6}$/.test(String(block.shading))) {
          page.rect(xPx * PT, y * PT, widthPx * PT, heightPx * PT, { fill: `#${String(block.shading).replace('#', '')}` });
        }
        drawParagraphBorders(page, block.borders, { xPx, yPx: y, widthPx, heightPx });
      }
      // The line numbers down the left margin: every line counted, every
      // countBy-th shown, in the gap Word leaves before the text.
      if (numbering && fr.lines && !(block && block.container)) {
        const gapPx = numbering.distancePx != null ? numbering.distancePx : 24;
        fr.lines.forEach((line, i) => {
          const n = lineNo++;
          if (n % numbering.countBy !== 0) return;
          const label = String(n);
          const size = 8;
          const baseline = (y + i * fr.lineHeightPx + fr.lineHeightPx * BASELINE) * PT;
          page.text(label, (xPx - gapPx) * PT - label.length * size * 0.556, baseline, { font: 'Helvetica', size, colour: '#666666' });
        });
      }
      y += drawParagraphLines(page, doc, {
        lines: fr.lines, fragment: fr, runs: block ? block.runs : null,
        xPx: xPx + (fr.indent || 0), yPx: y, widthPx: widthPx - (fr.indent || 0),
        listLabel: fr.listLabel || null, lastIsFinal: Boolean(fr.last),
      });
      y += fr.spaceAfter || 0;
    }

    // The page's footnotes, at its foot: a short rule, then each note with
    // its number in the gutter — in the room the paginator kept for them.
    if (sheet.notes?.length) {
      let ny = m.top + (sheet.contentHeightPx ?? (section.heightPx - m.top - m.bottom)) - sheet.notesHeightPx + 4;
      page.line(xPx * PT, ny * PT, (xPx + widthPx / 3) * PT, ny * PT, { width: 0.6, colour: '#333333' });
      ny += NOTE_RULE_PX - 4;
      for (const note of sheet.notes) {
        for (const p of note.paragraphs) {
          ny += p.spaceBefore || 0;
          if (p === note.paragraphs[0]) drawNoteNumber(page, doc, note.n, xPx, ny, p);
          ny += drawParagraphLines(page, doc, { lines: p.lines, fragment: p, runs: p.runs, xPx: xPx + NOTE_INDENT_PX, yPx: ny, widthPx: widthPx - NOTE_INDENT_PX });
          ny += p.spaceAfter || 0;
        }
        ny += 2;
      }
    }
  }
  return { buffer: doc.toBuffer(), pages: frame.pages.pages.length };
}

const NOTE_INDENT_PX = 18;
const NOTE_RULE_PX = 14;
const BOX_PAD_PX = 7;

/** A note's number, small and raised, in the gutter before its words. */
function drawNoteNumber(page, doc, n, xPx, yPx, fragment) {
  const size = Math.max(6, ((fragment.sizePx || BAND_SIZE_PX) * PT) * 0.7);
  const baseline = (yPx + (fragment.lineHeightPx || lineHeightOf(fragment.sizePx || BAND_SIZE_PX)) * BASELINE - 3) * PT;
  page.text(String(n), xPx * PT, baseline, { font: 'Helvetica', size, colour: fragment.colour || null });
}

/**
 * A text box: its frame — filled and outlined as the shape says, placed by
 * its alignment — and its paragraphs inside the padding, painted like the
 * body's from the lines the paginator laid.
 */
function drawTextBox(page, doc, fr, { xPx, yPx, widthPx }) {
  const bx = fr.hAlign === 'center' ? xPx + (widthPx - fr.widthPx) / 2 : fr.hAlign === 'right' ? xPx + widthPx - fr.widthPx : xPx;
  const fill = fr.fill && /^#[0-9a-fA-F]{6}$/.test(fr.fill) ? fr.fill : null;
  const stroke = fr.line && /^#[0-9a-fA-F]{6}$/.test(fr.line) ? fr.line : null;
  if (fill || stroke) page.rect(bx * PT, yPx * PT, fr.widthPx * PT, fr.heightPx * PT, { fill, stroke, width: 0.6 });
  let y = yPx + BOX_PAD_PX;
  for (const p of fr.paragraphs || []) {
    y += p.spaceBefore || 0;
    y += drawParagraphLines(page, doc, {
      lines: p.lines, fragment: p, runs: p.runs, xPx: bx + BOX_PAD_PX + (p.indent || 0), yPx: y, widthPx: fr.widthPx - 2 * BOX_PAD_PX - (p.indent || 0),
    });
    y += p.spaceAfter || 0;
  }
}

/** The watermark: big, grey, rising across the middle of the page. */
function drawWatermark(page, doc, watermark, section) {
  const text = String(watermark.text || '').trim();
  if (!text) return;
  const size = Math.min(150, Math.max(40, (section.widthPx * PT * 1.1) / Math.max(4, text.length)));
  const width = doc.widthOf(text, { font: 'Helvetica', size });
  const rotate = 45;
  const rad = (rotate * Math.PI) / 180;
  const cx = (section.widthPx * PT) / 2;
  const cy = (section.heightPx * PT) / 2;
  // The baseline runs up and to the right; start it half a width back along
  // that line so the words sit centred on the page.
  const x = cx - (width / 2) * Math.cos(rad);
  const y = cy + (width / 2) * Math.sin(rad) + size * 0.35;
  const colour = watermark.colour && /^#[0-9a-fA-F]{6}$/.test(watermark.colour) ? watermark.colour : '#C8C8C8';
  page.text(text, x, y, { font: 'Helvetica', size, colour, rotate });
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
