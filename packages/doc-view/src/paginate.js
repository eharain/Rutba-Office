/**
 * Pagination — the last thing between this and a word processor.
 *
 * A document has been rendering as one endless page. Everything else about it
 * was right — the paper had the author's own size and margins, the tables drew,
 * the text was editable — but content ran off the bottom of the first sheet and
 * kept going, which is the difference between "a document viewer" and "a page of
 * A4 you could print".
 *
 * THE SERVER OWNS LINE BREAKING. That is the decision everything else follows
 * from. The obvious alternative — let the browser wrap, then ask it where the
 * lines fell — puts layout in the one place we cannot test in Node, and makes
 * pagination a different answer in every browser. So the paginator breaks lines
 * itself and hands the shell the lines it must draw. Measurement and rendering
 * then agree BY CONSTRUCTION rather than by luck: there is no second opinion to
 * disagree with.
 *
 * The cost is honest and worth naming: our metrics come from a width table
 * (`@rutba/drawing/measure`), not from the font the reader actually has. Line
 * breaks are therefore close to, not identical to, what Word would choose. For a
 * screen that is unnoticeable. For print-exact fidelity it is not good enough,
 * and closing that gap means shaping text with real font metrics — a much bigger
 * machine, and not one to build before anyone has asked to print.
 *
 * Nothing here touches the model. Pagination is a pure function of (flow,
 * section, metrics) and the caret's address space is untouched — a paragraph
 * split across a page boundary is still one paragraph with one index, and each
 * on-screen fragment carries the character offset it starts at so the shell can
 * map a click back.
 */
import { measureText, wrapText, wrapFirstLine, lineHeight } from '@rutba/drawing';

/** Points to CSS pixels, at 96dpi. Word's sizes are in half-points. */
export const ptToPx = (pt) => (Number(pt) || 0) * 96 / 72;

/**
 * How a paragraph style renders. Mirrors the shell's stylesheet, and must keep
 * mirroring it: if these disagree, text is measured at one size and drawn at
 * another, and the page either overflows or ends early.
 */
export const PARAGRAPH_STYLES = {
  Heading1: { sizePx: 22, weight: 'bold', spaceBefore: 18, spaceAfter: 10 },
  ListParagraph: { sizePx: 14, weight: 'normal', spaceBefore: 0, spaceAfter: 10, indent: 26 },
  default: { sizePx: 14, weight: 'normal', spaceBefore: 0, spaceAfter: 10 },
};

/**
 * Resolve a style name, preferring the DOCUMENT's own resolved styles.
 *
 * `styles` is what the backend read out of `word/styles.xml` — flattened chains,
 * CSS pixels, plain data. The hardcoded table above is now only the fallback for
 * a backend with no style part (an email body), and for the day-one documents
 * written before styles were resolved at all. Without this, a file using
 * `Heading2` or a house style rendered as body text — and was MEASURED as body
 * text, so pagination put its page breaks in the wrong places.
 */
const styleOf = (name, styles = null) => {
  const doc = styles && (styles[name] ?? styles['*default*']);
  if (doc) {
    return {
      sizePx: doc.sizePx ?? PARAGRAPH_STYLES.default.sizePx,
      weight: doc.bold ? 'bold' : 'normal',
      italic: Boolean(doc.italic),
      colour: doc.colour ?? null,
      align: doc.align ?? null,
      spaceBefore: doc.spaceBeforePx ?? 0,
      spaceAfter: doc.spaceAfterPx ?? PARAGRAPH_STYLES.default.spaceAfter,
      indent: doc.indentPx ?? 0,
    };
  }
  return PARAGRAPH_STYLES[name] ?? PARAGRAPH_STYLES.default;
};

/**
 * Wrap to slightly less than the column.
 *
 * Our widths come from a table, and measured against a real browser the error is
 * about half a percent — a 96-character line came out 604.3px in a 602px column.
 * Half a percent sounds ignorable and is not: a line 2px too long re-wraps in the
 * browser, which adds a whole line box, and a page that is one line taller than
 * planned overflows. Every paragraph on the page then compounds it.
 *
 * So the column we measure against is a little narrower than the one we draw
 * into. The cost is a slightly deeper right margin; the benefit is that the
 * height the paginator computed is the height the browser produces.
 */
export const WRAP_SAFETY = 0.985;

/**
 * How many wrapped paragraphs to remember.
 *
 * Bounded because a long editing session would otherwise keep every version of
 * every paragraph the user has ever typed.
 */
const LINE_CACHE_LIMIT = 4000;

/**
 * Break one paragraph into lines that fit `width`.
 *
 * Returns lines as {text, start, end} so a fragment split across a page knows
 * the character range it covers — which is what keeps the caret working when the
 * same paragraph index appears on two sheets.
 *
 * Wrapping is the expensive half of pagination and a keystroke changes exactly
 * one paragraph, so with a cache the other nine hundred are not re-wrapped for
 * nothing. On a fifty-eight-page report that is the difference between 117ms a
 * keystroke and 31ms. Short paragraphs skip the cache: hashing a heading costs
 * more than wrapping it.
 */

export function layoutParagraph(block, width, { style = null, cache = null, styles = null, extraIndentPx = 0 } = {}) {
  // Resolve the style FIRST, then key the cache on the measurable inputs — the
  // resolved size and the usable column — rather than on the style's name: two
  // documents can give the same name different looks.
  const s = styleOf(style ?? block.style, styles);
  // The paragraph's own indent rides on top of the style's — the ribbon's
  // indent buttons write `w:ind` directly, and an indent the layout ignores
  // is an indent the page never shows. Part of the cache key by construction:
  // the key hashes the usable width, and the indent narrows it.
  const indent = (s.indent ?? 0) + (block.indentPx ?? 0) + extraIndentPx;
  const key = cache && block.text.length > 40
    ? s.sizePx + ':' + s.weight + ':' + Math.round(width - indent) + '\u0000' + block.text
    : null;
  if (key) {
    const hit = cache.get(key);
    // Re-inserting moves it to the end of the Map's order, so the entries
    // evicted below are the ones nobody has touched.
    if (hit) { cache.delete(key); cache.set(key, hit); return withDirectLineHeight(hit, block); }
  }
  const result = layoutParagraphUncached(block, width, s, indent);
  if (key) {
    cache.set(key, result);
    if (cache.size > LINE_CACHE_LIMIT) cache.delete(cache.keys().next().value);
  }
  return withDirectLineHeight(result, block);
}

/**
 * A paragraph's own line spacing beats the style's — `w:spacing` set directly
 * on the paragraph, 1.5-line or an exact height. Applied OUTSIDE the line
 * cache on purpose: the cache is keyed on what changes the BREAKS (size,
 * weight, width, text), and line height is not one of those — a cached wrap
 * must not pin the spacing its paragraph had the first time it was measured.
 */
function withDirectLineHeight(result, block) {
  const direct = block.spacing ?? null;
  if (!direct || (direct.lineFactor == null && direct.lineExactPx == null)) return result;
  const lineHeightPx = direct.lineExactPx != null
    ? Math.max(4, direct.lineExactPx)
    : result.lineHeightPx * direct.lineFactor;
  return { ...result, lineHeightPx };
}

function layoutParagraphUncached(block, width, s, indent) {
  const usable = Math.max(24, (width - indent) * WRAP_SAFETY);
  const text = block.text ?? '';

  // An empty paragraph is still a line: it is the blank line the author typed,
  // and swallowing it would close up the space they left.
  // A line break in the paragraph (`w:br`, a "\n" in its text) ends a line
  // wherever it falls — an address's lines stay lines on paper. Each piece
  // between breaks wraps on its own; a paragraph with none wraps as before.
  if (text.includes('\n')) {
    const lines = [];
    let base = 0;
    for (const piece of text.split('\n')) {
      const wrapped = piece === '' ? [''] : wrapText(piece, usable, { size: s.sizePx, weight: s.weight });
      let cursor = 0;
      for (const value of wrapped) {
        const at = value === '' ? cursor : piece.indexOf(value, cursor);
        const start = at < 0 ? cursor : at;
        const end = start + value.length;
        lines.push({ text: value, start: base + start, end: base + end, width: measureText(value, { size: s.sizePx, weight: s.weight }) });
        cursor = end;
      }
      base += piece.length + 1;
    }
    lines[lines.length - 1].end = text.length;
    return { lines, style: s, indentPx: indent, lineHeightPx: lineHeight(s.sizePx) };
  }
  const wrapped = text === '' ? [''] : wrapText(text, usable, { size: s.sizePx, weight: s.weight });

  const lines = [];
  let cursor = 0;
  for (const value of wrapped) {
    // wrapText drops the whitespace it broke on, so the offset is recovered by
    // searching forward rather than by adding lengths — otherwise every line
    // after the first would be off by the number of spaces consumed.
    const at = value === '' ? cursor : text.indexOf(value, cursor);
    const start = at < 0 ? cursor : at;
    const end = start + value.length;
    lines.push({ text: value, start, end, width: measureText(value, { size: s.sizePx, weight: s.weight }) });
    cursor = end;
  }
  // The last line must reach the end of the text, or a caret at the very end of
  // a paragraph would have nowhere to land.
  if (lines.length) lines[lines.length - 1].end = text.length;

  return { lines, style: s, indentPx: indent, lineHeightPx: lineHeight(s.sizePx) };
}

/**
 * Greedy word wrap where some characters are OBJECTS — an inline equation —
 * with a width of their own, keyed by the character's offset in the block.
 * Breaks on spaces as `wrapText` does; `base` is where `value` starts in the
 * block, and the lines come back in block offsets.
 */
export function wrapWithObjects(value, base, width, opts, objects) {
  const text = String(value ?? '');
  if (text === '') return [{ text: '', start: base, end: base, width: 0 }];
  const words = [...text.matchAll(/(?:[^\s]|\t)+/g)].map((m) => ({ at: m.index, word: m[0] }));
  const widthOf = (w, at) => {
    let total = 0;
    for (let j = 0; j < w.length; j++) {
      const box = objects.get(base + at + j);
      total += box ? box.widthPx : measureText(w[j], opts);
    }
    return total;
  };
  const space = measureText(' ', opts);
  const lines = [];
  let from = null;
  let to = 0;
  let used = 0;
  for (const { at, word } of words) {
    const w = widthOf(word, at);
    if (from !== null && used + space + w > width) {
      lines.push({ text: text.slice(from, to), start: base + from, end: base + to, width: used });
      from = null;
    }
    if (from === null) { from = at; used = w; } else used += space + w;
    to = at + word.length;
  }
  if (from !== null) lines.push({ text: text.slice(from, to), start: base + from, end: base + to, width: used });
  if (!lines.length) lines.push({ text: '', start: base, end: base, width: 0 });
  lines[lines.length - 1].end = base + text.length;
  return lines;
}

/** How tall a table row is: its tallest cell, wrapped to the column width. */
function rowHeight(row, table, width, cache) {
  // A row held to its height — a label's — is that tall, whatever it holds.
  if (row.heightRule === 'exact' && row.heightPx) return row.heightPx;
  const pad = cellPadding(table);
  const columns = table.columns.length
    ? table.columns
    : Array(Math.max(1, table.columnCount)).fill(width / Math.max(1, table.columnCount));
  let tallest = MIN_ROW_HEIGHT;
  let column = 0;
  for (const cell of row.cells) {
    const span = Math.max(1, cell.gridSpan);
    const cellWidth = columns.slice(column, column + span).reduce((a, b) => a + b, 0) || width;
    column += span;
    if (cell.vMerge === 'continue') continue;
    let height = pad.top + pad.bottom;
    for (const b of cell.blocks) {
      if (b.kind === 'table') { height += tableHeight(b.table, cellWidth, cache); continue; }
      const { lines, lineHeightPx } = layoutParagraph(b, cellWidth - pad.left - pad.right, { cache });
      height += lines.length * lineHeightPx;
    }
    tallest = Math.max(tallest, height);
  }
  return Math.max(row.heightPx ?? 0, tallest);
}

/**
 * A cell's inner margins: the file's own for a table laid out to fixed
 * widths (a sheet of labels is measured to the millimetre), the paginator's
 * usual padding — half of it above and below — for every other table, as
 * it always was.
 */
export function cellPadding(table) {
  if (table?.layoutFixed && table.cellMarginPx) return table.cellMarginPx;
  return { top: CELL_PADDING / 2, bottom: CELL_PADDING / 2, left: CELL_PADDING, right: CELL_PADDING };
}

const tableHeight = (table, width, cache = null) =>
  table.rows.reduce((total, row) => total + rowHeight(row, table, width, cache), 0);

const MIN_ROW_HEIGHT = 22;
const IMAGE_GAP = 8;
const CELL_PADDING = 8;
/** The room a floating picture keeps from the words when the file names none — Word's own defaults, near enough. */
const FLOAT_GAP_PX = 6;
/** The widest a picture is drawn: the column, or on screen 640 px, whichever is less. */
const FLOAT_MAX_PX = 640;

/**
 * Does a picture stand beside the words? Anchored, wrapped square, tight
 * or through, and not centred — the same rule the screen applies.
 */
export function floatsBeside(img) {
  if (!img?.anchored) return false;
  if (!['square', 'tight', 'through'].includes(img.wrap)) return false;
  return img.hAlign !== 'center';
}

/** A floating picture's drawn size: the file's, capped at the column. */
function floatBox(img, width) {
  const scale = Math.min(1, Math.min(width, FLOAT_MAX_PX) / Math.max(1, img.widthPx || 1));
  return { widthPx: (img.widthPx || 0) * scale, heightPx: (img.heightPx || 0) * scale };
}
/** A table needs its bottom rule; a paragraph does not. */
const TABLE_SPACE_AFTER = 12;

/**
 * Lay the whole flow out onto pages.
 *
 * @param {object}   input
 * @param {Array}    input.flow      document order: paragraph refs and tables
 * @param {Array}    input.blocks    the paragraphs, indexed as the caret sees them
 * @param {object}   input.section   page geometry, or null for continuous flow
 * @param {number}   [input.maxPages] a runaway guard; a measurement bug must not
 *   produce a million empty sheets and take the browser with it
 */
export function paginate({ flow, blocks, section: mainSection, sections = null, maxPages = 500, cache = null, styles = null, listLabels = null, notes = null, watermark = null, math = null }) {
  // No page geometry means no pages — an email body is a continuous flow, and
  // saying so is better than inventing A4 for it.
  if (!mainSection) return null;

  // Sections of different pages — an envelope added in front of a letter,
  // a landscape section in a portrait report — each lay their pages on
  // their own paper: `sections` is `[{ endsAt, geometry }]`, the geometry
  // shaped as `section` is. Only when two differ is any of this switched
  // on; a document of one page size comes out exactly as it always did.
  const geoms = distinctPages(sections) ? sections : null;
  const sectionAt = (index) => {
    if (!geoms || !Number.isFinite(index)) return -1;
    const k = geoms.findIndex((s) => s.endsAt === null || s.endsAt === undefined || index <= s.endsAt);
    return k < 0 ? geoms.length - 1 : k;
  };
  const entryIndex = (entry) => (entry?.kind === 'table' ? firstCellIndex(entry.table) : entry?.paragraphIndex);
  let sectionIdx = geoms ? sectionAt(entryIndex((flow || [])[0])) : -1;
  let section = geoms && sectionIdx >= 0 ? geoms[sectionIdx].geometry : mainSection;

  // More than one column: Word fills the first column top to bottom, then
  // the next, and starts a fresh PAGE only once the last column is full —
  // `newColumn` below is that rule. One column, the ordinary case and every
  // section that existed before this read `w:cols`, must come out of this
  // function BYTE-IDENTICAL to what it produced before, so every
  // column-aware path here is reached only through `columnBoxes`, which
  // stays null for one column.
  let columnBoxes = section.columns && section.columns.count > 1 ? section.columnBoxes : null;

  let width = columnBoxes ? columnBoxes[0].widthPx : section.contentWidthPx;
  let height = Math.max(
    120,
    section.heightPx - section.margins.top - section.margins.bottom,
  );
  /** The page a section lays on, as each page it starts carries it — see `geoms`. */
  const pageOf = () => ({
    widthPx: section.widthPx, heightPx: section.heightPx, orientation: section.orientation,
    margins: section.margins, contentWidthPx: section.contentWidthPx, envelope: Boolean(geoms?.[sectionIdx]?.envelope),
  });
  /** Lay what follows on another section's paper. */
  const useSection = (g) => {
    section = g;
    columnBoxes = section.columns && section.columns.count > 1 ? section.columnBoxes : null;
    width = columnBoxes ? columnBoxes[0].widthPx : section.contentWidthPx;
    height = Math.max(120, section.heightPx - section.margins.top - section.margins.bottom);
  };

  const byIndex = new Map(blocks.map((b) => [b.index, b]));
  const pages = [];
  let current = null;
  let used = 0;
  let colIdx = 0;
  // Notes finished in earlier columns of the page being laid out now —
  // `current.notes` below is always just the COLUMN being laid out, so
  // `remaining()` reserves the right room for that column alone, and a
  // finished column's notes wait here, each still knowing its column, until
  // the whole page is done and they can be published together.
  let notesDone = [];

  // A column's notes and the room they took, folded into `notesDone` — called
  // before moving off a column, to the next one or to a fresh page. A
  // one-column section never calls this: its `current.notes` IS the page's
  // notes, exactly as before `w:cols` was read at all.
  const finishColumn = () => {
    if (!columnBoxes) return;
    for (const n of current.notes) notesDone.push({ ...n, column: colIdx });
    current.notes = [];
    current.notesHeightPx = 0;
  };

  const newPage = () => {
    if (current && columnBoxes) {
      finishColumn();
      current.notes = notesDone;
      current.notesHeightPx = notesDone.reduce((s, n) => s + n.heightPx, 0);
    }
    // `notes` are the footnotes this page (or, mid-page, this column) carries
    // at its foot, and `notesHeightPx` the room they take — reserved from the
    // page's height the moment their reference lands here, so body text
    // never runs over them. The watermark rides every page. `floats` are the
    // pictures standing beside the words on this page, each with the band it
    // takes and the side it takes it on; a float ends with its column.
    current = { index: pages.length, number: pages.length + 1, fragments: [], contentHeightPx: height, notes: [], notesHeightPx: 0, watermark, floats: [] };
    if (geoms) current.section = pageOf();
    if (columnBoxes) current.columns = columnBoxes;
    pages.push(current);
    used = 0;
    colIdx = 0;
    notesDone = [];
    width = columnBoxes ? columnBoxes[0].widthPx : section.contentWidthPx;
    return current;
  };
  newPage();

  // Advance to the next column of the CURRENT page — Word's rule for "no room
  // here" in a multi-column section — or start a fresh page once the last
  // column is full. Every place below that used to mean "start a fresh page,
  // this does not fit" now means this instead; a one-column section has no
  // second column to advance to, so it falls straight through to `newPage`,
  // unchanged.
  const newColumn = () => {
    if (!columnBoxes || colIdx + 1 >= columnBoxes.length) { newPage(); return; }
    finishColumn();
    colIdx += 1;
    width = columnBoxes[colIdx].widthPx;
    used = 0;
    current.floats = [];
  };

  const remaining = () => height - used - current.notesHeightPx;
  const place = (fragment, cost) => {
    // The column a fragment landed in rides with it only in a multi-column
    // section — a one-column fragment must stay exactly the shape it always
    // was, for the byte-identical guarantee above.
    current.fragments.push(columnBoxes ? { ...fragment, column: colIdx } : fragment);
    used += cost;
  };

  /**
   * How much of the column the floats take from each side over a band of
   * the page — the lines of a paragraph that fall in that band are that much
   * shorter, and the ones beside a left float start that much further in.
   */
  const insetsFor = (top, bottom) => {
    let left = 0;
    let right = 0;
    for (const f of current.floats) {
      if (f.bottomPx <= top || f.topPx >= bottom) continue;
      if (f.side === 'left') left = Math.max(left, f.insetPx);
      else right = Math.max(right, f.insetPx);
    }
    return { left, right };
  };

  /**
   * A paragraph laid out line by line round the floats on the page, each
   * line as wide as the column leaves it at its own height. `y0` is where
   * the first line lands. Lines past the foot of the page are wrapped at
   * the full width, which is what they get on the next page.
   */
  const layoutAround = (block, y0, { extraIndentPx = 0 } = {}) => {
    const s = styleOf(block.style, styles);
    const indent = (s.indent ?? 0) + (block.indentPx ?? 0) + extraIndentPx;
    const opts = { size: s.sizePx, weight: s.weight };
    const lineHeightPx = withDirectLineHeight({ lineHeightPx: lineHeight(s.sizePx) }, block).lineHeightPx;
    const text = block.text ?? '';
    const lines = [];
    let rest = text;
    let cursor = 0;
    do {
      const top = y0 + lines.length * lineHeightPx;
      const { left, right } = insetsFor(top, top + lineHeightPx);
      const widthPx = Math.max(24, width - indent - left - right);
      const { line, rest: next } = wrapFirstLine(rest, widthPx * WRAP_SAFETY, opts);
      const at = line === '' ? cursor : text.indexOf(line, cursor);
      const start = at < 0 ? cursor : at;
      const end = start + line.length;
      lines.push({ text: line, start, end, width: measureText(line, opts), offsetPx: left, widthPx });
      cursor = end;
      rest = next;
    } while (rest.length);
    if (lines.length) lines[lines.length - 1].end = text.length;
    return { lines, style: s, indentPx: indent, lineHeightPx };
  };

  /**
   * A paragraph that holds an equation. A DISPLAY equation (m:oMathPara)
   * stands on a line of its own, centred unless the file says otherwise, as
   * tall as it draws: it is placed as an `equation` fragment, whole — half
   * a fraction is not a smaller fraction — and the words before and after
   * it are laid out as lines of their own. An INLINE equation is a word as
   * wide as it draws, and the lines of its piece are as tall as the tallest
   * equation in them needs (one height for the piece: the fragment carries
   * one line height, the honest simplification). Sizes come from `math`,
   * which answers the MathML's own measured box when the desktop measured
   * it, and the linear form's line otherwise.
   */
  const layEquations = (block, { spaceBefore: before, spaceAfter: after, extraIndentPx = 0, listing = null }) => {
    const s = styleOf(block.style, styles);
    const indent = (s.indent ?? 0) + (block.indentPx ?? 0) + extraIndentPx;
    const opts = { size: s.sizePx, weight: s.weight };
    const baseLine = withDirectLineHeight({ lineHeightPx: lineHeight(s.sizePx) }, block).lineHeightPx;
    const text = block.text ?? '';
    const runs = block.runs || [];
    const align = block.align ?? s.align ?? null;

    // The pieces: words between the display equations, and the equations.
    const pieces = [];
    const inline = new Map();
    let at = 0;
    let textFrom = 0;
    for (const r of runs) {
      const len = (r.text ?? '').length;
      if (r.math?.display) {
        if (at > textFrom) pieces.push({ kind: 'text', from: textFrom, to: at });
        pieces.push({ kind: 'eq', run: r, at });
        textFrom = at + len;
      } else if (r.math) inline.set(at, math(r, s.sizePx));
      at += len;
    }
    if (textFrom < text.length || !pieces.length) pieces.push({ kind: 'text', from: textFrom, to: text.length });
    // Words that are only white space between two display equations are
    // not a line of their own, as Word does not draw one.
    const shown = pieces.filter((p) => p.kind === 'eq' || pieces.length === 1 || text.slice(p.from, p.to).trim() !== '');

    let firstPlaced = true;
    shown.forEach((piece, k) => {
      const lastPiece = k === shown.length - 1;
      const sb = firstPlaced ? before : 0;
      const sa = lastPiece ? after : 0;
      if (piece.kind === 'eq') {
        const box = math(piece.run, s.sizePx);
        const gapPx = Math.round(s.sizePx * 0.3);
        const tall = box.heightPx + 2 * gapPx;
        if (sb + tall > remaining() && current.fragments.length) newColumn();
        const jc = box.jc === 'left' || box.jc === 'right' ? box.jc : 'center';
        place({
          kind: 'equation', paragraphIndex: block.index, start: piece.at, end: piece.at + 1,
          math: piece.run.math, box, gapPx, align: jc, indent, sizePx: s.sizePx, colour: s.colour ?? null,
          spaceBefore: current.fragments.length ? sb : 0, spaceAfter: sa,
        }, (current.fragments.length ? sb : 0) + tall + sa);
        firstPlaced = false;
        return;
      }
      const lines = wrapWithObjects(text.slice(piece.from, piece.to), piece.from, Math.max(24, (width - indent) * WRAP_SAFETY), opts, inline);
      let lineHeightPx = baseLine;
      for (const [offset, box] of inline) {
        if (offset >= piece.from && offset < piece.to) lineHeightPx = Math.max(lineHeightPx, box.heightPx + Math.round(s.sizePx * 0.3));
      }
      let cursor = 0;
      let spaceHere = sb;
      while (cursor < lines.length) {
        let fits = Math.floor((remaining() - spaceHere) / lineHeightPx);
        if (fits < 1) {
          if (current.fragments.length === 0) fits = 1;
          else { newColumn(); spaceHere = 0; continue; }
        }
        const slice = lines.slice(cursor, cursor + fits);
        const complete = cursor + fits >= lines.length;
        place({
          kind: 'paragraph', paragraphIndex: block.index, style: block.style ?? null,
          structural: Boolean(block.structural), structuralTags: block.structuralTags ?? [],
          lines: slice, start: slice[0].start, end: slice[slice.length - 1].end,
          first: firstPlaced && cursor === 0, last: complete && lastPiece,
          spaceBefore: spaceHere, spaceAfter: complete ? sa : 0,
          lineHeightPx, sizePx: s.sizePx, weight: s.weight ?? 'normal', italic: Boolean(s.italic), colour: s.colour ?? null,
          align, listLabel: firstPlaced && cursor === 0 && listing ? listing.label : null, indent,
        }, spaceHere + slice.length * lineHeightPx + (complete ? sa : 0));
        cursor += fits;
        spaceHere = 0;
        if (!complete) newColumn();
      }
      firstPlaced = false;
    });
  };

  /**
   * A DISPLAYED paragraph (a text box's, a note's) laid out at a width: the
   * lines and the look, the way a body fragment carries them, plus the runs
   * so the writer can paint the words in their own formatting.
   */
  const layShown = (p, at) => {
    const laid = layoutParagraph(p, at, { styles });
    return {
      lines: laid.lines, lineHeightPx: laid.lineHeightPx, indent: laid.indentPx ?? 0,
      sizePx: laid.style.sizePx, weight: laid.style.weight ?? 'normal', italic: Boolean(laid.style.italic), colour: laid.style.colour ?? null,
      align: p.align ?? laid.style.align ?? null, runs: p.runs || [],
      spaceBefore: p.spaceBeforePx ?? laid.style.spaceBefore ?? 0, spaceAfter: p.spaceAfterPx ?? laid.style.spaceAfter ?? 0,
    };
  };
  const heightOf = (laidParagraphs) => laidParagraphs.reduce((s, p) => s + p.spaceBefore + p.lines.length * p.lineHeightPx + p.spaceAfter, 0);

  // Footnotes by id, laid out once: the same note is never referenced twice.
  const noteById = new Map((notes?.footnotes || []).map((n) => [n.id, n]));
  const layNote = (note) => {
    const paragraphs = (note.paragraphs || []).map((p) => layShown(p, width - NOTE_INDENT_PX));
    return { n: note.n, id: note.id, paragraphs, heightPx: heightOf(paragraphs) + 2 };
  };

  const flowList = flow ?? [];
  for (let flowIdx = 0; flowIdx < flowList.length; flowIdx++) {
    const entry = flowList[flowIdx];
    if (pages.length > maxPages) break;

    // A section of other paper starts a page of its own paper.
    if (geoms) {
      const k = sectionAt(entryIndex(entry));
      if (k >= 0 && k !== sectionIdx) {
        sectionIdx = k;
        useSection(geoms[k].geometry);
        if (current.fragments.length) newPage();
        else { current.section = pageOf(); current.contentHeightPx = height; used = 0; }
      }
    }

    if (entry.kind === 'table') {
      // A table wider than the column it must now sit in — the file's own
      // width, or the file's own column widths, were set for the page before
      // it grew columns — is scaled down to fit, grid and all; the fragment
      // carries the scaled table, so the paint downstream needs no scale of
      // its own. Single-column tables are never wider than `width` is here
      // (it IS the file's content width), so this never touches them.
      let table = entry.table;
      if (columnBoxes && table.columns.length) {
        const natural = table.columns.reduce((a, b) => a + b, 0);
        if (natural > width && natural > 0) {
          const scale = width / natural;
          table = { ...table, columns: table.columns.map((c) => c * scale) };
        }
      }
      layTable(table, { width, height, place, remaining, newPage: newColumn, cache });
      continue;
    }

    const block = byIndex.get(entry.paragraphIndex);
    if (!block) continue;

    // A paragraph in a frame placed on the page — an envelope's delivery
    // address — is drawn where the frame says, and takes no room in the flow.
    if (block.frame) {
      const f = block.frame;
      const fw = Math.max(24, f.widthPx || width);
      const laid = layoutParagraph(block, fw, { cache, styles });
      const s = laid.style;
      place({
        kind: 'frame', paragraphIndex: block.index, xPx: f.xPx, yPx: f.yPx, widthPx: fw, heightPx: f.heightPx ?? null,
        lines: laid.lines, lineHeightPx: laid.lineHeightPx, sizePx: s.sizePx, weight: s.weight ?? 'normal', italic: Boolean(s.italic),
        colour: s.colour ?? null, align: block.align ?? s.align ?? null, indent: laid.indentPx ?? 0, first: true, last: true,
      }, 0);
      continue;
    }

    // An explicit page break is the author's instruction, not a suggestion.
    if (block.pageBreakBefore && current.fragments.length) newPage();

    // A list paragraph is indented past its label's gutter, so the label the
    // view computed changes where the text wraps — the indent has to reach the
    // layout, not just the paint.
    const listing = listLabels?.get(block.index) ?? null;
    // The numbering level's indent and the style's compete rather than stack:
    // Word applies whichever wins, and ListParagraph + a level indent summed
    // would march a flat list halfway across the page.
    const styleIndent = styleOf(block.style, styles).indent ?? 0;
    const extraIndentPx = listing ? Math.max(0, listing.indentPx - styleIndent) : 0;
    const blockStyle = styleOf(block.style, styles);
    // Direct paragraph spacing beats the style's, exactly as Word resolves it.
    let spaceBefore = block.spacing?.beforePx ?? blockStyle.spaceBefore;
    const spaceAfter = block.spacing?.afterPx ?? blockStyle.spaceAfter;

    // A drop cap — Word's own trick, the first letter split into its own
    // paragraph and framed to stand `lines` deep — is never laid out as a
    // paragraph of its own: it floats beside the words that follow it,
    // exactly as a floating picture does, and the paragraph after it lays
    // out round it via the same `floats`/`layoutAround` machinery. A drop
    // cap with nothing after it (the last paragraph in the flow) has no
    // words to stand beside, so it falls through and prints as its one
    // short ordinary line.
    const nextEntry = flowList[flowIdx + 1];
    const nextBlock = nextEntry && nextEntry.kind !== 'table' ? byIndex.get(nextEntry.paragraphIndex) : null;
    if (block.dropCap && nextBlock) {
      const run0 = (block.runs || [])[0] ?? {};
      const bold = Boolean(run0.bold) || blockStyle.weight === 'bold';
      const weight = bold ? 'bold' : 'normal';
      const sizePx = run0.fontSize ? Number(run0.fontSize) * (96 / 72) : blockStyle.sizePx;
      const text = block.text ?? '';
      const widthPx = measureText(text, { size: sizePx, weight });
      const { lineHeightPx: nextLineHeightPx } = layoutParagraph(nextBlock, width, { cache, styles });
      const heightPx = block.dropCap.lines * nextLineHeightPx;
      if (heightPx + spaceBefore > remaining() && current.fragments.length) { newColumn(); spaceBefore = 0; }
      const topPx = used + spaceBefore;
      // "In margin" hangs the letter in the left margin rather than taking
      // room from the column — Word's own second style — but only when the
      // margin actually has the room; a narrow margin falls back to standing
      // in the column like an ordinary drop, the honest simplification.
      const inMargin = block.dropCap.kind === 'margin' && section.margins.left >= widthPx;
      const insetPx = inMargin ? 6 : widthPx + 6;
      current.floats.push({ side: 'left', topPx, bottomPx: topPx + heightPx, insetPx });
      place({
        kind: 'dropcap', paragraphIndex: block.index, topPx, widthPx, heightPx, sizePx, weight, text,
        runs: block.runs, style: blockStyle, indentPx: (blockStyle.indent ?? 0) + (block.indentPx ?? 0),
        ...(inMargin ? { inMargin: true } : {}),
      }, 0);
      continue;
    }

    // A picture anchored to this paragraph that floats at the left or the
    // right stands beside the words, as it does on screen: it is placed at
    // the paragraph's top on the side it asks for, and the lines beside it
    // — this paragraph's and the next ones', for as far down as it reaches
    // — are laid out shorter. It goes whole onto the next page rather than
    // being cut, and its paragraph goes with it.
    const beside = (block.images ?? []).filter((img) => img.href && floatsBeside(img));
    if (beside.length) {
      const tallest = beside.reduce((h, img) => Math.max(h, floatBox(img, width).heightPx + (img.dist?.t || 0) + (img.dist?.b ?? FLOAT_GAP_PX)), 0);
      if (tallest + spaceBefore > remaining() && current.fragments.length) { newColumn(); spaceBefore = 0; }
      for (const img of beside) {
        const box = floatBox(img, width);
        const side = img.hAlign === 'right' || img.hAlign === 'outside' ? 'right' : 'left';
        const topPx = used + spaceBefore + (img.dist?.t || 0);
        const gap = side === 'left' ? (img.dist?.r ?? FLOAT_GAP_PX * 2) : (img.dist?.l ?? FLOAT_GAP_PX * 2);
        current.floats.push({ side, topPx, bottomPx: topPx + box.heightPx + (img.dist?.b ?? FLOAT_GAP_PX), insetPx: box.widthPx + gap });
        place({ kind: 'float', paragraphIndex: block.index, side, topPx, widthPx: box.widthPx, heightPx: box.heightPx, image: { ...img, ...box } }, 0);
      }
    }

    // A text box anchored to this paragraph that floats at the left or the
    // right — a pull quote, a sidebar — stands beside the words the same
    // way. Its size is the file's or its words', whichever is more; a box
    // that would leave the lines no room is drawn under the words instead.
    const boxesBeside = (block.textBoxes ?? []).filter((box) => floatsBeside(box) && Math.max(40, Math.min(width, box.widthPx || width)) <= width - 80);
    if (boxesBeside.length) {
      const laidBoxes = boxesBeside.map((box) => {
        const widthPx = Math.max(40, Math.min(width, box.widthPx || width));
        const paragraphs = (box.paragraphs || []).map((p) => layShown(p, widthPx - 2 * BOX_PAD_PX));
        const heightPx = Math.min(height, Math.max(box.heightPx || 0, heightOf(paragraphs) + 2 * BOX_PAD_PX));
        return { box, widthPx, heightPx, paragraphs };
      });
      const tallest = laidBoxes.reduce((h, b) => Math.max(h, b.heightPx + (b.box.dist?.t || 0) + (b.box.dist?.b ?? FLOAT_GAP_PX)), 0);
      if (tallest + spaceBefore > remaining() && current.fragments.length) { newColumn(); spaceBefore = 0; }
      for (const { box, widthPx, heightPx, paragraphs } of laidBoxes) {
        const side = box.hAlign === 'right' || box.hAlign === 'outside' ? 'right' : 'left';
        const topPx = used + spaceBefore + (box.dist?.t || 0);
        const gap = side === 'left' ? (box.dist?.r ?? FLOAT_GAP_PX * 2) : (box.dist?.l ?? FLOAT_GAP_PX * 2);
        current.floats.push({ side, topPx, bottomPx: topPx + heightPx + (box.dist?.b ?? FLOAT_GAP_PX), insetPx: widthPx + gap });
        place({ kind: 'floatbox', paragraphIndex: block.index, side, topPx, widthPx, heightPx, fill: box.fill || null, line: box.line || null, paragraphs }, 0);
      }
    }

    // A paragraph holding an equation is laid out by a path of its own —
    // `layEquations` below: a display equation is a block of its own
    // height, an inline one a word as wide and as tall as it draws — and
    // then its pictures follow it as any paragraph's do.
    const mathy = typeof math === 'function' && (block.runs || []).some((r) => r.math);
    let style = styleOf(block.style, styles);
    if (mathy) layEquations(block, { spaceBefore, spaceAfter, extraIndentPx, listing });
    else {
      const around = current.floats.some((f) => f.bottomPx > used + spaceBefore);
      const laid = around
        ? layoutAround(block, used + spaceBefore, { extraIndentPx })
        : layoutParagraph(block, width, { cache, styles, extraIndentPx });
      const { lines, indentPx, lineHeightPx } = laid;
      style = laid.style;
      let cursor = 0;

      // The footnotes this paragraph references go at the foot of the page its
      // first line lands on — Word's rule — so their room is reserved before
      // the line is placed, and a paragraph whose notes will not fit beside it
      // starts on the next page, notes and all.
      const pageNotes = (block.runs || [])
        .filter((r) => r.noteRef?.kind === 'footnote' && noteById.has(r.noteRef.id))
        .map((r) => layNote(noteById.get(r.noteRef.id)));
      if (pageNotes.length) {
        const cost = pageNotes.reduce((s, n) => s + n.heightPx, 0) + (current.notes.length ? 0 : NOTE_RULE_PX);
        if (remaining() - spaceBefore - cost < lineHeightPx && current.fragments.length) { newColumn(); spaceBefore = 0; }
        current.notesHeightPx += pageNotes.reduce((s, n) => s + n.heightPx, 0) + (current.notes.length ? 0 : NOTE_RULE_PX);
        current.notes.push(...pageNotes);
      }

      while (cursor < lines.length) {
        // Only the LINES have to fit. The space after a paragraph is empty room
        // below its last line, and at a page break there is nothing below it to
        // push — Word suppresses it there for the same reason. Reserving for it
        // would break a page early and leave a visible gap; charging for it
        // without reserving would run every page over by one margin. The shell
        // drops the trailing margin on each page to match.
        let fits = Math.floor((remaining() - spaceBefore) / lineHeightPx);

        if (fits < 1) {
          // Nothing fits. Start a page — unless we are already on an empty one, in
          // which case a single line is taller than the page and forcing another
          // break would loop for ever.
          if (current.fragments.length === 0) fits = 1;
          else { newColumn(); spaceBefore = 0; continue; }
        }

        // Widow and orphan control, the cheap version: never leave one line of a
        // paragraph alone on a page. One stranded line reads as a mistake, and
        // moving it costs nothing but a little whitespace.
        const left = lines.length - cursor;
        if (fits < left && left - fits === 1 && fits > 1) fits -= 1;
        if (fits === 1 && left > 2 && current.fragments.length) { newColumn(); spaceBefore = 0; continue; }

        const slice = lines.slice(cursor, cursor + fits);
        const complete = cursor + fits >= lines.length;
        place({
          kind: 'paragraph',
          paragraphIndex: block.index,
          style: block.style ?? null,
          structural: Boolean(block.structural),
          structuralTags: block.structuralTags ?? [],
          lines: slice,
          start: slice[0].start,
          end: slice[slice.length - 1].end,
          first: cursor === 0,
          last: complete,
          spaceBefore,
          // What the shell must draw BELOW the fragment — zero mid-paragraph
          // and, matching the cost charged, the full gap on the last slice.
          spaceAfter: complete ? spaceAfter : 0,
          lineHeightPx,
          sizePx: style.sizePx,
          // The look the style resolved to — the shell draws these rather than
          // consulting a stylesheet, so measurement and paint cannot drift.
          weight: style.weight ?? 'normal',
          italic: Boolean(style.italic),
          colour: style.colour ?? null,
          // The paragraph's own alignment beats the style's, as everywhere else.
          align: block.align ?? style.align ?? null,
          // The label only appears on the FIRST fragment; a continuation on the
          // next sheet is mid-item, and "3." repeated there would read as item 3
          // appearing twice.
          listLabel: cursor === 0 && listing ? listing.label : null,
          indent: indentPx ?? 0,
        }, spaceBefore + slice.length * lineHeightPx + (complete ? spaceAfter : 0));

        cursor += fits;
        spaceBefore = 0;
        if (!complete) newColumn();
      }
    }

    // Every other picture renders as a block under the paragraph's text: an
    // inline one, one wrapped top-and-bottom, one behind or in front of the
    // words (drawn in the flow rather than over it — the honest
    // simplification). Placed as a fragment of its own so a caret never
    // lands in one, and pushed whole onto the next sheet rather than sliced:
    // half a logo is not a smaller logo. A centred or right-aligned one
    // keeps its side.
    const images = (block.images ?? []).filter((img) => img.href && !floatsBeside(img));
    if (images.length) {
      const drawn = images.map((img) => {
        // To the page's width, and to its height: a picture taller than
        // the page's inside is drawn to fit it, its proportions kept — the
        // screen does the same, and a sheet cannot hold more.
        const scale = Math.min(1, width / Math.max(1, img.widthPx), (height - IMAGE_GAP) / Math.max(1, img.heightPx));
        const hAlign = img.anchored && img.hAlign === 'center' ? 'center' : img.anchored && (img.hAlign === 'right' || img.hAlign === 'outside') ? 'right' : 'left';
        return { ...img, widthPx: img.widthPx * scale, heightPx: img.heightPx * scale, hAlign };
      });
      // A paragraph of only pictures lays them in rows, as many to a row as
      // fit the column — as Word draws inline pictures and the screen does:
      // a scanner's four cards go two by two. Under words a picture is a row
      // of its own. Rows go down a page at a time: the ones that fit here
      // and the rest at the head of the next page, each row whole.
      const rowsOnly = !(block.text ?? '').length;
      const rows = [];
      for (const img of drawn) {
        const last = rows[rows.length - 1];
        if (rowsOnly && last && last.widthPx + img.widthPx <= width + 0.5) {
          last.images.push(img);
          last.widthPx += img.widthPx;
          last.heightPx = Math.max(last.heightPx, img.heightPx);
        } else rows.push({ images: [img], widthPx: img.widthPx, heightPx: img.heightPx });
      }
      const align = block.align ?? style.align ?? null;
      let batch = [];
      let cost = 0;
      const flush = () => {
        if (!batch.length) return;
        place({
          kind: 'images',
          paragraphIndex: block.index,
          images: batch.flatMap((row) => row.images),
          rows: batch.map((row) => ({ count: row.images.length, heightPx: row.heightPx })),
          align,
        }, cost);
        batch = [];
        cost = 0;
      };
      for (const row of rows) {
        const tall = row.heightPx + IMAGE_GAP;
        if (cost + tall > remaining()) {
          flush();
          if (tall > remaining() && current.fragments.length) newColumn();
        }
        batch.push(row);
        cost += tall;
      }
      flush();
    }

    // Text boxes ride under their paragraph like pictures do — the same
    // honest simplification of float layout the screen makes. The box is as
    // tall as the file says or as its words need, whichever is more, and is
    // pushed whole onto the next sheet rather than cut.
    for (const box of block.textBoxes || []) {
      if (boxesBeside.includes(box)) continue;
      const boxWidth = Math.max(40, Math.min(width, box.widthPx || width));
      const paragraphs = (box.paragraphs || []).map((p) => layShown(p, boxWidth - 2 * BOX_PAD_PX));
      const heightPx = Math.min(height, Math.max(box.heightPx || 0, heightOf(paragraphs) + 2 * BOX_PAD_PX));
      if (heightPx > remaining() && current.fragments.length) newColumn();
      place({
        kind: 'textbox', paragraphIndex: block.index, widthPx: boxWidth, heightPx,
        fill: box.fill || null, line: box.line || null, hAlign: box.hAlign || null, paragraphs,
      }, heightPx + IMAGE_GAP);
    }
  }

  // Endnotes come after everything, as their name says: each laid out as a
  // paragraph of its own, numbered where its mark is.
  for (const note of notes?.endnotes || []) {
    for (const p of note.paragraphs || []) {
      const laid = layShown(p, width - NOTE_INDENT_PX);
      const cost = laid.spaceBefore + laid.lines.length * laid.lineHeightPx + laid.spaceAfter;
      if (cost > remaining() && current.fragments.length) newPage();
      place({ kind: 'note', note: note.n, ...laid, first: true, last: true, structural: true, structuralTags: ['w:endnote'] }, cost);
    }
  }

  // The last column of the last page never gets a `newPage()` call after it
  // to publish its notes the way every earlier one does — this closes it out
  // the same way, so its footnotes are not left stranded in `notesDone`.
  if (columnBoxes) {
    finishColumn();
    current.notes = notesDone;
    current.notesHeightPx = notesDone.reduce((s, n) => s + n.heightPx, 0);
  }

  const count = pages.length;
  for (const page of pages) page.of = count;
  // `width` is the CURRENT COLUMN's width by the time the flow ends, not the
  // section's — the section's own is what a reader means by "how wide is the
  // page", so that is what is reported, not whatever column happened to be
  // laid out last.
  return { pages, count, contentWidthPx: mainSection.contentWidthPx, contentHeightPx: Math.max(120, mainSection.heightPx - mainSection.margins.top - mainSection.margins.bottom) };
}

/** Do the sections lie on more than one kind of paper? */
function distinctPages(sections) {
  if (!Array.isArray(sections) || sections.length < 2 || !sections.every((s) => s?.geometry)) return false;
  const key = (g) => [g.widthPx, g.heightPx, g.margins?.top, g.margins?.right, g.margins?.bottom, g.margins?.left].map((n) => Math.round(Number(n) || 0)).join(':');
  return new Set(sections.map((s) => key(s.geometry))).size > 1;
}

/** The edit address of a table's first cell paragraph, which says what section the table is in. */
function firstCellIndex(table) {
  for (const row of table?.rows || []) for (const cell of row.cells || []) for (const b of cell.blocks || []) if (Number.isFinite(b.blockIndex)) return b.blockIndex;
  return undefined;
}

/** The room a note takes from the page: its indent, and the rule above the first. */
const NOTE_INDENT_PX = 18;
const NOTE_RULE_PX = 14;
/** A text box's padding, the way Word insets a shape's words. */
const BOX_PAD_PX = 7;

/**
 * Place a table, breaking BETWEEN ROWS.
 *
 * Never inside one: half a row on each of two sheets is unreadable, and Word
 * only splits a row when it is taller than a whole page. A header row repeats on
 * every continuation, which is what `w:tblHeader` is for and the only way a
 * table that spans pages stays legible.
 */
/** More sheets than any one table can honestly need; past this it is a bug, and the floor. */
const maxPagesGuard = 2000;

function layTable(table, { width, height, place, remaining, newPage, cache = null }) {
  const headerRows = [];
  for (const row of table.rows) {
    if (!row.header) break;
    headerRows.push(row);
  }
  const headerHeight = headerRows.reduce((t, r) => t + rowHeight(r, table, width, cache), 0);

  let index = 0;
  let repeatHeader = false;
  let sheets = 0;
  while (index < table.rows.length) {
    const rows = [];
    let used = repeatHeader ? headerHeight : 0;

    while (index < table.rows.length) {
      const row = table.rows[index];
      const h = rowHeight(row, table, width, cache);
      // Break for a new page only once a BODY row is on this one. The repeated
      // header used to count as placed rows, so a header plus first row taller
      // than the page broke with nothing placed, opened a fresh page, and did
      // exactly the same again — an infinite loop that allocated a page object
      // each turn and took an 81 KB document to four gigabytes. A row that
      // does not fit on an otherwise empty page is placed anyway and clipped;
      // a clipped row at least shows the reader something is there.
      if (used + h > remaining() && rows.length > 0) break;
      rows.push(row);
      used += h;
      index += 1;
      if (used > remaining() && used > height) break;
    }

    if (!rows.length) {
      // Cannot happen now — the loop above always takes at least one row — but
      // a paginator's loops are exactly where "cannot happen" needs a floor.
      if (++sheets > maxPagesGuard) break;
      newPage();
      repeatHeader = headerRows.length > 0;
      continue;
    }

    const done = index >= table.rows.length;
    place({
      kind: 'table',
      table,
      rows: repeatHeader ? [...headerRows, ...rows] : rows,
      repeatedHeader: repeatHeader,
      continues: !done,
    }, used + (done ? TABLE_SPACE_AFTER : 0));

    if (!done) { newPage(); repeatHeader = headerRows.length > 0; }
  }
}

export { rowHeight, tableHeight };
