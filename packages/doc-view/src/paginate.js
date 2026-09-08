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
import { measureText, wrapText, lineHeight } from '@rutba/drawing';

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

/** How tall a table row is: its tallest cell, wrapped to the column width. */
function rowHeight(row, table, width, cache) {
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
    let height = CELL_PADDING;
    for (const b of cell.blocks) {
      if (b.kind === 'table') { height += tableHeight(b.table, cellWidth, cache); continue; }
      const { lines, lineHeightPx } = layoutParagraph(b, cellWidth - CELL_PADDING * 2, { cache });
      height += lines.length * lineHeightPx;
    }
    tallest = Math.max(tallest, height);
  }
  return Math.max(row.heightPx ?? 0, tallest);
}

const tableHeight = (table, width, cache = null) =>
  table.rows.reduce((total, row) => total + rowHeight(row, table, width, cache), 0);

const MIN_ROW_HEIGHT = 22;
const IMAGE_GAP = 8;
const CELL_PADDING = 8;
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
export function paginate({ flow, blocks, section, maxPages = 500, cache = null, styles = null, listLabels = null, notes = null, watermark = null }) {
  // No page geometry means no pages — an email body is a continuous flow, and
  // saying so is better than inventing A4 for it.
  if (!section) return null;

  const width = section.contentWidthPx;
  const height = Math.max(
    120,
    section.heightPx - section.margins.top - section.margins.bottom,
  );

  const byIndex = new Map(blocks.map((b) => [b.index, b]));
  const pages = [];
  let current = null;
  let used = 0;

  const newPage = () => {
    // `notes` are the footnotes this page carries at its foot, and
    // `notesHeightPx` the room they take — reserved from the page's height
    // the moment their reference lands here, so body text never runs over
    // them. The watermark rides every page.
    current = { index: pages.length, number: pages.length + 1, fragments: [], contentHeightPx: height, notes: [], notesHeightPx: 0, watermark };
    pages.push(current);
    used = 0;
    return current;
  };
  newPage();

  const remaining = () => height - used - current.notesHeightPx;
  const place = (fragment, cost) => {
    current.fragments.push(fragment);
    used += cost;
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

  for (const entry of flow ?? []) {
    if (pages.length > maxPages) break;

    if (entry.kind === 'table') {
      layTable(entry.table, { width, height, place, remaining, newPage, cache });
      continue;
    }

    const block = byIndex.get(entry.paragraphIndex);
    if (!block) continue;

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
    const { lines, style, indentPx, lineHeightPx } = layoutParagraph(block, width, {
      cache, styles, extraIndentPx: listing ? Math.max(0, listing.indentPx - styleIndent) : 0,
    });
    let cursor = 0;
    // Direct paragraph spacing beats the style's, exactly as Word resolves it.
    let spaceBefore = block.spacing?.beforePx ?? style.spaceBefore;
    const spaceAfter = block.spacing?.afterPx ?? style.spaceAfter;

    // The footnotes this paragraph references go at the foot of the page its
    // first line lands on — Word's rule — so their room is reserved before
    // the line is placed, and a paragraph whose notes will not fit beside it
    // starts on the next page, notes and all.
    const pageNotes = (block.runs || [])
      .filter((r) => r.noteRef?.kind === 'footnote' && noteById.has(r.noteRef.id))
      .map((r) => layNote(noteById.get(r.noteRef.id)));
    if (pageNotes.length) {
      const cost = pageNotes.reduce((s, n) => s + n.heightPx, 0) + (current.notes.length ? 0 : NOTE_RULE_PX);
      if (remaining() - spaceBefore - cost < lineHeightPx && current.fragments.length) { newPage(); spaceBefore = 0; }
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
        else { newPage(); spaceBefore = 0; continue; }
      }

      // Widow and orphan control, the cheap version: never leave one line of a
      // paragraph alone on a page. One stranded line reads as a mistake, and
      // moving it costs nothing but a little whitespace.
      const left = lines.length - cursor;
      if (fits < left && left - fits === 1 && fits > 1) fits -= 1;
      if (fits === 1 && left > 2 && current.fragments.length) { newPage(); spaceBefore = 0; continue; }

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
      if (!complete) newPage();
    }

    // Pictures render as a block under the paragraph's text — the honest
    // simplification recorded where they are parsed. They are placed as their
    // own fragment so a caret never lands in one, and pushed whole onto the
    // next sheet rather than sliced: half a logo is not a smaller logo.
    const images = (block.images ?? []).filter((img) => img.href);
    if (images.length) {
      const drawn = images.map((img) => {
        const scale = Math.min(1, width / Math.max(1, img.widthPx));
        return { ...img, widthPx: img.widthPx * scale, heightPx: img.heightPx * scale };
      });
      const cost = drawn.reduce((total, img) => total + img.heightPx + IMAGE_GAP, 0);
      if (cost > remaining() && current.fragments.length) newPage();
      place({ kind: 'images', paragraphIndex: block.index, images: drawn }, cost);
    }

    // Text boxes ride under their paragraph like pictures do — the same
    // honest simplification of float layout the screen makes. The box is as
    // tall as the file says or as its words need, whichever is more, and is
    // pushed whole onto the next sheet rather than cut.
    for (const box of block.textBoxes || []) {
      const boxWidth = Math.max(40, Math.min(width, box.widthPx || width));
      const paragraphs = (box.paragraphs || []).map((p) => layShown(p, boxWidth - 2 * BOX_PAD_PX));
      const heightPx = Math.min(height, Math.max(box.heightPx || 0, heightOf(paragraphs) + 2 * BOX_PAD_PX));
      if (heightPx > remaining() && current.fragments.length) newPage();
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

  const count = pages.length;
  for (const page of pages) page.of = count;
  return { pages, count, contentWidthPx: width, contentHeightPx: height };
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
