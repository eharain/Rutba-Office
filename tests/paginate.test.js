/**
 * Pagination — where the breaks fall, and what must not move when they do.
 *
 * The layout itself is the easy half. What these pin down is the set of things
 * that were wrong at least once while building it, each of which is invisible
 * until a document is long enough to notice:
 *
 *   - a page that overflows by one line, because the fit test forgot a margin
 *   - a caret that types on the wrong sheet, because a split paragraph's second
 *     half reported offsets from its own start rather than the paragraph's
 *   - a PAGE field that says the same number on every sheet
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { paginate, layoutParagraph, PARAGRAPH_STYLES, WRAP_SAFETY } from '@rutba/doc-view/paginate';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { parseBand } from '@rutba/ooxml';
import { resolveFields, bandForPage } from '@rutba/doc-view/bands';
import { buildComplexDocument } from './fixtures/complex-document.js';

const LETTER = buildComplexDocument();

/** A4 with one-inch margins, which is what the fixtures use. */
const A4 = {
  widthPx: 794, heightPx: 1123,
  margins: { top: 96, right: 96, bottom: 96, left: 96, header: 47, footer: 47, gutter: 0 },
  contentWidthPx: 602,
  titlePage: false, evenAndOdd: false, declared: true,
};

const para = (index, text, style = null) => ({ index, text, style, runs: [{ text }], structural: false, structuralTags: [] });
const flowOf = (blocks) => blocks.map((b) => ({ kind: 'paragraph', paragraphIndex: b.index }));

// ------------------------------------------------------------------ lines --

test('a line knows the character range it covers', () => {
  const text = 'one two three four five six seven eight nine ten eleven twelve';
  const { lines } = layoutParagraph(para(0, text), 120);
  assert.ok(lines.length > 1, 'that does not fit on one line at 120px');

  // Every line's slice must be exactly what the text says at that range —
  // this is what the caret relies on when a paragraph spans two sheets.
  for (const line of lines) {
    assert.equal(text.slice(line.start, line.end).trim(), line.text.trim());
  }
  assert.equal(lines[0].start, 0);
  assert.equal(lines[lines.length - 1].end, text.length,
    'the last line must reach the end, or a caret cannot land after the final character');
});

test('an empty paragraph is still one line', () => {
  // It is the blank line the author typed; swallowing it closes up their spacing.
  const { lines } = layoutParagraph(para(0, ''), 400);
  assert.equal(lines.length, 1);
  assert.deepEqual(lines[0], { text: '', start: 0, end: 0, width: 0 });
});

test('the measured column is narrower than the drawn one', () => {
  // Our widths come from a table and run about half a percent under. A line two
  // pixels too long re-wraps in the browser, which adds a whole line box and
  // pushes the page over — so we deliberately wrap short.
  assert.ok(WRAP_SAFETY < 1 && WRAP_SAFETY > 0.9);
  const wide = layoutParagraph(para(0, 'x'.repeat(400)), 600).lines;
  const narrow = layoutParagraph(para(0, 'x'.repeat(400)), 600 * WRAP_SAFETY).lines;
  assert.ok(wide.length <= narrow.length);
});

// ------------------------------------------------------------------ pages --

test('content longer than a sheet spills onto the next one', () => {
  const blocks = Array.from({ length: 60 }, (_, i) => para(i, 'Paragraph ' + (i + 1) + '. ' + 'word '.repeat(20)));
  const laid = paginate({ flow: flowOf(blocks), blocks, section: A4 });

  assert.ok(laid.count > 1, 'sixty paragraphs do not fit on one page');
  assert.equal(laid.pages[0].number, 1);
  assert.ok(laid.pages.every((p) => p.of === laid.count), 'every page knows the total');
});

test('no page is laid out taller than the page it is on', () => {
  // The bug this catches: the fit test counted lines but not the space after
  // them, so every page ran over by one margin and the error compounded.
  const blocks = Array.from({ length: 80 }, (_, i) => para(i, 'Line ' + i + '. ' + 'text '.repeat(24)));
  const laid = paginate({ flow: flowOf(blocks), blocks, section: A4 });

  for (const page of laid.pages) {
    const used = page.fragments.reduce((total, f) => {
      const style = PARAGRAPH_STYLES[f.style] ?? PARAGRAPH_STYLES.default;
      // The trailing space of the LAST paragraph is suppressed at a break, the
      // same way Word suppresses it and the shell drops it.
      const isLast = f === page.fragments[page.fragments.length - 1];
      return total + f.spaceBefore + f.lines.length * f.lineHeightPx
        + (f.last && !isLast ? style.spaceAfter : 0);
    }, 0);
    assert.ok(used <= page.contentHeightPx,
      'page ' + page.number + ' laid out ' + used + 'px into ' + page.contentHeightPx + 'px');
  }
});

test('a split paragraph keeps one index and reports where each half starts', () => {
  const long = 'sentence '.repeat(900);
  const blocks = [para(0, long)];
  const laid = paginate({ flow: flowOf(blocks), blocks, section: A4 });
  assert.ok(laid.count > 1, 'that paragraph is taller than a page');

  const fragments = laid.pages.flatMap((p) => p.fragments);
  assert.ok(fragments.every((f) => f.paragraphIndex === 0), 'still one paragraph, one index');
  assert.equal(fragments[0].first, true);
  assert.equal(fragments[0].last, false);
  assert.equal(fragments[fragments.length - 1].last, true);

  // No character is on two sheets, and no VISIBLE character is on none. The one
  // gap allowed is the space the line broke on: it is consumed by the wrap and
  // rendered nowhere, exactly as a soft wrap behaves in any editor.
  assert.equal(fragments[0].start, 0);
  for (let i = 1; i < fragments.length; i++) {
    const gap = long.slice(fragments[i - 1].end, fragments[i].start);
    assert.ok(gap.length <= 1, 'fragment ' + i + ' skipped ' + gap.length + ' characters');
    assert.equal(gap.trim(), '', 'only whitespace may fall in the gap, not ' + JSON.stringify(gap));
  }
  assert.equal(fragments[fragments.length - 1].end, long.length);
});

test('an explicit page break is obeyed', () => {
  const blocks = [para(0, 'One'), para(1, 'Two'), { ...para(2, 'Three'), pageBreakBefore: true }, para(3, 'Four')];
  const laid = paginate({ flow: flowOf(blocks), blocks, section: A4 });

  assert.equal(laid.count, 2);
  assert.deepEqual(laid.pages[0].fragments.map((f) => f.paragraphIndex), [0, 1]);
  assert.deepEqual(laid.pages[1].fragments.map((f) => f.paragraphIndex), [2, 3]);
});

test('a break on the very first paragraph does not make a blank first page', () => {
  const blocks = [{ ...para(0, 'One'), pageBreakBefore: true }, para(1, 'Two')];
  const laid = paginate({ flow: flowOf(blocks), blocks, section: A4 });
  assert.equal(laid.count, 1);
});

test('a single line taller than the page does not loop forever', () => {
  // The guard that matters: if nothing fits and we are already on an empty page,
  // place it anyway. Without that the paginator spins making empty sheets.
  const tiny = { ...A4, heightPx: 200, margins: { ...A4.margins, top: 96, bottom: 96 } };
  const blocks = [para(0, 'x '.repeat(2000))];
  const laid = paginate({ flow: flowOf(blocks), blocks, section: tiny });
  assert.ok(laid.count > 1 && laid.count <= 500);
  assert.ok(laid.pages.every((p) => p.fragments.length > 0), 'no empty sheets');
});

test('a runaway layout is bounded rather than left to hang the browser', () => {
  const blocks = Array.from({ length: 4000 }, (_, i) => para(i, 'x'.repeat(300)));
  const laid = paginate({ flow: flowOf(blocks), blocks, section: { ...A4, heightPx: 200 } });
  assert.ok(laid.count <= 501);
});

test('no page geometry means no pages', () => {
  // An email body is a continuous flow, and saying so beats inventing A4.
  assert.equal(paginate({ flow: [], blocks: [], section: null }), null);
});

// ---------------------------------------------------------------- the real --

test('the fixture letter lays out with its own header and footer', () => {
  const frame = openDocx(LETTER).render();
  assert.ok(frame.pages, 'a .docx has pages');
  assert.equal(frame.pages.count, 1);

  const page = frame.pages.pages[0];
  assert.deepEqual(page.header.map((p) => p.text), ['Rutba Trading Company']);
  assert.equal(page.footer[0].text, 'Page 1');
  assert.ok(page.fragments.some((f) => f.kind === 'table'), 'the table is on the sheet');
});

test('the PAGE field says something different on each sheet', () => {
  const band = parseBand('<w:ftr><w:p><w:r><w:t>Page </w:t></w:r>'
    + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>'
    + '<w:r><w:t> of </w:t></w:r>'
    + '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>');

  assert.equal(resolveFields(band, { page: 1, of: 7 })[0].text, 'Page 1 of 7');
  assert.equal(resolveFields(band, { page: 4, of: 7 })[0].text, 'Page 4 of 7');
});

test('a field we cannot evaluate keeps the value Word last wrote', () => {
  // A stale value the author has seen beats a blank, and beats a guess.
  const band = parseBand('<w:hdr><w:p><w:fldSimple w:instr=" DOCPROPERTY &quot;RefNo&quot; ">'
    + '<w:r><w:t>RB-2026-0001</w:t></w:r></w:fldSimple></w:p></w:hdr>');
  assert.equal(resolveFields(band, { page: 2, of: 9 })[0].text, 'RB-2026-0001');
});

test('a first-page header only applies when the document asked for one', () => {
  const bands = { default: { paragraphs: ['D'] }, first: { paragraphs: ['F'] }, even: { paragraphs: ['E'] } };
  // Without titlePg the first-page reference is present but unused — honouring
  // it anyway puts a letterhead on page one of a document that does not want it.
  assert.equal(bandForPage(bands, 1, {}).paragraphs[0], 'D');
  assert.equal(bandForPage(bands, 1, { titlePage: true }).paragraphs[0], 'F');
  assert.equal(bandForPage(bands, 2, { evenAndOdd: true }).paragraphs[0], 'E');
  assert.equal(bandForPage(bands, 3, { evenAndOdd: true }).paragraphs[0], 'D');
});

test('pagination changes nothing about the file or the caret', () => {
  const view = openDocx(LETTER);
  const before = view.render();
  const target = before.blocks.find((b) => !b.structural && b.text.length > 5).index;

  view.setSelection({ block: target, offset: 2 });
  view.insertText('X');
  const after = view.render();

  // The paragraph list — the caret's address space — is the same length and the
  // caret is where it was. Pages are a rendering, not a model change.
  assert.equal(after.blocks.length, before.blocks.length);
  assert.deepEqual(after.selection.focus, { block: target, offset: 3 });
  assert.ok(after.blocks[target].text.startsWith(before.blocks[target].text.slice(0, 2) + 'X'));
});

test('a table splits between rows, never inside one', () => {
  const view = openDocx(LETTER);
  const table = view.render().flow.find((e) => e.kind === 'table').table;

  // Force the table onto a page too short for it, so it has to break.
  const blocks = [];
  const flow = [{ kind: 'table', table }];
  const laid = paginate({ flow, blocks, section: { ...A4, heightPx: 300 } });

  const placed = laid.pages.flatMap((p) => p.fragments.filter((f) => f.kind === 'table'));
  const rowsPlaced = placed.reduce((n, f) => n + f.rows.filter((r) => !r.header || !f.repeatedHeader).length, 0);
  assert.ok(rowsPlaced >= table.rows.length, 'every row is placed somewhere');
  // A header row repeats on each continuation — the only way a split table stays
  // readable — and the fixture's first row is marked as one.
  if (laid.count > 1) {
    assert.ok(placed.slice(1).every((f) => f.repeatedHeader), 'the header repeats after a break');
  }
});

test('a keystroke does not re-wrap the whole document', () => {
  // Found by measuring rather than by reading: a fifty-eight-page report cost
  // 2.6 SECONDS a keystroke. Two causes, both invisible on a short document —
  // `paragraph(index)` rescanned the whole body (quadratic in paragraphs), and
  // pagination re-wrapped every paragraph when one had changed.
  const cache = new Map();
  const blocks = Array.from({ length: 400 }, (_, i) => para(i, 'Paragraph ' + i + '. ' + 'word '.repeat(30)));
  const flow = flowOf(blocks);

  paginate({ flow, blocks, section: A4, cache });
  const afterFirst = cache.size;
  assert.ok(afterFirst > 300, 'the first pass fills the cache');

  // edit one paragraph; every other one must come back from the cache, so the
  // cache grows by exactly one entry rather than by four hundred
  const edited = blocks.map((b, i) => (i === 7 ? para(7, b.text + ' changed') : b));
  paginate({ flow, blocks: edited, section: A4, cache });
  assert.equal(cache.size, afterFirst + 1, 'exactly one paragraph was re-wrapped');
});

test('the cache is bounded', () => {
  const cache = new Map();
  for (let round = 0; round < 3; round++) {
    const blocks = Array.from({ length: 2000 }, (_, i) => para(i, 'r' + round + ' p' + i + '. ' + 'word '.repeat(20)));
    paginate({ flow: flowOf(blocks), blocks, section: A4, cache });
  }
  assert.ok(cache.size <= 4000, 'a long session must not keep every version of every paragraph');
});
