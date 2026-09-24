// Drop caps — the engine's half.
//
// Word makes one by splitting the paragraph's first letter into its own
// paragraph, framed (`w:framePr`) to stand a few lines deep, its run sized
// to match; the body paragraph is untouched. The paginator lays that letter
// as a float, exactly as it lays a floating picture, so the body wraps round
// it; the PDF writer draws the one glyph at its full size.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const STORY = 'Once upon a time, in a quiet village, a baker opened her shop before dawn.';
const LONG_STORY = 'Once upon a time, in a quiet coastal village, the baker opened her shop '
  + 'long before the sun had cleared the harbour wall, and the smell of new bread drifted '
  + 'down every lane before the fishing boats came home to the quay. ';

test('setDropCap frames the letter\'s pPr in schema order — after pStyle, before jc', () => {
  const view = openDocx(buildDocx({
    styles: true,
    paragraphs: [{ text: STORY, style: 'Heading1', align: 'center' }],
  }));
  view.collapseTo({ block: 0, offset: 0 });
  view.setDropCap({ lines: 3 });

  const pPr = view.doc.doc.editParagraph(0).pPr || '';
  assert.ok(/<w:pStyle\b[^>]*w:val="Heading1"\/>/.test(pPr), `the style rides: ${pPr}`);
  assert.ok(
    /<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"\/>/.test(pPr),
    `the frame is written exactly as Word writes it: ${pPr}`
  );
  assert.ok(/<w:jc\b[^>]*w:val="center"\/>/.test(pPr), `the alignment rides too: ${pPr}`);
  const order = [pPr.indexOf('<w:pStyle'), pPr.indexOf('<w:framePr'), pPr.indexOf('<w:jc')];
  assert.ok(order[0] < order[1] && order[1] < order[2], `schema order pStyle/framePr/jc: ${order}`);
});

test('setDropCap sizes the run 38pt for a default 11pt paragraph over three lines, and reports its state', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: STORY }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.setDropCap({ lines: 3 });

  // 3 lines * 11pt base * 1.15, rounded to the nearest half point.
  const expected = Math.round(3 * 11 * 1.15 * 2) / 2;
  assert.equal(expected, 38, 'the formula this test holds the engine to');

  const block = view.render({ pages: false }).blocks[0];
  assert.deepEqual(block.dropCap, { kind: 'drop', lines: 3 });
  assert.equal(block.text, 'O');
  assert.equal(block.runs[0].fontSize, 38);

  const body = view.render({ pages: false }).blocks[1];
  assert.equal(body.text, STORY.slice(1));
  assert.equal(body.dropCap, undefined, 'the body carries no frame of its own');

  assert.equal(view.focus.block, 1, 'the caret ends at the start of the body');
  assert.deepEqual(view.formatAtCaret().dropCap, { kind: 'drop', lines: 3 }, 'the ribbon reads the pair from the body too');
});

test('a drop cap survives a save and reopen', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: STORY }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.setDropCap({ lines: 3 });

  const reopened = openDocx(view.save());
  const block = reopened.render({ pages: false }).blocks[0];
  assert.deepEqual(block.dropCap, { kind: 'drop', lines: 3 });
  assert.equal(block.text, 'O');
  assert.equal(block.runs[0].fontSize, 38);
  assert.equal(reopened.render({ pages: false }).blocks[1].text, STORY.slice(1));
});

test('render lays the drop cap as a float; the body wraps three lines beside it, clear on the fourth', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: LONG_STORY.repeat(2) }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.setDropCap({ lines: 3 });

  const laid = view.pages;
  const first = laid.pages[0];
  const drop = first.fragments.find((f) => f.kind === 'dropcap');
  assert.ok(drop, 'the drop cap is placed as a float, not a paragraph');
  assert.equal(drop.text, 'O');
  assert.equal(drop.paragraphIndex, 0);
  assert.ok(drop.sizePx > 30, `the letter is drawn well past ordinary text: ${drop.sizePx}px`);
  assert.ok(first.fragments.every((f) => !(f.kind === 'paragraph' && f.paragraphIndex === 0)), 'never laid out as a paragraph too');

  const lines = first.fragments.filter((f) => f.kind === 'paragraph' && f.paragraphIndex === 1).flatMap((f) => f.lines);
  assert.ok(lines.length >= 4, `the body wraps at least four lines: ${lines.length}`);
  assert.ok(lines.slice(0, 3).every((l) => (l.offsetPx || 0) > 0), `the first three lines stand beside the letter: ${lines.slice(0, 4).map((l) => l.offsetPx)}`);
  assert.equal(lines[3].offsetPx || 0, 0, 'the fourth line clears the letter');
});

test('a drop cap with nothing after it lays out as a plain paragraph', () => {
  // setDropCap always leaves a body paragraph behind it, so this shape only
  // arises from a file framed by hand (or by Word) with nothing following —
  // the last paragraph in the document. Framed directly, the way the reader
  // must still cope with one it did not write itself.
  const built = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Filler paragraph.' }, { text: 'Z' }] }));
  built.doc.setParagraphProp(1, 'dropCap', { kind: 'drop', lines: 3 });
  const view = openDocx(built.save());

  const dropIndex = view.blocks.findIndex((b) => b.dropCap);
  assert.equal(dropIndex, 1, 'the last paragraph carries the frame');
  const laid = view.pages;
  const fragments = laid.pages.flatMap((p) => p.fragments);
  assert.ok(fragments.some((f) => f.kind === 'paragraph' && f.paragraphIndex === dropIndex), 'laid as an ordinary paragraph');
  assert.ok(!fragments.some((f) => f.kind === 'dropcap'), 'never floated with nothing to stand beside');
});

test('setDropCap(null) merges the letter back, with no framePr and no run-size override', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: STORY }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.setDropCap({ lines: 3 });
  view.setDropCap(null);

  assert.equal(view.blocks.length, 1, 'the pair is one paragraph again');
  const block = view.render({ pages: false }).blocks[0];
  assert.equal(block.text, STORY);
  assert.equal(block.dropCap, undefined);
  assert.equal(block.runs[0].fontSize, undefined, 'the size override is gone');

  const pPr = view.doc.doc.editParagraph(0).pPr || '';
  assert.ok(!/w:framePr/.test(pPr), 'the frame is gone');
  const xml = view.doc.doc.editParagraph(0).xml;
  assert.ok(!/<w:sz\b/.test(xml), 'no explicit size left on the run');
});

test('setDropCap refuses an empty paragraph and a table cell', () => {
  const view = openDocx(buildDocx({ paragraphs: ['', 'Enough words here to hold a table after it.'] }));
  view.collapseTo({ block: 0, offset: 0 });
  assert.throws(() => view.setDropCap({ lines: 3 }), /Put the caret in a paragraph that starts with a letter/);

  view.collapseTo({ block: 1, offset: 0 });
  view.insertTable({ rows: 1, cols: 1 });
  view.insertText('Cell words');
  assert.throws(() => view.setDropCap({ lines: 3 }), /Put the caret in a paragraph that starts with a letter/);
});

test('a PDF export draws the letter once, at a size past an ordinary one', async () => {
  const { renderPdf } = await import('@rutba/doc-view/export/pdf');
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: LONG_STORY.repeat(2) }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.setDropCap({ lines: 3 });

  const { buffer } = renderPdf(view, { created: '2026-09-20T00:00:00Z' });
  const pdf = buffer.toString('latin1');
  const matches = [...pdf.matchAll(/\/(\S+) ([\d.]+) Tf[\s\S]{0,150}?\(O\) Tj/g)];
  assert.equal(matches.length, 1, `the letter is drawn exactly once: found ${matches.length}`);
  assert.ok(Number(matches[0][2]) > 30, `the letter prints past an ordinary size: ${matches[0][2]}pt`);
});

test('undo puts the paragraph back', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: STORY }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.setDropCap({ lines: 3 });
  assert.equal(view.blocks.length, 2);

  const undone = view.undo();
  assert.equal(undone, true);
  assert.equal(view.blocks.length, 1);
  assert.equal(view.block(0).text, STORY);
  assert.equal(view.block(0).dropCap, null, 'the frame is undone too');
});
