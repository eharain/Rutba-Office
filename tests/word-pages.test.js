// Pages on screen — the engine's half.
//
// The editor lays the flow onto sheets by measuring the drawn page (see
// apps/desktop/renderer/apps/word/pages.js, which needs a browser). What can be
// held here without one: the flags a paginator honours reach the frame the
// editor draws from, the builder writes them, and a paragraph's runs slice at a
// character offset the way a split paragraph needs.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { sliceRuns } from '../apps/desktop/renderer/apps/word/pages.js';

test('an explicit page break, keep-with-next and keep-lines reach the frame', () => {
  const bytes = buildDocx({
    styles: true,
    paragraphs: [
      { text: 'One' },
      { text: 'Two', pageBreakBefore: true },
      { text: 'Three', keepNext: true, keepLines: true, style: 'Heading1' },
      { text: 'Four' },
    ],
  });
  const view = openDocx(bytes);
  const { blocks } = view.render({ pages: false });
  assert.equal(blocks[0].pageBreakBefore, undefined, 'a plain paragraph carries no flag');
  assert.equal(blocks[1].pageBreakBefore, true, 'the break');
  assert.equal(blocks[2].keepNext, true, 'kept with the next');
  assert.equal(blocks[2].keepLines, true, 'kept together');
  assert.equal(blocks[3].keepNext, undefined);
});

test('a run-level page break (w:br type="page") marks the paragraph too', () => {
  const bytes = buildDocx({ paragraphs: [{ text: 'Before' }, { text: 'After' }] });
  // Word writes the break as a run inside the paragraph that follows it.
  const view = openDocx(bytes);
  const before = view.render({ pages: false }).blocks;
  assert.equal(before[1].pageBreakBefore, undefined);
  view.setSelection({ block: 1, offset: 0 });
  view.insertPageBreak();
  const after = view.render({ pages: false }).blocks;
  assert.equal(after.filter((b) => b.pageBreakBefore).length, 1, 'one paragraph starts a page');
});

test('sliceRuns cuts a paragraph at character offsets, keeping the formatting', () => {
  const runs = [
    { text: 'The quick ', bold: true },
    { text: 'brown fox ', italic: true },
    { text: '', noteMark: { n: 1 } },
    { text: 'jumps.' },
  ];
  assert.deepEqual(sliceRuns(runs, 0, Infinity), runs, 'the whole paragraph is the runs as they are');
  assert.deepEqual(sliceRuns(runs, 4, 15), [
    { text: 'quick ', bold: true },
    { text: 'brown', italic: true },
  ]);
  assert.deepEqual(sliceRuns(runs, 20, Infinity), [{ text: '', noteMark: { n: 1 } }, { text: 'jumps.' }], 'an empty run stays with the part its position falls in');
  assert.deepEqual(sliceRuns(runs, 0, 4), [{ text: 'The ', bold: true }]);
  assert.deepEqual(sliceRuns([], 0, 5), []);
  const first = sliceRuns(runs, 0, 10);
  assert.equal(first[0], runs[0], 'a run that is not cut keeps its identity');
});
