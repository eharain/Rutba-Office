// Sort — the engine's half: a run of body paragraphs in the order of their
// words, A to Z or Z to A, numbers in their order and case set aside, each
// paragraph moving whole with its look; the whole body when nothing is
// selected; a table's cells refused.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const texts = (view) => view.blocks.map((b) => b.text);

test('the selected paragraphs sort A to Z and Z to A, numbers in their order, the look moving with each, and undo in one step', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Fruit', style: 'Title' }, { text: 'pear' }, { text: 'Apple', style: 'Heading1' }, { text: 'item 10' }, { text: 'item 9' },
  ] }));
  view.setSelection({ block: 1, offset: 0 }, { block: 4, offset: 0 });
  view.sortParagraphs();
  assert.deepEqual(texts(view), ['Fruit', 'Apple', 'item 9', 'item 10', 'pear'], 'A to Z, case aside, 9 before 10');
  assert.equal(view.doc.doc.editParagraph(1).style, 'Heading1', 'the look moved with the words');
  assert.equal(view.focus.block, 1, 'the caret on the first of them');

  view.setSelection({ block: 1, offset: 0 }, { block: 4, offset: 0 });
  view.sortParagraphs({ descending: true });
  assert.deepEqual(texts(view), ['Fruit', 'pear', 'item 10', 'item 9', 'Apple']);

  view.undo();
  assert.deepEqual(texts(view), ['Fruit', 'Apple', 'item 9', 'item 10', 'pear'], 'one undo step');
  view.undo();
  assert.deepEqual(texts(view), ['Fruit', 'pear', 'Apple', 'item 10', 'item 9'], 'and another');

  // Nothing selected: the whole body.
  view.collapseTo({ block: 3, offset: 0 });
  view.sortParagraphs();
  assert.deepEqual(texts(view), ['Apple', 'Fruit', 'item 9', 'item 10', 'pear']);
});

test('a single paragraph is left alone, and a run with a table in it is refused', () => {
  const one = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'only' }] }));
  one.sortParagraphs();
  assert.deepEqual(texts(one), ['only']);
  assert.equal(one.canUndo, false, 'nothing to undo');

  const tabled = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'b' }, { table: { rows: [['x', 'y']] } }, { text: 'a' }] }));
  tabled.collapseTo({ block: 0, offset: 0 });
  assert.throws(() => tabled.sortParagraphs(), /table/);
  assert.equal(texts(tabled)[0], 'b', 'untouched');
});
