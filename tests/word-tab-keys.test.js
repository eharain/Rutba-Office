/**
 * Tab and Shift+Tab in a document (2.18.0, A3.2).
 *
 * In 2.17.1 Tab moved the keyboard off the page. Now it does what it does in
 * Word: at the start of a list item it moves the item a level in, Shift+Tab
 * a level out; in a table it walks to the next or the previous cell, and in
 * the last cell it adds a row; anywhere else it types a tab.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

const levelOf = (view, block) => view.doc.getParagraphProps(block)?.listLevel ?? 0;

test('Tab at the start of a list item moves it a level in, Shift+Tab a level out', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Opening jobs', 'Unlock the door', 'Turn the till on'] }));
  view.setSelection({ block: 1, offset: 0 }, { block: 2, offset: 4 });
  view.setParagraphFormat({ list: 'bullet' });
  view.collapseTo({ block: 2, offset: 0 });
  view.tabCell();
  assert.equal(levelOf(view, 2), 1, 'a level in');
  assert.equal(view.block(2).text, 'Turn the till on', 'and no tab typed');
  view.collapseTo({ block: 2, offset: 5 });
  view.tabCell({ back: true });
  assert.equal(levelOf(view, 2), 0, 'Shift+Tab anywhere in it, a level out');
  // Both items selected: both go in.
  view.setSelection({ block: 1, offset: 3 }, { block: 2, offset: 3 });
  view.tabCell();
  assert.deepEqual([levelOf(view, 1), levelOf(view, 2)], [1, 1]);
  view.undo();
  assert.deepEqual([levelOf(view, 1), levelOf(view, 2)], [0, 0], 'one undo puts them back');
});

test('Tab inside the words of a list item, or in any other paragraph, types a tab', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Name', 'Item'] }));
  view.collapseTo({ block: 0, offset: 4 });
  view.tabCell();
  view.insertText('Larkspur');
  assert.equal(view.block(0).text, 'Name\tLarkspur');
  view.collapseTo({ block: 1, offset: 0 });
  view.setParagraphFormat({ list: 'number' });
  view.collapseTo({ block: 1, offset: 4 });
  view.tabCell();
  assert.equal(view.block(1).text, 'Item\t');
  assert.equal(levelOf(view, 1), 0);
});

test('in a table Tab walks to the next cell, Shift+Tab to the one before, and the last cell adds a row', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Price list'] }));
  view.collapseTo({ block: 0, offset: 10 });
  view.insertTable({ rows: 2, cols: 2 });
  const first = view.focus.block;
  view.insertText('Tea');
  view.tabCell();
  assert.equal(view.focus.block, first + 1);
  view.insertText('2.40');
  view.tabCell({ back: true });
  assert.equal(view.focus.block, first, 'back to the first cell');
  view.tabCell();
  view.tabCell();
  view.tabCell();
  const count = view.blocks.length;
  assert.equal(view.focus.block, first + 3, 'the last cell');
  view.tabCell();
  assert.equal(view.blocks.length, count + 2, 'a row of two cells added');
  assert.equal(view.focus.block, first + 4, 'the caret in its first cell');
  assert.ok(view.blocks[first + 4].container, 'which is in the table');
});
