// Home → Format Painter paints.
//
// The button picked up the selection's formatting and nothing ever painted
// it: the window never asked for the paint. The frame now says while the
// brush is loaded, so the window paints the next selection made with the
// mouse; one press paints once, a double press keeps painting until the
// brush is put down.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { formatOf } from '../packages/sheet-view/src/styles-write.js';

const BOOK = () => buildXlsx({ sheets: [{ name: 'Data', rows: [['Head', 'Sum'], [1, 2], [3, 4], [5, 6]] }] });
const look = (view, row, col) => {
  const at = formatOf(view.pkg.text('xl/styles.xml'), view._styleIndexAt('Data', row, col) ?? 0);
  return [at.bold, at.italic, at.fill];
};
const styled = () => {
  const view = new SheetView(BOOK());
  view.select(0, 0);
  view.setFormat({ bold: true, fill: '#FFF2CC' });
  view.select(0, 1);
  view.setFormat({ italic: true });
  return view;
};

test('one press: the frame says the brush is loaded, the next selection is painted, and the brush is put down', () => {
  const view = styled();
  assert.equal(view.selectionFrame().brush, null);
  view.select(0, 0);
  view.markFormatBrush();
  assert.deepEqual(view.selectionFrame().brush, { keep: false });
  assert.deepEqual(view.render().brush, { keep: false }, 'in the whole frame too');

  view.select(2, 0);
  view.select(3, 1, { extend: true });
  view.paintFormat();
  for (const [r, c] of [[2, 0], [2, 1], [3, 0], [3, 1]]) assert.deepEqual(look(view, r, c), [true, false, 'FFFFF2CC'], 'row ' + r + ' column ' + c);
  assert.equal(view.selectionFrame().brush, null, 'used once');
  assert.equal(view.calc.getInput('Data', 2, 0), 3, 'the values are as they were');

  view.select(1, 0);
  view.paintFormat();
  assert.deepEqual(look(view, 1, 0), [false, false, null], 'nothing is painted without a brush');
  view.undo();
  assert.deepEqual(look(view, 2, 0), [false, false, null], 'the paint is one undo step');
});

test('a double press keeps the brush until it is put down', () => {
  const view = styled();
  view.select(0, 0);
  view.markFormatBrush({ keep: true });
  assert.deepEqual(view.selectionFrame().brush, { keep: true });
  view.select(1, 0);
  view.paintFormat();
  view.select(3, 1);
  view.paintFormat();
  assert.deepEqual([look(view, 1, 0), look(view, 3, 1)], [[true, false, 'FFFFF2CC'], [true, false, 'FFFFF2CC']]);
  assert.deepEqual(view.selectionFrame().brush, { keep: true });
  view.clearFormatBrush();
  assert.equal(view.selectionFrame().brush, null);
});

test('a single cell picked as the target takes the whole of what was picked up', () => {
  const view = styled();
  view.select(0, 0);
  view.select(0, 1, { extend: true });
  view.markFormatBrush();
  view.select(2, 0);
  view.paintFormat();
  assert.deepEqual([look(view, 2, 0), look(view, 2, 1)], [[true, false, 'FFFFF2CC'], [false, true, null]]);
  const r = view.selection.range;
  assert.deepEqual([r.top, r.left, r.bottom, r.right], [2, 0, 2, 1], 'the selection ends on what was painted');
});

test('a brush whose sheet has gone paints nothing and is put down', () => {
  const view = styled();
  view.markFormatBrush({ keep: true });
  view.formatBrush.sheet = 'Gone';
  view.select(2, 0);
  view.paintFormat();
  assert.deepEqual(look(view, 2, 0), [false, false, null]);
  assert.equal(view.selectionFrame().brush, null);
});
