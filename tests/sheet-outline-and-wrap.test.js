// Home → Borders → Outside borders, and Home → Wrap Text.
//
// Outside borders (and Thick outside borders) put the line round every cell
// of the selection, so it looked the same as All borders: the window sent
// the same four edges for each cell. Wrap Text left the row one line tall,
// so a wrapped cell's second line was squeezed out of sight.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const BOOK = () => buildXlsx({ sheets: [{ name: 'S', rows: [
  ['a', 'b', 'c', 'd'],
  ['e', 'f', 'g', 'h'],
  ['i', 'j', 'k', 'l'],
  ['m', 'n', 'o', 'p'],
] }] });
const THIN = { style: 'thin', colour: '#000000' };
const MEDIUM = { style: 'medium', colour: '#000000' };
// Which edges a cell wears, as letters: T, B, L, R. A thin edge is drawn one pixel wide, a medium one two.
const WIDTH = { thin: 1, medium: 2 };
const edges = (view, row, col, style = null) => {
  const b = view.styleFor(row, col)?.border ?? {};
  return ['top', 'bottom', 'left', 'right'].filter((k) => b[k] && b[k].style && b[k].style !== 'none' && (!style || b[k].widthPx === WIDTH[style])).map((k) => k[0].toUpperCase()).join('');
};
const pick = (view, top, left, bottom, right) => { view.select(top, left); view.select(bottom, right, { extend: true }); };

test('Outside borders line the outer edge of the selection and nothing inside it', () => {
  const view = new SheetView(BOOK());
  pick(view, 0, 0, 2, 2);
  view.setFormat({ outline: THIN });
  assert.deepEqual(
    [0, 1, 2].map((r) => [0, 1, 2].map((c) => edges(view, r, c))),
    [['TL', 'T', 'TR'], ['L', '', 'R'], ['BL', 'B', 'BR']],
  );
  assert.equal(edges(view, 3, 3), '', 'a cell outside is untouched');
  view.undo();
  assert.equal(edges(view, 0, 0), '', 'one undo step');
});

test('Thick outside borders keep the lines already inside', () => {
  const view = new SheetView(BOOK());
  pick(view, 0, 0, 1, 1);
  view.setFormat({ border: { top: THIN, bottom: THIN, left: THIN, right: THIN } });
  view.setFormat({ outline: MEDIUM });
  assert.equal(edges(view, 0, 0, 'medium'), 'TL');
  assert.equal(edges(view, 0, 0, 'thin'), 'BR', 'the inside edges are still thin');
  assert.equal(edges(view, 1, 1, 'medium'), 'BR');

  // One cell takes all four; a single row its top and bottom all along.
  view.select(3, 3);
  view.setFormat({ outline: THIN });
  assert.equal(edges(view, 3, 3), 'TBLR');
  pick(view, 2, 0, 2, 2);
  view.setFormat({ outline: THIN });
  assert.deepEqual([0, 1, 2].map((c) => edges(view, 2, c)), ['TBL', 'TB', 'TBR']);
});

test('the outline is in the saved file', () => {
  let view = new SheetView(BOOK());
  pick(view, 1, 1, 2, 2);
  view.setFormat({ outline: MEDIUM });
  view = SheetView.open(view.save());
  assert.deepEqual([[1, 1], [1, 2], [2, 1], [2, 2]].map(([r, c]) => edges(view, r, c, 'medium')), ['TL', 'TR', 'BL', 'BR']);
});

const LONG = 'These words are far too long to fit on one line of a narrow column';

test('Wrap Text makes the row as tall as the words need, and turning it off puts the row back', () => {
  const view = new SheetView(BOOK());
  const one = view.geo.rowHeight(1);
  view.setCell(1, 0, LONG);
  assert.equal(view.geo.rowHeight(1), one, 'unwrapped words leave the row alone');
  view.select(1, 0);
  view.setFormat({ wrap: true });
  const tall = view.geo.rowHeight(1);
  assert.ok(tall >= one * 3, 'the row holds several lines: ' + tall + ' against ' + one);
  assert.equal(view.geo.rowHeight(0), one, 'the other rows are as they were');

  const saved = SheetView.open(view.save());
  assert.equal(saved.geo.rowHeight(1), tall, 'the height is in the file');
  assert.ok(!/customHeight/.test(saved.workbook.snapshotParts([saved.workbook.partNameFor('S')])[saved.workbook.partNameFor('S')]), 'and is not marked as chosen by hand');

  view.undo();
  assert.equal(view.geo.rowHeight(1), one, 'undo takes the wrap and the height back together');
  assert.equal(view.formatState().wrap, false);
  view.redo();
  assert.equal(view.geo.rowHeight(1), tall);
  view.setFormat({ wrap: false });
  assert.equal(view.geo.rowHeight(1), one);
});

test('words typed into a wrapping cell grow its row; a height chosen by hand is kept', () => {
  const view = new SheetView(BOOK());
  const one = view.geo.rowHeight(2);
  view.select(2, 1);
  view.setFormat({ wrap: true });
  assert.equal(view.geo.rowHeight(2), one, 'one short word needs one line');
  view.setCell(2, 1, LONG);
  const tall = view.geo.rowHeight(2);
  assert.ok(tall >= one * 3);
  view.setCell(2, 1, 'short');
  assert.equal(view.geo.rowHeight(2), one, 'and shrinks again');
  view.setCell(2, 1, 'two\nlines');
  assert.ok(view.geo.rowHeight(2) > one && view.geo.rowHeight(2) < tall, 'a line break is a line');

  view.setRowHeight(3, 60);
  view.select(3, 0);
  view.setFormat({ wrap: true });
  view.setCell(3, 0, LONG);
  assert.equal(view.geo.rowHeight(3), 60, 'a row sized by hand stays that size');
});
