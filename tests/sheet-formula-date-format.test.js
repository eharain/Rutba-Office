// A formula that makes a date is shown as a date.
//
// =DATE(2026,11,30) typed in the empty cell under a column of dates showed
// the day count 46356: the cell had no number format of its own and the
// formula brought none. A typed formula whose answer is a date or a time
// now brings the format, as Excel's does; a format already on the cell is
// kept as it always was.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { readDatesAs } from '@rutba/formula';
import { formulaFormat } from '../packages/sheet-view/src/typed.js';

test.beforeEach(() => readDatesAs({ dayFirst: true, year: 2026 }));
test.after(() => readDatesAs());

const BOOK = () => buildXlsx({ sheets: [{ name: 'S', rows: [['Day', 'Qty'], [], [], [], []] }] });
const type = (view, row, col, text) => {
  view.select(row, col);
  view.beginEdit({ replace: true, initial: '' });
  view.updateDraft(text);
  view.commitEdit({ move: 'none' });
};
const shown = (view, row, col) => view.displayValue(row, col).text;

test('=DATE under a column of dates is shown as a date, and still is after a save', () => {
  let view = new SheetView(BOOK());
  type(view, 1, 0, '28/11/2026');
  type(view, 2, 0, '29/11/2026');
  type(view, 3, 0, '=DATE(2026,11,30)');
  assert.equal(view.calc.getValue('S', 3, 0), 46356);
  assert.equal(view.formatFor(3, 0), 'dd/mm/yyyy');
  assert.equal(shown(view, 3, 0), '30/11/2026');
  assert.equal(view.calc.getInput('S', 3, 0), '=DATE(2026,11,30)', 'the formula is the formula');

  view = SheetView.open(view.save());
  assert.equal(shown(view, 3, 0), '30/11/2026');
  view.undo();
});

test('a format already on the cell is kept under a formula', () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  view.setFormat({ numberFormat: 'd mmm yyyy' });
  type(view, 1, 0, '=DATE(2026,11,30)');
  assert.equal(view.formatFor(1, 0), 'd mmm yyyy');
  assert.equal(shown(view, 1, 0), '30 Nov 2026');

  view.select(2, 0);
  view.setFormat({ numberFormat: '0.00' });
  type(view, 2, 0, '=TODAY()');
  assert.equal(view.formatFor(2, 0), '0.00', 'a number format chosen by hand is not replaced');
});

test('a cell, or a cell with days added, takes the date format of the cell it reads', () => {
  const view = new SheetView(BOOK());
  type(view, 1, 0, '28 Nov 2026');
  type(view, 2, 0, '=A2+7');
  assert.equal(shown(view, 2, 0), '5 Dec 2026');
  type(view, 3, 0, '=$A$2');
  assert.equal(shown(view, 3, 0), '28 Nov 2026');
  type(view, 1, 1, '12');
  type(view, 2, 1, '=B2+7');
  assert.equal(view.formatFor(2, 1), 'General', 'a number plus seven is a number');
  assert.equal(shown(view, 2, 1), '19');
});

test('which formulas bring a format', () => {
  assert.equal(formulaFormat('=DATE(2026,11,30)'), 'dd/mm/yyyy');
  assert.equal(formulaFormat('= today()'), 'dd/mm/yyyy');
  assert.equal(formulaFormat('=EOMONTH(A1,0)'), 'dd/mm/yyyy');
  assert.equal(formulaFormat('=NOW()'), 'dd/mm/yyyy hh:mm');
  assert.equal(formulaFormat('=TIME(9,30,0)'), 'hh:mm');
  readDatesAs({ dayFirst: false, year: 2026 });
  assert.equal(formulaFormat('=DATE(2026,11,30)'), 'mm/dd/yyyy');
  // Not a date as a whole: a difference of days, a year, words.
  assert.equal(formulaFormat('=DATE(2026,11,30)-DATE(2026,1,1)'), null);
  assert.equal(formulaFormat('=YEAR(TODAY())'), null);
  assert.equal(formulaFormat('=TEXT(TODAY(),")")&"(x)"'), null);
  assert.equal(formulaFormat('=SUM(A1:A9)'), null);
  assert.equal(formulaFormat('=A1+7'), null, 'with no way to ask the cell, nothing');
  assert.equal(formulaFormat('=A1*2', () => 'dd/mm/yyyy'), null, 'twice a date is not a date');
  assert.equal(formulaFormat('30/11/2026'), null);
});
