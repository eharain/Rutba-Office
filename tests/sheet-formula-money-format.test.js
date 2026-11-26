// A formula that makes an amount of money is shown as one.
//
// =PMT(5%/12,60,10000) typed in a General cell showed -188.71233644011,
// and =IPMT and =PPMT the same way: right to the last digit and hard to
// read. Excel gives these a money format. A typed formula that is one loan
// or investment function now brings thousands grouped and two decimals; a
// format already on the cell is kept, and a longer formula brings none.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { formulaFormat } from '../packages/sheet-view/src/typed.js';

const BOOK = () => buildXlsx({ sheets: [{ name: 'S', rows: [['Loan', 'Payment'], [], [], [], [], []] }] });
const type = (view, row, col, text) => {
  view.select(row, col);
  view.beginEdit({ replace: true, initial: '' });
  view.updateDraft(text);
  view.commitEdit({ move: 'none' });
};
const shown = (view, row, col) => view.displayValue(row, col).text;

test('the functions that make money bring the format, and nothing else does', () => {
  for (const name of ['PMT', 'IPMT', 'PPMT', 'CUMIPMT', 'CUMPRINC', 'PV', 'FV', 'NPV']) {
    assert.equal(formulaFormat('=' + name + '(1,2,3)'), '#,##0.00', name);
  }
  assert.equal(formulaFormat('=pmt(5%/12,60,10000)'), '#,##0.00', 'in small letters too');
  assert.equal(formulaFormat('=-PMT(5%/12,60,10000)'), '#,##0.00', 'with its sign turned');
  assert.equal(formulaFormat('= - PMT(5%/12,60,10000)'), '#,##0.00');
  assert.equal(formulaFormat('=PMT(5%/12,60,10000)*2'), null, 'a longer formula is left alone');
  assert.equal(formulaFormat('=PMT(5%/12,60,10000)+PMT(5%/12,60,500)'), null);
  assert.equal(formulaFormat('=ROUND(PMT(5%/12,60,10000),0)'), null, 'the outer function decides');
  assert.equal(formulaFormat('=SUM(1,2)'), null);
  assert.equal(formulaFormat('=RATE(60,-188.71,10000)'), null, 'a rate is not money');
  assert.equal(formulaFormat('=NPER(5%/12,-188.71,10000)'), null, 'nor is a count of payments');
  assert.equal(formulaFormat('PMT(1,2,3)'), null, 'not a formula');
});

test('a payment typed in a General cell is shown with two decimals, and still is after a save', () => {
  let view = new SheetView(BOOK());
  type(view, 1, 1, '=PMT(5%/12,60,10000)');
  assert.ok(Math.abs(view.calc.getValue('S', 1, 1) + 188.7123364401099) < 1e-9, 'the value is kept whole');
  assert.equal(view.formatFor(1, 1), '#,##0.00');
  assert.equal(shown(view, 1, 1), '-188.71');
  type(view, 2, 1, '=IPMT(8%/12,1,36,10000)');
  assert.equal(shown(view, 2, 1), '-66.67');
  type(view, 3, 1, '=FV(6%/12,120,-250)');
  assert.equal(shown(view, 3, 1), '40,969.84', 'thousands are grouped');
  view = SheetView.open(view.save());
  assert.equal(view.formatFor(1, 1), '#,##0.00');
  assert.equal(shown(view, 1, 1), '-188.71');
  assert.equal(shown(view, 3, 1), '40,969.84');
});

test('a cell with a format of its own keeps it, and a longer formula stays General', () => {
  const view = new SheetView(BOOK());
  view.select(1, 1);
  view.setFormat({ numberFormat: '0.0000' });
  type(view, 1, 1, '=PMT(5%/12,60,10000)');
  assert.equal(view.formatFor(1, 1), '0.0000');
  assert.equal(shown(view, 1, 1), '-188.7123');
  type(view, 2, 1, '=PMT(5%/12,60,10000)*12');
  assert.equal(view.formatFor(2, 1), 'General');
});

test('one undo takes the payment and its format away together', () => {
  const view = new SheetView(BOOK());
  type(view, 1, 1, '=PPMT(8%/12,1,36,10000)');
  assert.equal(view.formatFor(1, 1), '#,##0.00');
  view.undo();
  assert.equal(view.calc.getInput('S', 1, 1) ?? '', '');
  assert.equal(view.formatFor(1, 1), 'General');
});
