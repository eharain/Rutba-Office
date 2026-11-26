/**
 * The Name Box takes typing. Before 2.28 it only showed the selection's
 * address: nobody could type B2:D9 to select it, type a name to reach its
 * range, or type a new word to name what was selected. Go To took one cell
 * and nothing wider. These tests hold the reading of an entry (an address,
 * a name that exists, a new name) and the going itself.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { parseAddress, nameBoxEntry } from '@rutba/sheet-view/name-box';
import { buildXlsx } from '@rutba/ooxml/build';

const book = () => new SheetView(buildXlsx({
  sheets: [
    { name: 'Takings', rows: [['Month', 'Shop', 'Web'], ['Jan', 120, 60], ['Feb', 80, 75], ['Mar', 95, 90]] },
    { name: 'May sales', rows: [['Item', 'Sold'], ['Tea', 4], ['Rice', 9]] },
    { name: 'Stock', rows: [['Item', 'Held'], ['Tea', 40]] },
  ],
}));
const range = (view) => ({ ...view.selection.range });

test('an address is read: a cell, a range, columns, rows, with or without a sheet', () => {
  assert.deepEqual(parseAddress('B7'), { sheet: null, kind: 'cells', top: 6, left: 1, bottom: 6, right: 1 });
  assert.deepEqual(parseAddress(' $b$2:$d$9 '), { sheet: null, kind: 'cells', top: 1, left: 1, bottom: 8, right: 3 });
  // Either corner first: the range is the same.
  assert.deepEqual(parseAddress('D9:B2'), parseAddress('B2:D9'));
  assert.deepEqual(parseAddress('A:C'), { sheet: null, kind: 'cols', top: 0, left: 0, bottom: 1048575, right: 2 });
  assert.deepEqual(parseAddress('3:7'), { sheet: null, kind: 'rows', top: 2, left: 0, bottom: 6, right: 16383 });
  assert.equal(parseAddress('Stock!B2').sheet, 'Stock');
  assert.equal(parseAddress("'May sales'!A1:B3").sheet, 'May sales');
  assert.equal(parseAddress("'Tom''s'!A1").sheet, "Tom's");
  assert.equal(parseAddress('XFD1048576').right, 16383);
});

test('what is not an address is not read as one', () => {
  for (const text of ['', 'Sales', 'Sales_May', 'A0', 'XFE1', 'A1048577', 'B2:', 'A:1', 'B2:D', '12', 'Stock!', 'Stock!Sales']) {
    assert.equal(parseAddress(text), null, text);
  }
});

test('an entry is decided in Excel\'s order: an address, a name that exists, a new name', () => {
  const names = [{ name: 'Sales_May' }, { name: 'Rate' }];
  assert.deepEqual(nameBoxEntry('b2:d9', names), { kind: 'ref', text: 'b2:d9' });
  // A name is found in any case and given back as the workbook spells it.
  assert.deepEqual(nameBoxEntry(' sales_may ', names), { kind: 'name', name: 'Sales_May' });
  assert.deepEqual(nameBoxEntry('Costs', names), { kind: 'define', name: 'Costs' });
  assert.equal(nameBoxEntry('   ', names), null);
  // A word that cannot be a name is said to be so, not sent to the workbook.
  assert.equal(nameBoxEntry('May sales', names).kind, 'bad');
  assert.equal(nameBoxEntry('9lives', names).kind, 'bad');
  assert.equal(nameBoxEntry('XFE1', names).kind, 'bad');
});

test('a typed cell is gone to and a typed range is selected, its first cell active', () => {
  const view = book();
  view.goToRef('C3');
  assert.deepEqual(range(view), { top: 2, left: 2, bottom: 2, right: 2 });
  view.goToRef('b2:c4');
  assert.deepEqual(range(view), { top: 1, left: 1, bottom: 3, right: 2 });
  assert.deepEqual({ ...view.selection.active }, { row: 1, col: 1 });
  assert.equal(view.render().selection.ref, 'B2:C4');
});

test('whole columns and whole rows are selected', () => {
  const view = book();
  view.goToRef('B:C');
  const cols = range(view);
  assert.equal(cols.left, 1);
  assert.equal(cols.right, 2);
  assert.equal(cols.top, 0);
  assert.ok(cols.bottom >= 3, 'the columns reach past the last row in use');
  view.goToRef('2:3');
  const rows = range(view);
  assert.equal(rows.top, 1);
  assert.equal(rows.bottom, 2);
  assert.equal(rows.left, 0);
  assert.ok(rows.right >= 2, 'the rows reach past the last column in use');
});

test('a sheet before the address is gone to, in any case, quoted or not', () => {
  const view = book();
  view.goToRef('stock!B2');
  assert.equal(view.activeSheet, 'Stock');
  assert.deepEqual(range(view), { top: 1, left: 1, bottom: 1, right: 1 });
  view.goToRef("'May sales'!A2:B3");
  assert.equal(view.activeSheet, 'May sales');
  assert.deepEqual(range(view), { top: 1, left: 0, bottom: 2, right: 1 });
});

test('a far cell is brought into sight', () => {
  const view = book();
  view.goToRef('Z500');
  assert.ok(view.scrollY > 0, 'the sheet has scrolled down');
  assert.ok(view.scrollX > 0, 'the sheet has scrolled across');
  assert.deepEqual(range(view), { top: 499, left: 25, bottom: 499, right: 25 });
});

test('a bad address, a sheet that is not there and a hidden sheet are refused, and nothing moves', () => {
  const view = book();
  view.goToRef('B2');
  assert.throws(() => view.goToRef('Sales'), /is not a cell or a range/);
  assert.throws(() => view.goToRef('Nowhere!B2'), /no such sheet/);
  view.hideSheet('Stock');
  assert.throws(() => view.goToRef('Stock!B2'), /is hidden/);
  assert.equal(view.activeSheet, 'Takings');
  assert.deepEqual(range(view), { top: 1, left: 1, bottom: 1, right: 1 });
});

test('going somewhere changes nothing in the file and makes no undo step', () => {
  const view = book();
  const before = view.canUndo;
  view.goToRef('B2:C4');
  assert.equal(view.canUndo, before);
  const again = SheetView.open(view.save());
  assert.equal(again.render().cells.find((c) => c.ref === 'B2')?.text, '120');
});

test('a new word names the selection, the name is reached by typing it, and the file keeps it', () => {
  const view = book();
  view.goToRef('B2:C4');
  // What the Name Box sends for a new word: a name for what is selected.
  const entry = nameBoxEntry('Takings_Q1', view.names());
  assert.equal(entry.kind, 'define');
  view.defineName(entry.name);
  view.goToRef("'May sales'!A1");
  const back = nameBoxEntry('takings_q1', view.names());
  assert.deepEqual(back, { kind: 'name', name: 'Takings_Q1' });
  view.gotoName(back.name);
  assert.equal(view.activeSheet, 'Takings');
  assert.deepEqual(range(view), { top: 1, left: 1, bottom: 3, right: 2 });
  // One undo step takes the name away again.
  view.undo();
  assert.equal(view.names().some((n) => n.name === 'Takings_Q1'), false);
  view.redo();
  const again = SheetView.open(view.save());
  again.gotoName('Takings_Q1');
  assert.deepEqual(range(again), { top: 1, left: 1, bottom: 3, right: 2 });
});
