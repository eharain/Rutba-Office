/**
 * The filter arrows' menu: what it lists, a value list applied, and its
 * Sort A to Z and Z to A, which sort the rows under the heading and keep the
 * filter's hidden rows the ones it excludes.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

const book = () => buildXlsx({
  sheets: [{
    name: 'Data',
    rows: [['Item', 'Region', 'Qty'], ['Ink', 'North', 4], ['Paper', 'South', 6], ['Pens', 'North', 2], ['Tape', '', 9]],
  }],
});

const filtered = () => {
  const view = SheetView.open(book());
  view.select(1, 0);
  view.toggleAutoFilter();
  return view;
};

test('the menu opens on a column with its values listed and none excluded, and closes', () => {
  const view = filtered();
  view.openFilterPanel('#sheet', 'Region');
  const panel = view.render().filterPanel;
  assert.deepEqual(panel.values, ['', 'North', 'South'], 'blank first, as text sorts');
  assert.equal(panel.selected, null);
  assert.equal(panel.column, 'Region');
  view.openFilterPanel(null);
  assert.equal(view.render().filterPanel, null);
});

test('the value list applied hides the rows it leaves out, and the arrow knows it filters', () => {
  const view = filtered();
  view.openFilterPanel('#sheet', 'Region');
  view.applyFilter('#sheet', 'Region', ['North']);
  assert.equal(view.render().filterPanel, null, 'applying closes the menu');
  assert.equal(view.geo.rowHeight(2), 0, 'South hidden');
  assert.equal(view.geo.rowHeight(4), 0, 'the blank hidden');
  assert.ok(view.geo.rowHeight(3) > 0, 'Pens, North, shows');
  assert.deepEqual(view.render().tables.find((t) => t.name === '#sheet').filtered, [1]);
  view.openFilterPanel('#sheet', 'Region');
  assert.deepEqual(view.render().filterPanel.selected, ['North'], 'it opens again on what passes');
});

test('Sort Z to A from the menu sorts the rows under the heading, keeps the heading, and hides by value again', () => {
  const view = filtered();
  view.applyFilter('#sheet', 'Region', ['North']);
  view.sortFilterColumn('#sheet', 'Qty', false);
  const items = [1, 2, 3, 4].map((r) => view.calc.getValue('Data', r, 0));
  assert.deepEqual(items, ['Tape', 'Paper', 'Ink', 'Pens'], 'largest Qty first');
  assert.equal(view.calc.getValue('Data', 0, 0), 'Item', 'the heading stays on top');
  // North is Ink and Pens, now on rows 4 and 5: those show, the rest hide.
  assert.deepEqual([1, 2, 3, 4].map((r) => view.geo.rowHeight(r) > 0), [false, false, true, true]);
  assert.match(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /<row[^>]*r="2"[^>]*hidden="1"/);

  view.undo();
  assert.deepEqual([1, 2, 3, 4].map((r) => view.calc.getValue('Data', r, 0)), ['Ink', 'Paper', 'Pens', 'Tape'], 'one undo puts the order back');
  assert.deepEqual([1, 2, 3, 4].map((r) => view.geo.rowHeight(r) > 0), [true, false, true, false], 'and the rows hidden with it');
});

test('a table\'s menu sorts the table\'s rows under its heading', () => {
  const view = SheetView.open(book());
  view.select(0, 0);
  view.select(4, 2, { extend: true });
  view.formatAsTable({ style: 'TableStyleMedium2' });
  const name = view.render().tables.find((t) => t.name !== '#sheet').name;
  view.sortFilterColumn(name, 'Item', false);
  assert.deepEqual([1, 2, 3, 4].map((r) => view.calc.getValue('Data', r, 0)), ['Tape', 'Pens', 'Paper', 'Ink']);
  assert.throws(() => view.sortFilterColumn(name, 'Ghost'), /no column/);
});
