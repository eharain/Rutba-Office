/**
 * Data → Advanced: a list filtered by a criteria range, in place or copied
 * elsewhere, with Excel's rules — rows OR'ed, columns AND'ed, text that
 * begins a field, comparisons and wildcards — and the names Excel keeps.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { condition, parseRange } from '@rutba/sheet-view/advanced-filter';

const LIST = [
  ['Region', 'Rep', 'Units', 'Sales'],
  ['East', 'Kim', 3, 300],
  ['East', 'Lee', 2, 250],
  ['North', 'Ann', 5, 400],
  ['West', 'Bo', 1, 90],
  ['Western', 'Cy', 4, 210],
  ['West', 'Bo', 1, 90],
];
/** The list in A1:D7, and a criteria block from F1 given row by row. */
const open = (criteria = []) => {
  const rows = LIST.map((r, i) => [...r, null, ...(criteria[i] ?? [])]);
  for (let i = LIST.length; i < criteria.length; i++) rows.push([null, null, null, null, null, ...criteria[i]]);
  return new SheetView(buildXlsx({ sheets: [{ name: 'Data', rows }] }));
};
const text = (view, row, col) => view.displayValue(row, col).text;
const shown = (view) => [1, 2, 3, 4, 5, 6].filter((r) => !view.geo.hiddenRows.has(r)).map((r) => text(view, r, 1) + ':' + text(view, r, 3));
const partXml = (view) => {
  const part = view.workbook.partNameFor(view.activeSheet);
  return view.workbook.snapshotParts([part])[part];
};
const workbookXml = (view) => view.pkg.text(view.workbook.mainPart);

test('A criteria cell reads as Excel reads it: text begins a field, comparisons, wildcards, blanks', () => {
  assert.equal(condition(null), null, 'a blank cell is no condition');
  assert.ok(condition('West')('Western'), 'bare text: begins with');
  assert.ok(!condition('West')('Northwest'));
  assert.ok(condition('west')('WEST'), 'letter case aside');
  assert.ok(condition('=West')('west') && !condition('=West')('Western'), '= is the whole value');
  assert.ok(condition('<>West')('East') && !condition('<>West')('WEST'));
  assert.ok(condition('=W*n')('Western') && !condition('=W*n')('West'), '* any run');
  assert.ok(condition('B?')('Bo') && condition('=B?')('Bo') && !condition('=B?')('Bob'), '? one character');
  assert.ok(condition('=~*')('*') && !condition('=~*')('x'), '~ takes the wildcard literally');
  assert.ok(condition('>250')(300) && !condition('>250')(250) && !condition('>250')('300 units'), 'numbers against numbers only');
  assert.ok(condition('<=90')(90) && condition('>=90')(90) && condition('<91')(90));
  assert.ok(condition('>M')('North') && !condition('>M')('East'), 'text orders as text');
  assert.ok(condition(5)(5) && !condition(5)(50), 'a number typed is that number');
  assert.ok(condition('=')('') && !condition('=')('x'), '= alone: blank');
  assert.ok(condition('<>')('x') && !condition('<>')(''), '<> alone: filled');
  assert.deepEqual(parseRange("'My Data'!$F$1:$G$3"), { top: 0, bottom: 2, left: 5, right: 6 });
});

test('In place, two criteria rows are OR\'ed: the rows that pass stay, the rest hide, and the sheet says so', () => {
  const view = open([['Region'], ['East'], ['North']]);
  const done = view.advancedFilter({ list: 'A1:D7', criteria: 'F1:F3' });
  assert.deepEqual(done, { action: 'filter', matched: 3, total: 6, range: 'A1:D7' });
  assert.deepEqual(shown(view), ['Kim:300', 'Lee:250', 'Ann:400']);
  assert.equal(view.geo.rowHeight(4), 0);
  const xml = partXml(view);
  assert.match(xml, /<row r="5"[^>]*hidden="1"/);
  assert.match(xml, /<sheetPr filterMode="1"/);
  const names = workbookXml(view);
  assert.match(names, /<definedName name="_xlnm\._FilterDatabase" localSheetId="0" hidden="1">Data!\$A\$1:\$D\$7<\/definedName>/);
  assert.match(names, /<definedName name="_xlnm\.Criteria" localSheetId="0">Data!\$F\$1:\$F\$3<\/definedName>/);

  view.undo();
  assert.equal(shown(view).length, 6, 'one undo step shows every row');
  assert.doesNotMatch(partXml(view), /filterMode|hidden=/);
});

test('Conditions on one criteria row are AND\'ed, and a heading can repeat for a range', () => {
  const view = open([['Region', 'Sales', 'Sales'], ['West', '>50', '<100']]);
  view.advancedFilter({ list: 'A1:D7', criteria: 'F1:H2' });
  assert.deepEqual(shown(view), ['Bo:90', 'Bo:90'], 'West* between 50 and 100 — Western\'s 210 is out');
  view.advancedFilter({ list: 'A1:D7', criteria: 'F1:H2', unique: true });
  assert.deepEqual(shown(view), ['Bo:90'], 'unique records only: the repeated row hides too');
});

test('Copy to another location writes the headings and the rows that pass, unique ones only, and names the extract', () => {
  const view = open([['Units'], ['>1']]);
  view.advancedFilter({ list: 'A1:D7', criteria: 'F1:F2', action: 'copy', copyTo: 'H1', unique: true });
  const out = [0, 1, 2, 3, 4, 5].map((r) => [7, 8, 9, 10].map((c) => text(view, r, c)).join('|'));
  assert.deepEqual(out, ['Region|Rep|Units|Sales', 'East|Kim|3|300', 'East|Lee|2|250', 'North|Ann|5|400', 'Western|Cy|4|210', '|||']);
  assert.equal(shown(view).length, 6, 'the list itself is untouched');
  assert.match(workbookXml(view), /<definedName name="_xlnm\.Extract" localSheetId="0">Data!\$H\$1:\$K\$1<\/definedName>/);

  // Headings typed at the destination pick the columns and their order.
  view.setCell(0, 12, 'Sales');
  view.setCell(0, 13, 'rep');
  view.advancedFilter({ list: 'A1:D7', criteria: 'F1:F2', action: 'copy', copyTo: 'M1:N1' });
  assert.deepEqual([0, 1, 2].map((r) => text(view, r, 12) + '|' + text(view, r, 13)), ['Sales|rep', '300|Kim', '250|Lee']);

  // A shorter extract clears what a longer one left below it.
  view.advancedFilter({ list: 'A1:D7', criteria: 'F1:F2', action: 'copy', copyTo: 'H1' });
  view.setCell(1, 5, '>4');
  view.advancedFilter({ list: 'A1:D7', criteria: 'F1:F2', action: 'copy', copyTo: 'H1' });
  assert.deepEqual([1, 2, 3].map((r) => text(view, r, 8)), ['Ann', '', '']);
  view.undo();
  assert.equal(text(view, 2, 8), 'Lee', 'one undo step puts the earlier extract back');
  assert.throws(() => view.advancedFilter({ list: 'A1:D7', criteria: 'F1:F2', action: 'copy', copyTo: 'B5' }), /land on the list/);
});

test('A formula under a heading that names no field is a condition of its own, read for every row', () => {
  const view = open([['Above average'], ['=D2>AVERAGE($D$2:$D$7)']]);
  view.advancedFilter({ list: 'A1:D7', criteria: 'F1:F2' });
  assert.deepEqual(shown(view), ['Kim:300', 'Lee:250', 'Ann:400'], 'above 223.33; the 210 is not');
});

test('Clear shows the rows again; a blank criteria row lets every row through; a heading must name a field', () => {
  const view = open([['Region'], ['North']]);
  view.advancedFilter({ list: 'A1:D7', criteria: 'F1:F2' });
  assert.equal(shown(view).length, 1);
  assert.deepEqual(view.clearAdvancedFilter(), { shown: 5 });
  assert.equal(shown(view).length, 6);
  assert.doesNotMatch(partXml(view), /filterMode/);
  assert.throws(() => view.clearAdvancedFilter(), /No advanced filter/);
  assert.equal(view.listFields().filter.criteria, 'F1:F2', 'the dialog is offered the ranges again');

  const blank = open([['Region'], [null], ['North']]);
  blank.advancedFilter({ list: 'A1:D7', criteria: 'F1:F3' });
  assert.equal(shown(blank).length, 6);
  const wrong = open([['Territory'], ['East']]);
  assert.throws(() => wrong.advancedFilter({ list: 'A1:D7', criteria: 'F1:F2' }), /names no field/);
});
