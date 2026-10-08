/**
 * Data → Get & Transform: a query's steps run on its source — columns
 * removed, renamed and typed, rows filtered, sorted and grouped — and in the
 * workbook: From Table/Range loads the result as a table on a sheet of its
 * own, the query is kept in the file, and Refresh brings in what the source
 * holds now.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { runSteps, applyStep, describeStep, parseDelimited } from '@rutba/sheet-view/queries';

const SALES = {
  columns: ['Region', 'Rep', 'Units', 'Sales'],
  rows: [['East', 'Kim', 3, 300], ['East', 'Lee', 2, 250], ['North', 'Ann', 5, 400], ['West', 'Bo', 1, 90], ['west', 'Cy', 4, 210], ['West', 'Bo', 1, 90]],
};

test('Steps run in order on the source: columns removed and renamed, rows filtered, sorted, deduplicated and grouped', () => {
  const t = runSteps(SALES, [
    { kind: 'removeDuplicates' },
    { kind: 'renameColumn', from: 'Rep', to: 'Person' },
    { kind: 'filterRows', column: 'Units', op: 'greater', value: 1 },
    { kind: 'sort', column: 'Sales', descending: true },
    { kind: 'removeColumns', columns: ['Units'] },
  ]);
  assert.deepEqual(t.columns, ['Region', 'Person', 'Sales']);
  assert.deepEqual(t.rows, [['North', 'Ann', 400], ['East', 'Kim', 300], ['East', 'Lee', 250], ['west', 'Cy', 210]]);
  const g = runSteps(SALES, [{ kind: 'groupBy', columns: ['Region'], aggregations: [{ fn: 'sum', column: 'Sales', name: 'Total' }, { fn: 'count', name: 'Orders' }] }]);
  assert.deepEqual(g.columns, ['Region', 'Total', 'Orders']);
  assert.deepEqual(g.rows, [['East', 550, 2], ['North', 400, 1], ['West', 390, 3]], 'groups ignore letter case, as Power Query\'s do by default here');
  assert.equal(runSteps(SALES, [{ kind: 'keepTopRows', count: 2 }]).rows.length, 2);
  assert.deepEqual(runSteps(SALES, [{ kind: 'addIndex', name: 'No', start: 1 }]).rows.map((r) => r[4]), [1, 2, 3, 4, 5, 6]);
  assert.equal(runSteps(SALES, [], { upTo: 0 }).rows.length, 6);
});

test('Types, text, replacing and splitting change the values a step at a time', () => {
  const t = { columns: ['Name', 'Amount', 'When'], rows: [['  ann  novak ', '1,200', '2026-10-01'], ['bo lund', 'n/a', 'not a date']] };
  const r = runSteps(t, [
    { kind: 'changeType', column: 'Amount', type: 'number' },
    { kind: 'changeType', column: 'When', type: 'date' },
    { kind: 'transformText', column: 'Name', how: 'trim' },
    { kind: 'transformText', column: 'Name', how: 'proper' },
    { kind: 'splitColumn', column: 'Name', delimiter: 'space' },
  ]);
  assert.deepEqual(r.columns, ['Name.1', 'Name.2', 'Amount', 'When']);
  assert.deepEqual(r.rows[0], ['Ann', 'Novak', 1200, 46296]);
  assert.deepEqual(r.rows[1], ['Bo', 'Lund', null, null], 'what cannot be a number or a date is blank, not an error');
  assert.deepEqual(applyStep({ columns: ['A'], rows: [['x'], ['y']] }, { kind: 'replaceValues', column: 'A', find: 'x', replace: 'z' }).rows, [['z'], ['y']]);
  assert.throws(() => applyStep(SALES, { kind: 'renameColumn', from: 'Rep', to: 'Units' }), /already a column/);
  assert.throws(() => applyStep(SALES, { kind: 'removeColumns', columns: ['Ghost'] }), /no column "Ghost"/);
  assert.equal(describeStep({ kind: 'filterRows', column: 'Units', op: 'greater', value: 1 }), 'Rows where "Units" is more than 1');
});

test('A CSV is read into rows, quotes honoured, numbers as numbers', () => {
  assert.deepEqual(parseDelimited('Name,Qty\r\n"Lund, Bo",4\r\nAnn,"1"\r\n'), [['Name', 'Qty'], ['Lund, Bo', 4], ['Ann', 1]]);
  assert.deepEqual(parseDelimited('a\tb\n1\t2'), [['a', 'b'], [1, 2]]);
});

test('From Table/Range loads a query as a table on a sheet of its own, kept in the file, and Refresh brings in the source as it is now', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Data', rows: [SALES.columns, ...SALES.rows] }] }));
  view.select(1, 0);
  const q = view.addQuery({ name: 'Big sales', source: { kind: 'range', sheet: 'Data', ref: 'A1:D7' }, steps: [{ kind: 'filterRows', column: 'Sales', op: 'greaterOrEqual', value: 250 }, { kind: 'removeColumns', columns: ['Units'] }] });
  assert.equal(q.sheet, 'Big sales');
  assert.equal(view.activeSheet, 'Big sales');
  const text = (r, c) => view.displayValue(r, c).text;
  assert.deepEqual([0, 1, 2].map((c) => text(0, c)), ['Region', 'Rep', 'Sales']);
  assert.deepEqual([1, 2, 3].map((r) => text(r, 2)), ['300', '250', '400']);
  assert.equal(view.workbook.tables().find((t) => t.sheet === 'Big sales')?.ref, 'A1:C4', 'loaded as a table');
  // Kept in the file.
  const again = new SheetView(view.save());
  assert.deepEqual(again.queries().map((x) => [x.name, x.steps.length]), [['Big sales', 2]]);
  // The source changes; Refresh brings it in.
  view.selectSheet('Data');
  view.select(4, 3);
  view.setCell(4, 3, 600);
  view.refreshQueries();
  view.selectSheet('Big sales');
  assert.deepEqual([1, 2, 3, 4].map((r) => text(r, 2)), ['300', '250', '400', '600'], 'Bo\'s sale now counts');
  assert.equal(view.workbook.tables().find((t) => t.sheet === 'Big sales')?.ref, 'A1:C5', 'the table grown to fit');
  view.undo();
  view.selectSheet('Big sales');
  assert.equal(text(4, 2), '', 'one undo step');
  view.removeQuery({ id: q.id });
  assert.deepEqual(view.queries(), []);
});

test('A CSV source is read from its file each time the query runs', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-query-'));
  const file = path.join(dir, 'stock.csv');
  fs.writeFileSync(file, 'Item,Qty\nInk,4\nPaper,6\n');
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Sheet1', rows: [[]] }] }));
  view.addQuery({ name: 'Stock', source: { kind: 'csv', path: file }, steps: [{ kind: 'promoteHeaders' }], read: (p) => fs.readFileSync(p, 'utf8') });
  assert.equal(view.displayValue(2, 0).text, 'Paper');
  fs.writeFileSync(file, 'Item,Qty\nInk,4\nPaper,6\nToner,1\n');
  view.refreshQueries({ read: (p) => fs.readFileSync(p, 'utf8') });
  assert.equal(view.displayValue(3, 0).text, 'Toner');
});
