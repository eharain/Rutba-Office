// A query's column types: given by Change Type, kept through every step that
// keeps the column, under its name now, and a column of dates loaded showing
// dates. The types went missing through append, merge, group, rename and the
// rest, and a date column loaded as bare day numbers.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { runSteps, applyStep } from '@rutba/sheet-view/queries';

const ORDERS = {
  columns: ['Id', 'Region', 'Placed', 'Units'],
  rows: [[1, 'East', '2026-01-05', '3'], [2, 'West', '2026-02-10', '5'], [3, 'East', '2026-03-15', '2']],
};
const typedOrders = () => runSteps(ORDERS, [{ kind: 'changeType', column: 'Placed', type: 'date' }, { kind: 'changeType', column: 'Units', type: 'integer' }]);

test('a type follows its column through rename, remove, keep, split and index', () => {
  const t = typedOrders();
  assert.deepEqual(t.types, { Placed: 'date', Units: 'integer' });
  const renamed = applyStep(t, { kind: 'renameColumn', from: 'Placed', to: 'Ordered' });
  assert.deepEqual(renamed.types, { Ordered: 'date', Units: 'integer' }, 'under its new name');
  assert.deepEqual(applyStep(t, { kind: 'removeColumns', columns: ['Units'] }).types, { Placed: 'date' });
  assert.deepEqual(applyStep(t, { kind: 'keepColumns', columns: ['Region', 'Placed'] }).types, { Placed: 'date' });
  assert.equal(applyStep(t, { kind: 'keepColumns', columns: ['Region'] }).types, undefined, 'none left, none given');
  assert.deepEqual(applyStep(t, { kind: 'splitColumn', column: 'Region', delimiter: 'a' }).types, { Placed: 'date', Units: 'integer' });
  assert.deepEqual(applyStep(t, { kind: 'addIndex' }).types, { Placed: 'date', Units: 'integer' });
  assert.deepEqual(applyStep(t, { kind: 'filterRows', column: 'Region', op: 'equals', value: 'East' }).types, { Placed: 'date', Units: 'integer' });
});

test('a type follows its column through append, merge and group', () => {
  const t = typedOrders();
  const more = { columns: ['Id', 'Region', 'Placed', 'Note'], rows: [[4, 'North', 46100, 'rush']] };
  const appended = applyStep(t, { kind: 'appendQuery', with: 'more' }, { table: () => more });
  assert.deepEqual(appended.types, { Placed: 'date', Units: 'integer' }, 'kept where the other table gives no other type');
  const clash = applyStep(t, { kind: 'appendQuery', with: 'more' }, { table: () => ({ ...more, types: { Placed: 'text', Note: 'text' } }) });
  assert.deepEqual(clash.types, { Units: 'integer', Note: 'text' }, 'two types for one column: neither; a column of its own keeps its own');

  const people = { columns: ['Region', 'Manager', 'Since'], rows: [['East', 'Kim', 45000], ['West', 'Lee', 45500]], types: { Since: 'date' } };
  const merged = applyStep(t, { kind: 'mergeQueries', with: 'people', on: 'Region', withOn: 'Region' }, { table: () => people });
  assert.deepEqual(merged.types, { Placed: 'date', Units: 'integer', Since: 'date' }, 'the other table\'s types come with its columns');
  const anti = applyStep(t, { kind: 'mergeQueries', with: 'people', on: 'Region', withOn: 'Region', how: 'leftAnti' }, { table: () => people });
  assert.deepEqual(anti.types, { Placed: 'date', Units: 'integer' });

  const grouped = applyStep(t, { kind: 'groupBy', columns: ['Placed'], aggregations: [{ fn: 'count', name: 'Orders' }, { fn: 'sum', column: 'Units', name: 'Total' }] });
  assert.deepEqual(grouped.types, { Placed: 'date', Orders: 'integer' }, 'a key keeps its type; a count is a whole number');
});

test('a column of dates loads showing dates, and one undo takes the load back', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Data', rows: [ORDERS.columns, ...ORDERS.rows] }] }));
  view.select(1, 0);
  view.addQuery({ name: 'Orders', source: { kind: 'range', sheet: 'Data', ref: 'A1:D4' }, steps: [{ kind: 'changeType', column: 'Placed', type: 'date' }, { kind: 'renameColumn', from: 'Placed', to: 'Ordered' }] });
  assert.equal(view.activeSheet, 'Orders');
  const shown = view.displayValue(1, 2).text;
  assert.doesNotMatch(shown, /^\d{5}$/, `a date, not a day number: ${shown}`);
  assert.match(shown, /2026|26/, shown);
  assert.equal(view.displayValue(1, 0).text, '1', 'the other columns as they were');
  const again = new SheetView(view.save());
  again.selectSheet('Orders');
  assert.equal(again.displayValue(1, 2).text, shown, 'kept in the file');
  view.undo();
  assert.equal(view.sheetNames().includes('Orders'), false, 'one undo step, the dates\' format with it');
  view.redo();
  view.selectSheet('Orders');
  assert.equal(view.displayValue(1, 2).text, shown, 'and back again');
});
