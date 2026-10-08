// Power Query — edges the first tests leave alone: a loaded sheet renamed
// before Refresh, an Append whose other table has two columns of one
// heading, a Refresh run twice, and a Merge on a key that is not there.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { applyStep } from '@rutba/sheet-view/queries';

const open = () => new SheetView(buildXlsx({ sheets: [{ name: 'Data', rows: [['Key', 'Value'], ['a', 1], ['b', 2]] }] }));
const query = (view, ref = 'A1:B3') => view.addQuery({ name: 'Copy', source: { kind: 'range', sheet: 'Data', ref }, steps: [] });

test('Refresh after the loaded sheet was renamed fills that sheet, rather than building a second sheet and table', () => {
  const view = open();
  query(view, 'A1:B4');
  view.renameSheet('Copy', 'Mine');
  view.selectSheet('Data');
  view.setCell(3, 0, 'c');
  view.setCell(3, 1, 3);
  view.refreshQueries();
  assert.deepEqual(view.sheetNames(), ['Data', 'Mine']);
  assert.deepEqual(view.workbook.tables().map((t) => [t.sheet, t.ref]), [['Mine', 'A1:B4']], 'one table, grown to fit');
  view.selectSheet('Mine');
  assert.equal(view.displayValue(3, 0).text, 'c');
});

test('Append keeps the values of a second column that has the same heading as the first', () => {
  const t = { columns: ['Key', 'Value'], rows: [['a', 1]] };
  const r = applyStep(t, { kind: 'appendQuery', with: {} }, { table: () => ({ columns: ['Key', 'Key'], rows: [['x', 'y']] }) });
  assert.equal(r.columns.length, 3, 'the second Key is named apart');
  assert.deepEqual(r.rows[1], ['x', null, 'y'], 'neither value is lost');
});

test('Refresh twice in a row changes nothing the second time, and one undo takes back only the last', () => {
  const view = open();
  query(view);
  view.selectSheet('Data');
  view.setCell(1, 1, 10);
  view.refreshQueries();
  view.selectSheet('Copy');
  const grid = () => [0, 1, 2].map((r) => [0, 1].map((c) => view.displayValue(r, c).text));
  const once = grid();
  view.refreshQueries();
  assert.deepEqual(grid(), once);
  assert.equal(view.queries().length, 1);
  view.undo();
  view.selectSheet('Copy');
  assert.deepEqual(grid(), once, 'the second refresh was its own step and changed nothing');
  const reopened = new SheetView(view.save());
  assert.deepEqual(reopened.queries().map((q) => q.name), ['Copy']);
});

test('Merge on a column the other table lacks says which, and a blank key never matches another blank', () => {
  const mine = { columns: ['Key', 'Value'], rows: [['a', 1], [null, 2]] };
  const other = { columns: ['Id', 'Note'], rows: [['A', 'x'], [null, 'y']] };
  assert.throws(() => applyStep(mine, { kind: 'mergeQueries', on: 'Key', withOn: 'Key', with: {} }, { table: () => other }), /no column "Key"/);
  assert.throws(() => applyStep(mine, { kind: 'mergeQueries', on: 'Nope', withOn: 'Id', with: {} }, { table: () => other }), /no column "Nope"/);
  const joined = applyStep(mine, { kind: 'mergeQueries', on: 'Key', withOn: 'Id', how: 'left', with: {} }, { table: () => other });
  assert.deepEqual(joined.rows, [['a', 1, 'x'], [null, 2, null]]);
});
