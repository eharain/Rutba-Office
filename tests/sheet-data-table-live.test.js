/**
 * A Data Table that stays live. Before 2.28 the grid was filled once with
 * plain values: a cost or a price changed in the model and the table kept
 * the old answers until it was run again. Excel's {=TABLE()} follows the
 * model, and now so does this one.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { Workbook } from '@rutba/ooxml/workbook';
import { SheetView } from '@rutba/sheet-view';

// A loan: B1 the rate, B2 the amount, B3 the interest. Rates to try down A6:A8, the formula named in B5.
const oneWay = () => {
  const view = SheetView.open(buildXlsx({
    sheets: [{
      name: 'Loan',
      rows: [['Rate', 0.05], ['Amount', 1000], ['Interest', '=B1*B2'], [''], ['', '=B3'], [0.01, ''], [0.02, ''], [0.03, '']],
    }],
  }));
  assert.deepEqual(view.dataTable({ range: 'A5:B8', colInput: 'B1' }), { cells: 3, live: true });
  return view;
};
const column = (view, col, from, to, sheet = 'Loan') => {
  const out = [];
  for (let r = from; r <= to; r++) out.push(view.calc.getValue(sheet, r, col));
  return out;
};
const marker = (bytes, sheet = 'Loan') => {
  const wb = Workbook.open(bytes);
  return wb.dataTableMarkers(sheet);
};

test('the grid follows the model: the amount changes and every answer with it', () => {
  const view = oneWay();
  assert.deepEqual(column(view, 1, 5, 7), [10, 20, 30]);
  view.setCell(1, 1, 2000);
  assert.deepEqual(column(view, 1, 5, 7), [20, 40, 60], 'no second run of the table');
  assert.equal(view.calc.getValue('Loan', 0, 1), 0.05, 'the rate in the model is as it was');
  assert.equal(view.calc.getValue('Loan', 2, 1), 100, 'and so is the model\'s own answer');
  // A formula the model's formula reads, changed.
  view.setCell(2, 1, '=B1*B2+5');
  assert.deepEqual(column(view, 1, 5, 7), [25, 45, 65]);
});

test('a candidate changed on the edge of the table is worked out too', () => {
  const view = oneWay();
  view.setCell(6, 0, 0.1);
  assert.deepEqual(column(view, 1, 5, 7), [10, 100, 30]);
});

test('undo and redo of the change put the grid back and forward with the model', () => {
  const view = oneWay();
  view.setCell(1, 1, 2000);
  view.undo();
  assert.deepEqual(column(view, 1, 5, 7), [10, 20, 30]);
  view.redo();
  assert.deepEqual(column(view, 1, 5, 7), [20, 40, 60]);
  // Undo back past the table itself: the grid is empty and nothing fills it again.
  view.undo();
  view.undo();
  assert.deepEqual(column(view, 1, 5, 7), ['', '', '']);
  view.setCell(1, 1, 3000);
  assert.deepEqual(column(view, 1, 5, 7), ['', '', '']);
});

test('the file is saved with the answers the grid shows and with its marker, and is live when opened again', () => {
  const view = oneWay();
  view.setCell(1, 1, 2000);
  const bytes = view.save();
  assert.deepEqual(marker(bytes).map((m) => [m.master, m.ref, m.colInput, m.rowInput]), [['B6', 'B6:B8', 'B1', null]]);
  const wb = Workbook.open(bytes);
  assert.deepEqual(['B6', 'B7', 'B8'].map((r) => wb.getCell('Loan', r)), [20, 40, 60], 'the saved answers are the new ones');
  const again = SheetView.open(bytes);
  assert.deepEqual(column(again, 1, 5, 7), [20, 40, 60]);
  again.setCell(1, 1, 500);
  assert.deepEqual(column(again, 1, 5, 7), [5, 10, 15], 'a table read from a file follows the model as well');
});

test('a two-way table follows the model', () => {
  const view = SheetView.open(buildXlsx({
    sheets: [{
      name: 'Loan',
      rows: [['Rate', 0.05], ['Amount', 1000], ['Fee', 0], ['Interest', '=B1*B2+B3'], [''], ['=B4', 1000, 2000], [0.01, '', ''], [0.02, '', '']],
    }],
  }));
  assert.deepEqual(view.dataTable({ range: 'A6:C8', rowInput: 'B2', colInput: 'B1' }), { cells: 4, live: true });
  assert.deepEqual([column(view, 1, 6, 7), column(view, 2, 6, 7)], [[10, 20], [20, 40]]);
  view.setCell(2, 1, 7);
  assert.deepEqual([column(view, 1, 6, 7), column(view, 2, 6, 7)], [[17, 27], [27, 47]]);
  assert.deepEqual([view.calc.getValue('Loan', 0, 1), view.calc.getValue('Loan', 1, 1)], [0.05, 1000], 'both inputs are as they were');
  const again = SheetView.open(view.save());
  again.setCell(2, 1, 0);
  assert.deepEqual([column(again, 1, 6, 7), column(again, 2, 6, 7)], [[10, 20], [20, 40]]);
});

test('typed over at its first cell, or cleared, a table is a plain grid from then on', () => {
  const view = oneWay();
  view.setCell(5, 1, 99);
  view.setCell(1, 1, 2000);
  assert.deepEqual(column(view, 1, 5, 7), [99, 20, 30], 'nothing is worked out again');
  assert.deepEqual(marker(view.save()), [], 'and the file no longer marks it');
});

test('rows moved on its sheet: the table becomes the plain answers it shows, with no marker left pointing at the old cells', () => {
  const view = oneWay();
  view.setCell(1, 1, 2000);
  view.insertRows(0, 1);
  assert.deepEqual(column(view, 1, 6, 8), [20, 40, 60], 'the answers moved down with their rows');
  view.setCell(2, 1, 9000);
  assert.deepEqual(column(view, 1, 6, 8), [20, 40, 60], 'a plain grid does not change');
  assert.deepEqual(marker(view.save()), []);
  // Run again where it now sits, it is live again.
  assert.equal(view.dataTable({ range: 'A6:B9', colInput: 'B2' }).live, true);
  view.setCell(2, 1, 1000);
  assert.deepEqual(column(view, 1, 6, 8), [10, 20, 30]);
});

test('a table on another sheet follows a change made on this one', () => {
  const view = SheetView.open(buildXlsx({
    sheets: [
      { name: 'Loan', rows: [['Rate', 0.05], ['Amount', '=Costs!A1'], ['Interest', '=B1*B2'], [''], ['', '=B3'], [0.01, ''], [0.02, '']] },
      { name: 'Costs', rows: [[1000]] },
    ],
  }));
  view.dataTable({ range: 'A5:B7', colInput: 'B1' });
  assert.deepEqual(column(view, 1, 5, 6), [10, 20]);
  view.selectSheet('Costs');
  view.setCell(0, 0, 3000);
  assert.deepEqual(column(view, 1, 5, 6), [30, 60]);
});
