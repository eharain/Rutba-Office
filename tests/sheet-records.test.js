// Data → From Picture → Receipt or Invoice puts its rows in with putRecords:
// typed values at the active cell, each with its number format, a date as a
// date and amounts as numbers, an invoice number kept as the text it is, as
// one step that Undo takes back whole and Redo puts back.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { recordCells } from '@rutba/ocr/record';

const shown = (v, r, c) => v.displayValue(r, c)?.text ?? '';
const BOOK = () => buildXlsx({ sheets: [{ name: 'Expenses', rows: [['Date', 'Supplier', 'Number', 'Subtotal', 'Tax', 'Total', 'Currency']] }] });

test('putRecords: receipts as rows from the selected row, dates as dates, amounts as numbers, numbers as text', () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  const rows = [
    recordCells({ date: '2025-09-04', supplier: 'Harbour Print Co.', number: '342457', subtotal: '181.08', tax: '9.05', total: '190.13', currency: 'USD' }),
    recordCells({ date: '2026-10-14', supplier: 'Northgate Supplies Ltd', number: '00042', subtotal: '65.00', tax: '13.00', total: '78.00', currency: 'GBP' }),
  ];
  const narrow = view.geo.colWidth(1);
  view.putRecords(rows);
  assert.ok(view.geo.colWidth(1) > narrow + 60, 'the supplier column widened to its names');
  assert.ok(view.geo.colWidth(0) >= 80, 'the date column to its dates');
  assert.equal(view.calc.getInput('Expenses', 1, 0), 45904, 'the date is its serial number');
  assert.equal(shown(view, 1, 0), '2025-09-04', 'and shows as a date');
  assert.equal(view.styleFor(1, 0)?.numFmt, 'yyyy-mm-dd');
  assert.equal(view.calc.getInput('Expenses', 1, 5), 190.13, 'the total a number');
  assert.equal(shown(view, 1, 5), '190.13');
  assert.equal(view.styleFor(2, 3)?.numFmt, '#,##0.00');
  assert.equal(shown(view, 2, 2), '00042', 'the invoice number keeps its zeros');
  assert.equal(shown(view, 1, 2), '342457');
  assert.equal(typeof view.calc.getInput('Expenses', 1, 2), 'string', 'and is text');
  assert.equal(shown(view, 2, 1), 'Northgate Supplies Ltd');
  assert.equal(shown(view, 2, 6), 'GBP');
  assert.equal(view.selection.toString(), 'A2:G3', 'the rows put in are selected');

  view.undo();
  assert.equal(shown(view, 1, 0), '', 'one Undo takes every row back');
  assert.equal(shown(view, 2, 5), '');
  assert.equal(view.styleFor(1, 0)?.numFmt ?? 'General', 'General', 'and the formats with them');
  assert.equal(shown(view, 0, 5), 'Total', 'the headings stay');
  assert.equal(view.geo.colWidth(1), narrow, 'and the column its width');
  view.redo();
  assert.equal(shown(view, 1, 0), '2025-09-04');
  assert.equal(view.calc.getInput('Expenses', 2, 5), 78);

  // Saved and opened again, the date is still a date and the amounts numbers.
  const again = new SheetView(view.save());
  assert.equal(shown(again, 1, 0), '2025-09-04');
  assert.equal(again.styleFor(1, 0)?.numFmt, 'yyyy-mm-dd');
  assert.equal(again.calc.getInput('Expenses', 2, 3), 65);
});

test('putRecords: an empty value leaves its cell empty; text that is not a figure goes in as text; nothing is no step', () => {
  const view = new SheetView(BOOK());
  view.select(4, 1);
  view.putRecords([recordCells({ date: 'unclear', total: '¥15,840', currency: 'JPY' })]);
  assert.equal(shown(view, 4, 1), 'unclear');
  assert.equal(shown(view, 4, 2), '');
  assert.equal(view.calc.getInput('Expenses', 4, 6), 15840);
  assert.equal(view.styleFor(4, 6)?.numFmt, '#,##0');
  const before = view.history.past.length;
  view.putRecords([]);
  assert.equal(view.history.past.length, before, 'no rows, no step');
  // A column already wide enough is left as it is.
  view.setColWidth(3, 300);
  view.select(6, 0);
  view.putRecords([recordCells({ date: '2026-10-14', supplier: 'X', total: '1.00' })]);
  assert.equal(view.geo.colWidth(3), 300, 'never narrowed');
});
