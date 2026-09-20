// Format as Table, written as Excel keeps a table.
//
// The table part carries the range, the columns, the autofilter and the
// style asked for by name; the sheet's rels point at it and its
// `<tableParts>` names it. The header row is the range's first row, and a
// header cell with nothing in it is given the column's name so the file and
// the sheet agree. The window paints the header and the banding at once, in
// its own palette; an undo takes the table away again.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const BOOK = () => buildXlsx({ sheets: [{ name: 'Sales', rows: [['Item', 'Qty', 'Price'], ['Pens', 12, 1.5], ['Ink', 3, 9], ['Clips', 40, 0.2]] }] });

const fillOf = (frame, ref) => frame.cells.find((c) => c.ref === ref)?.style?.fill?.colour ?? null;
const boldOf = (frame, ref) => frame.cells.find((c) => c.ref === ref)?.style?.font?.bold ?? false;

test('from one cell, the block of data round it becomes a table: the part, the rels, tableParts, the paint', () => {
  const view = new SheetView(BOOK());
  view.select(1, 1);
  view.formatAsTable({ style: 'TableStyleMedium2', stripes: true });

  const frame = view.render();
  assert.ok(fillOf(frame, 'A1') && boldOf(frame, 'A1'), 'the header row is painted and bold at once');
  assert.ok(fillOf(frame, 'C1'), 'across the whole header row');
  assert.ok(fillOf(frame, 'A3'), 'the second data row is banded');
  assert.equal(fillOf(frame, 'A2'), null, 'the first data row is not');
  assert.deepEqual(view.sheetTables().map((t) => [t.top, t.left, t.bottom, t.right]), [[0, 0, 3, 2]], 'A1:C4');

  const saved = SheetView.open(view.save());
  const sheetPart = saved.workbook.partNameFor('Sales');
  const rel = saved.pkg.rels(sheetPart).find((r) => r.Type.endsWith('/table'));
  assert.ok(rel, 'the sheet has a table relationship');
  assert.equal(rel.Target, '../tables/table1.xml');
  const xml = saved.pkg.text(sheetPart);
  assert.match(xml, new RegExp('<tableParts count="1"><tablePart r:id="' + rel.Id + '"/></tableParts></worksheet>$'), 'tableParts closes the sheet, before nothing but the root');
  assert.match(xml, /<worksheet[^>]*\sxmlns:r=/, 'the r prefix is declared');
  const table = saved.pkg.text('xl/tables/table1.xml');
  assert.match(table, /<table xmlns="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main" id="1" name="Table1" displayName="Table1" ref="A1:C4" totalsRowShown="0">/);
  assert.match(table, /<autoFilter ref="A1:C4"\/>/);
  assert.match(table, /<tableColumns count="3"><tableColumn id="1" name="Item"\/><tableColumn id="2" name="Qty"\/><tableColumn id="3" name="Price"\/><\/tableColumns>/);
  assert.match(table, /<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"\/>/);
  assert.equal(saved.pkg.contentTypeOf('xl/tables/table1.xml'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml');
  const read = saved.workbook.tables();
  assert.equal(read.length, 1);
  assert.equal(read[0].ref, 'A1:C4');
  assert.equal(read[0].styleName, 'TableStyleMedium2');
  assert.equal(read[0].showRowStripes, true);
  assert.deepEqual(read[0].columns, ['Item', 'Qty', 'Price']);
  assert.ok(fillOf(saved.render(), 'A1'), 'the reopened file paints the header');
});

test('a selection is the table; empty and repeated header cells get names, written into the cells too', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: [['Qty', '', 'Qty'], [1, 2, 3], [4, 5, 6]] }] }));
  view.select(0, 0);
  view.select(2, 2, { extend: true });
  view.formatAsTable({ style: 'TableStyleLight9', stripes: false });
  const t = view.workbook.tables()[0];
  assert.equal(t.ref, 'A1:C3');
  assert.deepEqual(t.columns, ['Qty', 'Column2', 'Qty2']);
  assert.equal(t.showRowStripes, false);
  const frame = view.render();
  assert.equal(frame.cells.find((c) => c.ref === 'B1').text, 'Column2', 'the empty header cell now says its name');
  assert.equal(frame.cells.find((c) => c.ref === 'C1').text, 'Qty2', 'and the repeated one its unique name');
  assert.equal(fillOf(frame, 'A3'), null, 'no banding when none was asked for');
});

test('a second table takes the next id and name, a clash is refused, and an undo takes a table away', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: [['A', 'B'], [1, 2], [], ['C', 'D'], [3, 4]] }] }));
  view.select(0, 0);
  view.formatAsTable();
  view.select(3, 0);
  view.formatAsTable();
  const both = view.workbook.tables();
  assert.deepEqual(both.map((t) => [t.id, t.name, t.ref]), [[1, 'Table1', 'A1:B2'], [2, 'Table2', 'A4:B5']]);
  // The sheet part is serialised on save; the package's copy is the file's until then.
  const written = SheetView.open(view.save());
  assert.match(written.pkg.text(written.workbook.partNameFor('Sales')), /<tableParts count="2"><tablePart r:id="[^"]+"\/><tablePart r:id="[^"]+"\/><\/tableParts>/);

  view.select(1, 0);
  assert.throws(() => view.formatAsTable(), /already has a table/);

  view.undo();
  assert.deepEqual(view.workbook.tables().map((t) => t.name), ['Table1'], 'the second table is gone');
  assert.equal(fillOf(view.render(), 'A4'), null, 'and its paint with it');
  assert.ok(!view.pkg.has('xl/tables/table2.xml'), 'its part too');
  view.undo();
  assert.equal(view.workbook.tables().length, 0);
  assert.equal(fillOf(view.render(), 'A1'), null);
});
