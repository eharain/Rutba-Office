// A whole row's or a whole column's format, as Excel keeps it: pressed on a
// row's or a column's header, Font colour (or any format) goes on the row or
// the column itself (`<row s customFormat="1">`, `<col style>`), so the cells
// past the ones in use wear it too, a cell typed there later takes it, and
// the file says so the way Excel writes it. Before, only the cells already in
// use were coloured, and on an empty sheet a "row" was its first cell.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

const book = (rows) => buildXlsx({ sheets: [{ name: 'S', rows }] });
const colourOf = (view, row, col) => view.styleFor(row, col)?.font?.colour ?? null;
const fillOf = (view, row, col) => view.styleFor(row, col)?.fill?.colour ?? null;
const sheetXml = async (view) => OoxmlPackage.read(await view.save()).text('xl/worksheets/sheet1.xml');
const RED = '#e03131';

test('a row coloured from its header colours the cells past the ones in use, and a cell typed there later', () => {
  const view = new SheetView(book([['a', 'b', 'c'], ['d', 'e', 'f']]));
  view.selectRow(1);
  view.setFormat({ fontColour: RED });
  assert.equal(colourOf(view, 1, 0), RED, 'a cell in use');
  assert.equal(colourOf(view, 1, 2), RED);
  assert.equal(colourOf(view, 1, 9), RED, 'a cell past the used range');
  assert.equal(colourOf(view, 0, 9), null, 'the row above is untouched');
  view.setCell(1, 5, 'typed later');
  assert.equal(colourOf(view, 1, 5), RED, 'the words typed later');
  assert.match(String(view.formatState().fontColour), /E03131$/i, 'the toolbar reads it');
});

test('a column likewise, down past the last row in use', () => {
  const view = new SheetView(book([['a', 'b', 'c'], ['d', 'e', 'f']]));
  view.selectColumn(1);
  view.setFormat({ fontColour: RED });
  assert.equal(colourOf(view, 0, 1), RED);
  assert.equal(colourOf(view, 500, 1), RED);
  assert.equal(colourOf(view, 0, 2), null, 'the next column is untouched');
  view.setCell(9, 1, 'typed later');
  assert.equal(colourOf(view, 9, 1), RED);
});

test('on an empty sheet a whole row is the whole row, not its first cell', () => {
  const view = new SheetView(book([[]]));
  view.selectRow(3);
  view.setFormat({ fontColour: RED, fill: '#FFF3BF' });
  view.setCell(3, 2, 'x');
  assert.equal(colourOf(view, 3, 2), RED);
  assert.equal(fillOf(view, 3, 7), '#fff3bf', 'an empty cell of the row shows the fill');
  const drawn = (view.render().cells || []).filter((c) => c.row === 3).length;
  assert.ok(drawn > 3, `the row's empty cells are drawn with the fill (${drawn})`);
});

test('the file carries the row and column styles as Excel writes them, and a typed cell carries its own', async () => {
  const view = new SheetView(book([['a', 'b'], ['c', 'd']]));
  view.selectRow(4);
  view.setFormat({ fontColour: RED });
  view.selectColumn(3);
  view.setFormat({ bold: true });
  view.setCell(4, 6, 'later');
  const xml = await sheetXml(view);
  const row = /<row\b[^>]*\br="5"[^>]*>/.exec(xml)?.[0] || '';
  assert.match(row, /\bs="\d+"/, 'the row names its style');
  assert.match(row, /\bcustomFormat="1"/, 'and says it is the row\'s own');
  const col = /<col\b[^>]*\bmin="4"[^>]*\/>/.exec(xml)?.[0] || '';
  assert.match(col, /\bmax="4"/);
  assert.match(col, /\bstyle="\d+"/, 'the column names its style');
  assert.match(col, /\bwidth="[\d.]+"/, 'and keeps a width, or Excel draws it at nothing');
  assert.match(xml, /<c r="G5"[^>]*\bs="1"/, 'the cell typed later carries the row\'s style on itself');

  const back = new SheetView(await view.save());
  assert.equal(colourOf(back, 4, 20), RED, 'opened again, the row is still red');
  assert.equal(back.styleFor(70, 3)?.font?.bold, true, 'and the column still bold');
  back.setCell(80, 3, 'x');
  assert.equal(back.styleFor(80, 3)?.font?.bold, true, 'and a cell typed into it after opening takes it');
});

test('a row style a file already has (Excel wrote it) is drawn on the row\'s empty cells', () => {
  const view = new SheetView(book([['a'], ['b']]));
  view.select(1, 0);
  view.setFormat({ fontColour: RED });
  const idx = view._styleIndexAt('S', 1, 0);
  // The second row made the row's own, as Excel saves a formatted row.
  view.workbook.setRowStyle('S', 2, idx);
  assert.equal(colourOf(view, 2, 4), RED);
  assert.equal(colourOf(view, 3, 4), null);
});

test('cells chosen by dragging are cells, not the row, even when they run the width of the cells in use', () => {
  const view = new SheetView(book([['a', 'b', 'c'], ['d', 'e', 'f']]));
  view.select(1, 0);
  view.select(1, 2, { extend: true });
  view.setFormat({ fontColour: RED });
  view.setCell(1, 3, 'next');
  assert.equal(colourOf(view, 1, 2), RED);
  assert.equal(colourOf(view, 1, 3), null, 'D2 was not chosen');
  // A row header pressed, then the selection moved on: forgotten.
  view.selectRow(0);
  view.select(0, 1);
  view.setFormat({ fontColour: RED });
  assert.equal(colourOf(view, 0, 7), null);
});

test('undo takes the row\'s look away again, and redo puts it back', () => {
  const view = new SheetView(book([['a', 'b'], ['c', 'd']]));
  view.selectRow(1);
  view.setFormat({ fontColour: RED });
  assert.equal(colourOf(view, 1, 8), RED);
  view.undo();
  assert.equal(colourOf(view, 1, 8), null);
  assert.equal(colourOf(view, 1, 0), null);
  view.redo();
  assert.equal(colourOf(view, 1, 8), RED);
});

test('a column\'s look under a row\'s: the row wins on an empty cell, as in Excel, and a formatted row keeps the column\'s colour on a cell in use', () => {
  const view = new SheetView(book([['a', 'b'], ['c', 'd']]));
  view.selectColumn(1);
  view.setFormat({ fontColour: RED });
  view.selectRow(1);
  view.setFormat({ bold: true });
  assert.equal(colourOf(view, 1, 1), RED, 'B2 was in use: bold added to its red');
  assert.equal(view.styleFor(1, 1)?.font?.bold, true);
  assert.equal(view.styleFor(1, 6)?.font?.bold, true, 'the row past the cells in use');
});
