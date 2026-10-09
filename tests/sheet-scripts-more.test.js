// Scripts: what the API offered and the workbook refused (a column's width, a
// row's height, AutoFit), merging, strikethrough and vertical alignment —
// and Record Actions writing lines for widths, merges, clearing and sheets.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { runScript } from '@rutba/sheet-view/scripts';
import { recordLines } from '../apps/desktop/renderer/apps/sheets/record-lines.js';

const book = () => new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: [['Item', 'Qty', 'Price'], ['A much longer item name', 4, 2.5], ['Paper', 6, 1.25]] }] }));
const run = (view, body) => {
  const out = runScript(`function main(workbook) { const sheet = workbook.getActiveWorksheet(); ${body} }`, view.scriptSnapshot());
  assert.equal(out.error, null, out.error);
  view.applyScriptEdits(out.edits);
  return out;
};

test('a column\'s width and a row\'s height in points, and AutoFit, as a script sets them', () => {
  const view = book();
  const before = view.geo.colWidth(1);
  run(view, `sheet.getRange('B1').getFormat().setColumnWidth(150); sheet.getRange('A3').getFormat().setRowHeight(30); sheet.getRange('A1').getFormat().autofitColumns();`);
  assert.equal(view.geo.colWidth(1), 200, '150 points is 200 pixels');
  assert.equal(view.geo.rowHeight(2), 40);
  assert.ok(view.geo.colWidth(0) > 150, `AutoFit widens A to its longest text: ${view.geo.colWidth(0)}`);
  assert.notEqual(before, 200);
  const again = new SheetView(view.serialize());
  assert.equal(again.geo.colWidth(1), 200, 'kept in the file');
  view.undo();
  assert.equal(view.geo.colWidth(1), before, 'one undo');
});

test('merge and unmerge, strikethrough and vertical alignment', () => {
  const view = book();
  run(view, `sheet.getRange('A1:C1').merge(); sheet.getRange('A2').getFormat().getFont().setStrikethrough(true); sheet.getRange('A2').getFormat().setVerticalAlignment('Top');`);
  assert.deepEqual((view.merges.get('Sales') || []).map((m) => m.ref), ['A1:C1']);
  assert.equal(view.displayValue(0, 1).text, '', 'the cells but the first emptied');
  view.select(1, 0);
  assert.equal(view.formatState().strike, true);
  assert.equal(view.formatState().valign, 'top');
  run(view, `sheet.getRange('B1').unmerge();`);
  assert.deepEqual(view.merges.get('Sales') || [], []);
  const across = book();
  run(across, `sheet.getRange('A2:C3').merge(true);`);
  assert.deepEqual((across.merges.get('Sales') || []).map((m) => m.ref).sort(), ['A2:C2', 'A3:C3'], 'across: a merge for each row');
  const out = runScript(`function main(workbook) { const s = workbook.getActiveWorksheet(); s.getRange('A2:C3').merge(); s.getRange('B2:D2').merge(); }`, book().scriptSnapshot());
  assert.throws(() => book().applyScriptEdits(out.edits), /across part of a merged region/);
});

test('Record Actions writes lines for widths, heights, merges, clearing, strikethrough and sheets', () => {
  const model = { selection: { top: 0, left: 0, bottom: 0, right: 2, active: { row: 0, col: 0 } } };
  const lines = recordLines([
    { op: 'colWidth', col: 1, width: 200 },
    { op: 'rowHeight', row: 4, height: 40 },
    { op: 'merge' },
    { op: 'setFormat', delta: { strike: true, valign: 'center', wrap: true, fontName: 'Georgia' } },
    { op: 'clear' },
    { op: 'addSheet', name: 'Totals' },
    { op: 'renameSheet', from: 'Totals', to: 'Summary' },
  ], model);
  assert.deepEqual(lines, [
    '  sheet.getRange("B1").getFormat().setColumnWidth(150);',
    '  sheet.getRange("A5").getFormat().setRowHeight(30);',
    '  sheet.getRange("A1:C1").merge();',
    '  sheet.getRange("A1:C1").getFormat().getFont().setStrikethrough(true);',
    '  sheet.getRange("A1:C1").getFormat().getFont().setName("Georgia");',
    '  sheet.getRange("A1:C1").getFormat().setWrapText(true);',
    '  sheet.getRange("A1:C1").getFormat().setVerticalAlignment("center");',
    '  sheet.getRange("A1:C1").clear();',
    '  workbook.addWorksheet("Totals");',
    '  workbook.getWorksheet("Totals").setName("Summary");',
  ]);
  // And what it records runs.
  const view = book();
  run(view, lines.join('\n'));
  assert.deepEqual(view.sheetNames(), ['Sales', 'Summary']);
});
