/**
 * View → Custom Views: the way a workbook looks, kept under a name and shown
 * again — written as Excel writes it: `<customWorkbookViews>` in
 * workbook.xml, a `<customSheetView>` with the same GUID in every sheet, and
 * Excel's hidden `Z_<GUID>_.wvu.*` names for the rows, columns, print area
 * and titles a view keeps. Excel greys the command in a workbook with a
 * table, and so does this.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { writePageSetup, readPageSetup } from '@rutba/sheet-view/print';

const grid = (rows, cols) => Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => (r === 0 ? 'H' + c : r * 10 + c)));
const open = () => new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: grid(40, 6) }, { name: 'Other data', rows: grid(5, 3) }] }), { viewportWidth: 1200, viewportHeight: 700 });
const xmlOf = (view, part) => OoxmlPackage.read(view.save()).text(part);

/** Rows 3 to 5 and column C hidden through the outline, as a person would with Group and Hide Detail. */
function hideSome(view) {
  const { part } = view.workbook._sheetPart('Sales');
  for (const r of [2, 3, 4]) part.setRowOutline(r, { hidden: true });
  part.setColOutline(2, { hidden: true });
  view._rebuildDerivedState();
}

test('Add writes customWorkbookViews and a customSheetView in every sheet, with Excel\'s hidden names', () => {
  const view = open();
  hideSome(view);
  writePageSetup(view, 'Sales', { ...readPageSetup(view, 'Sales'), orientation: 'landscape', area: 'A1:F30', repeatRows: 1 });
  view.select(6, 1);
  view.select(9, 3, { extend: true });
  view.addCustomView({ name: 'Quarter review', zoom: 1.25, windowWidth: 1400, windowHeight: 860 });
  const wb = xmlOf(view, 'xl/workbook.xml');
  const m = /<customWorkbookViews><customWorkbookView name="Quarter review" guid="(\{[0-9A-F-]{36}\})"([^>]*)\/><\/customWorkbookViews>/.exec(wb);
  assert.ok(m, 'the workbook view, with a GUID in braces');
  const guid = m[1];
  assert.match(m[2], /activeSheetId="1"/);
  assert.match(m[2], /windowWidth="1400" windowHeight="860"/);
  assert.doesNotMatch(m[2], /includePrintSettings|includeHiddenRowCol/, 'both kept: the defaults are not written');
  const key = 'Z_' + guid.slice(1, -1).replace(/-/g, '_') + '_.wvu.';
  assert.match(wb, new RegExp('<definedName name="' + key.replace(/\./g, '\\.') + 'Rows" localSheetId="0" hidden="1">Sales!\\$3:\\$5</definedName>'));
  assert.match(wb, new RegExp('<definedName name="' + key.replace(/\./g, '\\.') + 'Cols" localSheetId="0" hidden="1">Sales!\\$C:\\$C</definedName>'));
  assert.match(wb, new RegExp(key.replace(/\./g, '\\.') + 'PrintArea" localSheetId="0" hidden="1">Sales!\\$A\\$1:\\$F\\$30<'));
  assert.match(wb, new RegExp(key.replace(/\./g, '\\.') + 'PrintTitles" localSheetId="0" hidden="1">Sales!\\$1:\\$1<'));

  const sales = xmlOf(view, 'xl/worksheets/sheet1.xml');
  const csv = /<customSheetViews><customSheetView ([^>]*)>([\s\S]*?)<\/customSheetView><\/customSheetViews>/.exec(sales);
  assert.ok(csv, 'the sales sheet\'s view');
  assert.match(csv[1], new RegExp('guid="' + guid.replace(/[{}]/g, '\\$&') + '"'));
  assert.match(csv[1], /scale="125"/);
  assert.match(csv[1], /hiddenRows="1"/);
  assert.match(csv[1], /hiddenColumns="1"/);
  assert.match(csv[1], /printArea="1"/);
  assert.match(csv[2], /^<selection activeCell="B7" sqref="B7:D10"\/>/);
  assert.match(csv[2], /<pageMargins\b[^>]*\/><pageSetup\b[^>]*orientation="landscape"/, 'the print settings, in schema order');
  assert.ok(sales.indexOf('<customSheetViews>') < sales.indexOf('<pageMargins'), 'customSheetViews before the page setup in the tail');
  assert.match(xmlOf(view, 'xl/worksheets/sheet2.xml'), new RegExp('<customSheetView guid="' + guid.replace(/[{}]/g, '\\$&') + '"'), 'and in the other sheet');
  assert.deepEqual(new SheetView(view.save()).customViews(), [{ name: 'Quarter review', printSettings: true, hiddenRowCol: true, sheet: 'Sales' }], 'read back');
});

test('Show puts the view back: sheet, zoom, selection, hidden rows and columns, print settings and view', () => {
  const view = open();
  hideSome(view);
  writePageSetup(view, 'Sales', { ...readPageSetup(view, 'Sales'), orientation: 'landscape' });
  view.setViewMode('pageBreakPreview');
  view.select(6, 1);
  view.addCustomView({ name: 'Review', zoom: 0.8 });

  // Everything changed after.
  const { part } = view.workbook._sheetPart('Sales');
  for (const r of [2, 3, 4]) part.setRowOutline(r, { hidden: false });
  part.setColOutline(2, { hidden: false });
  part.setRowOutline(20, { hidden: true });
  view._rebuildDerivedState();
  writePageSetup(view, 'Sales', { ...readPageSetup(view, 'Sales'), orientation: 'portrait' });
  view.setViewMode('normal');
  view.selectSheet('Other data');

  const shown = new SheetView(view.save());
  const back = shown.showCustomView('Review');
  assert.deepEqual(back, { sheet: 'Sales', zoom: 0.8 });
  assert.equal(shown.activeSheet, 'Sales');
  assert.deepEqual([shown.selection.active.row, shown.selection.active.col], [6, 1]);
  assert.deepEqual([...shown.geo.hiddenRows].sort((a, b) => a - b), [2, 3, 4], 'rows 3 to 5 hidden again, row 21 shown');
  assert.deepEqual([...shown.geo.hiddenCols], [2]);
  assert.equal(readPageSetup(shown, 'Sales').orientation, 'landscape', 'the print settings');
  assert.equal(shown.viewMode(), 'pageBreakPreview', 'and the view');
});

test('a view that leaves out print settings and hidden rows says so, and Show leaves those alone', () => {
  const view = open();
  hideSome(view);
  view.addCustomView({ name: 'Plain', printSettings: false, hiddenRowCol: false });
  const wb = xmlOf(view, 'xl/workbook.xml');
  assert.match(wb, /<customWorkbookView name="Plain"[^>]*includePrintSettings="0" includeHiddenRowCol="0"/);
  assert.doesNotMatch(wb, /\.wvu\.Rows/, 'no hidden names kept');
  const { part } = view.workbook._sheetPart('Sales');
  for (const r of [2, 3, 4]) part.setRowOutline(r, { hidden: false });
  view._rebuildDerivedState();
  writePageSetup(view, 'Sales', { ...readPageSetup(view, 'Sales'), orientation: 'landscape' });
  view.showCustomView('Plain');
  assert.equal(view.geo.hiddenRows.size, 0, 'rows as they are now');
  assert.equal(readPageSetup(view, 'Sales').orientation, 'landscape', 'print settings as they are now');
});

test('Delete takes the view, its sheet views and its names away; a view of the same name is replaced', () => {
  const view = open();
  hideSome(view);
  view.addCustomView({ name: 'One' });
  view.addCustomView({ name: 'one' });
  assert.equal(view.customViews().length, 1, 'replaced, as Excel offers to');
  view.addCustomView({ name: 'Two', hiddenRowCol: false });
  view.deleteCustomView('one');
  assert.deepEqual(view.customViews().map((v) => v.name), ['Two']);
  view.deleteCustomView('Two');
  const wb = xmlOf(view, 'xl/workbook.xml');
  assert.doesNotMatch(wb, /customWorkbookView|\.wvu\./);
  assert.doesNotMatch(xmlOf(view, 'xl/worksheets/sheet1.xml'), /customSheetView/);
});

test('in a workbook with a table Custom Views are greyed, with Excel\'s reason', () => {
  const view = open();
  assert.equal(view.customViewsBlocked(), null);
  view.select(0, 0);
  view.select(10, 3, { extend: true });
  view.formatAsTable({ style: 'TableStyleMedium2' });
  assert.match(view.customViewsBlocked(), /contains a table/);
  assert.throws(() => view.addCustomView({ name: 'X' }), /contains a table/);
  assert.equal(view.render().customViewsBlocked, view.customViewsBlocked(), 'the frame says why');
});
