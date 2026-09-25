// Review → Check Accessibility in Worksheets: pictures and charts without a
// description (written on `xdr:cNvPr descr`), sheets still called "Sheet1",
// merged cells, a table saved without a header row, pale text on a cell's
// fill, and a link whose words are its address.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { checkAccessibility, describeSheet, setSheetAltText, setSheetTableHeader } from '@rutba/proofing';
import { gradientPng } from '../apps/desktop/main/sample-picture.js';

function fixture() {
  const view = new SheetView(buildXlsx({ sheets: [
    { name: 'Sales', rows: [['Item', 'Qty', 'Notes'], ['Pens', 12, 'blue'], ['Ink', 3, 'black'], [], ['Total', 15]] },
    { name: 'Sheet2', rows: [['Region', 'Revenue'], ['North', 10], ['South', 20]] },
  ] }));
  view.select(6, 0);
  view.insertPicture({ name: 'Logo', contentType: 'image/png', data: gradientPng(20, 20, [30, 30, 200], [200, 30, 30]), widthPx: 60, heightPx: 40 });
  // A table saved with no header row: the list below its first line.
  view.select(1, 0);
  view.select(2, 1, { extend: true });
  view.formatAsTable({});
  const table = view.workbook.tables()[0];
  view.pkg.write_(table.part, view.pkg.text(table.part).replace('<table ', '<table headerRowCount="0" '));
  // Merged cells, pale text, a bare address for a link.
  view.select(4, 0);
  view.select(4, 1, { extend: true });
  view.mergeSelection();
  view.select(0, 2);
  view.setFormat({ fontColour: '#D9D9D9' });
  view.setHyperlink({ row: 2, col: 2, href: 'https://example.com/ink' });
  view.setCell(2, 2, 'https://example.com/ink');
  return view;
}

test('a workbook is described sheet by sheet and flagged', () => {
  const view = fixture();
  const { issues } = checkAccessibility(describeSheet(view));
  const by = (rule) => issues.filter((i) => i.rule === rule);
  assert.equal(by('altText').length, 1);
  assert.deepEqual(by('altText')[0].where, { sheet: 'Sales', row: 6, col: 0, drawing: 0 }, 'found at its anchor cell');
  assert.deepEqual(by('defaultSheetName').map((i) => i.label), ['Sheet2']);
  assert.deepEqual(by('mergedCells').map((i) => i.label), ['Sales!A5:B5']);
  assert.equal(by('mergedCells')[0].fixes[0].kind, 'unmerge');
  assert.equal(by('tableHeader').length, 1);
  assert.equal(by('contrast').length, 1);
  assert.equal(by('contrast')[0].label, 'Sales!C1');
  assert.equal(by('linkText').length, 1);
});

test('alt text and decorative are written on xdr:cNvPr and survive a save; undo takes them back', () => {
  const view = fixture();
  setSheetAltText(view, { sheet: 'Sales', anchor: 0, descr: 'The company logo' });
  let drawing = view.pkg.text('xl/drawings/drawing1.xml');
  assert.match(drawing, /<xdr:cNvPr [^>]*descr="The company logo"/);
  const reopened = new SheetView(view.save());
  assert.equal(describeSheet(reopened).objects[0].alt, 'The company logo');

  setSheetAltText(view, { sheet: 'Sales', anchor: 0, decorative: true });
  drawing = view.pkg.text('xl/drawings/drawing1.xml');
  assert.match(drawing, /<xdr:cNvPr [^>]*>[\s\S]*<adec:decorative [^>]*val="1"\/>[\s\S]*<\/xdr:cNvPr>/);
  assert.ok(!checkAccessibility(describeSheet(view)).issues.some((i) => i.rule === 'altText'));
  view.undo();
  assert.match(view.pkg.text('xl/drawings/drawing1.xml'), /descr="The company logo"/);
});

test('Use First Row as Header names the columns from the first row and puts the filter back', () => {
  const view = fixture();
  const { part } = view.workbook.tables()[0];
  setSheetTableHeader(view, { part });
  const xml = view.pkg.text(part);
  assert.ok(!/headerRowCount="0"/.test(xml));
  assert.deepEqual([...xml.matchAll(/<tableColumn\b[^>]*name="([^"]*)"/g)].map((m) => m[1]), ['Pens', '12']);
  assert.match(xml, /<autoFilter ref="A2:B3"\/>/);
  assert.ok(!checkAccessibility(describeSheet(view)).issues.some((i) => i.rule === 'tableHeader'));
  view.undo();
  assert.match(view.pkg.text(part), /headerRowCount="0"/, 'one undo step');
});
