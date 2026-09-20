// Hyperlinks in a worksheet: put on a cell, read back with their target,
// kept through a save, followed to a place in the workbook, taken off —
// and the notes a file carries, counted for the status bar.
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { buildXlsx } from '@rutba/ooxml/build';

const book = () => buildXlsx({
  sheets: [
    { name: 'Sales', rows: [['Region', 'Q1'], ['North', 1420], ['South', 860], ['See the ledger']] },
    { name: 'Ledger', rows: [['Ledger', 'Amount'], ['North', 1420]] },
  ],
});
const cellOf = (view, ref) => view.render().cells.find((c) => c.ref === ref);

test('an address on a cell is a relationship marked external, read back with its target and kept through a save', () => {
  const view = SheetView.open(book());
  view.select(1, 0);
  assert.equal(view.render().link, null, 'no link on A2 yet');
  view.setHyperlink({ row: 1, col: 0, href: 'https://rutba.io', tooltip: 'Home' });
  assert.deepEqual(cellOf(view, 'A2').link, { ref: 'A2', href: 'https://rutba.io', location: null, display: null, tooltip: 'Home' });
  assert.equal(view.render().link.href, 'https://rutba.io', 'the active cell\'s link rides with the model');
  assert.equal(cellOf(view, 'A3').link, null);

  const part = view.workbook.partNameFor('Sales');
  const reopened = SheetView.open(view.save());
  assert.deepEqual(reopened.workbook.hyperlinks('Sales'), [{ ref: 'A2', href: 'https://rutba.io', location: null, display: null, tooltip: 'Home' }]);
  const rels = reopened.pkg.text(part.replace(/worksheets\//, 'worksheets/_rels/') + '.rels');
  assert.match(rels, /Type="[^"]*\/hyperlink" Target="https:\/\/rutba.io" TargetMode="External"/, 'Excel wants an external target said so');
  const xml = reopened.pkg.text(part);
  assert.match(xml, /<hyperlinks><hyperlink ref="A2" r:id="rId\d+" tooltip="Home"\/><\/hyperlinks>/, 'one block, after the data, as the schema orders it');
  assert.ok(xml.indexOf('<hyperlinks>') < xml.indexOf('<pageMargins') || !xml.includes('<pageMargins'), 'before the page margins');
  assert.match(xml, /<worksheet[^>]*xmlns:r="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships"/, 'the r: prefix is declared on the root');
});

test('a place in the workbook is a location, not a relationship; a second link replaces the first; removing empties the block', () => {
  const view = SheetView.open(book());
  view.setHyperlink({ row: 3, col: 0, location: 'Ledger!A2', tooltip: 'The ledger row' });
  assert.deepEqual(cellOf(view, 'A4').link, { ref: 'A4', href: null, location: 'Ledger!A2', display: null, tooltip: 'The ledger row' });
  assert.equal(view.pkg.has(view.workbook.partNameFor('Sales').replace(/worksheets\//, 'worksheets/_rels/') + '.rels'), false, 'no rels part was made for a place');

  view.setHyperlink({ row: 3, col: 0, href: 'mailto:hello@rutba.io' });
  assert.equal(view.workbook.hyperlinks('Sales').length, 1, 'the second link replaced the first');
  assert.equal(cellOf(view, 'A4').link.href, 'mailto:hello@rutba.io');
  assert.equal(cellOf(view, 'A4').link.location, null);

  assert.equal(view.removeHyperlink({ row: 3, col: 0 }), view, 'chainable');
  assert.equal(cellOf(view, 'A4').link, null);
  assert.equal(view.workbook.hyperlinks('Sales').length, 0);
  assert.equal(SheetView.open(view.save()).pkg.text(view.workbook.partNameFor('Sales')).includes('<hyperlinks'), false, 'an emptied block goes with its last link');
  assert.throws(() => view.setHyperlink({ row: 0, col: 0 }), /address or a place/);
});

test('a link is one undo step, and the notes a sheet carries are counted', () => {
  const view = SheetView.open(book());
  view.setHyperlink({ row: 1, col: 1, href: 'https://rutba.io/office' });
  assert.equal(cellOf(view, 'B2').link.href, 'https://rutba.io/office');
  view.undo();
  assert.equal(cellOf(view, 'B2').link, null, 'undone');
  assert.equal(view.workbook.hyperlinks('Sales').length, 0);
  view.redo();
  assert.equal(cellOf(view, 'B2').link.href, 'https://rutba.io/office', 'redone');
  assert.equal(view.render().notes, 0);

  // A comments part, as Excel writes one: the corner mark and the tip come from it.
  const part = view.workbook.partNameFor('Sales');
  view.pkg.addPart('xl/comments1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Kim Lee</author></authors><commentList><comment ref="B2" authorId="0"><text><r><t>Check this figure against the ledger.</t></r></text></comment></commentList></comments>', 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
  view.pkg.addRelationshipTo(part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', '../comments1.xml');
  const noted = SheetView.open(view.save());
  assert.equal(noted.render().notes, 1);
  assert.deepEqual(cellOf(noted, 'B2').note, { author: 'Kim Lee', text: 'Check this figure against the ledger.' });
  assert.equal(cellOf(noted, 'B2').link.href, 'https://rutba.io/office', 'the link survived beside the note');
});
