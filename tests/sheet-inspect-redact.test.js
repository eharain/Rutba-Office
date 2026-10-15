// Inspect Workbook and Redact cells: the inspector finds comments and
// notes, personal details, properties, headers and footers, hidden
// worksheets, hidden rows and columns and hidden names, and takes them out;
// Redact blacks out the words and numbers marked in every cell, comment,
// header and property, the package read back with none of them left.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx, OoxmlPackage } from '@rutba/ooxml';
import { redactPackage, partsHolding } from '@rutba/ooxml/redact';
import { inspectPackage, removeFromPackage } from '@rutba/ooxml/inspect';
import { SheetView } from '@rutba/sheet-view';

/** A workbook with a name in cells, a note, a header, a hidden sheet, a hidden row and column, a hidden name. */
export function secretWorkbook() {
  const pkg = OoxmlPackage.read(buildXlsx({ sheets: [
    { name: 'Staff', rows: [['Name', 'Phone', 'Pay'], ['Ayesha Khan', '+92 300 1234567', 120000], ['Bilal Ahmed', '+92 321 7654321', 95000]] },
    { name: 'Salaries', rows: [['Ayesha Khan', 120000]] },
  ] }));
  const s1 = 'xl/worksheets/sheet1.xml';
  let x = pkg.text(s1);
  x = x.replace(/<row r="3"/, '<row r="3" hidden="1"').replace('<sheetData>', '<cols><col min="3" max="3" width="10" hidden="1" customWidth="1"/></cols><sheetData>');
  x = x.replace('</worksheet>', '<headerFooter><oddHeader>&amp;LPrepared by Ayesha Khan</oddHeader></headerFooter></worksheet>');
  pkg.write_(s1, x);
  let wb = pkg.text('xl/workbook.xml');
  wb = wb.replace(/(<sheet\b[^>]*name="Salaries"[^>]*)\/>/, '$1 state="hidden"/>');
  wb = wb.includes('<definedNames>') ? wb.replace('<definedNames>', '<definedNames><definedName name="Secret" hidden="1">Staff!$C$2</definedName>') : wb.replace('</sheets>', '</sheets><definedNames><definedName name="Secret" hidden="1">Staff!$C$2</definedName></definedNames>');
  pkg.write_('xl/workbook.xml', wb);
  pkg.addPart('docProps/core.xml', '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Ayesha Khan</dc:creator><dc:title>Payroll</dc:title></cp:coreProperties>', 'application/vnd.openxmlformats-package.core-properties+xml');
  pkg.write_('_rels/.rels', pkg.text('_rels/.rels').replace('</Relationships>', '<Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'));
  // A note on A2, written the way the workbook engine writes one.
  const view = SheetView.open(pkg.write());
  view.setNote?.({ row: 1, col: 0, text: 'Ayesha Khan joined in May', author: 'Sana Iqbal' });
  return Buffer.from(view.save());
}

test('Inspect Workbook finds what Excel\'s inspector does', () => {
  const found = Object.fromEntries(inspectPackage(OoxmlPackage.read(secretWorkbook()), 'workbook').map((k) => [k.id, k]));
  assert.equal(found.comments.found, 1);
  assert.deepEqual(found.hiddenSheets.items, ['Salaries']);
  assert.equal(found.hiddenRowsColumns.rows, 1);
  assert.equal(found.hiddenRowsColumns.columns, 1);
  assert.deepEqual(found.hiddenNames.items, ['Secret']);
  assert.equal(found.headers.found, 1);
  assert.ok(found.personal.items.includes('Ayesha Khan'));
});

test('Remove All takes them out and the workbook still opens', () => {
  const pkg = OoxmlPackage.read(secretWorkbook());
  removeFromPackage(pkg, 'workbook', ['comments', 'hiddenSheets', 'hiddenNames', 'headers', 'personal', 'properties']);
  const after = Object.fromEntries(inspectPackage(pkg, 'workbook').map((k) => [k.id, k.found]));
  for (const id of ['comments', 'hiddenSheets', 'hiddenNames', 'headers', 'personal']) assert.equal(after[id], 0, id);
  const view = SheetView.open(pkg.write());
  assert.deepEqual(view.workbook.sheetNames(), ['Staff']);
});

test('Redact blacks out the marked cells, numbers and all, wherever the workbook keeps them', () => {
  const pkg = OoxmlPackage.read(secretWorkbook());
  redactPackage(pkg, { marks: [{ text: 'Ayesha Khan' }, { pattern: 'phone' }, { regex: '^120000$' }] });
  const out = Buffer.from(pkg.write());
  assert.deepEqual(partsHolding(out, ['Ayesha Khan', '300 1234567', '321 7654321']), []);
  const view = SheetView.open(out);
  assert.equal(view.workbook.sheetNames().length, 2);
  assert.doesNotMatch(pkg.text('xl/worksheets/sheet1.xml'), /<v>120000<\/v>/);
});
