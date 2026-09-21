/**
 * The tabs: a sheet added, renamed and taken out — the engine's half.
 *
 * A new sheet is a part, a content type, a relationship and a `<sheet>` entry
 * with the next sheetId; a rename follows every formula and defined name that
 * reads the sheet; a removal takes the entry, the relationship and the part
 * away and keeps the workbook's names pointing at the right sheets.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

const open = () => new SheetView(buildXlsx({ sheets: [
  { name: 'Data', rows: [['x', 1], ['y', 2]] },
  { name: 'Summary', rows: [['=SUM(Data!B1:B2)', "='Data'!A1"]] },
] }));

test('a new sheet is a real part with a relationship and the next sheetId, and undo takes it out again', () => {
  const view = open();
  const name = view.addSheet();
  assert.equal(name, 'Sheet3', 'named the way Excel names one');
  assert.deepEqual(view.sheetNames(), ['Data', 'Summary', 'Sheet3']);
  assert.equal(view.activeSheet, 'Sheet3', 'and the view is on it');

  const pkg = OoxmlPackage.read(view.save());
  assert.ok(pkg.has('xl/worksheets/sheet3.xml'), 'the part');
  assert.match(pkg.text('xl/workbook.xml'), /<sheet name="Sheet3" sheetId="3" r:id="rId\d+"\/><\/sheets>/, 'the entry, last');
  assert.match(pkg.text('[Content_Types].xml'), /sheet3\.xml/, 'its content type');
  assert.equal(new SheetView(view.save()).sheetNames().length, 3, 'and it survives a save');

  view.selectSheet('Data');
  view.undo();
  assert.deepEqual(view.sheetNames(), ['Data', 'Summary'], 'undo takes the sheet out');
  assert.ok(!view.pkg.has('xl/worksheets/sheet3.xml'), 'part and all');

  assert.throws(() => view.addSheet('Data'), /already a sheet/);
  assert.throws(() => view.addSheet('a/b'), /cannot contain/);
  assert.throws(() => view.addSheet(''.padEnd(32, 'x')), /1 to 31/);
});

test('a rename follows the formulas and names that read the sheet', () => {
  const view = open();
  view.renameSheet('Data', 'Figures 2024');
  assert.deepEqual(view.sheetNames(), ['Figures 2024', 'Summary']);
  view.selectSheet('Summary');
  assert.equal(view.calc.getInput('Summary', 0, 0), "=SUM('Figures 2024'!B1:B2)", 'a bare reference is quoted for the space');
  assert.equal(view.calc.getInput('Summary', 0, 1), "='Figures 2024'!A1");
  assert.equal(view.displayValue(0, 0).text, '3', 'and still computes');

  const reopened = new SheetView(view.save());
  assert.deepEqual(reopened.sheetNames(), ['Figures 2024', 'Summary'], 'kept in the file');
  reopened.selectSheet('Summary');
  assert.equal(reopened.displayValue(0, 0).text, '3');

  view.undo();
  assert.deepEqual(view.sheetNames(), ['Data', 'Summary'], 'one undo step');
  assert.equal(view.calc.getInput('Summary', 0, 0), '=SUM(Data!B1:B2)');
});

test('a removed sheet leaves no entry, relationship or part, and the last sheet cannot go', () => {
  const view = open();
  view.defineName('Total', 'Summary!$A$1');
  view.selectSheet('Data');
  view.removeSheet('Data');
  assert.deepEqual(view.sheetNames(), ['Summary']);
  assert.equal(view.activeSheet, 'Summary', 'the view moved to the sheet that is left');
  const pkg = OoxmlPackage.read(view.save());
  assert.ok(!pkg.has('xl/worksheets/sheet1.xml'), 'the part is gone');
  assert.doesNotMatch(pkg.text('xl/workbook.xml'), /name="Data"/, 'and the entry');
  assert.doesNotMatch(pkg.text('xl/_rels/workbook.xml.rels'), /sheet1\.xml/, 'and the relationship');
  assert.match(pkg.text('xl/workbook.xml'), /<definedName name="Total">Summary!\$A\$1<\/definedName>/, 'a workbook name stays');
  assert.throws(() => view.removeSheet('Summary'), /at least one sheet/);
});
