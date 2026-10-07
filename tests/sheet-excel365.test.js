// Excel 365's formulas, in the file — the engine's half.
//
// A function newer than Excel 2007 is written `_xlfn.NAME` (FILTER, SORT and
// SORTBY `_xlfn._xlws.NAME`), a name LET or LAMBDA binds `_xlpm.name`, and
// the spill operator `A1#` `_xlfn.ANCHORARRAY(A1)` — written without them,
// Excel shows #NAME?. A formula that spills carries `cm` naming the
// workbook's dynamic-array metadata, or Excel 365 reads it as an old array
// formula. Read back, the prefixes come off again: the formula bar shows
// what was typed.
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { toFileFormula, fromFileFormula } from '../packages/ooxml/src/xlfn.js';

const book = (rows) => buildXlsx({ sheets: [{ name: 'Data', rows }] });
const sheetXml = (view) => Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');

test('a formula is written as Excel writes it, and read back as it was typed', () => {
  const cases = [
    ['XLOOKUP(1,A1:A2,B1:B2)', '_xlfn.XLOOKUP(1,A1:A2,B1:B2)'],
    ['SORT(A1:A3)', '_xlfn._xlws.SORT(A1:A3)'],
    ['FILTER(A1:B9,A1:A9>2)', '_xlfn._xlws.FILTER(A1:B9,A1:A9>2)'],
    ['sum(A1:A3)+IFS(A1>1,"x",TRUE,"y")', 'sum(A1:A3)+_xlfn.IFS(A1>1,"x",TRUE,"y")'],
    ['LET(a,1,b,a+1,a*b)', '_xlfn.LET(_xlpm.a,1,_xlpm.b,_xlpm.a+1,_xlpm.a*_xlpm.b)'],
    ['LET(f,LAMBDA(n,n*2),f(3))', '_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2),_xlpm.f(3))'],
    ['SUM(A1#)+COUNT(Sheet2!$B$3#)', 'SUM(_xlfn.ANCHORARRAY(A1))+COUNT(_xlfn.ANCHORARRAY(Sheet2!$B$3))'],
    ["'My Sheet'!A1#", "_xlfn.ANCHORARRAY('My Sheet'!A1)"],
    ['LEN("XLOOKUP(A1#")&CONCAT(A1,B1)', 'LEN("XLOOKUP(A1#")&_xlfn.CONCAT(A1,B1)'],
    ['IF(ISNA(A1),#N/A,VLOOKUP(A1,B:C,2,FALSE))', 'IF(ISNA(A1),#N/A,VLOOKUP(A1,B:C,2,FALSE))'],
  ];
  for (const [typed, file] of cases) {
    assert.equal(toFileFormula(typed), file, typed);
    assert.equal(fromFileFormula(file), typed, file);
  }
});

test('a sheet saves its newer functions prefixed, and opens an Excel file\'s prefixed ones as typed', () => {
  const view = SheetView.open(book([[3, 'a'], [1, 'b'], [2, 'c']]));
  view.setCell(0, 3, '=XLOOKUP(2,A1:A3,B1:B3)');
  view.setCell(0, 4, '=LET(x,A1,x*2)');
  assert.equal(view.calc.getValue('Data', 0, 3), 'c');
  assert.equal(view.calc.getValue('Data', 0, 4), 6);
  const xml = sheetXml(view);
  assert.match(xml, /<f>_xlfn\.XLOOKUP\(2,A1:A3,B1:B3\)<\/f>/);
  assert.match(xml, /<f>_xlfn\.LET\(_xlpm\.x,A1,_xlpm\.x\*2\)<\/f>/);
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.editValue(0, 3), '=XLOOKUP(2,A1:A3,B1:B3)', 'the formula bar shows what was typed');
  assert.equal(reopened.editValue(0, 4), '=LET(x,A1,x*2)');
  assert.equal(reopened.calc.getValue('Data', 0, 4), 6);
});

test('a spilling formula is saved as a dynamic array Excel 365 reads, its metadata written once', () => {
  const view = SheetView.open(book([[3], [1], [2]]));
  view.setCell(0, 2, '=SORT(A1:A3)');
  view.setCell(0, 4, '=SUM(C1#)');
  assert.equal(view.calc.getValue('Data', 2, 2), 3, 'spilled');
  assert.equal(view.calc.getValue('Data', 0, 4), 6);
  const wb = Workbook.open(view.save());
  const xml = wb.pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<c r="C1"[^>]*\bcm="1"[^>]*><f t="array" ref="C1:C3">_xlfn\._xlws\.SORT\(A1:A3\)<\/f>/);
  assert.match(xml, /<f>SUM\(_xlfn\.ANCHORARRAY\(C1\)\)<\/f>/);
  assert.match(wb.pkg.text('xl/metadata.xml'), /<metadataType name="XLDAPR"[\s\S]*<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"\/>[\s\S]*<cellMetadata count="1"><bk><rc t="1" v="0"\/><\/bk><\/cellMetadata>/);
  assert.ok(wb.pkg.rels('xl/workbook.xml').some((r) => /\/sheetMetadata$/.test(r.Type) && r.Target === 'metadata.xml'));
  assert.match(wb.pkg.text('[Content_Types].xml'), /PartName="\/xl\/metadata\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheetMetadata\+xml"/);
  // Saved again: one metadata part, the same cm.
  const again = SheetView.open(view.save());
  again.setCell(0, 3, '=SORT(A1:A3,1,-1)');
  const wb2 = Workbook.open(again.save());
  assert.equal(wb2.pkg.partNames().filter((p) => /metadata/.test(p)).length, 1);
  assert.match(wb2.pkg.text('xl/worksheets/sheet1.xml'), /<c r="D1"[^>]*\bcm="1"/);
  assert.equal(again.editValue(0, 4), '=SUM(C1#)', 'the spill operator reads back as typed');
});
