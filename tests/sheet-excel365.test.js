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

test('the new array functions work out as Excel does', async () => {
  const { calculate } = await import('@rutba/formula');
  const cells = { A1: 'apple', A2: 'banana', A3: 'cherry', B1: 3, B2: 1, B3: 2 };
  const L = (n) => String.fromCharCode(65 + n);
  const resolver = {
    getCell: (s, r, c) => cells[L(c) + (r + 1)] ?? '',
    getRange: (s, a, b) => Array.from({ length: b.row - a.row + 1 }, (_, i) => Array.from({ length: b.col - a.col + 1 }, (_, j) => cells[L(a.col + j) + (a.row + i + 1)] ?? '')),
    getName: () => null,
  };
  const show = (v) => (Array.isArray(v) ? v.map((l) => l.map((x) => (x?.type ? x.type : x))) : (v?.type ? v.type : v));
  const at = (f) => show(calculate(f, resolver, { spill: true }));
  assert.equal(at('XMATCH("banana",A1:A3)'), 2);
  assert.equal(at('XMATCH(2.5,B1:B3,-1)'), 3);
  assert.equal(at('XMATCH("c*",A1:A3,2)'), 3);
  assert.deepEqual(at('TEXTSPLIT("x,y;z",",",";")'), [['x', 'y'], ['z', '#N/A']]);
  assert.equal(at('TEXTBEFORE("one-two-three","-",-1)'), 'one-two');
  assert.equal(at('TEXTAFTER("one-two-three","-",2)'), 'three');
  assert.equal(at('TEXTAFTER("abc","x",1,0,0,"none")'), 'none');
  assert.deepEqual(at('VSTACK({1,2},{3;4})'), [[1, 2], [3, '#N/A'], [4, '#N/A']]);
  assert.deepEqual(at('HSTACK({1;2},{3,4})'), [[1, 3, 4], [2, '#N/A', '#N/A']]);
  assert.deepEqual(at('TAKE(A1:B3,-1,1)'), [['cherry']]);
  assert.deepEqual(at('DROP(A1:B3,1,1)'), [[1], [2]]);
  assert.deepEqual(at('CHOOSECOLS(A1:B3,2,-2)'), [[3, 'apple'], [1, 'banana'], [2, 'cherry']]);
  assert.deepEqual(at('TOROW({1,2;3,4},0,TRUE)'), [[1, 3, 2, 4]]);
  assert.deepEqual(at('TOCOL({1,#N/A,3},2)'), [[1], [3]]);
  assert.deepEqual(at('WRAPROWS({1,2,3,4,5},2)'), [[1, 2], [3, 4], [5, '#N/A']]);
  assert.deepEqual(at('EXPAND({1,2},2,3,"-")'), [[1, 2, '-'], ['-', '-', '-']]);
  assert.deepEqual(at('SORTBY(A1:A3,B1:B3,-1)'), [['apple'], ['cherry'], ['banana']]);
  assert.equal(at('ARRAYTOTEXT({1,"a";TRUE,2},1)'), '{1,"a";TRUE,2}');
  assert.equal(at('ROWS(RANDARRAY(3,2))'), 3);
  assert.equal(at('TAKE(A1:B3,0)'), '#CALC!');
  // LAMBDA, called where written, from LET, and through each helper.
  assert.equal(at('LAMBDA(x,y,x*y)(3,4)'), 12);
  assert.equal(at('LET(f,LAMBDA(n,n*2),f(21))'), 42);
  assert.equal(at('LAMBDA(x,LAMBDA(y,x+y))(1)(2)'), 3);
  assert.deepEqual(at('MAP(B1:B3,LAMBDA(v,v*10))'), [[30], [10], [20]]);
  assert.equal(at('REDUCE(0,B1:B3,LAMBDA(acc,v,acc+v))'), 6);
  assert.deepEqual(at('SCAN(0,B1:B3,LAMBDA(acc,v,acc+v))'), [[3], [4], [6]]);
  assert.deepEqual(at('BYROW(B1:B3,LAMBDA(r,r*2))'), [[6], [2], [4]]);
  assert.deepEqual(at('MAKEARRAY(2,2,LAMBDA(r,c,r*10+c))'), [[11, 12], [21, 22]]);
  assert.equal(at('LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))(5)'), 5);
  assert.equal(at('LAMBDA(x,x)'), '#CALC!', 'a LAMBDA left uncalled');
  assert.equal(at('LAMBDA(x,x)(1,2)'), '#VALUE!');
});

test('a LAMBDA kept in a defined name is called by name, recalculates with what it reads, and is saved as Excel writes it', () => {
  const view = SheetView.open(book([[2], [5]]));
  view.defineName('TWICE', '=LAMBDA(n, n*2)');
  view.defineName('RATE', '=Data!$A$2/100');
  view.setCell(0, 2, '=TWICE(A1)');
  view.setCell(1, 2, '=A1*RATE');
  assert.equal(view.calc.getValue('Data', 0, 2), 4);
  assert.equal(view.calc.getValue('Data', 1, 2), 0.1);
  view.setCell(1, 0, 10);
  assert.equal(view.calc.getValue('Data', 1, 2), 0.2, 'what the name reads changed, and the cell with it');
  const wb = Workbook.open(view.save());
  assert.match(wb.pkg.text('xl/workbook.xml'), /<definedName name="TWICE">_xlfn\.LAMBDA\(_xlpm\.n, _xlpm\.n\*2\)<\/definedName>/);
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.calc.getValue('Data', 0, 2), 4);
  assert.deepEqual(reopened.names().find((n) => n.name === 'TWICE'), { name: 'TWICE', ref: '=LAMBDA(n, n*2)', target: null, formula: true });
  assert.throws(() => view.defineName('BAD', '=SUM('), /not a range or a formula/);
});
