// Faults the third review of Sheets and Documents found, 10 Nov 2026, each reproduced on the code as it stood.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Spreadsheet, FormulaError } from '@rutba/formula';
import { buildXlsx } from '@rutba/ooxml/build';
import { Workbook } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { writeCachedValue } from '../packages/ooxml/src/recalc.js';
import { textOf } from '../packages/ooxml/src/runs.js';

const book = (rows, extra = {}) => buildXlsx({ sheets: [{ name: 'Sales', rows }], ...extra });
const TABLE = [['Item', 'Qty', 'Price', 'Note'], ['Pens', 12, 1.5, 'a'], ['Ink', 3, 9, 'b'], ['Clips', 40, 0.5, 'c']];
function tabled(rows = TABLE) {
  const view = new SheetView(book(rows));
  view.select(1, 1);
  view.formatAsTable({ style: 'TableStyleMedium2', stripes: true });
  return SheetView.open(view.save());
}
const one = (calc) => (calc && typeof calc === 'object' ? calc.type : calc);

test('a saved #CALC! or #SPILL! is one Excel can open, and the file is worked out again when it is', () => {
  for (const kind of ['#CALC!', '#SPILL!', '#CIRCULAR!']) {
    const wb = Workbook.open(book([[1, '=A1']]));
    writeCachedValue(wb, 'Sales', 'B1', '=A1', new FormulaError(kind));
    const saved = Workbook.open(wb.save());
    const xml = saved.pkg.text(saved.sheets()[0].part);
    assert.doesNotMatch(xml, new RegExp(`<v>${kind.replace(/[#!]/g, '\\$&')}</v>`), kind);
    assert.match(xml, /<c r="B1"[^>]*t="e"[^>]*>.*?<v>#VALUE!<\/v>/);
    assert.match(saved.pkg.text(saved.mainPart), /fullCalcOnLoad="1"/, 'Excel works it out again on opening');
  }
});

test('a number past what a file can hold is saved as #NUM!, never as Infinity', () => {
  const wb = Workbook.open(book([[1, '=A1']]));
  writeCachedValue(wb, 'Sales', 'B1', '=A1', Infinity);
  const saved = Workbook.open(wb.save());
  const xml = saved.pkg.text(saved.sheets()[0].part);
  assert.doesNotMatch(xml, /Infinity/);
  assert.match(xml, /<c r="B1"[^>]*t="e"[^>]*>.*?<v>#NUM!<\/v>/);
});

test('a tracked deletion is no longer the paragraph\'s words, whatever it held', () => {
  const kept = '<w:r><w:t>Kept words</w:t></w:r>';
  const gone = '<w:del w:id="1" w:author="A"><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:delText>gone</w:delText></w:r>'
    + '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:del>';
  assert.equal(textOf(kept + gone), textOf(kept));
});

test('every data row of a table deleted leaves one, empty, and pulls nothing below into the table', () => {
  const view = tabled([...TABLE, [], ['Below the table']]);
  const wb = view.workbook;
  wb.deleteRows('Sales', 1, 3);
  const t = wb.tables()[0];
  assert.equal(t.ref, 'A1:D2', 'a header and one empty row');
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.editValue(1, 0), '', 'the row kept is empty');
  assert.equal(reopened.editValue(3, 0), 'Below the table', 'and what was below is still outside it');
});

test('a row inserted in a table takes its calculated column', () => {
  const view = tabled();
  const part = view.workbook.tables()[0].part;
  view.workbook.pkg.write_(part, view.workbook.pkg.text(part).replace(/<tableColumn id="4"[^>]*\/>/, (m) => m.replace('/>', '><calculatedColumnFormula>B2*2</calculatedColumnFormula></tableColumn>')));
  const again = SheetView.open(view.save());
  again.workbook.insertRows('Sales', 2, 1);
  const reopened = SheetView.open(again.save());
  assert.match(reopened.editValue(2, 3), /^=B3\*2$/);
});

test('a defined name with a relative part is left where it is by an insert, and an absolute one moves', () => {
  const wb = Workbook.open(book([[1], [2]], { definedNames: [{ name: 'Here', ref: 'Sales!A1' }, { name: 'Fixed', ref: 'Sales!$A$1' }] }));
  wb.insertRows('Sales', 0, 1);
  const names = Object.fromEntries(wb.definedNames().map((n) => [n.name, n.ref ?? n.value ?? n.formula]));
  assert.equal(names.Here, 'Sales!A1');
  assert.equal(names.Fixed, 'Sales!$A$2');
});

test('SECOND of a negative time is #NUM!, as the other parts of a date are', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setCell('S', 0, 0, '=SECOND(-1)');
  s.setCell('S', 0, 1, '=MINUTE(-1)');
  s.recalculate?.();
  assert.equal(one(s.getValue('S', 0, 0)), '#NUM!');
  assert.equal(one(s.getValue('S', 0, 1)), '#NUM!');
});

test('ADDRESS gives R1C1 addresses when asked, as Excel writes them', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  const cases = [['=ADDRESS(2,3,1,FALSE)', 'R2C3'], ['=ADDRESS(2,3,2,FALSE)', 'R2C[3]'], ['=ADDRESS(2,3,3,FALSE)', 'R[2]C3'], ['=ADDRESS(2,3,4,FALSE)', 'R[2]C[3]'], ['=ADDRESS(2,3,1,FALSE,"Other")', 'Other!R2C3'], ['=ADDRESS(2,3)', '$C$2']];
  cases.forEach(([f], i) => s.setCell('S', i, 0, f));
  s.recalculate?.();
  assert.deepEqual(cases.map((_, i) => s.getValue('S', i, 0)), cases.map(([, want]) => want));
});

test('a whole row is read as one, typed or in a file', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setCell('S', 0, 1, 7);
  s.setCell('S', 0, 2, 5);
  s.setCell('S', 2, 0, '=SUM(1:1)');
  s.setCell('S', 2, 1, '=SUM($1:$1)');
  s.setCell('S', 3, 1, '=@1:1');
  s.recalculate?.();
  assert.equal(s.getValue('S', 2, 0), 12);
  assert.equal(s.getValue('S', 2, 1), 12);
  assert.equal(s.getValue('S', 3, 1), 7, 'the cell of row 1 in line with it');
});

test('SUMIF and COUNTIF with a range as their condition are read as Excel reads them, one value in line', () => {
  const rows = [[10, 1], [20, 2], [30, 3], [null, null, '=SUMIF(A1:A3,A1:A3)'], [null, null, '=COUNTIF(A1:A3,A1:A3)']];
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'S', rows }] }));
  assert.equal(view.editValue(3, 2), '=SUMIF(A1:A3,@A1:A3)');
  assert.equal(view.editValue(4, 2), '=COUNTIF(A1:A3,@A1:A3)');
});
