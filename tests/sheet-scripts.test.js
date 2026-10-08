/**
 * Automate → scripts: a script in the shape of Office Scripts run on a
 * snapshot of the workbook — reading what is there, writing values,
 * formulas and formats, adding and renaming sheets — its edits applied to
 * the workbook as one undo step; and a script that fails says why and
 * changes nothing.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { runScript, parseAddress, addressOf } from '@rutba/sheet-view/scripts';

const book = () => new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: [['Item', 'Qty', 'Price'], ['Ink', 4, 2.5], ['Paper', 6, 1.25]] }] }));

test('addresses are read as Office Scripts reads them', () => {
  assert.deepEqual(parseAddress('B2', 'S'), { sheet: 'S', top: 1, left: 1, bottom: 1, right: 1 });
  assert.deepEqual(parseAddress("'My Sheet'!$A$1:C3", 'S'), { sheet: 'My Sheet', top: 0, left: 0, bottom: 2, right: 2 });
  assert.equal(addressOf({ top: 0, left: 0, bottom: 2, right: 2 }), 'A1:C3');
  assert.throws(() => parseAddress('nonsense!!', 'S'), /not a cell or a range/);
});

test('a script reads the workbook, writes values, formulas and formats and adds a sheet, all one undo step', () => {
  const view = book();
  const code = `function main(workbook) {
    const sheet = workbook.getActiveWorksheet();
    const rows = sheet.getUsedRange().getValues();
    console.log('rows', rows.length);
    sheet.getRange('D1').setValue('Total');
    for (let r = 2; r <= rows.length; r++) sheet.getRange('D' + r).setFormula('=B' + r + '*C' + r);
    sheet.getRange('A1:D1').getFormat().getFont().setBold(true);
    sheet.getRange('A1:D1').getFormat().getFill().setColor('yellow');
    const summary = workbook.addWorksheet('Summary');
    summary.getRange('A1:B1').setValues([['Lines', rows.length - 1]]);
  }`;
  const out = runScript(code, view.scriptSnapshot());
  assert.equal(out.error, null);
  assert.deepEqual(out.logs, ['rows 3']);
  const cells = view.applyScriptEdits(out.edits);
  assert.equal(cells, 5, 'the heading, two formulas and the two summary cells');
  const text = (r, c) => view.displayValue(r, c).text;
  view.selectSheet('Sales');
  assert.deepEqual([0, 1, 2].map((r) => text(r, 3)), ['Total', '10', '7.5']);
  view.select(0, 0);
  assert.equal(view.formatState().bold, true);
  assert.match(String(view.formatState().fill), /FFFF00$/i, 'yellow, in whichever form the styles keep it');
  view.selectSheet('Summary');
  assert.equal(text(0, 1), '2');
  view.undo();
  assert.deepEqual(view.sheetNames(), ['Sales'], 'one undo takes every change back');
  view.selectSheet('Sales');
  assert.equal(text(0, 3), '');
});

test('a script that fails says why and changes nothing; one without main is refused', () => {
  const view = book();
  const bad = runScript(`function main(workbook) { workbook.getActiveWorksheet().getRange('A1').setValue('x'); workbook.getWorksheet('Nope').getRange('A1'); }`, view.scriptSnapshot());
  assert.equal(bad.edits.length, 0);
  assert.match(bad.error, /Cannot read properties of undefined|not a function/);
  assert.match(runScript('const x = 1;', view.scriptSnapshot()).error, /needs a function called main/);
  assert.match(runScript('function main(w) { w.getActiveWorksheet().getRange("A1:B2").setValues([[1]]); }', view.scriptSnapshot()).error, /needs 2 rows of 2 values/);
  assert.equal(view.displayValue(0, 0).text, 'Item');
});

test('reads see the script\'s own writes over the workbook as it was', () => {
  const view = book();
  const out = runScript(`function main(workbook) {
    const s = workbook.getActiveWorksheet();
    s.getRange('E1').setValue(41);
    console.log(s.getRange('E1').getValue() + 1, s.getRange('B2').getValue(), s.getRange('A1:C1').getValues().length);
    s.setName('Takings');
  }`, view.scriptSnapshot());
  assert.deepEqual(out.logs, ['42 4 1']);
  view.applyScriptEdits(out.edits);
  assert.deepEqual(view.sheetNames(), ['Takings']);
});
