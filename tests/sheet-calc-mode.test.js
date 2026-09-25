/**
 * Formulas → Calculation Options and Calculate Now / Calculate Sheet.
 *
 * Excel keeps the mode on `workbook.xml` as `<calcPr calcMode="manual"/>`
 * (or "autoNoTable"; automatic is the attribute's absence). In manual mode
 * an edit calculates a formula typed into the edited cell and nothing that
 * depends on the edit; the status bar says "Calculate" until F9 (every
 * sheet) or Shift+F9 (this sheet) brings the formulas up to date.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

const open = () => new SheetView(buildXlsx({
  sheets: [
    { name: 'One', rows: [[1, '=A1*2'], [5, '=SUM(A1:A2)']] },
    { name: 'Two', rows: [['=One!A1+100']] },
  ],
}));
const workbookXml = (view) => view.pkg.text(view.workbook.mainPart);
const value = (view, sheet, a1) => view.calc.getByRef(sheet, a1);

test('the mode is written to calcPr and read back; automatic takes the attribute away', () => {
  const view = open();
  assert.equal(view.calcMode(), 'auto');
  view.setCalcMode('manual');
  assert.match(workbookXml(view), /<calcPr\b[^>]*\bcalcMode="manual"/);
  const again = new SheetView(view.save());
  assert.equal(again.calcMode(), 'manual');
  assert.equal(again.calc.manual, true, 'a workbook saved in manual mode opens in manual mode');
  again.setCalcMode('autoNoTable');
  assert.match(workbookXml(again), /calcMode="autoNoTable"/);
  again.setCalcMode('auto');
  assert.doesNotMatch(workbookXml(again), /calcMode=/);
  assert.throws(() => again.setCalcMode('sometimes'), /auto, autoNoTable or manual/);
});

test('calcPr goes where the schema puts it — before extLst, after definedNames', () => {
  const view = open();
  const main = view.workbook.mainPart;
  view.pkg.write_(main, view.pkg.text(main).replace('</workbook>', '<extLst><ext uri="{x}"/></extLst></workbook>'));
  view.setCalcMode('manual');
  const xml = workbookXml(view);
  assert.ok(xml.indexOf('<calcPr') < xml.indexOf('<extLst'), 'calcPr before extLst');
  assert.ok(xml.indexOf('</sheets>') < xml.indexOf('<calcPr'), 'calcPr after sheets');
});

test('in manual mode an edit waits for Calculate Now, and a typed formula shows its answer', () => {
  const view = open();
  view.setCalcMode('manual');
  assert.equal(view.calcState().pending, false);
  view.select(0, 0);
  view.setCell(0, 0, '10');
  assert.equal(value(view, 'One', 'A1'), 10);
  assert.equal(value(view, 'One', 'B1'), 2, 'the dependent keeps its old answer');
  assert.equal(value(view, 'One', 'B2'), 6);
  assert.equal(view.calcState().pending, true, 'the status bar says Calculate');
  assert.deepEqual(view.render().calc, { mode: 'manual', pending: true });

  view.setCell(2, 0, '=A1+A2');
  assert.equal(value(view, 'One', 'A3'), 15, 'a new formula is calculated as it is entered');

  const n = view.calculate();
  assert.ok(n >= 3);
  assert.equal(value(view, 'One', 'B1'), 20);
  assert.equal(value(view, 'One', 'B2'), 15);
  assert.equal(value(view, 'Two', 'A1'), 110, 'Calculate Now reaches every sheet');
  assert.equal(view.calcState().pending, false);
});

test('Calculate Sheet works out this sheet only; the rest waits', () => {
  const view = open();
  view.setCalcMode('manual');
  view.setCell(0, 0, '7');
  view.calculate({ scope: 'sheet' });
  assert.equal(value(view, 'One', 'B1'), 14);
  assert.equal(value(view, 'Two', 'A1'), 101, 'the other sheet waits');
  assert.equal(view.calcState().pending, true);
  view.calculate();
  assert.equal(value(view, 'Two', 'A1'), 107);
  assert.equal(view.calcState().pending, false);
});

test('back to automatic calculates what waited; undo puts manual back', () => {
  const view = open();
  view.setCalcMode('manual');
  view.setCell(0, 0, '3');
  assert.equal(value(view, 'One', 'B1'), 2);
  view.setCalcMode('auto');
  assert.equal(value(view, 'One', 'B1'), 6);
  assert.equal(view.calc.manual, false);
  view.undo();
  assert.equal(view.calcMode(), 'manual');
  assert.equal(view.calc.manual, true);
});

test('saving in manual mode calculates first, as Excel does; a recovery draft does not', () => {
  const view = open();
  view.setCalcMode('manual');
  view.setCell(0, 0, '4');
  view.serialize();
  assert.equal(value(view, 'One', 'B1'), 2, 'the draft left the waiting alone');
  const bytes = view.save();
  assert.equal(value(view, 'One', 'B1'), 8);
  const pkg = OoxmlPackage.read(bytes);
  const sheet = pkg.text('xl/worksheets/sheet1.xml');
  assert.match(sheet, /<c r="B1"[^>]*>.*?<v>8<\/v>/s, 'the saved cache is the calculated answer');
});

test('in automatic mode nothing waits', () => {
  const view = open();
  view.setCell(0, 0, '9');
  assert.equal(value(view, 'One', 'B1'), 18);
  assert.equal(view.calcState().pending, false);
});
