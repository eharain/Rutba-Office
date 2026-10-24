/**
 * Conditional formatting and data validation from their dialogs: every
 * choice each offers reaches the engine as a rule the engine writes. Before
 * 2.18 the dialogs sent `{ type, format }` and `{ type, formula1 }` where the
 * engine reads `{ kind, look }` and `{ kind, values }`, and every rule was
 * refused.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { CONDITION_RULES, LOOKS, ruleSpec, VALIDATION_KINDS, validationSpec } from '../apps/desktop/renderer/apps/sheets/rule-specs.js';

const book = () => buildXlsx({ sheets: [{ name: 'Data', rows: [[5], [12], [12], [30], [7]] }] });

const overColumn = () => {
  const view = SheetView.open(book());
  view.select(0, 0);
  view.select(4, 0, { extend: true });
  return view;
};

test('every condition the dialog offers is a rule the engine writes', () => {
  for (const c of CONDITION_RULES) {
    const view = overColumn();
    const spec = ruleSpec({ rule: c.id, a: c.text ? 'x' : '10', b: '20', style: 0 });
    assert.doesNotThrow(() => view.addConditionalRule(spec), c.id);
    assert.match(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /<cfRule /, c.id);
  }
});

test('each look the dialog offers is a look the engine knows, bold among them', () => {
  assert.equal(LOOKS.length, 4);
  const view = overColumn();
  view.addConditionalRule(ruleSpec({ rule: 'greaterThan', a: '10', style: LOOKS.indexOf('bold') }));
  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.equal(cells.get('A4').style.font.bold, true, '30 is over 10, so bold');
  assert.equal(cells.get('A4').style.fill ?? null, null, 'bold only: no fill');
  assert.match(Workbook.open(view.save()).pkg.text('xl/styles.xml'), /<dxf><font><b\/><\/font><\/dxf>/);
});

test('Not equal to and Unique values paint the cells they name', () => {
  const view = overColumn();
  view.addConditionalRule(ruleSpec({ rule: 'notEqual', a: '12', style: 0 }));
  let cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.ok(cells.get('A1').style.fill, '5 is not 12');
  assert.equal(cells.get('A2')?.style?.fill ?? null, null, '12 is 12');

  const other = overColumn();
  other.addConditionalRule(ruleSpec({ rule: 'uniqueValues', style: 2 }));
  cells = new Map(other.render().cells.map((c) => [c.ref, c]));
  assert.ok(cells.get('A1').style.fill, '5 appears once');
  assert.equal(cells.get('A2')?.style?.fill ?? null, null, '12 appears twice');
  assert.match(Workbook.open(other.save()).pkg.text('xl/worksheets/sheet1.xml'), /<cfRule type="uniqueValues" dxfId="0" priority="1"\/>/);
});

test('colour scales and data bars come from the dialog, each in the colours it names', () => {
  const view = overColumn();
  view.addConditionalRule(ruleSpec({ rule: 'colorScaleReverse' }));
  view.addConditionalRule(ruleSpec({ rule: 'dataBarGreen' }));
  const xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<color rgb="FF63BE7B"\/><color rgb="FFFFEB84"\/><color rgb="FFF8696B"\/>/, 'green first: low is good');
  assert.match(xml, /<dataBar><cfvo type="min"\/><cfvo type="max"\/><color rgb="FF63C384"\/><\/dataBar>/);
  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.ok(cells.get('A4').bar, 'the bar is drawn');
  assert.ok(cells.get('A4').style.fill, 'and the scale shades it');
  assert.throws(() => view.addConditionalRule({ kind: 'dataBar', colour: 'mauve' }), /not a data bar colour/);
});

test('every kind of value the validation dialog allows is a rule the engine writes and enforces', () => {
  for (const kind of VALIDATION_KINDS) {
    const view = SheetView.open(book());
    view.select(0, 1);
    view.addValidationRule(validationSpec({ kind, list: 'Draft, Sent', min: '1', max: '4', prompt: 'Pick one' }));
    const xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
    assert.match(xml, new RegExp('<dataValidation type="' + kind + '"[^>]* prompt="Pick one"'), kind);
  }
  const view = SheetView.open(book());
  view.select(0, 1);
  view.addValidationRule(validationSpec({ kind: 'textLength', min: '', max: '3' }));
  assert.match(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /type="textLength" operator="lessThanOrEqual"/);
  assert.ok(view.checkValidation(0, 1, 'longer'), 'six letters is too long');
  assert.equal(view.checkValidation(0, 1, 'ok'), null);

  const ranged = SheetView.open(book());
  ranged.select(0, 1);
  ranged.addValidationRule(validationSpec({ kind: 'list', list: '$A$1:$A$5' }));
  assert.match(Workbook.open(ranged.save()).pkg.text('xl/worksheets/sheet1.xml'), /<formula1>\$A\$1:\$A\$5<\/formula1>/);
});
