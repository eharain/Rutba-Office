// Data Validation acting on what is typed.
//
// A rule made in the dialog was written to the file and then did nothing:
// no entry was checked, and the frame's list of choices had no reader. Now a
// committed entry is judged by the cell's rule. One the rule refuses does not
// go in and the selection stays on the cell; a warning or an information
// rule lets it in and says so; a rule with its alert off judges nothing. The
// dialog also makes a rule for dates, and a list can come from another sheet.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { readDatesAs } from '@rutba/formula';
import { VALIDATION_KINDS, validationSpec } from '../apps/desktop/renderer/apps/sheets/rule-specs.js';

const BOOK = () => buildXlsx({ sheets: [
  { name: 'Orders', rows: [['Drink', 'When'], ['Tea', ''], ['', '']] },
  { name: 'Products', rows: [['Product'], ['Pens'], ['Ink'], ['Clips']] },
  { name: 'Price list', rows: [['Item'], ['Paper'], ['Card']] },
] });

const type = (view, row, col, text, move = 'down') => {
  view.select(row, col);
  view.beginEdit({ replace: true, initial: '' });
  view.updateDraft(text);
  view.commitEdit({ move });
  return view.refusal;
};
const input = (view, row, col) => view.calc.getInput(view.activeSheet, row, col);

/** The view with a rule written straight into the sheet, as a file made elsewhere holds one. */
const withRule = (entry) => {
  const view = new SheetView(BOOK());
  view.workbook.addDataValidation('Orders', entry);
  view._structuralDirty = true;
  return SheetView.open(view.save());
};

test.beforeEach(() => readDatesAs({ dayFirst: true, year: 2026 }));
test.after(() => readDatesAs());

test('a list rule refuses an entry that is not on it, and the cell and the selection stay as they were', () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  view.addValidationRule(validationSpec({ kind: 'list', list: 'Tea, Coffee, Milk', prompt: 'Pick a drink' }));

  const said = type(view, 1, 0, 'Bananas');
  assert.equal(said.style, 'stop');
  assert.deepEqual(said.options, ['Tea', 'Coffee', 'Milk']);
  assert.equal(said.ref, 'A2');
  assert.match(said.said, /pick one of: Tea, Coffee, Milk/);
  assert.equal(input(view, 1, 0), 'Tea', 'the cell keeps what it held');
  assert.equal(view.editing, null, 'the edit is over');
  assert.deepEqual([view.selection.active.row, view.selection.active.col], [1, 0], 'the selection has not moved on');
  assert.equal(view.canUndo, true, 'only the rule is there to undo');

  assert.equal(type(view, 1, 0, 'coffee'), null, 'an entry on the list goes in, whatever its capitals');
  assert.equal(input(view, 1, 0), 'coffee');
  assert.deepEqual([view.selection.active.row, view.selection.active.col], [2, 0], 'and Enter moves on');
});

test('the selected cell carries its list for the arrow, and the message its rule shows', () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  view.addValidationRule(validationSpec({ kind: 'list', list: 'Tea, Coffee, Milk', prompt: 'Pick a drink' }));
  const frame = view.selectionFrame();
  assert.deepEqual(frame.validation.list, ['Tea', 'Coffee', 'Milk']);
  assert.equal(frame.validation.prompt, 'Pick a drink');
  assert.deepEqual(view.render().validation.list, ['Tea', 'Coffee', 'Milk'], 'in the whole frame too');
  view.select(2, 0);
  assert.equal(view.selectionFrame().validation, null, 'a cell with no rule has neither');

  const quiet = withRule('<dataValidation type="list" allowBlank="1" showInputMessage="0" showErrorMessage="1" prompt="Hidden" sqref="A2"><formula1>"Yes,No"</formula1></dataValidation>');
  quiet.select(1, 0);
  assert.equal(quiet.selectionFrame().validation.prompt, null, 'a message switched off is not shown');
  assert.deepEqual(quiet.selectionFrame().validation.list, ['Yes', 'No']);
});

test('a list can come from another sheet, and is written as Excel writes it', () => {
  assert.deepEqual(validationSpec({ kind: 'list', list: 'Products!A2:A4' }), { kind: 'list', source: 'Products!A2:A4' });
  assert.deepEqual(validationSpec({ kind: 'list', list: "='Price list'!$A$2:$A$3" }), { kind: 'list', source: "'Price list'!$A$2:$A$3" });
  assert.deepEqual(validationSpec({ kind: 'list', list: '=Drinks' }), { kind: 'list', source: 'Drinks' });
  assert.deepEqual(validationSpec({ kind: 'list', list: 'Products, Services' }), { kind: 'list', items: ['Products', 'Services'] }, 'words with a comma are still a typed list');

  const view = new SheetView(BOOK());
  view.select(1, 0);
  view.addValidationRule(validationSpec({ kind: 'list', list: 'Products!A2:A4' }));
  assert.deepEqual(view.selectionFrame().validation.list, ['Pens', 'Ink', 'Clips']);
  assert.equal(type(view, 1, 0, 'Staples').style, 'stop');
  assert.equal(type(view, 1, 0, 'Ink'), null);

  view.select(2, 0);
  view.addValidationRule(validationSpec({ kind: 'list', list: "'Price list'!A2:A3" }));
  assert.deepEqual(view.selectionFrame().validation.list, ['Paper', 'Card']);

  const again = SheetView.open(view.save());
  const xml = again.pkg.text(again.workbook.partNameFor('Orders'));
  assert.match(xml, /<dataValidation type="list"[^>]*sqref="A2"><formula1>Products!A2:A4<\/formula1><\/dataValidation>/);
  assert.match(xml, /<formula1>'Price list'!A2:A3<\/formula1>/);
  again.select(1, 0);
  assert.deepEqual(again.selectionFrame().validation.list, ['Pens', 'Ink', 'Clips'], 'and acts the same once the file is opened again');

  assert.throws(() => view.addValidationRule({ kind: 'list', source: 'Nowhere!A1:A3' }), /no sheet called Nowhere/);
  assert.throws(() => view.addValidationRule({ kind: 'list', source: 'Drinks' }), /must be a range/, 'a name the workbook does not define');
});

test('a list can come from a name the workbook defines', () => {
  const view = new SheetView(BOOK());
  view.selectSheet('Products');
  view.select(1, 0);
  view.selection.extendTo(3, 0);
  view.defineName('Stock', 'Products!$A$2:$A$4');
  view.selectSheet('Orders');
  view.select(1, 0);
  view.addValidationRule(validationSpec({ kind: 'list', list: '=Stock' }));
  assert.deepEqual(view.selectionFrame().validation.list, ['Pens', 'Ink', 'Clips']);
});

test('the dialog offers dates: typed bounds are kept as days, and a typed date is judged as one', () => {
  assert.ok(VALIDATION_KINDS.includes('date'));
  const view = new SheetView(BOOK());
  view.select(1, 1);
  view.addValidationRule(validationSpec({ kind: 'date', min: '01/11/2026', max: '30/11/2026' }));
  const rule = view.validationAt(1, 1);
  assert.deepEqual([rule.type, rule.formula1, rule.formula2], ['date', '46327', '46356']);

  assert.equal(type(view, 1, 1, '15/11/2026'), null);
  assert.equal(input(view, 1, 1), 46341, 'the date went in as a date');
  assert.equal(view.displayValue(1, 1).text, '15/11/2026');
  const late = type(view, 1, 1, '15/12/2026');
  assert.equal(late.style, 'stop');
  assert.match(late.said, /between 1 Nov 2026 and 30 Nov 2026/, 'the bounds are said as dates');
  assert.equal(type(view, 1, 1, 'soon').style, 'stop', 'words are not a date');
  assert.equal(input(view, 1, 1), 46341);

  view.select(2, 1);
  view.addValidationRule(validationSpec({ kind: 'date', min: '1 Nov 2026', max: '' }));
  assert.equal(view.validationAt(2, 1).operator, 'greaterThanOrEqual');
  assert.throws(() => view.addValidationRule(validationSpec({ kind: 'date', min: 'soon', max: '' })), /needs a date/);
  assert.match(SheetView.open(view.save()).pkg.text(view.workbook.partNameFor('Orders')), /<dataValidation type="date"[^>]*sqref="B2"><formula1>46327<\/formula1><formula2>46356<\/formula2>/);
});

test('a warning rule lets the entry in and says so; the author\'s own words are carried', () => {
  const view = withRule('<dataValidation type="whole" errorStyle="warning" allowBlank="1" showErrorMessage="1" errorTitle="Check" error="Usually 1 to 10." sqref="A3"><formula1>1</formula1><formula2>10</formula2></dataValidation>');
  const said = type(view, 2, 0, '25');
  assert.deepEqual([said.style, said.title, said.message], ['warning', 'Check', 'Usually 1 to 10.']);
  assert.equal(input(view, 2, 0), 25, 'the entry went in');
  assert.equal(type(view, 2, 0, '7'), null);
});

test('a rule with its alert switched off judges nothing', () => {
  const view = withRule('<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="A3"><formula1>"Yes,No"</formula1></dataValidation>');
  assert.equal(type(view, 2, 0, 'Perhaps'), null);
  assert.equal(input(view, 2, 0), 'Perhaps');
});

test('an entry left as it was is not judged, and an emptied cell is allowed', () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  view.addValidationRule(validationSpec({ kind: 'list', list: 'Coffee, Milk' }));
  view.select(1, 0);
  view.beginEdit();
  view.commitEdit({ move: 'none' });
  assert.equal(view.refusal, null, 'Tea was there before the rule; opening and closing the cell changes nothing');
  assert.equal(input(view, 1, 0), 'Tea');
  assert.equal(type(view, 1, 0, ''), null);
  assert.equal(view.editValue(1, 0), '');
});

test('a rule whose file leaves its alert out stops nothing, as Excel reads it (the schema says off)', () => {
  const view = withRule('<dataValidation type="whole" sqref="A3"><formula1>1</formula1><formula2>10</formula2></dataValidation>');
  assert.equal(type(view, 2, 0, '50'), null, 'nothing refused');
  assert.equal(input(view, 2, 0), 50);
});
