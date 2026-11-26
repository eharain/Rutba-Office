/**
 * Conditional formatting: top and bottom rules, above and below average, a
 * rule written as a formula, and the rules seen and removed one by one.
 * Before 2.28 the dialog had none of the first three, and Clear rules took
 * every rule off the sheet with no way to see or remove one.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { CONDITION_RULES, ruleSpec, ruleReady } from '../apps/desktop/renderer/apps/sheets/rule-specs.js';

const book = () => buildXlsx({
  sheets: [{ name: 'Data', rows: [['Ink', 5], ['Paper', 12], ['Pens', 12], ['Tape', 30], ['Glue', 7]] }],
});

/** The view with B1:B5 selected. */
const overFigures = () => {
  const view = SheetView.open(book());
  view.select(0, 1);
  view.select(4, 1, { extend: true });
  return view;
};
const painted = (view, col = 'B') => {
  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  return [1, 2, 3, 4, 5].filter((r) => cells.get(col + r)?.style?.fill).map((r) => col + r);
};
const sheetXml = (view) => Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');

test('the dialog offers top and bottom, average and formula rules, each one the engine writes', () => {
  const ids = CONDITION_RULES.map((c) => c.id);
  for (const id of ['top10', 'top10Percent', 'bottom10', 'bottom10Percent', 'aboveAverage', 'belowAverage', 'formula']) {
    assert.ok(ids.includes(id), id);
    const view = overFigures();
    assert.doesNotThrow(() => view.addConditionalRule(ruleSpec({ rule: id, a: id === 'formula' ? '=B1>10' : '', style: 0 })), id);
  }
  assert.deepEqual(ruleSpec({ rule: 'top10', a: '', style: 0 }), { kind: 'top10', rank: 10, bottom: false, percent: false, look: 'red' });
  assert.deepEqual(ruleSpec({ rule: 'bottom10Percent', a: '25', style: 2 }), { kind: 'top10', rank: 25, bottom: true, percent: true, look: 'green' });
  assert.deepEqual(ruleSpec({ rule: 'belowAverage', style: 1 }), { kind: 'aboveAverage', below: true, look: 'yellow' });
  assert.deepEqual(ruleSpec({ rule: 'formula', a: '=$B1>10', style: 0 }), { kind: 'expression', formula: '=$B1>10', look: 'red' });
});

test('the dialog knows when a rule is filled: a count may be empty, a formula may not', () => {
  assert.equal(ruleReady({ rule: 'top10', a: '' }), true);
  assert.equal(ruleReady({ rule: 'top10', a: '3' }), true);
  assert.equal(ruleReady({ rule: 'top10', a: '0' }), false);
  assert.equal(ruleReady({ rule: 'top10', a: '2.5' }), false);
  assert.equal(ruleReady({ rule: 'top10Percent', a: '101' }), false);
  assert.equal(ruleReady({ rule: 'aboveAverage' }), true);
  assert.equal(ruleReady({ rule: 'formula', a: '  ' }), false);
  assert.equal(ruleReady({ rule: 'formula', a: '=B1>10' }), true);
  assert.equal(ruleReady({ rule: 'between', a: '1', b: '' }), false);
  assert.equal(ruleReady({ rule: 'duplicateValues' }), true);
});

test('the top two paints the largest, ties kept, and is written as Excel writes it', () => {
  const view = overFigures();
  view.addConditionalRule(ruleSpec({ rule: 'top10', a: '2', style: 0 }));
  assert.deepEqual(painted(view), ['B2', 'B3', 'B4'], '30, and both 12s');
  assert.match(sheetXml(view), /<conditionalFormatting sqref="B1:B5"><cfRule type="top10" dxfId="0" priority="1" rank="2"\/><\/conditionalFormatting>/);
  assert.deepEqual(painted(SheetView.open(view.save())), ['B2', 'B3', 'B4'], 'and the same when opened again');
});

test('the bottom 40 per cent paints the smallest two of five', () => {
  const view = overFigures();
  view.addConditionalRule(ruleSpec({ rule: 'bottom10Percent', a: '40', style: 2 }));
  assert.deepEqual(painted(view), ['B1', 'B5'], '5 and 7');
  assert.match(sheetXml(view), /<cfRule type="top10" dxfId="0" priority="1" percent="1" bottom="1" rank="40"\/>/);
});

test('above and below average', () => {
  const above = overFigures();
  above.addConditionalRule(ruleSpec({ rule: 'aboveAverage', style: 0 }));
  assert.deepEqual(painted(above), ['B4'], 'the mean is 13.2: only 30 is above');
  assert.match(sheetXml(above), /<cfRule type="aboveAverage" dxfId="0" priority="1"\/>/);
  const below = overFigures();
  below.addConditionalRule(ruleSpec({ rule: 'belowAverage', style: 0 }));
  assert.deepEqual(painted(below), ['B1', 'B2', 'B3', 'B5']);
  assert.match(sheetXml(below), /<cfRule type="aboveAverage" dxfId="0" priority="1" aboveAverage="0"\/>/);
});

test('a count out of range is refused and leaves no undo step', () => {
  const view = overFigures();
  assert.throws(() => view.addConditionalRule({ kind: 'top10', rank: 0, look: 'red' }), /count from 1 to 1000/);
  assert.throws(() => view.addConditionalRule({ kind: 'top10', rank: 150, percent: true, look: 'red' }), /percentage from 1 to 100/);
  assert.equal(view.canUndo, false);
});

test('a rule from a formula paints the whole row of each cell it is true for', () => {
  const view = SheetView.open(book());
  view.select(0, 0);
  view.select(4, 1, { extend: true });
  // Written for the first cell: the column is held by its dollar sign, the row moves.
  view.addConditionalRule(ruleSpec({ rule: 'formula', a: '=$B1>10', style: 1 }));
  assert.deepEqual(painted(view, 'A'), ['A2', 'A3', 'A4']);
  assert.deepEqual(painted(view, 'B'), ['B2', 'B3', 'B4']);
  assert.match(sheetXml(view), /<conditionalFormatting sqref="A1:B5"><cfRule type="expression" dxfId="0" priority="1"><formula>\$B1&gt;10<\/formula><\/cfRule><\/conditionalFormatting>/);
  const again = SheetView.open(view.save());
  assert.deepEqual(painted(again, 'A'), ['A2', 'A3', 'A4']);
  // The rule follows the cells: a figure changed changes what is painted.
  again.setCell(0, 1, 99);
  assert.deepEqual(painted(again, 'A'), ['A1', 'A2', 'A3', 'A4']);
});

test('a formula that is empty or cannot be read is refused, and nothing is written', () => {
  const view = overFigures();
  assert.throws(() => view.addConditionalRule({ kind: 'expression', formula: ' ', look: 'red' }), /needs a formula/);
  assert.throws(() => view.addConditionalRule({ kind: 'expression', formula: '=B1>', look: 'red' }), /cannot be read/);
  assert.equal(view.canUndo, false);
  assert.doesNotMatch(sheetXml(view), /cfRule/);
});

test('the rules are listed, the first to be judged first, each with what it tests and its look', () => {
  const view = overFigures();
  view.addConditionalRule(ruleSpec({ rule: 'top10', a: '2', style: 0 }));
  view.addConditionalRule(ruleSpec({ rule: 'formula', a: '=B1<6', style: 2 }));
  view.addConditionalRule(ruleSpec({ rule: 'dataBarBlue' }));
  const rules = view.render().conditionalRules;
  assert.deepEqual(rules.map((r) => [r.priority, r.type, r.sqref]), [[1, 'top10', 'B1:B5'], [2, 'expression', 'B1:B5'], [3, 'dataBar', 'B1:B5']]);
  assert.equal(rules[0].rank, 2);
  assert.equal(rules[0].bottom, false);
  assert.match(rules[0].fill, /^#/, 'the look a rule gives comes with it');
  assert.deepEqual(rules[1].formulas, ['B1<6']);
  assert.equal(rules[2].fill, null);
  assert.equal(rules[2].colours.length, 1, 'a data bar has its own colour');
});

test('one rule is removed and the others stay; undo puts it back', () => {
  const view = overFigures();
  view.addConditionalRule(ruleSpec({ rule: 'top10', a: '1', style: 0 }));
  view.addConditionalRule(ruleSpec({ rule: 'formula', a: '=B1<6', style: 2 }));
  assert.deepEqual(painted(view), ['B1', 'B4']);
  const [top, formula] = view.conditionalRules();
  view.removeConditionalRule(top);
  assert.deepEqual(painted(view), ['B1'], 'only the formula rule is left');
  assert.deepEqual(view.conditionalRules().map((r) => r.type), ['expression']);
  const xml = sheetXml(view);
  assert.doesNotMatch(xml, /type="top10"/);
  assert.match(xml, /type="expression"/);

  view.undo();
  assert.deepEqual(painted(view), ['B1', 'B4'], 'one undo puts the rule back');
  view.redo();
  view.removeConditionalRule(formula);
  assert.deepEqual(view.conditionalRules(), []);
  assert.doesNotMatch(sheetXml(view), /conditionalFormatting/, 'the last rule of a range takes its block with it');
  assert.throws(() => view.removeConditionalRule(formula), /no longer on the sheet/);
});

test('two rules written in one block: removing one leaves the other in the block', () => {
  const wb = Workbook.open(book());
  wb.addConditionalFormatting('Data', 'B1:B5',
    '<cfRule type="top10" dxfId="0" priority="1" rank="1"/><cfRule type="aboveAverage" dxfId="0" priority="2"/>');
  const view = SheetView.open(wb.save());
  assert.equal(view.conditionalRules().length, 2);
  view.removeConditionalRule({ sqref: 'B1:B5', priority: 1 });
  assert.match(sheetXml(view), /<conditionalFormatting sqref="B1:B5"><cfRule type="aboveAverage" dxfId="0" priority="2"\/><\/conditionalFormatting>/);
});

test('a data bar from Excel leaves no half behind: its 2010 partner goes with it', () => {
  const wb = Workbook.open(book());
  const id = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}';
  wb.addConditionalFormatting('Data', 'B1:B5',
    '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>'
    + '<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>' + id + '</x14:id></ext></extLst></cfRule>');
  wb._sheetPart('Data').part.setTailElement('extLst',
    '<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">'
    + '<x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    + '<x14:cfRule type="dataBar" id="' + id + '"><x14:dataBar minLength="0" maxLength="100" border="1"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule>'
    + '<xm:sqref>B1:B5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>');
  const view = SheetView.open(wb.save());
  assert.match(sheetXml(view), /x14:cfRule/);
  const [bar] = view.conditionalRules();
  assert.equal(bar.type, 'dataBar');
  view.removeConditionalRule(bar);
  const xml = sheetXml(view);
  assert.doesNotMatch(xml, /cfRule/);
  assert.doesNotMatch(xml, /extLst/, 'the emptied extension list goes too');
});

test('a protected sheet keeps its rules', () => {
  const view = overFigures();
  view.addConditionalRule(ruleSpec({ rule: 'aboveAverage', style: 0 }));
  view.protect();
  assert.throws(() => view.removeConditionalRule(view.conditionalRules()[0]), /protected/);
  assert.equal(view.conditionalRules().length, 1);
});

test('one of two data bars removed: the other keeps its 2010 half, which Excel holds in the same block', () => {
  const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'Data', rows: [[1], [2], [3]] }] }));
  const part = wb.partNameFor('Data');
  const bar = (n) => '<cfRule type="dataBar" priority="' + n + '"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>'
    + '<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{0000000' + n + '-0000-0000-0000-000000000000}</x14:id></ext></extLst></cfRule>';
  const half = (n) => '<x14:cfRule type="dataBar" id="{0000000' + n + '-0000-0000-0000-000000000000}"><x14:dataBar minLength="0" maxLength="100" negativeBarColorSameAsPositive="1"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule>';
  const ext = '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings>'
    + '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' + half(1) + half(2) + '<xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting>'
    + '</x14:conditionalFormattings></ext></extLst>';
  wb.pkg.write_(part, wb.pkg.text(part).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1:A3">' + bar(1) + bar(2) + '</conditionalFormatting>').replace('</worksheet>', ext + '</worksheet>'));
  const back = Workbook.open(wb.save());
  assert.equal(back.removeConditionalRule('Data', 'A1:A3', 1), true);
  const xml = Workbook.open(back.save()).pkg.text(part);
  assert.doesNotMatch(xml, /\{00000001-/, 'the rule removed is gone, both halves');
  assert.equal((xml.match(/<x14:cfRule\b/g) || []).length, 1, 'the other\'s 2010 half is kept');
  assert.match(xml, /<x14:cfRule type="dataBar" id="\{00000002-/);
});
