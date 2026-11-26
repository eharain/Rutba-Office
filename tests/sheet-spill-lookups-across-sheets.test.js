/**
 * XLOOKUP answering with a whole row, COUNTIF and SUMIF given several
 * criteria, and a sum across a run of sheets. Before 2.28 the first was
 * #VALUE!, the second gave one answer in the top cell, and
 * =SUM(Week1:Week4!F8) was #VALUE!.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { buildXlsx } from '@rutba/ooxml/build';
import { moveReferences } from '@rutba/formula/parser';

const shown = (view, r, c) => view.displayValue(r, c).text;
const row = (view, sheet, r, from, to) => {
  const out = [];
  for (let c = from; c <= to; c++) out.push(view.calc.getValue(sheet, r, c));
  return out;
};
const col = (view, sheet, c, from, to) => {
  const out = [];
  for (let r = from; r <= to; r++) out.push(view.calc.getValue(sheet, r, c));
  return out;
};

// ---- XLOOKUP ------------------------------------------------------------------

const marks = () => SheetView.open(buildXlsx({
  sheets: [{ name: 'Marks', rows: [['Name', 'Maths', 'Art', 'Music'], ['Ali', 70, 80, 90], ['Bea', 60, 65, 75], ['Cy', 50, 55, 58]] }],
}));

test('XLOOKUP with a return range of several columns spills the matching row', () => {
  const view = marks();
  view.setCell(6, 0, '=XLOOKUP("Bea",A2:A4,B2:D4)');
  assert.deepEqual(row(view, 'Marks', 6, 0, 2), [60, 65, 75]);
  assert.equal(view.editValue(6, 0), '=XLOOKUP("Bea",A2:A4,B2:D4)', 'no @ is put in front of it');
  // And TRANSPOSE turns that row into a column.
  view.setCell(8, 0, '=TRANSPOSE(XLOOKUP("Cy",A2:A4,B2:D4))');
  assert.deepEqual(col(view, 'Marks', 0, 8, 10), [50, 55, 58]);
  // The file keeps it as a formula that spills.
  const again = SheetView.open(view.save());
  assert.deepEqual(row(again, 'Marks', 6, 0, 2), [60, 65, 75]);
  assert.equal(again.editValue(6, 0), '=XLOOKUP("Bea",A2:A4,B2:D4)');
});

test('XLOOKUP along a heading row spills the column under the match', () => {
  const view = marks();
  view.setCell(1, 6, '=XLOOKUP("Art",B1:D1,B2:D4)');
  assert.deepEqual(col(view, 'Marks', 6, 1, 3), [80, 65, 55]);
});

test('XLOOKUP as it always was: one column gives one cell, no match says so, a wrong size is #VALUE!', () => {
  const view = marks();
  view.setCell(6, 0, '=XLOOKUP("Bea",A2:A4,C2:C4)');
  assert.equal(view.calc.getValue('Marks', 6, 0), 65);
  assert.equal(view.isFilled(6, 1), false);
  view.setCell(7, 0, '=XLOOKUP("Zed",A2:A4,B2:D4,"none")');
  assert.equal(view.calc.getValue('Marks', 7, 0), 'none');
  view.setCell(8, 0, '=XLOOKUP("Zed",A2:A4,B2:D4)');
  assert.equal(shown(view, 8, 0), '#N/A');
  view.setCell(9, 0, '=XLOOKUP("Bea",A2:A4,B2:D3)');
  assert.equal(shown(view, 9, 0), '#VALUE!');
  // INDEX and MATCH still give the row, as before.
  view.setCell(10, 0, '=INDEX(B2:D4,MATCH("Bea",A2:A4,0),0)');
  assert.deepEqual(row(view, 'Marks', 10, 0, 2), [60, 65, 75]);
});

// ---- COUNTIF and SUMIF ------------------------------------------------------------

const orders = () => SheetView.open(buildXlsx({
  sheets: [{ name: 'Orders', rows: [['Name', 'Amount'], ['Ann', 5], ['Bob', 7], ['Ann', 3], ['Cy', 1], ['Bob', 2]] }],
}));

test('COUNTIF and SUMIF given a spill range answer once for each of its cells', () => {
  const view = orders();
  view.setCell(0, 3, '=UNIQUE(A2:A6)');
  assert.deepEqual(col(view, 'Orders', 3, 0, 2), ['Ann', 'Bob', 'Cy']);
  view.setCell(0, 4, '=COUNTIF(A2:A6,D1#)');
  assert.deepEqual(col(view, 'Orders', 4, 0, 2), [2, 2, 1]);
  view.setCell(0, 5, '=SUMIF(A2:A6,D1#,B2:B6)');
  assert.deepEqual(col(view, 'Orders', 5, 0, 2), [8, 9, 1]);
  // A new name in the list: the spill grows, and the counts with it.
  view.setCell(6, 0, 'Dee');
  view.setCell(0, 3, '=UNIQUE(A2:A7)');
  view.setCell(0, 4, '=COUNTIF(A2:A7,D1#)');
  assert.deepEqual(col(view, 'Orders', 4, 0, 3), [2, 2, 1, 1]);
  const again = SheetView.open(view.save());
  assert.deepEqual(col(again, 'Orders', 5, 0, 2), [8, 9, 1], 'and the file opens with the same answers');
});

test('criteria written across spill across; one criterion is still one answer', () => {
  const view = orders();
  view.setCell(8, 0, '=COUNTIF(A2:A6,{"Ann","Cy"})');
  assert.deepEqual(row(view, 'Orders', 8, 0, 1), [2, 1]);
  view.setCell(9, 0, '=COUNTIF(A2:A6,"Bob")');
  assert.equal(view.calc.getValue('Orders', 9, 0), 2);
  assert.equal(view.isFilled(10, 0), false);
  view.setCell(0, 3, 'Ann');
  view.setCell(11, 0, '=SUMIF(A2:A6,D1,B2:B6)');
  assert.equal(view.calc.getValue('Orders', 11, 0), 8);
  view.setCell(12, 0, '=SUMIF(A2:A6,">2")');
  assert.equal(view.calc.getValue('Orders', 12, 0), 0, 'names are not numbers');
  view.setCell(13, 0, '=SUMIF(B2:B6,">2")');
  assert.equal(view.calc.getValue('Orders', 13, 0), 15);
});

// ---- a run of sheets ---------------------------------------------------------------

const weeks = (names = ['Week1', 'Week2', 'Week3']) => SheetView.open(buildXlsx({
  sheets: [
    ...names.map((name, i) => ({ name, rows: [[10 * (i + 1)], [i + 1]] })),
    { name: 'Total', rows: [['']] },
  ],
}));

test('a sum across sheets: =SUM(Week1:Week3!A1)', () => {
  const view = weeks();
  view.selectSheet('Total');
  view.setCell(0, 0, '=SUM(Week1:Week3!A1)');
  assert.equal(view.calc.getValue('Total', 0, 0), 60);
  view.setCell(1, 0, '=AVERAGE(Week1:Week3!A1)');
  view.setCell(2, 0, '=COUNT(Week1:Week3!A1)');
  view.setCell(3, 0, '=MAX(Week1:Week3!A1)');
  view.setCell(4, 0, '=SUM(Week1:Week3!A1:A2)');
  view.setCell(5, 0, '=SUM(Week3:Week1!A1)+1');
  view.setCell(6, 0, '=SUM(Week2:Week3!$A$1,5)');
  assert.deepEqual(col(view, 'Total', 0, 1, 6), [20, 3, 30, 66, 61, 55]);
  assert.equal(view.editValue(0, 0), '=SUM(Week1:Week3!A1)', 'kept as typed');
});

test('a change on one of the sheets is added up at once, and the file keeps the formula', () => {
  const view = weeks();
  view.selectSheet('Total');
  view.setCell(0, 0, '=SUM(Week1:Week3!A1)');
  view.selectSheet('Week2');
  view.setCell(0, 0, 100);
  assert.equal(view.calc.getValue('Total', 0, 0), 140);
  const again = SheetView.open(view.save());
  assert.equal(again.calc.getValue('Total', 0, 0), 140);
  again.selectSheet('Total');
  assert.equal(again.editValue(0, 0), '=SUM(Week1:Week3!A1)');
});

test('sheet names with spaces, in quotes; a sheet that is not there is #REF!; alone in a cell it is #VALUE!', () => {
  const view = weeks(['Wk 1', 'Wk 2', 'Wk 3']);
  view.selectSheet('Total');
  view.setCell(0, 0, "=SUM('Wk 1:Wk 3'!A1)");
  assert.equal(view.calc.getValue('Total', 0, 0), 60);
  view.setCell(1, 0, "=SUM('Wk 1:Nowhere'!A1)");
  assert.equal(shown(view, 1, 0), '#REF!');
  view.setCell(2, 0, "='Wk 1:Wk 3'!A1");
  assert.equal(shown(view, 2, 0), '#VALUE!');
  // A plain range between two cells that look like sheet names is still a range.
  view.setCell(3, 2, 4);
  view.setCell(4, 2, 6);
  view.setCell(5, 0, '=SUM(C4:C5)');
  assert.equal(view.calc.getValue('Total', 5, 0), 10);
});

test('a sheet at either end of the run renamed: the formula follows, quoted when the new name needs it', () => {
  const view = weeks();
  view.selectSheet('Total');
  view.setCell(0, 0, '=SUM(Week1:Week3!A1)');
  view.setCell(1, 0, '=Week1!A1+SUM(Week1:Week3!A2)');
  view.renameSheet('Week2', 'Middle');
  assert.equal(view.editValue(0, 0), '=SUM(Week1:Week3!A1)', 'a sheet inside the run is not named in it');
  view.renameSheet('Week1', 'First');
  assert.equal(view.editValue(0, 0), '=SUM(First:Week3!A1)');
  assert.equal(view.editValue(1, 0), '=First!A1+SUM(First:Week3!A2)');
  assert.equal(view.calc.getValue('Total', 0, 0), 60);
  assert.equal(view.calc.getValue('Total', 1, 0), 16);
  view.renameSheet('Week3', 'Last week');
  assert.equal(view.editValue(0, 0), "=SUM('First:Last week'!A1)");
  assert.equal(view.calc.getValue('Total', 0, 0), 60);
  view.renameSheet('First', 'Wk1');
  assert.equal(view.editValue(0, 0), "=SUM('Wk1:Last week'!A1)");
  view.renameSheet('Last week', 'Wk3');
  assert.equal(view.editValue(0, 0), "=SUM('Wk1:Wk3'!A1)", 'Wk1 reads as the cell WK1: quoted, as Excel writes it');
  assert.equal(view.calc.getValue('Total', 0, 0), 60);
  view.undo();
  assert.equal(view.editValue(0, 0), "=SUM('Wk1:Last week'!A1)");
  assert.equal(view.calc.getValue('Total', 0, 0), 60);
});

test('the run follows the tabs: a sheet moved out of the run is left out', () => {
  const view = weeks();
  view.selectSheet('Total');
  view.setCell(0, 0, '=SUM(Week1:Week2!A1)');
  assert.equal(view.calc.getValue('Total', 0, 0), 30);
  // Week3 moved between Week1 and Week2 joins the run.
  view.moveSheet('Week3', 1);
  assert.equal(view.calc.getValue('Total', 0, 0), 60);
});

test('a sheet at either end of the run deleted: the run ends at the next sheet in it, as Excel has it', () => {
  const view = weeks(['Week1', 'Week2', 'Week3', 'Week4']);
  view.selectSheet('Total');
  view.setCell(0, 0, '=SUM(Week1:Week4!A1)');
  view.setCell(1, 0, '=SUM(Week1:Week2!A1)');
  view.setCell(2, 0, '=Sum(Week3!A1:Week4!A1)');
  view.removeSheet('Week4');
  assert.equal(view.editValue(0, 0), '=SUM(Week1:Week3!A1)');
  assert.equal(view.calc.getValue('Total', 0, 0), 60);
  assert.equal(view.editValue(1, 0), '=SUM(Week1:Week2!A1)', 'a run without it stays');
  view.removeSheet('Week1');
  assert.equal(view.editValue(0, 0), '=SUM(Week2:Week3!A1)');
  assert.equal(view.editValue(1, 0), '=SUM(Week2!A1)', 'a run of one sheet is that sheet');
  assert.equal(view.calc.getValue('Total', 0, 0), 50);
  view.removeSheet('Week2');
  assert.equal(view.editValue(1, 0), '=SUM(#REF!)', 'and with none left it is #REF!');
});

test('a run of sheets named like cells is not moved by a fill: Q1:Q4!B2 down one is Q1:Q4!B3', () => {
  const view = weeks(['Q1', 'Q2', 'Q3', 'Q4']);
  view.selectSheet('Total');
  view.setCell(0, 0, "=SUM('Q1:Q4'!A1)");
  view.setCell(0, 1, '=SUM(Q1:Q4!A1)+Q1!A1');
  view.select(0, 0);
  view.select(0, 1, { extend: true });
  view.fill({ top: 0, left: 0, bottom: 1, right: 1 });
  assert.equal(view.editValue(1, 0), "=SUM('Q1:Q4'!A2)");
  assert.equal(view.editValue(1, 1), '=SUM(Q1:Q4!A2)+Q1!A2');
  assert.equal(view.calc.getValue('Total', 1, 0), 10);
  view.renameSheet('Q4', 'Q5');
  assert.equal(view.editValue(1, 1), "=SUM('Q1:Q5'!A2)+Q1!A2", 'a name that reads as a cell is quoted, as Excel writes it');
});

test('a run deleted at its end whatever the case it was typed in, and a sheet named with a quote', () => {
  const view = weeks(["Bob's", 'Week2', 'Week3']);
  view.selectSheet('Total');
  view.setCell(0, 0, "=SUM('Bob''s:Week3'!A1)");
  view.setCell(1, 0, "='Bob''s'!A1");
  view.setCell(2, 0, '=SUM(week2:WEEK3!A1)');
  view.removeSheet("Bob's");
  assert.equal(view.editValue(0, 0), '=SUM(Week2:Week3!A1)');
  assert.equal(view.calc.getValue('Total', 0, 0), 50);
  assert.equal(view.editValue(1, 0), '=#REF!');
  view.removeSheet('Week3');
  assert.equal(view.editValue(2, 0), '=SUM(Week2!A1)', 'the case it was typed in is no matter');
  assert.equal(view.calc.getValue('Total', 2, 0), 20);
});

test('a cut and paste on the sheet leaves a run of sheets as it is', () => {
  const run = 'SUM(Q1:Q4!B2)+SUM(Week1:Week4!F8)';
  assert.equal(moveReferences(run, { home: 'Week4', from: { sheet: 'Week4', top: 0, left: 0, bottom: 9, right: 20 }, to: { sheet: 'Week4', row: 20, col: 0 } }), run);
  assert.equal(moveReferences('A1+Q1!B2', { home: 'Week4', from: { sheet: 'Week4', top: 0, left: 0, bottom: 0, right: 0 }, to: { sheet: 'Week4', row: 4, col: 0 } }), "A5+'Q1'!B2", 'a plain cell still moves, and a name that reads as a cell is quoted, as Excel writes it');
});
