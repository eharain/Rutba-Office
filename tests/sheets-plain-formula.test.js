// A formula stored without the array mark answers with one value, as Excel has
// always done. Every answer below is what Excel gave for this same workbook.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '../packages/ooxml/src/build.js';
import { Workbook } from '../packages/ooxml/src/workbook.js';
import { toSpreadsheet, writeCachedValue } from '../packages/ooxml/src/recalc.js';
import { Spreadsheet } from '../packages/formula/src/engine.js';

const rows = [
  [10, 11, 12, null, null, null, null],
  [20, 21, 22, null, '=A1:A3', '=INDIRECT("A1:A3")', '=IF(TRUE,A1:A3,0)'],
  [30, 31, 32, null, '=A1:C1', '=A1:C3', '=A5:A9'],
  [null, null, null, null, '=SUM(A1:A3)', '={7,8,9}', '=A:A'],
  [null, null, null, null, null, '=A1:C1', '=B1:B3'],
];
const opened = (data = rows) => toSpreadsheet(Workbook.open(buildXlsx({ sheets: [{ name: 'S', rows: data }] }))).sheet;
const shown = (v) => (v && typeof v === 'object' ? v.type : v);
const excel = { E2: 20, F2: 20, G2: 20, E3: '#VALUE!', F3: '#VALUE!', G3: '#VALUE!', E4: 60, F4: 7, G4: 0, F5: '#VALUE!', G5: '#VALUE!' };

for (const [ref, want] of Object.entries(excel)) {
  test(`stored with no array mark, ${ref} answers as Excel does`, () => {
    const s = opened();
    assert.equal(shown(s.getValue('S', Number(ref[1]) - 1, ref.charCodeAt(0) - 65)), want);
  });
}

test('a value in the next cell is no reason for #SPILL! when the formula has one answer', () => {
  const s = opened([[1, 2, 3], ['=A1:C1', 9, 9]]);
  assert.equal(shown(s.getValue('S', 1, 0)), 1);
  assert.equal(s.getValue('S', 1, 1), 9);
});

test('such a formula takes over no neighbouring cell', () => {
  const s = opened([[1], [2], [3], [null, '=A1:A3']]);
  assert.equal(shown(s.getValue('S', 3, 1)), '#VALUE!');
  assert.equal(s.spills.size, 0);
});

test('a formula typed here still spills', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setCell('S', 0, 0, 1);
  s.setCell('S', 1, 0, 2);
  s.setCell('S', 0, 2, '=A1:A2');
  s.recalculate?.();
  assert.equal(s.getValue('S', 0, 2), 1);
  assert.equal(s.spillValueAt('S', 1, 2), 2);
});

test('the same words brought back by an undo still mean one value', () => {
  const s = opened([[1, 2, 3], ['=A1:C1', 9, 9]]);
  s.setCell('S', 1, 0, 5);
  s.setCell('S', 1, 0, '=A1:C1');
  s.recalculate?.();
  assert.equal(shown(s.getValue('S', 1, 0)), 1);
});

test('a formula marked as an array in the file still spills', () => {
  const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'S', rows: [[1], [2], [null, '=A1:A2']] }] }));
  writeCachedValue(wb, 'S', 'B3', '=A1:A2', 1, { arrayRef: 'B3:B4' });
  const s = toSpreadsheet(Workbook.open(wb.save())).sheet;
  assert.equal(s.getValue('S', 2, 1), 1);
  assert.equal(s.spillValueAt('S', 3, 1), 2);
});

test('a spill typed here keeps spilling once saved, even when it is one cell or blocked for now', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setCell('S', 0, 0, 1);
  s.setCell('S', 0, 2, '=A1:A1');
  s.setCell('S', 1, 0, 2);
  s.setCell('S', 0, 3, '=A1:A2');
  s.setCell('S', 1, 3, 'in the way');
  s.recalculate?.();
  const cell = (r, c) => s.cell('S', r, c);
  assert.equal(shown(s.getValue('S', 0, 3)), '#SPILL!');
  assert.equal(cell(0, 3).arrayAnswer, true, 'blocked is still an array answer, and is saved with the mark');
  assert.equal(cell(0, 2).single, false);
});
