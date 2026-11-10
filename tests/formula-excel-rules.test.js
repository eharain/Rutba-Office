// What Excel does that its own answers cannot be asked for: a sheet that is
// not there would open a file dialog, and the clock and recalculation are ours.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Spreadsheet, isError } from '@rutba/formula';
import { FUNCTIONS } from '@rutba/formula/functions';

const book = (now) => new Spreadsheet(now ? { now } : {});

test('a sheet that is not there is #REF!, and so is everything built on it', () => {
  const s = book();
  s.setByRef('One', 'A1', '=Nope!A1+1');
  s.setByRef('One', 'A2', '=SUM(Nope!A1:A3)');
  assert.equal(s.getByRef('One', 'A1').type, '#REF!');
  assert.equal(s.getByRef('One', 'A2').type, '#REF!');
});

test('a sheet written in another case is the same sheet, and an edit there reaches the formula', () => {
  const s = book();
  s.addSheet('Sales');
  s.setByRef('Sales', 'A1', 5);
  s.setByRef('One', 'A1', '=sales!A1*2');
  s.setByRef('One', 'A2', '=SUM(SALES!A1:A2)');
  assert.equal(s.getByRef('One', 'A1'), 10);
  s.setByRef('Sales', 'A1', 7);
  assert.equal(s.getByRef('One', 'A1'), 14);
  assert.equal(s.getByRef('One', 'A2'), 7);
});

test('a sheet named in Cyrillic needs no quotes, and an edit there reaches the formula', () => {
  const s = book();
  s.addSheet('Дані');
  s.setByRef('Дані', 'A1', 3);
  s.setByRef('One', 'A1', '=Дані!A1+1');
  assert.equal(s.getByRef('One', 'A1'), 4);
  s.setByRef('Дані', 'A1', 9);
  assert.equal(s.getByRef('One', 'A1'), 10);
});

test('TODAY and NOW follow this computer, whatever the hour in Greenwich', () => {
  const at = new Date(2026, 5, 15, 3, 30, 0); // 03:30 here
  const s = book(() => at);
  s.setByRef('One', 'A1', '=TODAY()');
  s.setByRef('One', 'A2', '=NOW()-TODAY()');
  s.setByRef('One', 'A3', '=DAY(TODAY())');
  assert.equal(s.getByRef('One', 'A3'), 15);
  assert.ok(Math.abs(s.getByRef('One', 'A2') - 3.5 / 24) < 1e-9);
});

test('INDIRECT reads R1C1, counting brackets from its own cell', () => {
  const s = book();
  s.setByRef('One', 'B2', 42);
  s.setByRef('One', 'C3', '=INDIRECT("R[-1]C[-1]",FALSE)');
  s.setByRef('One', 'C4', '=INDIRECT("R2C2",FALSE)');
  s.setByRef('One', 'C5', '=INDIRECT("One!R2C2",FALSE)');
  s.setByRef('One', 'C6', '=INDIRECT("banana",FALSE)');
  assert.equal(s.getByRef('One', 'C3'), 42);
  assert.equal(s.getByRef('One', 'C4'), 42);
  assert.equal(s.getByRef('One', 'C5'), 42);
  assert.equal(s.getByRef('One', 'C6').type, '#REF!');
});

test('a spilled blank shows 0, and a cleared cell still counts as nothing', () => {
  const s = book();
  s.setByRef('One', 'A1', 1);
  s.setByRef('One', 'B1', '=A1:A2');
  assert.equal(s.getByRef('One', 'B2'), 0);
  s.setByRef('One', 'C1', '=ISBLANK(A2)');
  assert.equal(s.getByRef('One', 'C1'), true);
  s.setByRef('One', 'A2', '');
  assert.equal(s.getByRef('One', 'C1'), true);
});

test('a range keeps the shape it was written with, past the cells in use', () => {
  const s = book();
  s.setByRef('One', 'A1', 1);
  s.addSheet('Long');
  for (let r = 1; r <= 5; r++) s.setByRef('Long', 'A' + r, r);
  s.setByRef('One', 'C1', '=ROWS(A1:A100)');
  s.setByRef('One', 'C2', '=COUNTBLANK(A1:A100)');
  s.setByRef('One', 'C3', '=ROWS(A:A)');
  s.setByRef('One', 'C4', '=COLUMNS(A:B)');
  s.setByRef('One', 'C5', '=COUNTBLANK(D:D)');
  s.setByRef('One', 'C6', '=INDEX(A:A,500)');
  s.setByRef('One', 'C7', '=COUNTIFS(One!A1:A5,1,Long!A1:A5,">0")');
  s.setByRef('One', 'C8', '=COUNTIFS(One!A:A,1,Long!A:A,">0")');
  s.setByRef('One', 'C9', '=SUMPRODUCT(Long!A1:A5,One!E1:E5+1)');
  assert.equal(s.getByRef('One', 'C1'), 100);
  assert.equal(s.getByRef('One', 'C2'), 99);
  assert.equal(s.getByRef('One', 'C3'), 1048576);
  assert.equal(s.getByRef('One', 'C4'), 2);
  assert.equal(s.getByRef('One', 'C5'), 1048576);
  assert.equal(s.getByRef('One', 'C6'), 0);
  assert.equal(s.getByRef('One', 'C7'), 1);
  assert.equal(s.getByRef('One', 'C8'), 1);
  assert.equal(s.getByRef('One', 'C9'), 15);
});

test('a column of 150,000 numbers has a largest and a smallest', () => {
  const list = Array.from({ length: 150000 }, (_, i) => [i]);
  assert.equal(FUNCTIONS.MAX.fn(list), 149999);
  assert.equal(FUNCTIONS.MIN.fn(list), 0);
  assert.equal(FUNCTIONS.MAXA.fn(list), 149999);
  assert.equal(FUNCTIONS.MINA.fn(list), 0);
});

test('AGGREGATE can leave out hidden rows, and SUBTOTAL leaves it out in turn', () => {
  const s = book();
  s.rowHidden = (_sheet, row) => (row === 1 ? 'manual' : null);
  for (const [a1, v] of [['A1', 1], ['A2', 10], ['A3', 100]]) s.setByRef('One', a1, v);
  s.setByRef('One', 'B1', '=AGGREGATE(9,5,A1:A3)');
  s.setByRef('One', 'B2', '=AGGREGATE(9,4,A1:A3)');
  s.setByRef('One', 'C1', '=SUBTOTAL(9,B1:B2)');
  assert.equal(s.getByRef('One', 'B1'), 101);
  assert.equal(s.getByRef('One', 'B2'), 111);
  assert.equal(s.getByRef('One', 'C1'), 0);
  assert.ok(!isError(s.getByRef('One', 'C1')));
});
