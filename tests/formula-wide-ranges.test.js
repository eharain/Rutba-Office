/**
 * Ranges read whole: a column of thousands of rows is kept as one range in
 * the dependency graph, its grid is read once per pass, and a long table's
 * first column is indexed for exact lookups. Each shortcut must answer what
 * the slow way would have.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { Spreadsheet, isError, ERR, compareValues } from '@rutba/formula';
import { FUNCTIONS } from '@rutba/formula/functions';
import { firstError } from '../packages/formula/src/values.js';

const ROWS = 5000;

function column() {
  const s = new Spreadsheet();
  s.addSheet('S');
  for (let r = 1; r <= ROWS; r++) {
    s.setByRef('S', 'A' + r, r);
    s.setByRef('S', 'B' + r, 'v' + r);
  }
  return s;
}

test('an edit inside a wide range reaches the formulas reading it', () => {
  const s = column();
  s.setByRef('S', 'D1', '=SUM(A1:A' + ROWS + ')');
  s.recalculate();
  assert.equal(s.getByRef('S', 'D1'), (ROWS * (ROWS + 1)) / 2);
  s.setByRef('S', 'A4000', 0);
  s.recalculate();
  assert.equal(s.getByRef('S', 'D1'), (ROWS * (ROWS + 1)) / 2 - 4000);
});

test('a formula inside a wide range is worked out before the range is read', () => {
  const s = column();
  s.setByRef('S', 'D1', '=SUM(A1:A' + ROWS + ')');
  s.setByRef('S', 'A3000', '=C1*2');
  s.setByRef('S', 'C1', 10);
  s.recalculate();
  const base = (ROWS * (ROWS + 1)) / 2 - 3000;
  assert.equal(s.getByRef('S', 'D1'), base + 20);
  s.setByRef('S', 'C1', 50);
  s.recalculate();
  assert.equal(s.getByRef('S', 'A3000'), 100);
  assert.equal(s.getByRef('S', 'D1'), base + 100, 'the change flows through the formula into the range');
});

test('a wide range that holds its own reader is circular', () => {
  const s = column();
  s.setByRef('S', 'A4500', '=SUM(A1:A' + ROWS + ')');
  const { cycles } = s.recalculate();
  assert.ok(cycles.includes('S!A4500'));
  assert.equal(s.getByRef('S', 'A4500').type, '#CIRCULAR!');
});

test('a lookup into a wide range sees a cell changed since the last pass', () => {
  const s = column();
  s.setByRef('S', 'D1', '=VLOOKUP(4321, A1:B' + ROWS + ', 2, FALSE)');
  s.setByRef('S', 'D2', '=VLOOKUP(4321, A1:B' + ROWS + ', 2, FALSE)&"!"');
  s.recalculate();
  assert.equal(s.getByRef('S', 'D1'), 'v4321');
  s.setByRef('S', 'B4321', 'changed');
  s.recalculate();
  assert.equal(s.getByRef('S', 'D1'), 'changed');
  assert.equal(s.getByRef('S', 'D2'), 'changed!');
});

test('a lookup reads a formula in its range after that formula settles in the same pass', () => {
  const s = column();
  s.setByRef('S', 'C1', 'first');
  s.setByRef('S', 'B2500', '=C1&"-x"');
  s.setByRef('S', 'D1', '=VLOOKUP(2500, A1:B' + ROWS + ', 2, FALSE)');
  s.setByRef('S', 'D2', '=VLOOKUP(10, A1:B' + ROWS + ', 2, FALSE)');
  s.recalculate();
  assert.equal(s.getByRef('S', 'D1'), 'first-x');
  s.setByRef('S', 'C1', 'second');
  s.recalculate();
  assert.equal(s.getByRef('S', 'D1'), 'second-x');
  assert.equal(s.getByRef('S', 'D2'), 'v10');
});

// ------------------------------------------------------- exact lookups --

/** The walk the index stands in for. */
function walk(table, key, idx) {
  for (const row of table) {
    const cmp = compareValues(row[0], key);
    if (isError(cmp)) return cmp;
    if (cmp === 0) return row[idx - 1] ?? '';
  }
  return ERR.NA('no match');
}

const same = (a, b) => (isError(a) || isError(b) ? isError(a) && isError(b) && a.type === b.type : a === b);

test('an indexed exact VLOOKUP answers what the walk would', () => {
  let seed = 7;
  const rand = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
  const pick = () => {
    const x = rand();
    if (x < 0.3) return Math.floor(rand() * 40);
    if (x < 0.55) return ['ab', 'AB', 'Ab', 'cd', '5', '0', 'x'][Math.floor(rand() * 7)];
    if (x < 0.65) return rand() < 0.5;
    if (x < 0.8) return '';
    if (x < 0.82) return ERR.DIV0();
    return -Math.floor(rand() * 5) / 2;
  };
  const keys = [0, 1, 5, 39, -0.5, -2, 'ab', 'aB', 'CD', '5', '0', 'x', 'missing', true, false, '', null, 1000];
  for (let t = 0; t < 30; t++) {
    const table = Array.from({ length: 300 + Math.floor(rand() * 400) }, (_, r) => [pick(), r]);
    for (const key of keys) {
      const got = FUNCTIONS.VLOOKUP.fn(key, table, 2, false);
      assert.ok(same(got, walk(table, key, 2)), `table ${t}, key ${JSON.stringify(key)}: ${got}`);
    }
  }
});

test('an exact VLOOKUP over a long column with no blanks or errors finds the first match', () => {
  const table = Array.from({ length: 1000 }, (_, r) => [r % 100, 'row ' + r]);
  assert.equal(FUNCTIONS.VLOOKUP.fn(42, table, 2, false), 'row 42');
  assert.equal(FUNCTIONS.VLOOKUP.fn(0, table, 2, false), 'row 0');
  assert.equal(FUNCTIONS.VLOOKUP.fn(100, table, 2, false).type, '#N/A');
  assert.equal(FUNCTIONS.VLOOKUP.fn('42', table, 2, false).type, '#N/A', 'text "42" is not the number 42');
});

test('a long column holding dates is walked, not indexed', () => {
  const when = new Date(Date.UTC(2026, 0, 1));
  const table = Array.from({ length: 400 }, (_, r) => [r === 300 ? when : r, r]);
  assert.ok(same(FUNCTIONS.VLOOKUP.fn(350, table, 2, false), walk(table, 350, 2)));
});

// ---------------------------------------------------------- first error --

test('a long array remembers its first error', () => {
  const clean = Array.from({ length: 1000 }, (_, r) => [r, r * 2]);
  assert.equal(firstError([clean]), null);
  assert.equal(firstError([clean]), null, 'asked twice, the same answer');
  const dirty = Array.from({ length: 1000 }, (_, r) => [r, r === 700 ? ERR.REF() : r]);
  assert.equal(firstError([1, dirty]).type, '#REF!');
  assert.equal(firstError([dirty]).type, '#REF!');
  assert.equal(firstError([ERR.NA(), dirty]).type, '#N/A', 'an earlier error still wins');
  const short = [[1, ERR.NUM()]];
  assert.equal(firstError([short]).type, '#NUM!');
});
