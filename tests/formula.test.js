/**
 * The calculation engine.
 *
 * These assert spreadsheet semantics, not JavaScript semantics. Where the two
 * differ the test says so, because those differences are exactly what makes a
 * sheet built in Excel produce the same numbers here.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  Spreadsheet, calculate, parse, tokenize, isError, ERR,
  toNumber, toText, compareValues, roundHalfAwayFromZero, dateToSerial, serialToDate,
} from '@rutba/formula';

/** A resolver over a plain object keyed by A1, for testing formulas alone. */
function sheetOf(cells, now = new Date('2026-08-20T12:00:00Z')) {
  const at = (row, col) => {
    const letters = (n) => { let s = ''; let v = n + 1; while (v > 0) { const r = (v - 1) % 26; s = String.fromCharCode(65 + r) + s; v = Math.floor((v - 1) / 26); } return s; };
    return cells[letters(col) + (row + 1)] ?? '';
  };
  return {
    now: () => now,
    usedBounds: () => ({ maxRow: 200, maxCol: 50 }),
    getCell: (_s, row, col) => at(row, col),
    getRange: (_s, start, end) => {
      const grid = [];
      for (let r = start.row; r <= end.row; r++) {
        const line = [];
        for (let c = start.col; c <= end.col; c++) line.push(at(r, c));
        grid.push(line);
      }
      return grid;
    },
    getName: () => null,
  };
}

const evalIn = (formula, cells) => calculate(formula, sheetOf(cells ?? {}), { sheet: 'Sheet1' });

// ---------------------------------------------------------------- parsing --

test('tokenizer distinguishes unary from binary minus', () => {
  assert.equal(evalIn('=-2'), -2);
  assert.equal(evalIn('=3-2'), 1);
  assert.equal(evalIn('=3--2'), 5);
  assert.equal(evalIn('=-(3-5)'), 2);
});

test('power binds tighter than unary minus, as in a spreadsheet', () => {
  // This is the case that differs from most programming languages.
  assert.equal(evalIn('=-2^2'), -4);
  assert.equal(evalIn('=(-2)^2'), 4);
  assert.equal(evalIn('=2^-1'), 0.5);
  assert.equal(evalIn('=2^3^2'), 512, 'power is right-associative');
});

test('percent is a postfix operator', () => {
  assert.equal(evalIn('=50%'), 0.5);
  assert.equal(evalIn('=200*10%'), 20);
  assert.equal(evalIn('=A1%', { A1: 250 }), 2.5);
});

test('operator precedence overall', () => {
  assert.equal(evalIn('=1+2*3'), 7);
  assert.equal(evalIn('=(1+2)*3'), 9);
  assert.equal(evalIn('="a"&1+2'), 'a3', 'concat is looser than arithmetic');
  assert.equal(evalIn('=1+1=2'), true, 'comparison is loosest');
});

test('strings, escaped quotes and error literals parse', () => {
  assert.equal(evalIn('="hello"'), 'hello');
  assert.equal(evalIn('="say ""hi"""'), 'say "hi"');
  assert.ok(isError(evalIn('=#N/A')));
  assert.equal(evalIn('=#N/A').type, '#N/A');
});

test('malformed formulas become an error value, not an exception', () => {
  for (const bad of ['=1+', '=SUM(', '=)', '=1 2 3 4', '="unterminated']) {
    const r = evalIn(bad);
    assert.ok(isError(r), bad + ' should be an error');
  }
});

test('references parse in all their forms', () => {
  assert.equal(evalIn('=A1', { A1: 5 }), 5);
  assert.equal(evalIn('=$A$1', { A1: 5 }), 5);
  assert.equal(evalIn('=$A1', { A1: 5 }), 5);
  assert.equal(evalIn('=A$1', { A1: 5 }), 5);
  assert.equal(evalIn('=SUM(A1:A3)', { A1: 1, A2: 2, A3: 3 }), 6);
  assert.equal(evalIn('=SUM(A1:C1)', { A1: 1, B1: 2, C1: 3 }), 6);
  assert.equal(evalIn('=AA1', { AA1: 9 }), 9);
  assert.deepEqual(parse('=Sheet2!B7').sheet, 'Sheet2');
  assert.deepEqual(parse("='My Sheet'!B7").sheet, 'My Sheet');
});

// ------------------------------------------------------------- coercion ----

test('coercion follows spreadsheet rules, not JavaScript rules', () => {
  assert.equal(evalIn('="5"+1'), 6, 'numeric text is a number in arithmetic');
  assert.ok(isError(evalIn('="abc"+1')), 'non-numeric text is #VALUE!');
  assert.equal(evalIn('=TRUE+1'), 2, 'TRUE is 1 in arithmetic');
  assert.equal(evalIn('=A1+1', {}), 1, 'a blank cell is 0');
  assert.equal(evalIn('="x"&A1', {}), 'x', 'a blank cell is "" in concatenation');
  assert.equal(evalIn('=TRUE&""'), 'TRUE', 'TRUE is "TRUE" in concatenation');
  assert.equal(evalIn('=1&2'), '12');
  assert.equal(toNumber(''), 0);
  assert.equal(toText(1.5), '1.5');
});

test('comparison is case-insensitive and orders numbers before text', () => {
  assert.equal(evalIn('="ABC"="abc"'), true);
  assert.equal(evalIn('="a"<"b"'), true);
  assert.equal(compareValues(1, 'a'), -1, 'numbers sort before text');
  assert.equal(evalIn('=1<>2'), true);
  assert.equal(evalIn('=A1=0', {}), true, 'a blank equals 0');
  assert.equal(evalIn('=A1=""', {}), true, 'a blank also equals empty text');
});

test('floating point noise is hidden the way a spreadsheet hides it', () => {
  assert.equal(evalIn('=0.1+0.2'), 0.30000000000000004, 'the value is still binary float');
  assert.equal(toText(evalIn('=0.1+0.2')), '0.3', 'but it renders as 0.3');
  assert.equal(roundHalfAwayFromZero(2.675, 2), 2.68, 'not 2.67 — a penny matters');
  assert.equal(roundHalfAwayFromZero(-2.5), -3, 'half away from zero, not banker rounding');
});

// ---------------------------------------------------------------- errors ---

test('errors are values that propagate, not exceptions', () => {
  const r = evalIn('=A1+1', { A1: ERR.DIV0() });
  assert.ok(isError(r));
  assert.equal(r.type, '#DIV/0!');
  assert.equal(evalIn('=1/0').type, '#DIV/0!');
  assert.equal(evalIn('=SQRT(-1)').type, '#NUM!');
  assert.equal(evalIn('=NOSUCHFUNC(1)').type, '#NAME?');
  assert.equal(evalIn('=SUM(A1:A3)', { A2: ERR.VALUE() }).type, '#VALUE!', 'an error inside a range propagates');
});

test('IFERROR and the IS family look at an error without spreading it', () => {
  assert.equal(evalIn('=IFERROR(1/0, "n/a")'), 'n/a');
  assert.equal(evalIn('=IFERROR(4/2, "n/a")'), 2);
  assert.equal(evalIn('=ISERROR(1/0)'), true);
  assert.equal(evalIn('=ISERROR(1)'), false);
  assert.equal(evalIn('=ISNA(NA())'), true);
  assert.equal(evalIn('=ISNA(1/0)'), false, '#DIV/0! is not #N/A');
});

test('IF does not evaluate the branch it does not take', () => {
  // The classic guard. If the false branch were evaluated eagerly this is #DIV/0!.
  assert.equal(evalIn('=IF(A1=0, 0, 1/A1)', { A1: 0 }), 0);
  assert.equal(evalIn('=IF(A1=0, 0, 1/A1)', { A1: 4 }), 0.25);
});

// -------------------------------------------------------------- functions --

test('aggregates skip text and blanks in ranges but not in literals', () => {
  const cells = { A1: 1, A2: 'hello', A3: 3, A4: '' };
  assert.equal(evalIn('=SUM(A1:A4)', cells), 4);
  assert.equal(evalIn('=COUNT(A1:A4)', cells), 2, 'COUNT counts numbers');
  assert.equal(evalIn('=COUNTA(A1:A4)', cells), 3, 'COUNTA counts non-blanks');
  assert.equal(evalIn('=AVERAGE(A1:A4)', cells), 2);
  assert.ok(isError(evalIn('=SUM(1,"hello")')), 'a literal argument must be a number');
});

test('the maths functions behave', () => {
  assert.equal(evalIn('=ROUND(2.5)'), 3);
  assert.equal(evalIn('=ROUND(3.14159, 2)'), 3.14);
  assert.equal(evalIn('=ROUNDUP(1.01, 0)'), 2);
  assert.equal(evalIn('=ROUNDDOWN(1.99, 0)'), 1);
  assert.equal(evalIn('=MOD(-3, 2)'), 1, 'MOD takes the sign of the divisor, unlike JS %');
  assert.equal(evalIn('=INT(-1.5)'), -2);
  assert.equal(evalIn('=TRUNC(-1.5)'), -1, 'TRUNC differs from INT for negatives');
  assert.equal(evalIn('=ABS(-4)'), 4);
  assert.equal(evalIn('=POWER(2,10)'), 1024);
  assert.equal(evalIn('=MIN(A1:A3)', { A1: 5, A2: 2, A3: 9 }), 2);
  assert.equal(evalIn('=MAX(A1:A3)', { A1: 5, A2: 2, A3: 9 }), 9);
});

test('the text functions behave', () => {
  assert.equal(evalIn('=LEN("hello")'), 5);
  assert.equal(evalIn('=LEFT("hello", 2)'), 'he');
  assert.equal(evalIn('=RIGHT("hello", 3)'), 'llo');
  assert.equal(evalIn('=MID("hello", 2, 3)'), 'ell');
  assert.equal(evalIn('=UPPER("abc")'), 'ABC');
  assert.equal(evalIn('=TRIM("  a   b  ")'), 'a b');
  assert.equal(evalIn('=CONCATENATE("a","b","c")'), 'abc');
  assert.equal(evalIn('=SUBSTITUTE("a-b-c","-","+")'), 'a+b+c');
  assert.equal(evalIn('=VALUE("42")'), 42);
  assert.equal(evalIn('=TEXT(1234.5, "#,##0.00")'), '1,234.50');
});

test('conditional aggregation with operators and wildcards', () => {
  const cells = { A1: 'apple', A2: 'apricot', A3: 'banana', B1: 10, B2: 20, B3: 30 };
  assert.equal(evalIn('=COUNTIF(A1:A3, "ap*")', cells), 2);
  assert.equal(evalIn('=SUMIF(A1:A3, "ap*", B1:B3)', cells), 30);
  assert.equal(evalIn('=SUMIF(B1:B3, ">15")', cells), 50);
  assert.equal(evalIn('=COUNTIF(B1:B3, ">=20")', cells), 2);
  assert.equal(evalIn('=SUMIF(A1:A3, "banana", B1:B3)', cells), 30);
});

test('lookup functions', () => {
  const cells = { A1: 'SKU-1', B1: 100, A2: 'SKU-2', B2: 200, A3: 'SKU-3', B3: 300 };
  assert.equal(evalIn('=VLOOKUP("SKU-2", A1:B3, 2, FALSE)', cells), 200);
  assert.equal(evalIn('=VLOOKUP("SKU-9", A1:B3, 2, FALSE)').type, '#N/A');
  assert.equal(evalIn('=VLOOKUP("SKU-2", A1:B3, 5, FALSE)', cells).type, '#REF!');
  assert.equal(evalIn('=INDEX(A1:B3, 2, 2)', cells), 200);
  assert.equal(evalIn('=MATCH("SKU-3", A1:A3, 0)', cells), 3);
});

test('logic functions', () => {
  assert.equal(evalIn('=AND(TRUE, 1, "TRUE")'), true);
  assert.equal(evalIn('=AND(TRUE, FALSE)'), false);
  assert.equal(evalIn('=OR(FALSE, 0, 1)'), true);
  assert.equal(evalIn('=NOT(TRUE)'), false);
});

test('date serials use the 1900 epoch, leap-year bug included', () => {
  // The format carries the bug; reproducing the arithmetic is the point.
  assert.equal(dateToSerial(new Date(Date.UTC(1900, 0, 1))), 1);
  assert.equal(dateToSerial(new Date(Date.UTC(1900, 1, 28))), 59);
  assert.equal(dateToSerial(new Date(Date.UTC(1900, 2, 1))), 61, 'serial 60 is the day that never existed');
  assert.equal(dateToSerial(new Date(Date.UTC(2024, 0, 1))), 45292);
  assert.equal(dateToSerial(new Date(Date.UTC(2026, 7, 20))), 46254);
  assert.equal(serialToDate(46254).toISOString().slice(0, 10), '2026-08-20');
  assert.equal(evalIn('=YEAR(DATE(2026,8,20))'), 2026);
  assert.equal(evalIn('=MONTH(DATE(2026,8,20))'), 8);
  assert.equal(evalIn('=DAY(DATE(2026,8,20))'), 20);
  assert.equal(evalIn('=DATE(2026,8,21)-DATE(2026,8,20)'), 1);
  assert.equal(evalIn('=YEAR(TODAY())'), 2026);
});

// ----------------------------------------------------------------- engine --

test('recalculation follows dependency order, not cell order', () => {
  const s = new Spreadsheet();
  s.addSheet('Sheet1');
  // deliberately defined so that evaluating top-to-bottom would be wrong
  s.setByRef('Sheet1', 'C1', '=B1*2');
  s.setByRef('Sheet1', 'B1', '=A1+1');
  s.setByRef('Sheet1', 'A1', 10);

  const { calculated } = s.recalculate();
  assert.equal(calculated, 2);
  assert.equal(s.getByRef('Sheet1', 'B1'), 11);
  assert.equal(s.getByRef('Sheet1', 'C1'), 22);

  s.setByRef('Sheet1', 'A1', 100);
  s.recalculate();
  assert.equal(s.getByRef('Sheet1', 'C1'), 202, 'a change flows through the chain');
});

test('a deep chain recalculates correctly', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setByRef('S', 'A1', 1);
  for (let i = 2; i <= 50; i++) s.setByRef('S', 'A' + i, '=A' + (i - 1) + '+1');
  s.recalculate();
  assert.equal(s.getByRef('S', 'A50'), 50);
});

test('circular references are reported, not hung on', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setByRef('S', 'A1', '=B1');
  s.setByRef('S', 'B1', '=A1');
  const { cycles } = s.recalculate();
  assert.equal(cycles.length, 2);
  assert.ok(cycles.includes('S!A1'));
  assert.equal(s.getByRef('S', 'A1').type, '#CIRCULAR!');
});

test('a self-reference is circular', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setByRef('S', 'A1', '=A1+1');
  const { cycles } = s.recalculate();
  assert.deepEqual(cycles, ['S!A1']);
});

test('a cycle does not stop the rest of the sheet calculating', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setByRef('S', 'A1', '=B1');
  s.setByRef('S', 'B1', '=A1');
  s.setByRef('S', 'D1', 5);
  s.setByRef('S', 'D2', '=D1*3');
  s.recalculate();
  assert.equal(s.getByRef('S', 'D2'), 15);
  assert.equal(s.getByRef('S', 'A1').type, '#CIRCULAR!');
});

test('cross-sheet references work', () => {
  const s = new Spreadsheet();
  s.addSheet('Data');
  s.addSheet('Report');
  s.setByRef('Data', 'A1', 40);
  s.setByRef('Data', 'A2', 2);
  s.setByRef('Report', 'B1', '=SUM(Data!A1:A2)');
  s.recalculate();
  assert.equal(s.getByRef('Report', 'B1'), 42);
});

test('defined names resolve to ranges', () => {
  const s = new Spreadsheet();
  s.addSheet('Stock');
  s.setByRef('Stock', 'A1', 10);
  s.setByRef('Stock', 'A2', 20);
  s.setByRef('Stock', 'A3', 30);
  s.defineName('RUTBA_STOCK', { sheet: 'Stock', start: { row: 0, col: 0 }, end: { row: 2, col: 0 } });
  s.setByRef('Stock', 'C1', '=SUM(RUTBA_STOCK)');
  s.recalculate();
  assert.equal(s.getByRef('Stock', 'C1'), 60);
});

test('an unknown name is #NAME?, and says which name', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setByRef('S', 'A1', '=SUM(NOT_DEFINED)');
  s.recalculate();
  const v = s.getByRef('S', 'A1');
  assert.equal(v.type, '#NAME?');
  assert.match(v.detail, /NOT_DEFINED/);
});

test('the formula bar shows what was typed, the cell shows the result', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setByRef('S', 'A1', 6);
  s.setByRef('S', 'A2', '=A1*7');
  s.recalculate();
  assert.equal(s.getInput('S', 1, 0), '=A1*7');
  assert.equal(s.getByRef('S', 'A2'), 42);
});

test('clearing a cell removes it and dependents see a blank', () => {
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setByRef('S', 'A1', 5);
  s.setByRef('S', 'B1', '=A1+1');
  s.recalculate();
  assert.equal(s.getByRef('S', 'B1'), 6);
  s.setByRef('S', 'A1', '');
  s.recalculate();
  assert.equal(s.getByRef('S', 'B1'), 1, 'a blank counts as 0');
});

test('a realistic ERP sheet computes end to end', () => {
  const s = new Spreadsheet();
  s.addSheet('Stock');
  const rows = [
    ['SKU-1001', 1420, 120],
    ['SKU-1002', 860, 40],
    ['SKU-1140', 240, 0],
    ['SKU-2202', 12, 12],
  ];
  s.setByRef('Stock', 'A1', 'SKU');
  s.setByRef('Stock', 'B1', 'On hand');
  s.setByRef('Stock', 'C1', 'Reserved');
  s.setByRef('Stock', 'D1', 'Available');
  rows.forEach(([sku, onHand, reserved], i) => {
    const r = i + 2;
    s.setByRef('Stock', 'A' + r, sku);
    s.setByRef('Stock', 'B' + r, onHand);
    s.setByRef('Stock', 'C' + r, reserved);
    s.setByRef('Stock', 'D' + r, '=B' + r + '-C' + r);
  });
  s.setByRef('Stock', 'B6', '=SUM(B2:B5)');
  s.setByRef('Stock', 'D6', '=SUM(D2:D5)');
  s.setByRef('Stock', 'F1', '=COUNTIF(D2:D5,"=0")');
  s.setByRef('Stock', 'F2', '=IF(F1>0, "Restock needed", "OK")');
  s.recalculate();

  assert.equal(s.getByRef('Stock', 'D2'), 1300);
  assert.equal(s.getByRef('Stock', 'B6'), 2532);
  assert.equal(s.getByRef('Stock', 'D6'), 2360);
  assert.equal(s.getByRef('Stock', 'F1'), 1, 'SKU-2202 is fully reserved');
  assert.equal(s.getByRef('Stock', 'F2'), 'Restock needed');
});

test('tokenize is exported and usable on its own', () => {
  const tokens = tokenize('=SUM(A1:A3)*2');
  assert.equal(tokens[0].type, 'func');
  assert.ok(tokens.some((t) => t.type === 'ref'));
});

// ---------------------------------------------------------------------------
// Incremental recalculation — only what the edit can have changed.
// ---------------------------------------------------------------------------

test('editing a cell recalculates its dependents and leaves the rest alone', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, 10);            // A1
  s.setCell('S', 0, 1, '=A1*2');       // B1  depends on A1
  s.setCell('S', 0, 2, '=B1+1');       // C1  depends on B1
  s.setCell('S', 5, 0, 100);           // A6  unrelated
  s.setCell('S', 5, 1, '=A6*3');       // B6  unrelated
  const first = s.recalculate();
  assert.equal(first.calculated, 3, 'the first pass evaluates every formula');

  s.setCell('S', 0, 0, 20);
  const second = s.recalculate();
  assert.equal(second.calculated, 2, 'only B1 and C1 are downstream of A1');
  assert.equal(s.getValue('S', 0, 1), 40);
  assert.equal(s.getValue('S', 0, 2), 41);
  assert.equal(s.getValue('S', 5, 1), 300, 'the unrelated formula kept its value');
});

test('changing a formula recalculates that cell even if nothing it reads moved', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, 7);
  s.setCell('S', 0, 1, '=A1*2');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 1), 14);

  s.setCell('S', 0, 1, '=A1*10');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 1), 70);
});

test('a value typed into a literal cell reaches the formulas that read it', () => {
  // The edge that makes incremental recalculation possible at all: the
  // dependency graph has to carry literal-to-formula edges, or typing a number
  // reaches nothing, because a number is not a formula.
  const s = new Spreadsheet();
  s.setCell('S', 0, 1, '=A1+1');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 1), 1, 'an empty A1 is zero');

  s.setCell('S', 0, 0, 41);
  s.recalculate();
  assert.equal(s.getValue('S', 0, 1), 42);
});

test('a new cell inside a range updates the formula that sums it', () => {
  // Subtle: a range clamps to the used bounds, so a cell added BEYOND them
  // extends the range. The dependency edges are rebuilt each pass, so the new
  // cell is linked to the SUM before reachability is computed — miss that and
  // the total silently stops moving.
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, 1);
  s.setCell('S', 1, 0, 2);
  s.setCell('S', 0, 3, '=SUM(A1:A100)');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 3), 3);

  s.setCell('S', 2, 0, 4);   // a row that did not exist before
  s.recalculate();
  assert.equal(s.getValue('S', 0, 3), 7);
});

test('a cycle is still caught, and stops being reported once it is broken', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, '=B1');
  s.setCell('S', 0, 1, '=A1');
  const cyclic = s.recalculate();
  assert.deepEqual(cyclic.cycles.sort(), ['S!A1', 'S!B1']);
  assert.equal(isError(s.getValue('S', 0, 0)), true);

  s.setCell('S', 0, 1, 5);
  const fixed = s.recalculate();
  assert.deepEqual(fixed.cycles, [], 'breaking the cycle clears it');
  assert.equal(s.getValue('S', 0, 0), 5);
});

test('a full recalculation can still be asked for explicitly', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, 2);
  s.setCell('S', 0, 1, '=A1*2');
  s.recalculate();
  assert.equal(s.recalculate({ all: true }).calculated, 1, 'every formula, regardless of dirt');
});

// ------------------------------------------------- the wider function set --

test('CEILING, FLOOR and MROUND round to multiples with sign discipline', () => {
  assert.equal(evalIn('=CEILING(6.3)'), 7);
  assert.equal(evalIn('=CEILING(6.3, 0.5)'), 6.5);
  assert.equal(evalIn('=FLOOR(6.7, 0.5)'), 6.5);
  assert.equal(evalIn('=MROUND(7.3, 2)'), 8, 'nearest multiple, half away from zero');
  assert.equal(evalIn('=MROUND(-7, 2)').type, '#NUM!', 'opposite signs are refused');
  assert.equal(evalIn('=CEILING(5, 0)'), 0);
});

test('EVEN and ODD round away from zero; ISEVEN and ISODD read the integer part', () => {
  assert.equal(evalIn('=EVEN(1.5)'), 2);
  assert.equal(evalIn('=EVEN(-1)'), -2);
  assert.equal(evalIn('=ODD(2)'), 3);
  assert.equal(evalIn('=ODD(-1.5)'), -3);
  assert.equal(evalIn('=ISEVEN(2.7)'), true, 'the fraction is cut, not rounded');
  assert.equal(evalIn('=ISODD(-3)'), true);
});

test('logs and exponentials guard their domains', () => {
  assert.equal(evalIn('=LN(EXP(2))'), 2);
  assert.equal(evalIn('=LOG(1000)'), 3, 'LOG defaults to base 10');
  assert.equal(evalIn('=LOG(8, 2)'), 3);
  assert.equal(evalIn('=LOG10(100)'), 2);
  assert.equal(evalIn('=LN(0)').type, '#NUM!');
  assert.equal(Math.round(evalIn('=PI()') * 10000), 31416);
});

test('SUMPRODUCT multiplies position by position, text counting as zero', () => {
  const cells = { A1: 2, A2: 3, A3: 'x', B1: 10, B2: 100, B3: 1000 };
  assert.equal(evalIn('=SUMPRODUCT(A1:A3, B1:B3)', cells), 320);
  assert.equal(evalIn('=SUMPRODUCT(A1:A2, B1:B3)', cells).type, '#VALUE!', 'sizes must agree');
  assert.equal(evalIn('=SUMSQ(3, 4)'), 25);
});

test('MEDIAN, MODE, LARGE, SMALL and RANK read a data column', () => {
  const cells = { A1: 10, A2: 40, A3: 20, A4: 40, A5: 30 };
  assert.equal(evalIn('=MEDIAN(A1:A5)', cells), 30);
  assert.equal(evalIn('=MEDIAN(1, 2, 3, 4)'), 2.5, 'an even count averages the middle two');
  assert.equal(evalIn('=MODE(A1:A5)', cells), 40);
  assert.equal(evalIn('=MODE(1, 2, 3)').type, '#N/A', 'nothing repeats');
  assert.equal(evalIn('=LARGE(A1:A5, 2)', cells), 40);
  assert.equal(evalIn('=SMALL(A1:A5, 2)', cells), 20);
  assert.equal(evalIn('=LARGE(A1:A5, 9)', cells).type, '#NUM!');
  assert.equal(evalIn('=RANK(30, A1:A5)', cells), 3, 'descending by default');
  assert.equal(evalIn('=RANK(30, A1:A5, 1)', cells), 3, 'ascending when asked');
  assert.equal(evalIn('=RANK(10, A1:A5)', cells), 5);
});

test('standard deviation and variance, sample and population', () => {
  const cells = { A1: 2, A2: 4, A3: 4, A4: 4, A5: 5, A6: 5, A7: 7, A8: 9 };
  assert.equal(evalIn('=VAR.P(A1:A8)', cells), 4);
  assert.equal(evalIn('=STDEV.P(A1:A8)', cells), 2);
  assert.equal(Math.round(evalIn('=VAR.S(A1:A8)', cells) * 1000) / 1000, 4.571);
  assert.equal(evalIn('=STDEV(5)').type, '#DIV/0!', 'a sample needs two values');
  assert.equal(evalIn('=STDEVP(5)'), 0, 'a population of one has no spread');
});

test('the *IFS family filters on every criterion at once', () => {
  const cells = {
    A1: 'east', A2: 'west', A3: 'east', A4: 'east',
    B1: 'red', B2: 'red', B3: 'blue', B4: 'red',
    C1: 10, C2: 20, C3: 30, C4: 40,
  };
  assert.equal(evalIn('=SUMIFS(C1:C4, A1:A4, "east", B1:B4, "red")', cells), 50);
  assert.equal(evalIn('=COUNTIFS(A1:A4, "east", C1:C4, ">15")', cells), 2);
  assert.equal(evalIn('=AVERAGEIFS(C1:C4, A1:A4, "east")', cells), 80 / 3);
  assert.equal(evalIn('=MAXIFS(C1:C4, B1:B4, "red")', cells), 40);
  assert.equal(evalIn('=MINIFS(C1:C4, B1:B4, "red")', cells), 10);
  assert.equal(evalIn('=MAXIFS(C1:C4, B1:B4, "green")', cells), 0, 'no match is 0, not an error');
  assert.equal(evalIn('=SUMIFS(C1:C4, A1:A2, "east")', cells).type, '#VALUE!', 'ranges must be the same size');
});

test('IFS, SWITCH and XOR', () => {
  assert.equal(evalIn('=IFS(1>2, "a", 2>1, "b")'), 'b');
  assert.equal(evalIn('=IFS(1>2, "a")').type, '#N/A', 'nothing was true');
  assert.equal(evalIn('=IFS(TRUE, "x", 1/0, "boom")'), 'x', 'later conditions are never evaluated');
  assert.equal(evalIn('=SWITCH(2, 1, "one", 2, "two", "many")'), 'two');
  assert.equal(evalIn('=SWITCH(9, 1, "one", "many")'), 'many', 'the odd trailing value is the default');
  assert.equal(evalIn('=SWITCH(9, 1, "one")').type, '#N/A');
  assert.equal(evalIn('=XOR(TRUE, TRUE, TRUE)'), true);
  assert.equal(evalIn('=XOR(TRUE, TRUE)'), false);
});

test('FIND is case-sensitive, SEARCH is not and takes wildcards', () => {
  assert.equal(evalIn('=FIND("b", "abc")'), 2);
  assert.equal(evalIn('=FIND("B", "abc")').type, '#VALUE!');
  assert.equal(evalIn('=SEARCH("B", "abc")'), 2);
  assert.equal(evalIn('=SEARCH("c?n", "second")'), 3);
  assert.equal(evalIn('=FIND("a", "banana", 3)'), 4, 'start position honoured');
});

test('REPLACE, REPT, PROPER, EXACT, CHAR, CODE and TEXTJOIN', () => {
  assert.equal(evalIn('=REPLACE("abcdef", 2, 3, "XY")'), 'aXYef');
  assert.equal(evalIn('=REPT("ab", 3)'), 'ababab');
  assert.equal(evalIn('=REPT("x", -1)').type, '#VALUE!');
  assert.equal(evalIn('=PROPER("hello world-two")'), 'Hello World-Two');
  assert.equal(evalIn('=EXACT("Word", "word")'), false);
  assert.equal(evalIn('=EXACT("word", "word")'), true);
  assert.equal(evalIn('=CHAR(65)'), 'A');
  assert.equal(evalIn('=CODE("A")'), 65);
  const cells = { A1: 'a', A2: '', A3: 'c' };
  assert.equal(evalIn('=TEXTJOIN("-", TRUE, A1:A3)', cells), 'a-c');
  assert.equal(evalIn('=TEXTJOIN("-", FALSE, A1:A3)', cells), 'a--c');
});

test('HLOOKUP mirrors VLOOKUP along the other axis', () => {
  const cells = { A1: 'q1', B1: 'q2', C1: 'q3', A2: 10, B2: 20, C2: 30 };
  assert.equal(evalIn('=HLOOKUP("q2", A1:C2, 2, FALSE)', cells), 20);
  assert.equal(evalIn('=HLOOKUP("q9", A1:C2, 2, FALSE)', cells).type, '#N/A');
});

test('XLOOKUP finds exactly, approximately, by wildcard, and falls back', () => {
  const cells = {
    A1: 'ant', A2: 'bee', A3: 'cat',
    C1: 10, C2: 20, C3: 30,
  };
  assert.equal(evalIn('=XLOOKUP("bee", A1:A3, C1:C3)', cells), 20);
  assert.equal(evalIn('=XLOOKUP("dog", A1:A3, C1:C3)', cells).type, '#N/A');
  assert.equal(evalIn('=XLOOKUP("dog", A1:A3, C1:C3, "none")', cells), 'none');
  assert.equal(evalIn('=XLOOKUP(25, C1:C3, A1:A3, "x", -1)', cells), 'bee', 'next smaller');
  assert.equal(evalIn('=XLOOKUP(25, C1:C3, A1:A3, "x", 1)', cells), 'cat', 'next larger');
  assert.equal(evalIn('=XLOOKUP("b*", A1:A3, C1:C3, "x", 2)', cells), 20, 'wildcard mode');
  assert.equal(evalIn('=XLOOKUP("bee", A1:A3, C1:C2)', cells).type, '#VALUE!', 'sizes must agree');
});

test('CHOOSE picks lazily by index', () => {
  assert.equal(evalIn('=CHOOSE(2, "a", "b", "c")'), 'b');
  assert.equal(evalIn('=CHOOSE(2, 1/0, 5)'), 5, 'the branch not taken never runs');
  assert.equal(evalIn('=CHOOSE(4, "a", "b")').type, '#VALUE!');
});

test('EDATE and EOMONTH walk months, clamping the day', () => {
  const jan31 = evalIn('=DATE(2026, 1, 31)');
  assert.equal(evalIn('=EDATE(DATE(2026,1,31), 1)'), evalIn('=DATE(2026, 2, 28)'), 'day clamped into February');
  assert.equal(evalIn('=EOMONTH(DATE(2026,2,10), 0)'), evalIn('=DATE(2026, 2, 28)'));
  assert.equal(evalIn('=EOMONTH(DATE(2026,1,15), 1)'), evalIn('=DATE(2026, 2, 28)'));
  assert.equal(evalIn('=EDATE(DATE(2026,1,31), 12)') - jan31, 365);
  assert.equal(evalIn('=DAYS(DATE(2026,3,1), DATE(2026,2,1))'), 28);
});

test('WEEKDAY, TIME and the time parts agree with the serial arithmetic', () => {
  // 2026-08-24 is a Monday.
  assert.equal(evalIn('=WEEKDAY(DATE(2026,8,24))'), 2, 'type 1: Sunday is 1');
  assert.equal(evalIn('=WEEKDAY(DATE(2026,8,24), 2)'), 1, 'type 2: Monday is 1');
  assert.equal(evalIn('=WEEKDAY(DATE(2026,8,24), 3)'), 0, 'type 3: Monday is 0');
  assert.equal(evalIn('=HOUR(TIME(13, 45, 30))'), 13);
  assert.equal(evalIn('=MINUTE(TIME(13, 45, 30))'), 45);
  assert.equal(evalIn('=SECOND(TIME(13, 45, 30))'), 30);
  assert.equal(evalIn('=TIME(25, 0, 0)'), 1 / 24, 'wraps past midnight');
});

test('the annuity functions agree with each other', () => {
  // A 60-month loan of 10,000 at 0.5% per month.
  const pmt = evalIn('=PMT(0.005, 60, 10000)');
  assert.equal(Math.round(pmt * 100) / 100, -193.33);
  // Paying that back has a present value of the principal…
  const pv = evalIn('=PV(0.005, 60, ' + pmt + ')');
  assert.equal(Math.round(pv), 10000);
  // …and takes 60 periods.
  const nper = evalIn('=NPER(0.005, ' + pmt + ', 10000)');
  assert.equal(Math.round(nper), 60);
  assert.equal(evalIn('=PMT(0, 10, 1000)'), -100, 'zero rate is straight division');
  const fv = evalIn('=FV(0.01, 12, -100)');
  assert.equal(Math.round(fv * 100) / 100, 1268.25);
  const npv = evalIn('=NPV(0.1, 100, 100)');
  assert.equal(Math.round(npv * 100) / 100, 173.55);
});

test('N and the volatile RAND family', () => {
  assert.equal(evalIn('=N("text")'), 0);
  assert.equal(evalIn('=N(TRUE)'), 1);
  assert.equal(evalIn('=N(7)'), 7);
  const r = evalIn('=RAND()');
  assert.equal(r >= 0 && r < 1, true);
  assert.equal(evalIn('=RANDBETWEEN(5, 5)'), 5);
  assert.equal(evalIn('=RANDBETWEEN(9, 1)').type, '#NUM!');
});

test('shiftFormula moves relative references and honours every $', async () => {
  const { shiftFormula } = await import('@rutba/formula');
  assert.equal(shiftFormula('=A1+B2', 1, 0), '=A2+B3');
  assert.equal(shiftFormula('=A1+B2', 0, 2), '=C1+D2');
  assert.equal(shiftFormula('=$A$1+B2', 3, 3), '=$A$1+E5');
  assert.equal(shiftFormula('=$A1+A$1', 1, 1), '=$A2+B$1', 'mixed anchors pin one axis each');
  assert.equal(shiftFormula('=SUM(A1:B2)*2', 1, 1), '=SUM(B2:C3)*2');
  assert.equal(shiftFormula('=SUM(A:A)+SUM(1:1)', 1, 1), '=SUM(B:B)+SUM(2:2)', 'whole columns and rows shift too');
  assert.equal(shiftFormula('=LOG10(A1)', 1, 0), '=LOG10(A2)', 'a function name is not a reference');
  assert.equal(shiftFormula('="A1 is "&A1', 1, 0), '="A1 is "&A2', 'quoted text passes through');
  assert.equal(shiftFormula("='Q1 2024'!B2", 1, 0), "='Q1 2024'!B3", 'a quoted sheet name is untouched, its cell shifts');
  assert.equal(shiftFormula('=A1', -1, 0), '=#REF!', 'pushed off the sheet is an honest error');
  assert.equal(shiftFormula('=B1', 0, -1), '=A1');
});

test('ROW and COLUMN know where they are asked from, and read references', () => {
  const s = new Spreadsheet();
  s.setCell('S', 4, 2, '=ROW()');
  s.setCell('S', 4, 3, '=COLUMN()');
  s.setCell('S', 0, 0, '=ROW(B9)');
  s.setCell('S', 0, 1, '=COLUMN(D1)');
  s.setCell('S', 0, 2, '=ROW(A3:A7)');
  s.setCell('S', 0, 3, '=ROWS(A1:A5)+COLUMNS(A1:C1)');
  s.recalculate();
  assert.equal(s.getValue('S', 4, 2), 5, 'ROW() is the asking cell, 1-based');
  assert.equal(s.getValue('S', 4, 3), 4);
  assert.equal(s.getValue('S', 0, 0), 9, 'ROW(B9) reads the reference, not its value');
  assert.equal(s.getValue('S', 0, 1), 4);
  assert.equal(s.getValue('S', 0, 2), 3, 'a range answers with its first row');
  assert.equal(s.getValue('S', 0, 3), 8);
});

test('OFFSET moves and resizes from its anchor, reading the resolver', () => {
  const cells = { A1: 10, A2: 20, A3: 30, B1: 1, B2: 2, B3: 3, C5: 'x' };
  assert.equal(evalIn('=OFFSET(A1,1,0)', cells), 20, 'one cell down');
  assert.equal(evalIn('=OFFSET(A1,0,1)', cells), 1, 'one cell right');
  assert.equal(evalIn('=SUM(OFFSET(A1,0,0,3,1))', cells), 60, 'a resized reference feeds an aggregate');
  assert.equal(evalIn('=SUM(OFFSET(A1:A2,1,1))', cells), 5, 'a range anchor keeps its shape');
  assert.equal(evalIn('=OFFSET(A1,-1,0)', cells).type, '#REF!', 'off the top of the sheet');
  assert.equal(evalIn('=OFFSET(A1,0,0,0,1)', cells).type, '#REF!', 'zero height is not a range');
  assert.equal(evalIn('=OFFSET(B2,0,0)', cells), 2, 'zero offset is the anchor itself');
});

test('OFFSET in its own anchor cell is not a circular reference', () => {
  const s = new Spreadsheet();
  s.setCell('S', 1, 0, 42);            // A2
  s.setCell('S', 0, 0, '=OFFSET(A1,1,0)'); // A1 reads A2, not itself
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0), 42);
});

test('INDIRECT follows a reference built from text', () => {
  const cells = { A1: 'B2', B1: 7, B2: 99, C1: 2 };
  assert.equal(evalIn('=INDIRECT("B1")', cells), 7);
  assert.equal(evalIn('=INDIRECT(A1)', cells), 99, 'the text can come from a cell');
  assert.equal(evalIn('=INDIRECT("B"&C1)', cells), 99, 'or be concatenated');
  assert.equal(evalIn('=SUM(INDIRECT("B1:B2"))', cells), 106, 'a range in text is a range');
  assert.equal(evalIn('=INDIRECT("no such thing!")', cells).type, '#REF!');
  assert.equal(evalIn('=INDIRECT("B1",FALSE)', cells).type, '#REF!', 'R1C1 style is refused honestly');
});

test('INDIRECT recalculates when what it points at changes — the volatile scope', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 1, 5);                    // B1
  s.setCell('S', 0, 0, '=INDIRECT("B1")');    // A1 — no static edge to B1
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0), 5);
  s.setCell('S', 0, 1, 12);                   // an incremental pass, not a full one
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0), 12,
    'the graph cannot see this dependency, so volatility must carry it');
  // And a DEPENDENT of the volatile cell rides along.
  s.setCell('S', 1, 0, '=A1*2');
  s.recalculate();
  s.setCell('S', 0, 1, 50);
  s.recalculate();
  assert.equal(s.getValue('S', 1, 0), 100);
});

test('ADDRESS builds an address in all four anchor modes', () => {
  assert.equal(evalIn('=ADDRESS(2,3)'), '$C$2');
  assert.equal(evalIn('=ADDRESS(2,3,2)'), 'C$2');
  assert.equal(evalIn('=ADDRESS(2,3,3)'), '$C2');
  assert.equal(evalIn('=ADDRESS(2,3,4)'), 'C2');
  assert.equal(evalIn('=ADDRESS(1,1,1,TRUE,"Data")'), 'Data!$A$1');
  assert.equal(evalIn('=ADDRESS(1,1,1,TRUE,"Q1 2024")'), "'Q1 2024'!$A$1", 'a sheet with a space is quoted');
  assert.equal(evalIn('=INDIRECT(ADDRESS(1,2))', { B1: 8 }), 8, 'the pair composes');
  assert.equal(evalIn('=ADDRESS(0,1)').type, '#VALUE!');
});

test('the function catalog and the engine agree exactly', async () => {
  const { FUNCTIONS } = await import('@rutba/formula/functions');
  const { FUNCTION_CATALOG, CATEGORIES, catalogByCategory } = await import('@rutba/formula/catalog');
  const implemented = new Set(Object.keys(FUNCTIONS));
  const listed = new Set(FUNCTION_CATALOG.map((f) => f.name));
  for (const name of implemented) assert.equal(listed.has(name), true, name + ' is implemented but not in the catalog');
  for (const name of listed) assert.equal(implemented.has(name), true, name + ' is in the catalog but not implemented');
  for (const f of FUNCTION_CATALOG) assert.equal(CATEGORIES.includes(f.category), true, f.name + ' has an unknown category');
  const grouped = catalogByCategory();
  assert.equal(grouped.reduce((n, g) => n + g.functions.length, 0), FUNCTION_CATALOG.length);
});

// -------------------------------------------------------- structured references --

const stockTable = () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, 'Item'); s.setCell('S', 0, 1, 'Qty'); s.setCell('S', 0, 2, 'Price');
  s.setCell('S', 1, 0, 'Ink'); s.setCell('S', 1, 1, 4); s.setCell('S', 1, 2, 2.5);
  s.setCell('S', 2, 0, 'Paper'); s.setCell('S', 2, 1, 6); s.setCell('S', 2, 2, 3);
  s.defineTable('Stock', {
    sheet: 'S', top: 0, left: 0, bottom: 2, right: 2,
    headerRows: 1, totalsRows: 0, columns: ['Item', 'Qty', 'Price'],
  });
  return s;
};

test('structured references parse in the forms Excel writes', () => {
  assert.deepEqual(parse('=Table1[Qty]'), {
    type: 'structref', table: 'Table1', area: 'data', thisRow: false, startCol: 'Qty', endCol: 'Qty',
  });
  assert.deepEqual(parse('=Table1[]').startCol, null, 'empty brackets are the whole data body');
  assert.equal(parse('=Table1[#Headers]').area, 'headers');
  assert.deepEqual(parse('=Table1[[#All],[Qty]]'), {
    type: 'structref', table: 'Table1', area: 'all', thisRow: false, startCol: 'Qty', endCol: 'Qty',
  }, 'the two-part form');
  const span = parse('=Table1[[Qty]:[Price]]');
  assert.equal(span.startCol, 'Qty');
  assert.equal(span.endCol, 'Price');
  const bare = parse('=[@Price]');
  assert.equal(bare.table, null, 'a bare bracket means "this table"');
  assert.equal(bare.thisRow, true);
  assert.equal(bare.startCol, 'Price');
  assert.equal(parse("=T['[odd'] col]").startCol, "[odd] col", 'quotes escape the awkward characters');
});

test('structured references resolve against the table registry', () => {
  const s = stockTable();
  s.setCell('S', 0, 4, '=SUM(Stock[Qty])');
  s.setCell('S', 1, 4, '=SUM(Stock[])');
  s.setCell('S', 2, 4, '=COUNTA(Stock[#Headers])');
  s.setCell('S', 3, 4, '=SUM(Stock[[Qty]:[Price]])');
  s.setCell('S', 1, 3, '=Stock[@Qty]*Stock[@Price]');   // D2, inside the table row
  s.setCell('S', 4, 4, '=SUM(Stock[Nope])');
  s.setCell('S', 5, 4, '=SUM(Ghost[Qty])');
  s.setCell('S', 6, 4, '=SUM(Stock[#Totals])');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 4), 10, 'a column names its data, not its header');
  assert.equal(s.getValue('S', 1, 4), 15.5, 'the data body is every column');
  assert.equal(s.getValue('S', 2, 4), 3);
  assert.equal(s.getValue('S', 3, 4), 15.5);
  assert.equal(s.getValue('S', 1, 3), 10, '@ means this row');
  assert.equal(s.getValue('S', 4, 4).type, '#REF!', 'an unknown column is a broken reference');
  assert.equal(s.getValue('S', 5, 4).type, '#NAME?', 'an unknown table is an unknown name');
  assert.equal(s.getValue('S', 6, 4).type, '#REF!', 'this table has no totals row');
  assert.equal(calculate('=[@Qty]', s.resolver(), { sheet: 'S', row: 5, col: 3 }).type, '#VALUE!',
    'a bare @ outside any table says so');
});

test('editing a table cell recalculates through the structured reference', () => {
  const s = stockTable();
  s.setCell('S', 0, 4, '=SUM(Stock[Qty])');
  s.setCell('S', 1, 3, '=Stock[@Qty]*2');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 4), 10);
  s.setCell('S', 2, 1, 16); // Paper: 6 -> 16, an incremental pass
  s.recalculate();
  assert.equal(s.getValue('S', 0, 4), 20, 'the sum saw the edit through the table');
  s.setCell('S', 1, 1, 5);
  s.recalculate();
  assert.equal(s.getValue('S', 1, 3), 10, 'the @ row formula saw its own row change');
});

test('shiftFormula leaves bracketed column labels alone', async () => {
  const { shiftFormula } = await import('@rutba/formula');
  assert.equal(shiftFormula('=SUM(Table1[Qty])', 2, 3), '=SUM(Table1[Qty])');
  assert.equal(shiftFormula('=[@B2]*A1', 1, 0), '=[@B2]*A2', 'a column NAMED B2 is a label, the cell A1 shifts');
  assert.equal(shiftFormula('=T[[#Data],[C4]]+B1', 0, 1), '=T[[#Data],[C4]]+C1');
});

// ------------------------------------------------------------ dynamic arrays --

test('an array result SPILLS into empty neighbours, and ghosts read back', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, '=SEQUENCE(3)');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0), 1);
  assert.equal(s.getValue('S', 1, 0), 2, 'a ghost displays its spilled value');
  assert.equal(s.getValue('S', 2, 0), 3);
  assert.equal(s.getInput('S', 1, 0), '', 'but holds no input of its own');
  assert.equal(s.usedBounds('S').maxRow, 2, 'the spill is used space');
  s.setCell('S', 0, 2, '=SUM(A1:A3)');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 2), 6, 'a range over ghosts sums them');
});

test('a reader of a ghost cell tracks the spill through its source', () => {
  const s = new Spreadsheet();
  s.setCell('S', 4, 1, 10); // B5
  s.setCell('S', 0, 0, '=SEQUENCE(3,1,B5)');
  s.setCell('S', 1, 2, '=A2*10'); // C2 reads the ghost A2
  s.recalculate();
  assert.equal(s.getValue('S', 1, 2), 110);
  s.setCell('S', 4, 1, 20);
  s.recalculate();
  assert.equal(s.getValue('S', 1, 0), 21, 'the ghost moved with its source');
  assert.equal(s.getValue('S', 1, 2), 210, 'and the reader of the ghost moved with it');
});

test('a blocked spill is #SPILL!, and clearing the blocker lets it flow again', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, '=SEQUENCE(3)');
  s.recalculate();
  assert.equal(s.getValue('S', 2, 0), 3);

  s.setCell('S', 1, 0, 'in the way');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0).type, '#SPILL!', 'the anchor says what happened');
  assert.equal(s.getValue('S', 2, 0), '', 'the ghosts are gone, not half-painted');

  s.setCell('S', 1, 0, '');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0), 1, 'unblocked, it spills again');
  assert.equal(s.getValue('S', 2, 0), 3);
});

test('operators broadcast over arrays, which is what array formulas are', () => {
  const cells = { A1: 1, A2: 5, A3: 2, A4: 8, A5: 3 };
  assert.equal(evalIn('=SUM((A1:A5>2)*1)', cells), 3, 'the classic counting pattern');
  assert.equal(evalIn('=SUM(A1:A5*10)', cells), 190, 'scalar extends across the array');
  assert.equal(evalIn('=SUM(A1:A2*A3:A4)', cells), 42, '1*2 + 5*8, aligned elementwise');

  const s = new Spreadsheet();
  s.setCell('S', 0, 0, '=SEQUENCE(3)*10');
  s.recalculate();
  assert.equal(s.getValue('S', 1, 0), 20, 'a broadcast result spills');
});

test('SORT, FILTER, UNIQUE and TRANSPOSE spill their answers', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, 'b'); s.setCell('S', 0, 1, 2);
  s.setCell('S', 1, 0, 'a'); s.setCell('S', 1, 1, 9);
  s.setCell('S', 2, 0, 'b'); s.setCell('S', 2, 1, 4);
  s.setCell('S', 0, 3, '=SORT(A1:B3)');            // D1: by first column
  s.setCell('S', 0, 6, '=FILTER(A1:B3,B1:B3>3)');  // G1
  s.setCell('S', 4, 0, '=UNIQUE(A1:A3)');          // A5
  s.setCell('S', 4, 3, '=TRANSPOSE(B1:B3)');       // D5
  s.recalculate();
  assert.equal(s.getValue('S', 0, 3), 'a', 'sorted rows travel whole');
  assert.equal(s.getValue('S', 0, 4), 9);
  assert.equal(s.getValue('S', 2, 3), 'b');
  assert.equal(s.getValue('S', 0, 6), 'a', 'FILTER kept the rows over 3');
  assert.equal(s.getValue('S', 1, 6), 'b');
  assert.equal(s.getValue('S', 1, 7), 4);
  assert.equal(s.getValue('S', 4, 0), 'b');
  assert.equal(s.getValue('S', 5, 0), 'a');
  assert.equal(s.getValue('S', 6, 0), '', 'two distinct values, two rows');
  assert.equal(s.getValue('S', 4, 5), 4, 'transposed sideways');
  assert.equal(evalIn('=FILTER(A1:A3,A1:A3>99)', { A1: 1, A2: 2, A3: 3 }).type, '#CALC!', 'nothing kept says so');
});

test('the two new error literals parse and print', () => {
  assert.equal(evalIn('=#SPILL!').type, '#SPILL!');
  assert.equal(evalIn('=IFERROR(#CALC!,"ok")'), 'ok');
});

test('A1# names the range a cell spills, and follows it as it grows', async () => {
  const { shiftFormula } = await import('@rutba/formula');
  const s = new Spreadsheet();
  s.setCell('S', 4, 1, 3); // B5 drives the spill height
  s.setCell('S', 0, 0, '=SEQUENCE(B5)');
  s.setCell('S', 0, 2, '=SUM(A1#)');
  s.setCell('S', 1, 2, '=COUNT(A1#)');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 2), 6, '1+2+3');
  assert.equal(s.getValue('S', 1, 2), 3);

  // The spill grows; the reader of its # range follows without being edited.
  s.setCell('S', 4, 1, 5);
  s.recalculate();
  assert.equal(s.getValue('S', 0, 2), 15, 'the sum saw the grown range');
  assert.equal(s.getValue('S', 1, 2), 5);

  // A cell that is not a spill anchor has no # range to give.
  s.setCell('S', 6, 0, 42);
  s.setCell('S', 6, 2, '=SUM(A7#)');
  s.recalculate();
  assert.equal(s.getValue('S', 6, 2).type, '#REF!');

  // A BLOCKED spill has none either — the anchor is #SPILL!, its # is #REF!.
  s.setCell('S', 2, 0, 'in the way');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0).type, '#SPILL!');
  assert.equal(s.getValue('S', 0, 2).type, '#REF!');

  // And shiftFormula moves the anchor, keeping the operator.
  assert.equal(shiftFormula('=SUM(A1#)*2', 1, 1), '=SUM(B2#)*2');
});

test('IF broadcasts over an array condition, and stays lazy for a scalar one', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, 1); s.setCell('S', 1, 0, 5); s.setCell('S', 2, 0, 3);
  s.setCell('S', 0, 2, '=IF(A1:A3>2,"big","small")');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 2), 'small', '1 is not big');
  assert.equal(s.getValue('S', 1, 2), 'big', 'the condition spilled its answers');
  assert.equal(s.getValue('S', 2, 2), 'big');

  // Arrays on the BRANCHES broadcast too.
  s.setCell('S', 0, 4, '=SUM(IF(A1:A3>2,A1:A3,0))');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 4), 8, 'the classic conditional sum');

  // The scalar path still refuses to evaluate the branch it did not take.
  assert.equal(evalIn('=IF(A1=0, 0, 1/A1)', { A1: 0 }), 0);
});

test('UNIQUE and SORT run sideways with by_col', () => {
  // Column C is byte-for-byte column A — UNIQUE compares WHOLE columns.
  const cells = { A1: 'b', B1: 'a', C1: 'b', A2: 2, B2: 9, C2: 2 };
  const s = new Spreadsheet();
  for (const [ref, v] of Object.entries(cells)) s.setByRef('S', ref, v);
  s.setCell('S', 4, 0, '=UNIQUE(A1:C2,TRUE)');
  s.setCell('S', 7, 0, '=SORT(A1:C2,1,1,TRUE)');
  s.recalculate();
  assert.equal(s.getValue('S', 4, 0), 'b');
  assert.equal(s.getValue('S', 4, 1), 'a');
  assert.equal(s.getValue('S', 4, 2), '', 'two distinct columns, two spilled');
  assert.equal(s.getValue('S', 5, 1), 9, 'the column came down whole');
  // SORT by column, keyed on row 1: a, b, b.
  assert.equal(s.getValue('S', 7, 0), 'a');
  assert.equal(s.getValue('S', 8, 0), 9);
  assert.equal(s.getValue('S', 7, 1), 'b');
});

test('IRR finds the rate at which the flows net to nothing, and RATE solves PMT for the rate', () => {
  // A 1,000 outlay returning 500 a year for three years: about 23.4%.
  const irr = evalIn('=IRR(A1:A4)', { A1: -1000, A2: 500, A3: 500, A4: 500 });
  assert.ok(Math.abs(irr - 0.23375) < 1e-4, `IRR ${irr}`);
  assert.ok(Math.abs(evalIn('=NPV(IRR(A1:A4), A2:A4) + A1', { A1: -1000, A2: 500, A3: 500, A4: 500 })) < 1e-6, 'NPV at the IRR is zero');
  assert.equal(String(evalIn('=IRR(A1:A3)', { A1: 100, A2: 100, A3: 100 })), '#NUM!', 'no sign change, no rate');
  // The rate behind a 60-month, 10,000 loan at 193.33 a month: 0.5% a month.
  const rate = evalIn('=RATE(60, -193.33, 10000)');
  assert.ok(Math.abs(rate - 0.005) < 1e-5, `RATE ${rate}`);
  const back = evalIn('=PMT(RATE(60, -193.33, 10000), 60, 10000)');
  assert.ok(Math.abs(back + 193.33) < 1e-6, 'PMT at that rate gives the payment back');
  assert.equal(String(evalIn('=RATE(0, -100, 1000)')), '#NUM!');
});

test('DATEDIF counts the way Excel counts, unit by unit', () => {
  const d = (unit) => evalIn(`=DATEDIF(DATE(2024,1,31), DATE(2026,3,15), "${unit}")`);
  assert.equal(d('Y'), 2);
  assert.equal(d('M'), 25, 'complete months: the 15th is before the 31st');
  assert.equal(d('D'), 774);
  assert.equal(d('YM'), 1);
  assert.equal(d('MD'), 12, "days ignoring months and years: Excel borrows February's 28, so 28 - 31 + 15");
  assert.equal(d('YD'), 43);
  assert.equal(evalIn('=DATEDIF(DATE(2026,1,1), DATE(2026,1,1), "D")'), 0);
  assert.equal(String(evalIn('=DATEDIF(DATE(2026,2,1), DATE(2026,1,1), "D")')), '#NUM!', 'start after end');
  assert.equal(String(evalIn('=DATEDIF(DATE(2026,1,1), DATE(2026,2,1), "W")')), '#NUM!', 'no such unit');
});

test("Excel's _xlfn. prefix is the file's, not the function's, and CEILING.MATH rounds the 2013 way", () => {
  assert.equal(evalIn('=_xlfn.CEILING.MATH(450*12/220)'), 25, 'the prefix is stripped before the lookup');
  assert.equal(evalIn('=CEILING.MATH(4.3)'), 5);
  assert.equal(evalIn('=CEILING.MATH(-4.3)'), -4, 'a negative rounds toward zero');
  assert.equal(evalIn('=CEILING.MATH(-4.3, 1, 1)'), -5, 'unless mode says away from it');
  assert.equal(evalIn('=CEILING.MATH(7, -3)'), 9, 'the sign of the significance never matters');
  assert.equal(evalIn('=FLOOR.MATH(4.7)'), 4);
  assert.equal(evalIn('=FLOOR.MATH(-4.3)'), -5, 'a negative rounds away from zero');
  assert.equal(evalIn('=FLOOR.MATH(-4.3, 1, 1)'), -4, 'unless mode says toward it');
  assert.equal(String(evalIn('=_xlfn.NOSUCHFUNCTION(1)')), '#NAME?', 'an unknown function is still unknown');
});

test('a range wholly past the used cells is blank, not the last row that holds something', () => {
  // Ranges are clipped to the used area so that A:A does not walk a million
  // rows. Clipping the END of a range that starts beyond the used area pulled
  // it back before its start, and the backwards-range rule then normalised it
  // onto the boundary: on a sheet whose last used row is 5, =SUM(D7:D9) came
  // back with the contents of D5.
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setCell('S', 0, 3, 20); // D1
  s.setCell('S', 4, 3, 7); // D5, the last used row
  s.setCell('S', 0, 0, '=SUM(D7:D9)');
  s.setCell('S', 1, 0, '=COUNT(D7:D9)');
  s.setCell('S', 2, 0, '=SUM(F1:F3)'); // a column past the used ones
  s.setCell('S', 3, 0, '=SUM(D9:D7)'); // the same, written backwards
  s.setCell('S', 4, 0, '=SUM(D1:D9)'); // reaching past the end still sums what is there
  s.setCell('S', 5, 0, '=SUM(D:D)'); // and a whole column is still clipped
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0), 0, 'rows past the end hold nothing');
  assert.equal(s.getValue('S', 1, 0), 0, 'and count as nothing');
  assert.equal(s.getValue('S', 2, 0), 0, 'a column past the end likewise');
  assert.equal(s.getValue('S', 3, 0), 0, 'written backwards or forwards');
  assert.equal(s.getValue('S', 4, 0), 27, 'a range that does cover the data still sums it');
  assert.equal(s.getValue('S', 5, 0), 27, 'a whole column still sums what it holds');
});

test('a reference to an empty cell is nothing, not a zero', () => {
  // Blanks inside a range were skipped and a lone one was not: =COUNT(A1) on
  // an empty cell answered 1, and =AVERAGE(A1) answered 0 where Excel answers
  // #DIV/0!. The difference matters wherever a sheet asks whether a cell has
  // been filled in yet.
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setCell('S', 0, 0, 5); // A1
  const at = (formula) => {
    s.setCell('S', 2, 0, formula);
    s.recalculate();
    return s.getValue('S', 2, 0);
  };
  assert.equal(at('=COUNT(F1)'), 0, 'an empty cell counts as nothing');
  assert.equal(at('=COUNT(A1)'), 1, 'a number still counts');
  assert.equal(at('=COUNT(A1,F1)'), 1);
  assert.equal(String(at('=AVERAGE(F1)')), '#DIV/0!', 'there is nothing to average');
  assert.equal(at('=AVERAGE(A1,F1)'), 5, 'and a blank is not a zero in the mean');
  assert.equal(at('=SUM(A1,F1)'), 5);
  assert.equal(at('=MIN(A1,F1)'), 5, 'nor in the minimum');
  assert.equal(at('=SUM(1,2,3)'), 6, 'ordinary arguments are untouched');
});

test('an array constant is a grid, and a ragged one is padded the way Excel pads it', () => {
  // {1,2;3,4} — commas across, semicolons down. The semicolon is the awkward
  // part: outside braces it is what a great many locales type instead of a
  // comma between arguments, and it has meant that here since the beginning.
  assert.equal(evalIn('=SUM({1;2;3})'), 6);
  assert.equal(evalIn('=SUM({1,2;3,4})'), 10);
  assert.equal(evalIn('=COUNT({1;2;3})'), 3);
  assert.equal(evalIn('=MAX({1,9;3,4})'), 9);
  assert.equal(evalIn('=INDEX({1,2;3,4},2,1)'), 3, 'two rows down, one across');
  assert.equal(evalIn('=SUM({"a";"b"})'), 0, 'text in an array is skipped, as in a range');
  // A row shorter than the widest is padded with #N/A rather than with zero,
  // so the hole is visible.
  assert.equal(String(evalIn('=SUM({1,2;3})')), '#N/A');
  // And a semicolon between arguments still separates arguments.
  assert.equal(evalIn('=SUM(1;2;3)'), 6);
});

test('LET names a value and then uses it', () => {
  assert.equal(evalIn('=LET(x,2,x*3)'), 6);
  assert.equal(evalIn('=LET(x,2,y,x*5,x+y)'), 12, 'a later value may use an earlier name');
  assert.equal(evalIn('=LET(rate,0.2,total,50,total*rate)'), 10);
  // A LET inside a LET has its own names and gives them back afterwards.
  assert.equal(evalIn('=LET(x,1,LET(x,2,x)+x)'), 3);
  assert.equal(evalIn('=LET(x,2,x)+LET(y,3,y)'), 5);
  // The shape of the call is checked before anything is evaluated.
  assert.equal(String(evalIn('=LET(x,2)')), '#VALUE!', 'a name with no calculation');
  assert.equal(String(evalIn('=LET(x,2,y,3)')), '#VALUE!', 'an even number of arguments');
  assert.equal(String(evalIn('=LET(1,2,3)')), '#VALUE!', 'the name has to be a name');
});
