// Lotus 1-2-3 worksheets (the 2.8 plan's old spreadsheets): WK1 labels with
// their alignment, numbers and formats, column widths and named ranges, and
// formulas from Lotus's reverse Polish @functions into the spreadsheet's own,
// offsets counted from 1 and argument orders put right; WK3's sheets, 80-bit
// numbers and small numbers.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readLotus, lotusKind, wk1Formula, formatOf, float80, smallNumber } from '../packages/office-formats/src/lotus.js';

const rec = (type, body) => { const b = Buffer.alloc(4 + body.length); b.writeUInt16LE(type, 0); b.writeUInt16LE(body.length, 2); body.copy(b, 4); return b; };
const u16 = (n) => { const b = Buffer.alloc(2); b.writeUInt16LE(n & 0xffff); return b; };
const f64 = (n) => { const b = Buffer.alloc(8); b.writeDoubleLE(n); return b; };
const cell = (fmt, col, row) => Buffer.concat([Buffer.from([fmt]), u16(col), u16(row)]);
/** A WK1 coordinate: absolute, or relative by `off` from the formula's cell. */
const rel = (off) => ((off & 0x3fff) | 0x8000 | (off < 0 ? 0x4000 : 0)) & 0xffff;
const ref = (c, r) => Buffer.concat([u16(c), u16(r)]);
const fm = (...parts) => Buffer.concat(parts.map((p) => (Buffer.isBuffer(p) ? p : Buffer.from(Array.isArray(p) ? p : [p]))));
const str = (s) => Buffer.from(`${s}\0`, 'latin1');

test('WK1 formulas read into the spreadsheet\'s form', () => {
  // @SUM(A1..A2) in A3, relative.
  assert.equal(wk1Formula(fm(0x02, ref(rel(0), rel(-2)), ref(rel(0), rel(-1)), [0x50, 1], 0x03), 2, 0).formula, 'SUM(A1:A2)');
  // @IF($A$1>5,"big","small")
  assert.equal(wk1Formula(fm(0x01, ref(0, 0), 0x05, u16(5), 0x13, 0x06, str('big'), 0x06, str('small'), 0x3b, 0x03), 4, 4).formula, 'IF($A$1>5,"big","small")');
  // (B1+2)*3 with the parentheses Lotus kept.
  assert.equal(wk1Formula(fm(0x01, ref(rel(1), rel(0)), 0x05, u16(2), 0x09, 0x04, 0x05, u16(3), 0x0b, 0x03), 0, 0).formula, '(B1+2)*3');
  // @CHOOSE(1,"a","b") counts from 0 in Lotus.
  assert.equal(wk1Formula(fm(0x05, u16(1), 0x06, str('a'), 0x06, str('b'), [0x30, 3], 0x03), 0, 0).formula, 'CHOOSE(2,"a","b")');
  // @PMT(10000,0.01,12) is PMT(rate, term, -principal).
  assert.equal(wk1Formula(fm(0x05, u16(10000), 0x00, f64(0.01), 0x05, u16(12), 0x38, 0x03), 0, 0).formula, 'PMT(0.01,12,-10000)');
  // @VLOOKUP(x,range,1): the second column.
  assert.equal(wk1Formula(fm(0x01, ref(0, 0), 0x02, ref(1, 0), ref(2, 9), 0x05, u16(1), 0x55, 0x03), 0, 5).formula, 'VLOOKUP($A$1,$B$1:$C$10,2)');
  // #AND# and unary minus, and a power that keeps its order.
  assert.equal(wk1Formula(fm(0x01, ref(0, 0), 0x01, ref(1, 0), 0x14, 0x03), 0, 0).formula, 'AND($A$1,$B$1)');
  assert.equal(wk1Formula(fm(0x05, u16(2), 0x05, u16(3), 0x05, u16(2), 0x0d, 0x0d, 0x03), 0, 0).formula, '2^(3^2)');
  // A function the spreadsheet has no word for.
  assert.equal(wk1Formula(fm(0x06, str('row'), 0x01, ref(0, 0), 0x6c, 0x03), 0, 0).formula, null);
});

test('a WK1 file reads its labels, numbers, formats, formulas, widths and names', () => {
  const file = Buffer.concat([
    rec(0x00, u16(0x0406)),
    rec(0x08, Buffer.concat([u16(0), Buffer.from([18])])),
    rec(0x0f, Buffer.concat([cell(0xff, 0, 0), str("'Item")])),
    rec(0x0f, Buffer.concat([cell(0xff, 1, 0), str('"Price')])),
    rec(0x0f, Buffer.concat([cell(0xff, 0, 1), str('^Tea')])),
    rec(0x0e, Buffer.concat([cell(0x22, 1, 1), f64(3.5)])),
    rec(0x0d, Buffer.concat([cell(0x22, 1, 2), u16(12)])),
    rec(0x10, (() => { const f = fm(0x02, ref(rel(0), rel(-2)), ref(rel(0), rel(-1)), [0x50, 1], 0x03); return Buffer.concat([cell(0x22, 1, 3), f64(15.5), u16(f.length), f]); })()),
    rec(0x10, (() => { const f = fm(0x06, str('row'), 0x01, ref(0, 0), 0x6c, 0x03); return Buffer.concat([cell(0xff, 2, 3), f64(0), u16(f.length), f]); })()),
    rec(0x0b, Buffer.concat([Buffer.from('PRICES\0\0\0\0\0\0\0\0\0\0', 'latin1'), u16(1), u16(1), u16(1), u16(2)])),
    rec(0x01, Buffer.alloc(0)),
  ]);
  const bytes = new Uint8Array(file);
  assert.equal(lotusKind(bytes), 'wk1');
  const { sheets, names, notes, version } = readLotus(bytes);
  const [s] = sheets;
  assert.equal(version, 'Lotus 1-2-3 2');
  assert.deepEqual(s.rows[0], ['Item', 'Price']);
  assert.equal(s.styles.B1.align, 'right');
  assert.equal(s.styles.A2.align, 'center');
  assert.equal(s.rows[1][1], 3.5);
  assert.equal(s.styles.B2.numFmt, '$#,##0.00');
  assert.equal(s.rows[2][1], 12);
  assert.equal(s.rows[3][1], '=SUM(B2:B3)');
  assert.equal(s.rows[3][2], 0);
  assert.deepEqual(s.widths, { A: 18 });
  assert.deepEqual(names, [{ name: 'PRICES', ref: '$B$2:$B$3' }]);
  assert.match(notes[0], /1 formula was kept as its value: they use @CELL/);
});

test('number formats from the WK1 format byte', () => {
  assert.equal(formatOf(0x02), '0.00');
  assert.equal(formatOf(0x31), '0.0%');
  assert.equal(formatOf(0x40), '#,##0');
  assert.equal(formatOf(0x72), 'dd-mmm-yy');
  assert.equal(formatOf(0x71), null);
});

test('WK3: several sheets, 80-bit numbers, small numbers and labels', () => {
  const ext = (n) => { const b = Buffer.alloc(10); const e = n === 0 ? 0 : Math.floor(Math.log2(Math.abs(n))); const m = BigInt(Math.round((Math.abs(n) / 2 ** e) * 2 ** 63)); b.writeBigUInt64LE(m); b.writeUInt16LE((n < 0 ? 0x8000 : 0) | (e + 16383), 8); return b; };
  assert.equal(float80(ext(1234.5), 0), 1234.5);
  assert.equal(float80(ext(-0.25), 0), -0.25);
  assert.equal(smallNumber(42 << 1), 42);
  assert.equal(smallNumber(((-3) << 1) & 0xffff), -3);
  assert.equal(smallNumber((7 << 4) | (2 << 1) | 1), 0.35);
  const at = (row, sheet, col) => Buffer.concat([u16(row), Buffer.from([sheet, col])]);
  const file = Buffer.concat([
    rec(0x00, Buffer.concat([u16(0x1000), Buffer.alloc(24)])),
    rec(0x16, Buffer.concat([at(0, 0, 0), str("'North")])),
    rec(0x17, Buffer.concat([at(1, 0, 0), ext(1234.5)])),
    rec(0x16, Buffer.concat([at(0, 1, 0), str("'South")])),
    rec(0x18, Buffer.concat([at(1, 1, 0), u16(99 << 1)])),
    rec(0x19, Buffer.concat([at(2, 1, 0), ext(100), Buffer.from([0x03])])),
  ]);
  const bytes = new Uint8Array(file);
  assert.equal(lotusKind(bytes), 'wk3');
  const { sheets, notes } = readLotus(bytes);
  assert.deepEqual(sheets.map((s) => s.name), ['A', 'B']);
  assert.deepEqual(sheets[0].rows, [['North'], [1234.5]]);
  assert.deepEqual(sheets[1].rows, [['South'], [99], [100]]);
  assert.match(notes[0], /kept as the values 1-2-3 last worked out/);
  assert.equal(lotusKind(new Uint8Array([1, 2, 3, 4, 5, 6])), null);
});
