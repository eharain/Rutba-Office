// Lotus 1-2-3 worksheets: .wks (Release 1A, and Microsoft Works) and .wk1
// (Release 2), then .wk3 and .wk4 (Releases 3 to 5, several sheets in one
// file). A file is records, each a 16-bit type and length. A WK1 formula is
// kept in reverse Polish order, its functions Lotus's own (@SUM, @IF,
// @VLOOKUP ...); each is put into the spreadsheet's form here, an offset
// Lotus counts from 0 counted from 1 where the spreadsheet does (CHOOSE,
// VLOOKUP, MID, INDEX ...), and the argument orders that differ (PMT, IRR,
// ATAN2 ...) put right. Read into the old spreadsheet readers' sheets
// (legacy-grid.js).
//
// The record and formula layouts are those Lotus published for Release 2
// and those long since worked out for the later releases.

import { decodeIn } from './codepage.js';
import { Grid, cellRef } from './legacy-grid.js';

/** What a BOF record's version says wrote the file. */
const VERSIONS = { 0x0404: 'Lotus 1-2-3 1A', 0x0405: 'Symphony', 0x0406: 'Lotus 1-2-3 2', 0x1000: 'Lotus 1-2-3 3', 0x1002: 'Lotus 1-2-3 4', 0x1003: 'Lotus 1-2-3 4', 0x1005: 'Lotus 1-2-3 5' };

/** Whether bytes are a Lotus worksheet: a BOF record with a known version. */
export function lotusKind(bytes) {
  if (bytes.length < 6 || bytes[0] !== 0 || bytes[1] !== 0) return null;
  const length = bytes[2] | (bytes[3] << 8);
  const version = bytes[4] | (bytes[5] << 8);
  if (length < 2 || length > 32) return null;
  if (version >= 0x0404 && version <= 0x0406) return 'wk1';
  if (version >= 0x1000 && version <= 0x1005) return 'wk3';
  return null;
}

/* ── cell formats ─────────────────────────────────────────────────── */

/** A WK1 format byte as an .xlsx number format, or null for General. */
export function formatOf(byte) {
  const type = (byte >> 4) & 0x07;
  const n = byte & 0x0f;
  const dec = n ? `.${'0'.repeat(n)}` : '';
  switch (type) {
    case 0: return `0${dec}`;
    case 1: return `0${dec}E+00`;
    case 2: return `$#,##0${dec}`;
    case 3: return `0${dec}%`;
    case 4: return `#,##0${dec}`;
    case 7: return [null, null, 'dd-mmm-yy', 'dd-mmm', 'mmm-yy', null, ';;;', 'hh:mm:ss AM/PM', 'hh:mm AM/PM', 'dd/mm/yy', 'dd/mm', 'hh:mm:ss', 'hh:mm', null, null, null][n] ?? null;
    default: return null;
  }
}

/** A label's look from its prefix: ' left, " right, ^ centred, \ repeated. */
function label(text) {
  const p = text[0];
  if (p === "'") return { text: text.slice(1), align: null };
  if (p === '"') return { text: text.slice(1), align: 'right' };
  if (p === '^') return { text: text.slice(1), align: 'center' };
  if (p === '\\') return { text: text.slice(1).repeat(Math.max(1, Math.ceil(9 / Math.max(1, text.length - 1)))), align: null };
  if (p === '|') return { text: text.slice(1), align: null };
  return { text, align: null };
}

const cstring = (b, at, end, page) => {
  let e = at;
  while (e < end && b[e] !== 0) e++;
  return decodeIn(page, b.subarray(at, e));
};

/* ── WK1 formulas ─────────────────────────────────────────────────── */

// [Lotus opcode]: [name, argument count (-1: a count byte follows), how it is written]
const FUNCS = {
  0x1f: ['NA', 0], 0x20: ['ERR', 0], 0x21: ['ABS', 1], 0x22: ['INT', 1], 0x23: ['SQRT', 1], 0x24: ['LOG', 1], 0x25: ['LN', 1],
  0x26: ['PI', 0], 0x27: ['SIN', 1], 0x28: ['COS', 1], 0x29: ['TAN', 1], 0x2a: ['ATAN2', 2], 0x2b: ['ATAN', 1], 0x2c: ['ASIN', 1],
  0x2d: ['ACOS', 1], 0x2e: ['EXP', 1], 0x2f: ['MOD', 2], 0x30: ['CHOOSE', -1], 0x31: ['ISNA', 1], 0x32: ['ISERR', 1],
  0x33: ['FALSE', 0], 0x34: ['TRUE', 0], 0x35: ['RAND', 0], 0x36: ['DATE', 3], 0x37: ['TODAY', 0], 0x38: ['PMT', 3],
  0x39: ['PV', 3], 0x3a: ['FV', 3], 0x3b: ['IF', 3], 0x3c: ['DAY', 1], 0x3d: ['MONTH', 1], 0x3e: ['YEAR', 1], 0x3f: ['ROUND', 2],
  0x40: ['TIME', 3], 0x41: ['HOUR', 1], 0x42: ['MINUTE', 1], 0x43: ['SECOND', 1], 0x44: ['ISNUMBER', 1], 0x45: ['ISSTRING', 1],
  0x46: ['LENGTH', 1], 0x47: ['VALUE', 1], 0x48: ['STRING', 2], 0x49: ['MID', 3], 0x4a: ['CHAR', 1], 0x4b: ['CODE', 1],
  0x4c: ['FIND', 3], 0x4d: ['DATEVALUE', 1], 0x4e: ['TIMEVALUE', 1], 0x4f: ['CELLPOINTER', 1], 0x50: ['SUM', -1], 0x51: ['AVG', -1],
  0x52: ['COUNT', -1], 0x53: ['MIN', -1], 0x54: ['MAX', -1], 0x55: ['VLOOKUP', 3], 0x56: ['NPV', 2], 0x57: ['VAR', -1],
  0x58: ['STD', -1], 0x59: ['IRR', 2], 0x5a: ['HLOOKUP', 3], 0x5b: ['DSUM', 3], 0x5c: ['DAVG', 3], 0x5d: ['DCOUNT', 3],
  0x5e: ['DMIN', 3], 0x5f: ['DMAX', 3], 0x60: ['DVAR', 3], 0x61: ['DSTD', 3], 0x62: ['INDEX', 3], 0x63: ['COLS', 1],
  0x64: ['ROWS', 1], 0x65: ['REPEAT', 2], 0x66: ['UPPER', 1], 0x67: ['LOWER', 1], 0x68: ['LEFT', 2], 0x69: ['RIGHT', 2],
  0x6a: ['REPLACE', 4], 0x6b: ['PROPER', 1], 0x6c: ['CELL', 2], 0x6d: ['TRIM', 1], 0x6e: ['CLEAN', 1], 0x6f: ['S', 1],
  0x70: ['N', 1], 0x71: ['EXACT', 2], 0x73: ['@', 1], 0x74: ['RATE', 3], 0x75: ['TERM', 3], 0x76: ['CTERM', 3],
  0x77: ['SLN', 3], 0x78: ['SYD', 4], 0x79: ['DDB', 4],
};

/** Each Lotus function written in the spreadsheet's own: (args) => text. Missing ones keep the value. */
const plus1 = (s) => (/^\d+$/.test(s) ? String(Number(s) + 1) : `${s}+1`);
const neg = (s) => (/^\d+(\.\d+)?$/.test(s) ? `-${s}` : `-(${s})`);
const AS = {
  NA: () => 'NA()', ERR: () => 'NA()', PI: () => 'PI()', FALSE: () => 'FALSE', TRUE: () => 'TRUE', RAND: () => 'RAND()', TODAY: () => 'TODAY()',
  ISERR: ([a]) => `ISERROR(${a})`, ISSTRING: ([a]) => `ISTEXT(${a})`, LENGTH: ([a]) => `LEN(${a})`,
  CHOOSE: ([i, ...rest]) => `CHOOSE(${plus1(i)},${rest.join(',')})`,
  PMT: ([prin, rate, term]) => `PMT(${rate},${term},${neg(prin)})`,
  PV: ([pmt, rate, term]) => `PV(${rate},${term},${neg(pmt)})`,
  FV: ([pmt, rate, term]) => `FV(${rate},${term},${neg(pmt)})`,
  STRING: ([x, n]) => `FIXED(${x},${n},TRUE)`,
  MID: ([s, at, n]) => `MID(${s},${plus1(at)},${n})`,
  FIND: ([a, s, at]) => `(FIND(${a},${s},${plus1(at)})-1)`,
  AVG: (a) => `AVERAGE(${a.join(',')})`, COUNT: (a) => `COUNTA(${a.join(',')})`,
  VLOOKUP: ([x, r, off]) => `VLOOKUP(${x},${r},${plus1(off)})`, HLOOKUP: ([x, r, off]) => `HLOOKUP(${x},${r},${plus1(off)})`,
  VAR: (a) => `VARP(${a.join(',')})`, STD: (a) => `STDEVP(${a.join(',')})`, IRR: ([guess, r]) => `IRR(${r},${guess})`,
  DSUM: ([i, f, c]) => `DSUM(${i},${plus1(f)},${c})`, DAVG: ([i, f, c]) => `DAVERAGE(${i},${plus1(f)},${c})`,
  DCOUNT: ([i, f, c]) => `DCOUNTA(${i},${plus1(f)},${c})`, DMIN: ([i, f, c]) => `DMIN(${i},${plus1(f)},${c})`,
  DMAX: ([i, f, c]) => `DMAX(${i},${plus1(f)},${c})`, DVAR: ([i, f, c]) => `DVARP(${i},${plus1(f)},${c})`,
  DSTD: ([i, f, c]) => `DSTDEVP(${i},${plus1(f)},${c})`,
  INDEX: ([r, c, n]) => `INDEX(${r},${plus1(n)},${plus1(c)})`, COLS: ([a]) => `COLUMNS(${a})`, REPEAT: ([s, n]) => `REPT(${s},${n})`,
  REPLACE: ([s, at, n, t]) => `REPLACE(${s},${plus1(at)},${n},${t})`, S: ([a]) => `T(${a})`, '@': ([a]) => `INDIRECT(${a})`,
  RATE: ([fv, pv, term]) => `RATE(${term},0,${neg(pv)},${fv})`, TERM: ([pmt, rate, fv]) => `NPER(${rate},${neg(pmt)},0,${fv})`,
  CTERM: ([rate, fv, pv]) => `NPER(${rate},0,${neg(pv)},${fv})`,
};
const UNKNOWN = new Set(['CELLPOINTER', 'CELL']);

const PREC = { cmp: 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5, unary: 6, atom: 7 };
const BINARY = { 0x09: '+', 0x0a: '-', 0x0b: '*', 0x0c: '/', 0x0d: '^', 0x0e: '=', 0x0f: '<>', 0x10: '<=', 0x11: '>=', 0x12: '<', 0x13: '>', 0x18: '&' };

/** A WK1 cell or range coordinate: relative (bit 15) as an offset from the formula's cell, else absolute. */
function coord(v, here, mask) {
  if (v & 0x8000) {
    const off = v & 0x4000 ? ((v | 0x8000) << 16) >> 16 : v & mask;
    return { at: here + off, abs: false };
  }
  return { at: v & mask, abs: true };
}

/**
 * A WK1 formula's bytes as the spreadsheet's formula text (no leading '='),
 * from the cell at (row, col), or null with the reason where it holds what
 * the spreadsheet has no word for.
 */
export function wk1Formula(bytes, row, col) {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const stack = [];
  const push = (text, prec = PREC.atom) => stack.push({ text, prec });
  const wrap = (e, need) => (e.prec < need ? `(${e.text})` : e.text);
  const ref = (at) => {
    const c = coord(v.getUint16(at, true), col, 0xff);
    const r = coord(v.getUint16(at + 2, true), row, 0x3fff);
    if (c.at < 0 || r.at < 0) return '#REF!';
    return cellRef(r.at, c.at, { absRow: r.abs, absCol: c.abs });
  };
  let i = 0;
  while (i < bytes.length) {
    const op = bytes[i++];
    if (op === 0x03) break;
    if (op === 0x00) { push(String(+v.getFloat64(i, true).toPrecision(15))); i += 8; }
    else if (op === 0x01) { push(ref(i)); i += 4; }
    else if (op === 0x02) { push(`${ref(i)}:${ref(i + 4)}`); i += 8; }
    else if (op === 0x04) { const e = stack.pop(); push(`(${e?.text ?? ''})`); }
    else if (op === 0x05) { push(String(v.getInt16(i, true))); i += 2; }
    else if (op === 0x06) {
      let e = i;
      while (e < bytes.length && bytes[e] !== 0) e++;
      push(`"${decodeIn('cp437', bytes.subarray(i, e)).replace(/"/g, '""')}"`);
      i = e + 1;
    } else if (op === 0x08 || op === 0x17) {
      const e = stack.pop();
      push(`${op === 0x08 ? '-' : '+'}${wrap(e, PREC.unary)}`, PREC.unary);
    } else if (BINARY[op]) {
      const b = stack.pop();
      const a = stack.pop();
      const sym = BINARY[op];
      const p = PREC[sym] ?? PREC.cmp;
      push(`${wrap(a, p)}${sym}${wrap(b, sym === '-' || sym === '/' || sym === '^' ? p + 1 : p)}`, p);
    } else if (op === 0x14 || op === 0x15) {
      const b = stack.pop();
      const a = stack.pop();
      push(`${op === 0x14 ? 'AND' : 'OR'}(${a.text},${b.text})`);
    } else if (op === 0x16) {
      push(`NOT(${stack.pop().text})`);
    } else if (FUNCS[op]) {
      const [name, fixed] = FUNCS[op];
      const n = fixed < 0 ? bytes[i++] : fixed;
      const args = stack.splice(stack.length - n, n).map((e) => e.text);
      if (UNKNOWN.has(name)) return { formula: null, reason: `@${name}` };
      push(AS[name] ? AS[name](args) : `${name}(${args.join(',')})`);
    } else {
      return { formula: null, reason: `code ${op.toString(16)}` };
    }
  }
  if (stack.length !== 1) return { formula: null, reason: 'an unbalanced formula' };
  return { formula: stack[0].text, reason: null };
}

/* ── reading ──────────────────────────────────────────────────────── */

/** Lotus's 80-bit numbers (Releases 3 and later): sign, 15-bit exponent, 64-bit mantissa. */
export function float80(b, at) {
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const hi = v.getUint32(at + 4, true);
  const lo = v.getUint32(at, true);
  const se = v.getUint16(at + 8, true);
  const sign = se & 0x8000 ? -1 : 1;
  const exp = se & 0x7fff;
  if (exp === 0 && hi === 0 && lo === 0) return 0;
  const mant = hi * 2 ** 32 + lo;
  return sign * (mant / 2 ** 63) * 2 ** (exp - 16383);
}

/** WK3's small numbers in two bytes: a 15-bit integer, or a factor times a 12-bit one. */
export function smallNumber(n) {
  const s = (n << 16) >> 16;
  if (s & 1) return +([5000, 500, 0.05, 0.005, 0.0005, 0.00005, 0.0625, 0.015625][(s >> 1) & 7] * (s >> 4)).toPrecision(15);
  return s >> 1;
}

/** A sheet's name as 1-2-3 shows it: A, B ... Z, AA. */
const sheetLetter = (n) => {
  let s = '';
  for (let k = n + 1; k > 0; k = Math.floor((k - 1) / 26)) s = String.fromCharCode(65 + ((k - 1) % 26)) + s;
  return s;
};

/**
 * Read a Lotus worksheet: the old spreadsheet readers' model. `codePage`
 * decodes the labels (1-2-3 kept the Lotus International Character Set,
 * whose letters with accents are where code page 850 puts most of them).
 */
export function readLotus(bytes, { name = 'Sheet1', codePage = 'cp850' } = {}) {
  const kind = lotusKind(bytes);
  if (!kind) throw new Error('This is not a Lotus 1-2-3 worksheet.');
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const grids = new Map();
  const gridFor = (sheet) => {
    if (!grids.has(sheet)) grids.set(sheet, new Grid(kind === 'wk1' ? name : sheetLetter(sheet)));
    return grids.get(sheet);
  };
  const names = [];
  const lost = new Set();
  let keptValues = 0;
  let version = null;
  let pendingString = null;
  let at = 0;
  while (at + 4 <= bytes.length) {
    const type = v.getUint16(at, true);
    const length = v.getUint16(at + 2, true);
    const d = at + 4;
    const end = d + length;
    if (end > bytes.length) break;
    if (type === 0x0001 && kind === 'wk1') break;
    if (type === 0x0000) version = v.getUint16(d, true);
    if (kind === 'wk1') {
      const g = gridFor(0);
      const fmt = bytes[d];
      const col = v.getUint16(d + 1, true);
      const row = v.getUint16(d + 3, true);
      const look = () => { const f = formatOf(fmt); return f ? { numFmt: f } : null; };
      if (type === 0x0d) g.set(row, col, v.getInt16(d + 5, true), look());
      else if (type === 0x0e) g.set(row, col, v.getFloat64(d + 5, true), look());
      else if (type === 0x0f) {
        const l = label(cstring(bytes, d + 5, end, codePage));
        g.set(row, col, l.text, l.align ? { align: l.align } : null);
      } else if (type === 0x10) {
        const value = v.getFloat64(d + 5, true);
        const size = v.getUint16(d + 13, true);
        const { formula, reason } = wk1Formula(bytes.subarray(d + 15, Math.min(end, d + 15 + size)), row, col);
        if (formula) g.set(row, col, `=${formula}`, look());
        else { g.set(row, col, value, look()); lost.add(reason); keptValues++; }
        pendingString = { row, col, formula };
      } else if (type === 0x33 && pendingString) {
        // A formula's text result: kept when the formula itself could not be.
        const text = cstring(bytes, d + 5, end, codePage);
        if (!pendingString.formula) g.set(pendingString.row, pendingString.col, text);
        pendingString = null;
      } else if (type === 0x08 && length >= 3) {
        g.width(v.getUint16(d, true), bytes[d + 2]);
      } else if (type === 0x0b && length >= 24) {
        const nm = cstring(bytes, d, d + 16, codePage);
        const [c1, r1, c2, r2] = [0, 2, 4, 6].map((k) => v.getUint16(d + 16 + k, true));
        const a = cellRef(r1, c1, { absRow: true, absCol: true });
        const b = cellRef(r2, c2, { absRow: true, absCol: true });
        if (nm) names.push({ name: nm.replace(/[^A-Za-z0-9_.]/g, '_'), ref: a === b ? a : `${a}:${b}` });
      }
    } else {
      // Releases 3 and later: a cell is its row, its sheet and its column.
      const row = v.getUint16(d, true);
      const sheet = bytes[d + 2];
      const col = bytes[d + 3];
      if (type === 0x16) {
        const l = label(cstring(bytes, d + 4, end, codePage));
        gridFor(sheet).set(row, col, l.text, l.align ? { align: l.align } : null);
      } else if (type === 0x17) {
        gridFor(sheet).set(row, col, float80(bytes, d + 4));
      } else if (type === 0x18) {
        gridFor(sheet).set(row, col, smallNumber(v.getUint16(d + 4, true)));
      } else if (type === 0x19) {
        // A formula: its value as 1-2-3 last worked it out.
        gridFor(sheet).set(row, col, float80(bytes, d + 4));
        keptValues++;
      } else if (type === 0x1a) {
        // A formula's text result.
        gridFor(sheet).set(row, col, cstring(bytes, d + 4, end, codePage));
      }
    }
    at = end;
  }
  const notes = [];
  if (keptValues) {
    notes.push(kind === 'wk1'
      ? `${keptValues} formula${keptValues === 1 ? ' was' : 's were'} kept as ${keptValues === 1 ? 'its value' : 'their values'}: they use ${[...lost].join(', ')}, which the spreadsheet has no word for.`
      : `The ${keptValues} formula${keptValues === 1 ? ' is' : 's are'} kept as the values 1-2-3 last worked out.`);
  }
  const sheets = [...grids.keys()].sort((a, b) => a - b).map((k) => grids.get(k).toSheet());
  if (!sheets.length) sheets.push(new Grid(name).toSheet());
  return { sheets, names, notes, version: VERSIONS[version] || 'Lotus 1-2-3' };
}
