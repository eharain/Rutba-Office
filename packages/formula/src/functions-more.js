/**
 * The rest of the everyday library: trigonometry, integer maths, percentiles,
 * working days, number bases, and the functions that look at an error rather
 * than pass it on. Each was held to what Excel itself answers
 * (tests/excel-oracle.test.js).
 */
import {
  ERR, isError, isEmpty, firstError, toNumber, toText, serialToParts, partsToSerial,
  parseDateText, parseTimeText, roundHalfAwayFromZero,
} from './values.js';

const def = (fn, opts = {}) => ({ fn, ...opts });
const first = (v) => (Array.isArray(v) ? v.flat(Infinity)[0] : v);
const num = (v) => toNumber(first(v));
const txt = (v) => toText(first(v));
const vec = (v) => (Array.isArray(v) ? v.flat(Infinity) : [v]);
const opt = (v, fallback) => (v === undefined || isEmpty(first(v)) ? fallback : num(v));

/** A one-number function; an answer that is not a finite number is #NUM!. */
const math = (f) => def((v) => {
  const n = num(v);
  if (isError(n)) return n;
  const r = f(n);
  return Number.isFinite(r) ? r : ERR.NUM();
});

/** Numbers from a range: text, booleans and blanks there are not counted. */
function numbersIn(v) {
  const out = [];
  for (const x of vec(v)) {
    if (isError(x)) return x;
    if (typeof x === 'number') out.push(x);
  }
  return out;
}

/** Pairs where both sides are numbers; a pair with text or a blank is left out. */
function pairsIn(ys, xs) {
  const a = vec(ys);
  const b = vec(xs);
  if (a.length !== b.length) return ERR.NA('the two ranges must be the same size');
  const out = [];
  for (let i = 0; i < a.length; i++) {
    if (isError(a[i])) return a[i];
    if (isError(b[i])) return b[i];
    if (typeof a[i] === 'number' && typeof b[i] === 'number') out.push([a[i], b[i]]);
  }
  return out;
}

function line(ys, xs) {
  const pairs = pairsIn(ys, xs);
  if (isError(pairs)) return pairs;
  const n = pairs.length;
  if (!n) return ERR.DIV0();
  const my = pairs.reduce((s, p) => s + p[0], 0) / n;
  const mx = pairs.reduce((s, p) => s + p[1], 0) / n;
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (const [y, x] of pairs) {
    sxy += (x - mx) * (y - my);
    sxx += (x - mx) ** 2;
    syy += (y - my) ** 2;
  }
  return { my, mx, sxy, sxx, syy };
}

function percentile(range, k, exclusive) {
  const nums = numbersIn(range);
  if (isError(nums)) return nums;
  const p = num(k);
  if (isError(p)) return p;
  nums.sort((a, b) => a - b);
  const n = nums.length;
  if (!n) return ERR.NUM();
  const at = exclusive ? p * (n + 1) - 1 : p * (n - 1);
  if (p < 0 || p > 1 || at < 0 || at > n - 1) return ERR.NUM();
  const low = Math.floor(at);
  return low === at ? nums[low] : nums[low] + (at - low) * (nums[low + 1] - nums[low]);
}

function quartile(range, q, exclusive) {
  const n = num(q);
  if (isError(n)) return n;
  const k = Math.trunc(n);
  if (exclusive ? k < 1 || k > 3 : k < 0 || k > 4) return ERR.NUM();
  return percentile(range, k / 4, exclusive);
}

/** MAXA, MINA and AVERAGEA: in a range text is 0 and TRUE is 1; blanks are not counted. */
function allValues(args) {
  const out = [];
  for (const arg of args) {
    if (Array.isArray(arg)) {
      for (const v of arg.flat(Infinity)) {
        if (isError(v)) return v;
        if (isEmpty(v)) continue;
        if (typeof v === 'number') out.push(v);
        else if (typeof v === 'boolean') out.push(v ? 1 : 0);
        else out.push(0);
      }
    } else {
      if (isEmpty(arg)) continue;
      const n = toNumber(arg);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
}

const grouped = (ints) => ints.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

function fixedText(value, places, noCommas) {
  const r = roundHalfAwayFromZero(Math.abs(value), places);
  const [ints, frac] = r.toFixed(Math.max(0, places)).split('.');
  return (noCommas ? ints : grouped(ints)) + (frac ? '.' + frac : '');
}

/** Which days are the weekend, Sunday first, from a number or a Monday-first mask. */
function weekendOf(w) {
  const raw = first(w);
  if (typeof raw === 'string' && /^[01]{7}$/.test(raw)) {
    const off = Array(7).fill(false);
    for (let i = 0; i < 7; i++) off[(i + 1) % 7] = raw[i] === '1';
    return off.every(Boolean) ? ERR.VALUE('every day is a weekend') : off;
  }
  const n = opt(w, 1);
  if (isError(n)) return n;
  const off = Array(7).fill(false);
  if (n >= 1 && n <= 7) {
    off[(n + 5) % 7] = true;
    off[(n + 6) % 7] = true;
  } else if (n >= 11 && n <= 17) off[n - 11] = true;
  else return ERR.NUM('that weekend number does not exist');
  return off;
}

function holidaysIn(h) {
  if (h === undefined) return new Set();
  const out = new Set();
  for (const v of vec(h)) {
    if (isEmpty(v)) continue;
    const n = toNumber(v);
    if (isError(n)) return n;
    out.add(Math.trunc(n));
  }
  return out;
}

function networkDays(a, b, weekend, holidays) {
  const s = num(a);
  const e = num(b);
  const off = weekendOf(weekend);
  const hol = holidaysIn(holidays);
  const bad = firstError([s, e, off, hol]);
  if (bad) return bad;
  const from = Math.trunc(Math.min(s, e));
  const to = Math.trunc(Math.max(s, e));
  let count = 0;
  for (let d = from; d <= to; d++) if (!off[serialToParts(d).weekday] && !hol.has(d)) count += 1;
  return s > e ? -count : count;
}

function workDay(start, days, weekend, holidays) {
  const s = num(start);
  const n = num(days);
  const off = weekendOf(weekend);
  const hol = holidaysIn(holidays);
  const bad = firstError([s, n, off, hol]);
  if (bad) return bad;
  let d = Math.trunc(s);
  let left = Math.abs(Math.trunc(n));
  const step = n < 0 ? -1 : 1;
  while (left > 0) {
    d += step;
    if (d < 1) return ERR.NUM('that date is before the calendar starts');
    if (!off[serialToParts(d).weekday] && !hol.has(d)) left -= 1;
  }
  return d;
}

function days360(a, b, european) {
  const p = serialToParts(a);
  const q = serialToParts(b);
  let d1 = p.d;
  let d2 = q.d;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    const lastOfFeb = (x) => x.m === 1 && serialToParts(partsToSerial(x.y, 2, 1) - 1).d === x.d;
    if (lastOfFeb(p)) {
      if (lastOfFeb(q)) d2 = 30;
      d1 = 30;
    }
    if (d2 === 31 && d1 >= 30) d2 = 30;
    if (d1 === 31) d1 = 30;
  }
  return (q.y - p.y) * 360 + (q.m - p.m) * 30 + (d2 - d1);
}

const leap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;

function yearFrac(a, b, basis) {
  const from = Math.min(a, b);
  const to = Math.max(a, b);
  if (from === to) return 0;
  switch (basis) {
    case 0: return days360(from, to, false) / 360;
    case 2: return (to - from) / 360;
    case 3: return (to - from) / 365;
    case 4: return days360(from, to, true) / 360;
    default: break;
  }
  const s = serialToParts(from);
  const e = serialToParts(to);
  const withinAYear = s.y === e.y || (s.y + 1 === e.y && (s.m > e.m || (s.m === e.m && s.d >= e.d)));
  if (withinAYear) {
    let length = 365;
    if (s.y === e.y ? leap(s.y) : (leap(s.y) && (s.m < 2)) || (leap(e.y) && (e.m > 1 || (e.m === 1 && e.d === 29)))) length = 366;
    return (to - from) / length;
  }
  const years = e.y - s.y + 1;
  const days = partsToSerial(e.y + 1, 0, 1) - partsToSerial(s.y, 0, 1);
  return (to - from) / (days / years);
}

/** How many bits each base's two's-complement text holds, as Excel's functions take it. */
const BITS = { 2: 10, 8: 30, 16: 40 };

function fromBase(text, radix) {
  const s = txt(text);
  if (isError(s)) return s;
  const clean = s.trim().toUpperCase();
  const digits = '0123456789ABCDEF'.slice(0, radix);
  const bits = BITS[radix];
  if (!clean || clean.length > Math.ceil(bits / Math.log2(radix)) || [...clean].some((c) => !digits.includes(c))) return ERR.NUM(`"${s}" is not a base ${radix} number`);
  const n = parseInt(clean, radix);
  return n >= 2 ** (bits - 1) ? n - 2 ** bits : n;
}

function toBase(v, radix, places) {
  const n = num(v);
  const p = places === undefined ? undefined : num(places);
  const bad = firstError([n, p]);
  if (bad) return bad;
  const value = Math.trunc(n);
  const bits = BITS[radix];
  if (value < -(2 ** (bits - 1)) || value >= 2 ** (bits - 1)) return ERR.NUM('that number does not fit');
  const text = (value < 0 ? value + 2 ** bits : value).toString(radix).toUpperCase();
  if (value < 0 || p === undefined) return text;
  const width = Math.trunc(p);
  if (width < text.length || width > Math.ceil(bits / Math.log2(radix))) return ERR.NUM('too few places for that number');
  return text.padStart(width, '0');
}

const betweenBases = (from, to) => def((v, places) => {
  const n = fromBase(v, from);
  return isError(n) ? n : toBase(n, to, places);
});

const ROMAN = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
const ERROR_NUMBERS = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#SPILL!': 9, '#CALC!': 14 };

/** LOOKUP's search: the last value not above the key, among values of the key's own kind. */
function lastNotAbove(cells, key) {
  const k = isEmpty(key) ? 0 : key;
  const fold = (v) => (typeof v === 'string' ? v.toUpperCase() : v);
  let best = -1;
  for (let i = 0; i < cells.length; i++) {
    if (typeof cells[i] !== typeof k || cells[i] === null) continue;
    if (fold(cells[i]) <= fold(k)) best = i;
    else break;
  }
  return best;
}

const AGGREGATES = ['AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P', 'MEDIAN', 'MODE.SNGL', 'LARGE', 'SMALL'];

/** @param base the library's other functions, which AGGREGATE hands its work to. */
export const moreFunctions = (base) => ({
  ISERR: def((v) => { const r = first(v()); return isError(r) && r.type !== '#N/A'; }, { lazy: true }),
  TYPE: def((v) => {
    if (Array.isArray(v)) return 64;
    if (isError(v)) return 16;
    if (typeof v === 'string') return 2;
    if (typeof v === 'boolean') return 4;
    return 1;
  }, { keepsErrors: true }),
  'ERROR.TYPE': def((v) => {
    const r = first(v);
    return isError(r) ? (ERROR_NUMBERS[r.type] ?? ERR.NA()) : ERR.NA('that is not an error');
  }, { keepsErrors: true }),
  T: def((v) => { const r = first(v); return typeof r === 'string' ? r : ''; }),
  LOOKUP: def((needle, where, result) => {
    const key = first(needle);
    if (isError(key)) return key;
    const grid = Array.isArray(where) ? (Array.isArray(where[0]) ? where : [where]) : [[where]];
    if (result !== undefined) {
      const hit = lastNotAbove(vec(where), key);
      return hit < 0 ? ERR.NA('LOOKUP found no match') : (vec(result)[hit] ?? ERR.NA());
    }
    const wide = (grid[0]?.length ?? 0) > grid.length;
    const search = wide ? grid[0] : grid.map((r) => r[0]);
    const answers = wide ? grid[grid.length - 1] : grid.map((r) => r[r.length - 1]);
    const hit = lastNotAbove(search, key);
    return hit < 0 ? ERR.NA('LOOKUP found no match') : answers[hit];
  }, { keepsErrors: true }),
  AGGREGATE: def((which, options, ...rest) => {
    const f = Math.trunc(num(which));
    const o = Math.trunc(num(options));
    if (Number.isNaN(f) || Number.isNaN(o)) return ERR.VALUE();
    if (f < 1 || f > 19 || o < 0 || o > 7) return ERR.VALUE('AGGREGATE takes a function 1 to 19 and an option 0 to 7');
    const skipErrors = [2, 3, 6, 7].includes(o);
    const data = rest.length && f >= 14 ? rest.slice(0, -1) : rest;
    const filtered = data.map((a) => (Array.isArray(a) ? a.flat(Infinity) : [a]).filter((v) => !(skipErrors && isError(v))));
    if (!skipErrors) {
      const e = firstError(filtered);
      if (e) return e;
    }
    if (f >= 14) {
      if (rest.length < 2) return ERR.VALUE('this AGGREGATE needs k');
      const k = rest[rest.length - 1];
      const pool = filtered.flat();
      if (f === 14 || f === 15) return base[AGGREGATES[f - 1]].fn(pool, k);
      return f === 16 ? percentile(pool, k, false) : f === 17 ? quartile(pool, k, false) : f === 18 ? percentile(pool, k, true) : quartile(pool, k, true);
    }
    return base[AGGREGATES[f - 1]].fn(...filtered);
  }, { keepsErrors: true }),

  SIN: math(Math.sin),
  COS: math(Math.cos),
  TAN: math(Math.tan),
  ASIN: math(Math.asin),
  ACOS: math(Math.acos),
  ATAN: math(Math.atan),
  SINH: math(Math.sinh),
  COSH: math(Math.cosh),
  TANH: math(Math.tanh),
  ASINH: math(Math.asinh),
  ACOSH: math(Math.acosh),
  ATANH: math(Math.atanh),
  DEGREES: math((x) => (x * 180) / Math.PI),
  RADIANS: math((x) => (x * Math.PI) / 180),
  ATAN2: def((x, y) => {
    const a = num(x);
    const b = num(y);
    const e = firstError([a, b]);
    if (e) return e;
    return a === 0 && b === 0 ? ERR.DIV0() : Math.atan2(b, a);
  }),
  FACT: math((x) => {
    if (x < 0) return NaN;
    let r = 1;
    for (let i = 2; i <= Math.trunc(x); i++) r *= i;
    return r;
  }),
  GCD: def((...args) => {
    const all = args.flatMap(vec).map(toNumber);
    const e = firstError(all);
    if (e) return e;
    if (all.some((n) => n < 0)) return ERR.NUM();
    const gcd = (a, b) => (b ? gcd(b, a % b) : a);
    return all.map(Math.trunc).reduce(gcd, 0);
  }),
  LCM: def((...args) => {
    const all = args.flatMap(vec).map(toNumber);
    const e = firstError(all);
    if (e) return e;
    if (all.some((n) => n < 0)) return ERR.NUM();
    const gcd = (a, b) => (b ? gcd(b, a % b) : a);
    return all.map(Math.trunc).reduce((a, b) => (a === 0 || b === 0 ? 0 : (a / gcd(a, b)) * b), 1);
  }),
  COMBIN: def((n, k) => {
    const a = Math.trunc(num(n));
    const b = Math.trunc(num(k));
    if (Number.isNaN(a) || Number.isNaN(b)) return ERR.VALUE();
    if (a < 0 || b < 0 || b > a) return ERR.NUM();
    let r = 1;
    for (let i = 1; i <= Math.min(b, a - b); i++) r = (r * (a - i + 1)) / i;
    return Math.round(r);
  }),
  PERMUT: def((n, k) => {
    const a = Math.trunc(num(n));
    const b = Math.trunc(num(k));
    if (Number.isNaN(a) || Number.isNaN(b)) return ERR.VALUE();
    if (a < 0 || b < 0 || b > a) return ERR.NUM();
    let r = 1;
    for (let i = 0; i < b; i++) r *= a - i;
    return r;
  }),
  QUOTIENT: def((a, b) => {
    const x = num(a);
    const y = num(b);
    const e = firstError([x, y]);
    if (e) return e;
    return y === 0 ? ERR.DIV0() : Math.trunc(x / y);
  }),
  SUMX2MY2: def((a, b) => {
    const p = pairsIn(a, b);
    return isError(p) ? p : p.reduce((s, [x, y]) => s + x * x - y * y, 0);
  }),
  SUMX2PY2: def((a, b) => {
    const p = pairsIn(a, b);
    return isError(p) ? p : p.reduce((s, [x, y]) => s + x * x + y * y, 0);
  }),
  SUMXMY2: def((a, b) => {
    const p = pairsIn(a, b);
    return isError(p) ? p : p.reduce((s, [x, y]) => s + (x - y) ** 2, 0);
  }),

  PERCENTILE: def((r, k) => percentile(r, k, false)),
  'PERCENTILE.INC': def((r, k) => percentile(r, k, false)),
  'PERCENTILE.EXC': def((r, k) => percentile(r, k, true)),
  QUARTILE: def((r, q) => quartile(r, q, false)),
  'QUARTILE.INC': def((r, q) => quartile(r, q, false)),
  'QUARTILE.EXC': def((r, q) => quartile(r, q, true)),
  CORREL: def((a, b) => {
    const l = line(a, b);
    if (isError(l)) return l;
    const d = Math.sqrt(l.sxx * l.syy);
    return d === 0 ? ERR.DIV0() : l.sxy / d;
  }),
  SLOPE: def((y, x) => {
    const l = line(y, x);
    if (isError(l)) return l;
    return l.sxx === 0 ? ERR.DIV0() : l.sxy / l.sxx;
  }),
  INTERCEPT: def((y, x) => {
    const l = line(y, x);
    if (isError(l)) return l;
    return l.sxx === 0 ? ERR.DIV0() : l.my - (l.sxy / l.sxx) * l.mx;
  }),
  MAXA: def((...args) => {
    const all = allValues(args);
    return isError(all) ? all : all.reduce((a, b) => (b > a ? b : a), all.length ? all[0] : 0);
  }),
  MINA: def((...args) => {
    const all = allValues(args);
    return isError(all) ? all : all.reduce((a, b) => (b < a ? b : a), all.length ? all[0] : 0);
  }),
  AVERAGEA: def((...args) => {
    const all = allValues(args);
    if (isError(all)) return all;
    return all.length ? all.reduce((a, b) => a + b, 0) / all.length : ERR.DIV0();
  }),

  CLEAN: def((v) => {
    const s = txt(v);
    return isError(s) ? s : s.replace(/[\x00-\x1f]/g, '');
  }),
  UNICHAR: def((v) => {
    const n = Math.trunc(num(v));
    if (Number.isNaN(n)) return ERR.VALUE();
    if (n < 1 || n > 0x10ffff) return ERR.VALUE('no such character');
    return String.fromCodePoint(n);
  }),
  UNICODE: def((v) => {
    const s = txt(v);
    if (isError(s)) return s;
    return s === '' ? ERR.VALUE('there is no character') : s.codePointAt(0);
  }),
  DOLLAR: def((v, d) => {
    const n = num(v);
    const p = opt(d, 2);
    const e = firstError([n, p]);
    if (e) return e;
    const body = '$' + fixedText(n, Math.trunc(p), false);
    return n < 0 ? `(${body})` : body;
  }),
  FIXED: def((v, d, noCommas) => {
    const n = num(v);
    const p = opt(d, 2);
    const e = firstError([n, p]);
    if (e) return e;
    const plain = noCommas !== undefined && first(noCommas) !== false && first(noCommas) !== 0 && !isEmpty(first(noCommas));
    return (n < 0 ? '-' : '') + fixedText(n, Math.trunc(p), plain);
  }),

  WEEKNUM: def((v, type) => {
    const n = num(v);
    const t = opt(type, 1);
    const e = firstError([n, t]);
    if (e) return e;
    const start = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 }[t];
    if (t === 21) return isoWeek(n);
    if (start === undefined) return ERR.NUM('that week type does not exist');
    const jan1 = partsToSerial(serialToParts(n).y, 0, 1);
    const offset = (serialToParts(jan1).weekday - start + 7) % 7;
    return Math.floor((Math.trunc(n) - jan1 + offset) / 7) + 1;
  }),
  ISOWEEKNUM: def((v) => {
    const n = num(v);
    return isError(n) ? n : isoWeek(n);
  }),
  DATEVALUE: def((v) => {
    const s = txt(v);
    if (isError(s)) return s;
    const n = parseDateText(s, { dropTime: true });
    return n === null ? ERR.VALUE('"' + s + '" is not a date') : n;
  }),
  TIMEVALUE: def((v) => {
    const s = txt(v);
    if (isError(s)) return s;
    const n = parseTimeText(s);
    return n === null ? ERR.VALUE('"' + s + '" is not a time') : n;
  }),
  DAYS360: def((a, b, method) => {
    const x = num(a);
    const y = num(b);
    const e = firstError([x, y]);
    if (e) return e;
    return days360(x, y, method !== undefined && Boolean(first(method)));
  }),
  YEARFRAC: def((a, b, basis) => {
    const x = num(a);
    const y = num(b);
    const k = opt(basis, 0);
    const e = firstError([x, y, k]);
    if (e) return e;
    const which = Math.trunc(k);
    if (which < 0 || which > 4) return ERR.NUM('basis is 0 to 4');
    return yearFrac(Math.trunc(x), Math.trunc(y), which);
  }),
  NETWORKDAYS: def((a, b, holidays) => networkDays(a, b, 1, holidays)),
  'NETWORKDAYS.INTL': def((a, b, weekend, holidays) => networkDays(a, b, weekend, holidays)),
  WORKDAY: def((a, n, holidays) => workDay(a, n, 1, holidays)),
  'WORKDAY.INTL': def((a, n, weekend, holidays) => workDay(a, n, weekend, holidays)),

  BIN2DEC: def((v) => fromBase(v, 2)),
  OCT2DEC: def((v) => fromBase(v, 8)),
  HEX2DEC: def((v) => fromBase(v, 16)),
  DEC2BIN: def((v, p) => toBase(v, 2, p)),
  DEC2OCT: def((v, p) => toBase(v, 8, p)),
  DEC2HEX: def((v, p) => toBase(v, 16, p)),
  BIN2HEX: betweenBases(2, 16),
  BIN2OCT: betweenBases(2, 8),
  HEX2BIN: betweenBases(16, 2),
  HEX2OCT: betweenBases(16, 8),
  OCT2BIN: betweenBases(8, 2),
  OCT2HEX: betweenBases(8, 16),
  BASE: def((v, radix, minLength) => {
    const n = num(v);
    const r = num(radix);
    const m = opt(minLength, 0);
    const e = firstError([n, r, m]);
    if (e) return e;
    if (n < 0 || r < 2 || r > 36) return ERR.NUM();
    return Math.trunc(n).toString(Math.trunc(r)).toUpperCase().padStart(Math.trunc(m), '0');
  }),
  DECIMAL: def((v, radix) => {
    const s = txt(v);
    const r = num(radix);
    const e = firstError([s, r]);
    if (e) return e;
    if (r < 2 || r > 36) return ERR.NUM();
    const digits = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'.slice(0, Math.trunc(r));
    const clean = s.trim().toUpperCase();
    if ([...clean].some((c) => !digits.includes(c))) return ERR.NUM();
    return clean === '' ? 0 : parseInt(clean, Math.trunc(r));
  }),
  ROMAN: def((v) => {
    let n = Math.trunc(num(v));
    if (Number.isNaN(n)) return ERR.VALUE();
    if (n < 0 || n > 3999) return ERR.VALUE('ROMAN takes 0 to 3999');
    let out = '';
    for (const [value, sign] of ROMAN) while (n >= value) { out += sign; n -= value; }
    return out;
  }),
  ARABIC: def((v) => {
    const s = txt(v);
    if (isError(s)) return s;
    const values = { M: 1000, D: 500, C: 100, L: 50, X: 10, V: 5, I: 1 };
    const clean = s.trim().toUpperCase();
    if ([...clean].some((c) => !values[c])) return ERR.VALUE('that is not a Roman numeral');
    let total = 0;
    for (let i = 0; i < clean.length; i++) {
      const here = values[clean[i]];
      total += here < (values[clean[i + 1]] ?? 0) ? -here : here;
    }
    return total;
  }),
});

function isoWeek(serial) {
  const n = Math.trunc(serial);
  const thursday = n - ((serialToParts(n).weekday + 6) % 7) + 3;
  const jan1 = partsToSerial(serialToParts(thursday).y, 0, 1);
  return Math.floor((thursday - jan1) / 7) + 1;
}

/** Their lines for the function picker, in the catalogue's own shape. */
export const MORE_CATALOG = [
  ['Information', 'ISERR', '(value)', 'Is it an error other than #N/A'],
  ['Information', 'TYPE', '(value)', 'The kind of value, as a number'],
  ['Information', 'ERROR.TYPE', '(error)', 'Which error, as a number'],
  ['Text', 'T', '(value)', 'The text, or nothing if it is not text'],
  ['Lookup & Reference', 'LOOKUP', '(value, range, [results])', 'Last value not above the one asked for'],
  ['Math & Trig', 'AGGREGATE', '(function, options, range, [k])', 'An aggregate that can leave out errors and hidden rows'],
  ['Math & Trig', 'SIN', '(angle)', 'Sine of an angle in radians'],
  ['Math & Trig', 'COS', '(angle)', 'Cosine of an angle in radians'],
  ['Math & Trig', 'TAN', '(angle)', 'Tangent of an angle in radians'],
  ['Math & Trig', 'ASIN', '(number)', 'Inverse sine, in radians'],
  ['Math & Trig', 'ACOS', '(number)', 'Inverse cosine, in radians'],
  ['Math & Trig', 'ATAN', '(number)', 'Inverse tangent, in radians'],
  ['Math & Trig', 'ATAN2', '(x, y)', 'The angle of a point from the x axis'],
  ['Math & Trig', 'SINH', '(number)', 'Hyperbolic sine'],
  ['Math & Trig', 'COSH', '(number)', 'Hyperbolic cosine'],
  ['Math & Trig', 'TANH', '(number)', 'Hyperbolic tangent'],
  ['Math & Trig', 'ASINH', '(number)', 'Inverse hyperbolic sine'],
  ['Math & Trig', 'ACOSH', '(number)', 'Inverse hyperbolic cosine'],
  ['Math & Trig', 'ATANH', '(number)', 'Inverse hyperbolic tangent'],
  ['Math & Trig', 'DEGREES', '(angle)', 'Radians to degrees'],
  ['Math & Trig', 'RADIANS', '(angle)', 'Degrees to radians'],
  ['Math & Trig', 'FACT', '(number)', 'Factorial'],
  ['Math & Trig', 'GCD', '(number1, …)', 'Greatest common divisor'],
  ['Math & Trig', 'LCM', '(number1, …)', 'Least common multiple'],
  ['Math & Trig', 'COMBIN', '(n, chosen)', 'Combinations'],
  ['Math & Trig', 'PERMUT', '(n, chosen)', 'Permutations'],
  ['Math & Trig', 'QUOTIENT', '(number, divisor)', 'Whole part of a division'],
  ['Math & Trig', 'SUMX2MY2', '(x, y)', 'Sum of the differences of squares'],
  ['Math & Trig', 'SUMX2PY2', '(x, y)', 'Sum of the sums of squares'],
  ['Math & Trig', 'SUMXMY2', '(x, y)', 'Sum of the squares of differences'],
  ['Math & Trig', 'ROMAN', '(number)', 'A number in Roman numerals'],
  ['Math & Trig', 'ARABIC', '(text)', 'A Roman numeral as a number'],
  ['Math & Trig', 'BASE', '(number, radix, [length])', 'A number in another base'],
  ['Math & Trig', 'DECIMAL', '(text, radix)', 'Text in another base as a number'],
  ['Statistical', 'PERCENTILE', '(range, k)', 'The k-th percentile'],
  ['Statistical', 'PERCENTILE.INC', '(range, k)', 'The k-th percentile, 0 and 1 included'],
  ['Statistical', 'PERCENTILE.EXC', '(range, k)', 'The k-th percentile, 0 and 1 left out'],
  ['Statistical', 'QUARTILE', '(range, quart)', 'A quartile'],
  ['Statistical', 'QUARTILE.INC', '(range, quart)', 'A quartile, 0 and 4 included'],
  ['Statistical', 'QUARTILE.EXC', '(range, quart)', 'A quartile, 1 to 3 only'],
  ['Statistical', 'CORREL', '(range1, range2)', 'Correlation of two ranges'],
  ['Statistical', 'SLOPE', '(y, x)', 'Slope of the fitted line'],
  ['Statistical', 'INTERCEPT', '(y, x)', 'Where the fitted line crosses y'],
  ['Statistical', 'MAXA', '(value1, …)', 'Largest, counting text as 0 and TRUE as 1'],
  ['Statistical', 'MINA', '(value1, …)', 'Smallest, counting text as 0 and TRUE as 1'],
  ['Statistical', 'AVERAGEA', '(value1, …)', 'Mean, counting text as 0 and TRUE as 1'],
  ['Text', 'CLEAN', '(text)', 'Remove characters that do not print'],
  ['Text', 'UNICHAR', '(number)', 'The character with this code'],
  ['Text', 'UNICODE', '(text)', 'The code of the first character'],
  ['Text', 'DOLLAR', '(number, [decimals])', 'A number as currency text'],
  ['Text', 'FIXED', '(number, [decimals], [no_commas])', 'A number as text with fixed decimals'],
  ['Date & Time', 'WEEKNUM', '(date, [type])', 'Week of the year'],
  ['Date & Time', 'ISOWEEKNUM', '(date)', 'ISO week of the year'],
  ['Date & Time', 'DATEVALUE', '(text)', 'A date written as text, as a date'],
  ['Date & Time', 'TIMEVALUE', '(text)', 'A time written as text, as a time'],
  ['Date & Time', 'DAYS360', '(start, end, [european])', 'Days between dates on a 360-day year'],
  ['Date & Time', 'YEARFRAC', '(start, end, [basis])', 'The part of a year between two dates'],
  ['Date & Time', 'NETWORKDAYS', '(start, end, [holidays])', 'Working days between two dates'],
  ['Date & Time', 'NETWORKDAYS.INTL', '(start, end, [weekend], [holidays])', 'Working days, with your own weekend'],
  ['Date & Time', 'WORKDAY', '(start, days, [holidays])', 'The date some working days away'],
  ['Date & Time', 'WORKDAY.INTL', '(start, days, [weekend], [holidays])', 'The date some working days away, with your own weekend'],
  ['Math & Trig', 'BIN2DEC', '(text)', 'Binary to decimal'],
  ['Math & Trig', 'OCT2DEC', '(text)', 'Octal to decimal'],
  ['Math & Trig', 'HEX2DEC', '(text)', 'Hexadecimal to decimal'],
  ['Math & Trig', 'DEC2BIN', '(number, [places])', 'Decimal to binary'],
  ['Math & Trig', 'DEC2OCT', '(number, [places])', 'Decimal to octal'],
  ['Math & Trig', 'DEC2HEX', '(number, [places])', 'Decimal to hexadecimal'],
  ['Math & Trig', 'BIN2HEX', '(text, [places])', 'Binary to hexadecimal'],
  ['Math & Trig', 'BIN2OCT', '(text, [places])', 'Binary to octal'],
  ['Math & Trig', 'HEX2BIN', '(text, [places])', 'Hexadecimal to binary'],
  ['Math & Trig', 'HEX2OCT', '(text, [places])', 'Hexadecimal to octal'],
  ['Math & Trig', 'OCT2BIN', '(text, [places])', 'Octal to binary'],
  ['Math & Trig', 'OCT2HEX', '(text, [places])', 'Octal to hexadecimal'],
];
