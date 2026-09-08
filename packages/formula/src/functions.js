/**
 * The function library.
 *
 * Two conventions run through all of it, and they are what make results match
 * what a customer built in Excel:
 *
 * 1. **Aggregates ignore text and blanks in ranges, but not in literals.**
 *    `SUM(A1:A3)` where A2 holds "hello" is the sum of A1 and A3. But
 *    `SUM(1, "hello")` is `#VALUE!`, because a literal argument was meant to be
 *    a number. `COUNT` counts numbers; `COUNTA` counts anything non-blank.
 *
 * 2. **Errors propagate through everything except the handlers.** Any error in
 *    any argument becomes the result, except in `IFERROR`, `ISERROR`, `ISNA`
 *    and `IFNA`, which exist precisely to look at one without spreading it.
 *
 * `lazy: true` marks a function whose arguments must NOT be evaluated up front —
 * `IF` must not evaluate the branch it does not take, or `=IF(A1=0,0,1/A1)`
 * would still divide by zero.
 */
import {
  ERR, FormulaError, isError, isBlank, firstError,
  toNumber, toText, toBoolean, compareValues, roundHalfAwayFromZero,
  dateToSerial, serialToDate, numericText,
} from './values.js';
import { indexToCol } from './parser.js';

/** Flatten range results into a list of scalars. */
const flatten = (args) => args.flat(Infinity);

/** An array collapsed to its first element; a scalar untouched. */
const firstOf = (v) => (Array.isArray(v) ? (v.flat(Infinity)[0] ?? '') : v);

/** Any value as a rectangular 2-D array, the shape the array functions share. */
const gridOf = (v) => (Array.isArray(v) ? (Array.isArray(v[0]) ? v : [v]) : [[v]]);

/** Rows become columns — how the by_col modes reuse the row machinery. */
const transposeGrid = (g) => {
  const out = [];
  for (let j = 0; j < (g[0]?.length ?? 0); j++) out.push(g.map((line) => line[j] ?? ''));
  return out;
};

/** Numbers usable by an aggregate, applying convention 1 above. */
function aggregateNumbers(args) {
  const out = [];
  for (const arg of args) {
    if (Array.isArray(arg)) {
      for (const v of arg.flat(Infinity)) {
        if (isError(v)) return v;
        if (typeof v === 'number') out.push(v);
        else if (typeof v === 'boolean') continue; // booleans in ranges are skipped
        // text and blanks in a range are skipped
      }
    } else {
      // A reference to an empty cell is not a zero. Blanks inside a RANGE were
      // skipped above and a lone one was not, so =COUNT(A1) on an empty cell
      // answered 1 and =AVERAGE(A1) answered 0 where Excel answers #DIV/0! —
      // the difference between "nothing here" and "nothing here, counted".
      if (isBlank(arg)) continue;
      const n = toNumber(arg);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
}

const numericArgs = (args) => {
  const nums = aggregateNumbers(args);
  return isError(nums) ? nums : nums;
};

/** One required scalar, coerced to number. */
const num1 = (v) => toNumber(Array.isArray(v) ? v.flat(Infinity)[0] : v);
const text1 = (v) => toText(Array.isArray(v) ? v.flat(Infinity)[0] : v);

const def = (fn, opts = {}) => ({ fn, ...opts });

/**
 * Criteria matching for SUMIF/COUNTIF: ">100", "<=5", "<>x", "apple", or a
 * bare value. Text comparison is case-insensitive, and `*`/`?` are wildcards.
 */
function makeMatcher(criteria) {
  const raw = Array.isArray(criteria) ? criteria.flat(Infinity)[0] : criteria;
  if (isError(raw)) return () => raw;
  const s = toText(raw);
  const m = /^(<=|>=|<>|=|<|>)(.*)$/.exec(s);
  const op = m ? m[1] : '=';
  const operandText = m ? m[2] : s;
  const operandNum = numericText(operandText);
  const operand = operandNum === null ? operandText : operandNum;

  const hasWildcard = typeof operand === 'string' && /[*?]/.test(operand);
  const wildcard = hasWildcard
    ? new RegExp('^' + operand.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i')
    : null;

  return (value) => {
    if (isError(value)) return value;
    if (wildcard && (op === '=' || op === '<>')) {
      const hit = wildcard.test(toText(value));
      return op === '=' ? hit : !hit;
    }
    const cmp = compareValues(value, operand);
    if (isError(cmp)) return cmp;
    switch (op) {
      case '=': return cmp === 0;
      case '<>': return cmp !== 0;
      case '<': return cmp < 0;
      case '<=': return cmp <= 0;
      case '>': return cmp > 0;
      case '>=': return cmp >= 0;
      default: return false;
    }
  };
}

/** One argument as a flat vector of scalars, preserving order. */
const vec = (arg) => (Array.isArray(arg) ? arg.flat(Infinity) : [arg]);

/**
 * The *IFS family: pairs of (criteria range, criteria), all the same length as
 * the target range. Returns the matching indices, or the first error met.
 */
function ifsIndices(length, pairs) {
  if (pairs.length === 0 || pairs.length % 2 !== 0) return ERR.VALUE('criteria come in range/criteria pairs');
  const tests = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const cells = vec(pairs[i]);
    if (cells.length !== length) return ERR.VALUE('criteria ranges must be the same size');
    tests.push({ cells, match: makeMatcher(pairs[i + 1]) });
  }
  const hits = [];
  for (let i = 0; i < length; i++) {
    let all = true;
    for (const t of tests) {
      const hit = t.match(t.cells[i]);
      if (isError(hit)) return hit;
      if (!hit) { all = false; break; }
    }
    if (all) hits.push(i);
  }
  return hits;
}

/** Sample/population variance over aggregate numbers. */
function variance(args, { population }) {
  const nums = aggregateNumbers(args);
  if (isError(nums)) return nums;
  const n = nums.length;
  if (population ? n < 1 : n < 2) return ERR.DIV0();
  const mean = nums.reduce((a, b) => a + b, 0) / n;
  const ss = nums.reduce((a, b) => a + (b - mean) ** 2, 0);
  return ss / (population ? n : n - 1);
}

/** Wildcard-or-equality comparison used by SEARCH and XLOOKUP's mode 2. */
function wildcardRegex(pattern) {
  return new RegExp(
    '^' + String(pattern).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$',
    'i',
  );
}

/** Excel's date parts from a serial, via the shared 1900-epoch arithmetic. */
const partsOf = (serial) => {
  const d = serialToDate(serial);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth(), day: d.getUTCDate() };
};

/** Serial of (year, monthIndex, day), clamping day into the month like EDATE does. */
function serialOfClamped(y, m, day) {
  const lastDay = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return dateToSerial(new Date(Date.UTC(y, m, Math.min(day, lastDay))));
}

export const FUNCTIONS = {
  // ---- maths and aggregation --------------------------------------------
  SUM: def((...args) => {
    const nums = numericArgs(args);
    return isError(nums) ? nums : nums.reduce((a, b) => a + b, 0);
  }),
  PRODUCT: def((...args) => {
    const nums = numericArgs(args);
    if (isError(nums)) return nums;
    return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
  }),
  AVERAGE: def((...args) => {
    const nums = numericArgs(args);
    if (isError(nums)) return nums;
    return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : ERR.DIV0();
  }),
  MIN: def((...args) => {
    const nums = numericArgs(args);
    if (isError(nums)) return nums;
    return nums.length ? Math.min(...nums) : 0;
  }),
  MAX: def((...args) => {
    const nums = numericArgs(args);
    if (isError(nums)) return nums;
    return nums.length ? Math.max(...nums) : 0;
  }),
  COUNT: def((...args) => {
    const nums = numericArgs(args);
    return isError(nums) ? nums : nums.length;
  }),
  COUNTA: def((...args) => flatten(args).filter((v) => !isBlank(v)).length),
  COUNTBLANK: def((...args) => flatten(args).filter((v) => isBlank(v)).length),
  ABS: def((v) => { const n = num1(v); return isError(n) ? n : Math.abs(n); }),
  SIGN: def((v) => { const n = num1(v); return isError(n) ? n : Math.sign(n); }),
  SQRT: def((v) => {
    const n = num1(v);
    if (isError(n)) return n;
    return n < 0 ? ERR.NUM('SQRT of a negative number') : Math.sqrt(n);
  }),
  POWER: def((a, b) => {
    const x = num1(a);
    const y = num1(b);
    const e = firstError([x, y]);
    if (e) return e;
    const r = x ** y;
    return Number.isFinite(r) ? r : ERR.NUM();
  }),
  MOD: def((a, b) => {
    const x = num1(a);
    const y = num1(b);
    const e = firstError([x, y]);
    if (e) return e;
    if (y === 0) return ERR.DIV0();
    // Spreadsheet MOD takes the sign of the divisor, unlike JS %
    return x - y * Math.floor(x / y);
  }),
  INT: def((v) => { const n = num1(v); return isError(n) ? n : Math.floor(n); }),
  TRUNC: def((v, d) => {
    const n = num1(v);
    if (isError(n)) return n;
    const digits = d === undefined ? 0 : num1(d);
    if (isError(digits)) return digits;
    const f = 10 ** digits;
    return Math.trunc(n * f) / f;
  }),
  ROUND: def((v, d) => {
    const n = num1(v);
    const digits = d === undefined ? 0 : num1(d);
    const e = firstError([n, digits]);
    return e || roundHalfAwayFromZero(n, digits);
  }),
  ROUNDUP: def((v, d) => {
    const n = num1(v);
    const digits = d === undefined ? 0 : num1(d);
    const e = firstError([n, digits]);
    if (e) return e;
    const f = 10 ** digits;
    return (n < 0 ? -Math.ceil(-n * f) : Math.ceil(n * f)) / f;
  }),
  ROUNDDOWN: def((v, d) => {
    const n = num1(v);
    const digits = d === undefined ? 0 : num1(d);
    const e = firstError([n, digits]);
    if (e) return e;
    const f = 10 ** digits;
    return (n < 0 ? -Math.floor(-n * f) : Math.floor(n * f)) / f;
  }),

  // ---- conditional aggregation ------------------------------------------
  SUMIF: def((range, criteria, sumRange) => {
    const cells = Array.isArray(range) ? range.flat(Infinity) : [range];
    const targets = sumRange === undefined
      ? cells
      : (Array.isArray(sumRange) ? sumRange.flat(Infinity) : [sumRange]);
    const match = makeMatcher(criteria);
    let total = 0;
    for (let i = 0; i < cells.length; i++) {
      const hit = match(cells[i]);
      if (isError(hit)) return hit;
      if (!hit) continue;
      const v = targets[i];
      if (typeof v === 'number') total += v;
      else if (isError(v)) return v;
    }
    return total;
  }),
  COUNTIF: def((range, criteria) => {
    const cells = Array.isArray(range) ? range.flat(Infinity) : [range];
    const match = makeMatcher(criteria);
    let n = 0;
    for (const c of cells) {
      const hit = match(c);
      if (isError(hit)) return hit;
      if (hit) n += 1;
    }
    return n;
  }),
  AVERAGEIF: def((range, criteria, avgRange) => {
    const cells = Array.isArray(range) ? range.flat(Infinity) : [range];
    const targets = avgRange === undefined
      ? cells
      : (Array.isArray(avgRange) ? avgRange.flat(Infinity) : [avgRange]);
    const match = makeMatcher(criteria);
    const picked = [];
    for (let i = 0; i < cells.length; i++) {
      const hit = match(cells[i]);
      if (isError(hit)) return hit;
      if (hit && typeof targets[i] === 'number') picked.push(targets[i]);
    }
    return picked.length ? picked.reduce((a, b) => a + b, 0) / picked.length : ERR.DIV0();
  }),

  // ---- logic --------------------------------------------------------------
  IF: def((condition, whenTrue, whenFalse) => {
    // A broadcast comparison hands IF an ARRAY condition; the first element
    // decides, which is the legacy implicit intersection — full array-IF is
    // a later step.
    const c = toBoolean(firstOf(condition()));
    if (isError(c)) return c;
    if (c) return whenTrue ? whenTrue() : true;
    return whenFalse ? whenFalse() : false;
  }, { lazy: true }),
  IFERROR: def((value, fallback) => {
    const v = value();
    return isError(v) ? (fallback ? fallback() : '') : v;
  }, { lazy: true }),
  IFNA: def((value, fallback) => {
    const v = value();
    return isError(v) && v.type === '#N/A' ? (fallback ? fallback() : '') : v;
  }, { lazy: true }),
  AND: def((...args) => {
    for (const v of flatten(args)) {
      if (isBlank(v)) continue;
      const b = toBoolean(v);
      if (isError(b)) return b;
      if (!b) return false;
    }
    return true;
  }),
  OR: def((...args) => {
    let any = false;
    for (const v of flatten(args)) {
      if (isBlank(v)) continue;
      const b = toBoolean(v);
      if (isError(b)) return b;
      if (b) any = true;
    }
    return any;
  }),
  NOT: def((v) => {
    const b = toBoolean(Array.isArray(v) ? v.flat(Infinity)[0] : v);
    return isError(b) ? b : !b;
  }),
  TRUE: def(() => true),
  FALSE: def(() => false),

  // ---- type tests ---------------------------------------------------------
  ISBLANK: def((v) => isBlank(Array.isArray(v) ? v.flat(Infinity)[0] : v)),
  ISNUMBER: def((v) => typeof (Array.isArray(v) ? v.flat(Infinity)[0] : v) === 'number'),
  ISTEXT: def((v) => typeof (Array.isArray(v) ? v.flat(Infinity)[0] : v) === 'string'),
  ISLOGICAL: def((v) => typeof (Array.isArray(v) ? v.flat(Infinity)[0] : v) === 'boolean'),
  ISERROR: def((v) => isError(v()), { lazy: true }),
  ISNA: def((v) => { const r = v(); return isError(r) && r.type === '#N/A'; }, { lazy: true }),
  NA: def(() => ERR.NA()),

  // ---- text ---------------------------------------------------------------
  CONCATENATE: def((...args) => {
    const e = firstError(args);
    if (e) return e;
    return flatten(args).map(toText).join('');
  }),
  CONCAT: def((...args) => {
    const e = firstError(args);
    if (e) return e;
    return flatten(args).map(toText).join('');
  }),
  LEN: def((v) => { const s = text1(v); return isError(s) ? s : s.length; }),
  LEFT: def((v, n) => {
    const s = text1(v);
    const count = n === undefined ? 1 : num1(n);
    const e = firstError([s, count]);
    return e || s.slice(0, Math.max(0, count));
  }),
  RIGHT: def((v, n) => {
    const s = text1(v);
    const count = n === undefined ? 1 : num1(n);
    const e = firstError([s, count]);
    if (e) return e;
    return count <= 0 ? '' : s.slice(Math.max(0, s.length - count));
  }),
  MID: def((v, start, len) => {
    const s = text1(v);
    const from = num1(start);
    const count = num1(len);
    const e = firstError([s, from, count]);
    if (e) return e;
    if (from < 1) return ERR.VALUE('MID start must be at least 1');
    return s.substr(from - 1, Math.max(0, count));
  }),
  UPPER: def((v) => { const s = text1(v); return isError(s) ? s : s.toUpperCase(); }),
  LOWER: def((v) => { const s = text1(v); return isError(s) ? s : s.toLowerCase(); }),
  TRIM: def((v) => { const s = text1(v); return isError(s) ? s : s.replace(/\s+/g, ' ').trim(); }),
  SUBSTITUTE: def((v, find, replace) => {
    const s = text1(v);
    const f = text1(find);
    const r = text1(replace);
    const e = firstError([s, f, r]);
    if (e) return e;
    return f === '' ? s : s.split(f).join(r);
  }),
  VALUE: def((v) => {
    const s = text1(v);
    if (isError(s)) return s;
    const n = numericText(s);
    return n === null ? ERR.VALUE('"' + s + '" is not a number') : n;
  }),
  TEXT: def((v, fmt) => {
    const raw = Array.isArray(v) ? v.flat(Infinity)[0] : v;
    if (isError(raw)) return raw;
    const pattern = text1(fmt);
    if (isError(pattern)) return pattern;
    const n = toNumber(raw);
    if (isError(n)) return toText(raw);
    // A deliberately small subset: thousands separators and fixed decimals.
    // Anything else returns the plain rendering rather than guessing.
    const m = /^([#,]*)0*(?:\.(0+))?$/.exec(pattern.replace(/["']/g, ''));
    if (!m) return toText(raw);
    const decimals = m[2] ? m[2].length : 0;
    const grouped = pattern.includes(',');
    return n.toLocaleString('en-US', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
      useGrouping: grouped,
    });
  }),

  // ---- lookup -------------------------------------------------------------
  VLOOKUP: def((needle, table, colIndex, approximate) => {
    const key = Array.isArray(needle) ? needle.flat(Infinity)[0] : needle;
    if (isError(key)) return key;
    const idx = num1(colIndex);
    if (isError(idx)) return idx;
    if (!Array.isArray(table) || !Array.isArray(table[0])) return ERR.VALUE('VLOOKUP needs a range');
    if (idx < 1 || idx > table[0].length) return ERR.REF('VLOOKUP column ' + idx + ' is outside the range');
    const exact = approximate === undefined ? false : !toBoolean(Array.isArray(approximate) ? approximate.flat(Infinity)[0] : approximate);

    if (exact) {
      for (const row of table) {
        const cmp = compareValues(row[0], key);
        if (isError(cmp)) return cmp;
        if (cmp === 0) return row[idx - 1] ?? '';
      }
      return ERR.NA('VLOOKUP found no match');
    }
    // approximate: last row whose first column is <= key, assuming sorted
    let best = null;
    for (const row of table) {
      const cmp = compareValues(row[0], key);
      if (isError(cmp)) return cmp;
      if (cmp <= 0) best = row;
      else break;
    }
    return best ? (best[idx - 1] ?? '') : ERR.NA('VLOOKUP found no match');
  }),
  INDEX: def((table, rowNum, colNum) => {
    if (!Array.isArray(table)) return ERR.VALUE('INDEX needs a range');
    const grid = Array.isArray(table[0]) ? table : [table];
    const r = num1(rowNum);
    if (isError(r)) return r;
    const c = colNum === undefined ? 1 : num1(colNum);
    if (isError(c)) return c;
    const row = grid[r - 1];
    if (!row) return ERR.REF('INDEX row ' + r + ' is outside the range');
    const v = row[c - 1];
    if (v === undefined) return ERR.REF('INDEX column ' + c + ' is outside the range');
    return v;
  }),
  MATCH: def((needle, range, matchType) => {
    const key = Array.isArray(needle) ? needle.flat(Infinity)[0] : needle;
    if (isError(key)) return key;
    const cells = Array.isArray(range) ? range.flat(Infinity) : [range];
    const type = matchType === undefined ? 1 : num1(matchType);
    if (isError(type)) return type;
    if (type === 0) {
      for (let i = 0; i < cells.length; i++) {
        const cmp = compareValues(cells[i], key);
        if (isError(cmp)) return cmp;
        if (cmp === 0) return i + 1;
      }
      return ERR.NA('MATCH found no match');
    }
    let best = null;
    for (let i = 0; i < cells.length; i++) {
      const cmp = compareValues(cells[i], key);
      if (isError(cmp)) return cmp;
      if (type > 0 ? cmp <= 0 : cmp >= 0) best = i + 1;
    }
    return best ?? ERR.NA('MATCH found no match');
  }),

  // ---- dates --------------------------------------------------------------
  TODAY: def((_, ctx) => dateToSerial(ctx.now()), { wantsContext: true, volatile: true }),
  NOW: def((_, ctx) => {
    const d = ctx.now();
    const dayFraction = (d.getUTCHours() * 3600 + d.getUTCMinutes() * 60 + d.getUTCSeconds()) / 86400;
    return dateToSerial(d) + dayFraction;
  }, { wantsContext: true, volatile: true }),
  DATE: def((y, m, d) => {
    const yy = num1(y);
    const mm = num1(m);
    const dd = num1(d);
    const e = firstError([yy, mm, dd]);
    if (e) return e;
    return dateToSerial(new Date(Date.UTC(yy, mm - 1, dd)));
  }),
  YEAR: def((v) => { const n = num1(v); return isError(n) ? n : serialToDate(n).getUTCFullYear(); }),
  MONTH: def((v) => { const n = num1(v); return isError(n) ? n : serialToDate(n).getUTCMonth() + 1; }),
  DAY: def((v) => { const n = num1(v); return isError(n) ? n : serialToDate(n).getUTCDate(); }),

  // ---- maths, the wider set ----------------------------------------------
  PI: def(() => Math.PI),
  EXP: def((v) => { const n = num1(v); return isError(n) ? n : Math.exp(n); }),
  LN: def((v) => {
    const n = num1(v);
    if (isError(n)) return n;
    return n <= 0 ? ERR.NUM('LN of a non-positive number') : Math.log(n);
  }),
  LOG: def((v, base) => {
    const n = num1(v);
    const b = base === undefined ? 10 : num1(base);
    const e = firstError([n, b]);
    if (e) return e;
    if (n <= 0 || b <= 0 || b === 1) return ERR.NUM();
    // The ratio of logs carries float noise a spreadsheet does not show:
    // LOG(1000) must be 3, not 2.9999999999999996.
    return Number((b === 10 ? Math.log10(n) : Math.log(n) / Math.log(b)).toPrecision(15));
  }),
  LOG10: def((v) => {
    const n = num1(v);
    if (isError(n)) return n;
    return n <= 0 ? ERR.NUM() : Math.log10(n);
  }),
  // CEILING/FLOOR round to a multiple of the significance. A number and a
  // significance of opposite signs is #NUM!, as the spreadsheet convention has
  // it; a significance of 0 yields 0.
  CEILING: def((v, sig) => {
    const n = num1(v);
    const s = sig === undefined ? 1 : num1(sig);
    const e = firstError([n, s]);
    if (e) return e;
    if (s === 0) return 0;
    if (n > 0 && s < 0) return ERR.NUM('CEILING with opposite signs');
    return Math.ceil(n / s) * s;
  }),
  FLOOR: def((v, sig) => {
    const n = num1(v);
    const s = sig === undefined ? 1 : num1(sig);
    const e = firstError([n, s]);
    if (e) return e;
    if (s === 0) return 0;
    if (n > 0 && s < 0) return ERR.NUM('FLOOR with opposite signs');
    return Math.floor(n / s) * s;
  }),
  // CEILING.MATH / FLOOR.MATH are the 2013 forms: the significance defaults
  // to 1, its sign never matters, and a negative number rounds toward zero
  // (CEILING.MATH) or away from it (FLOOR.MATH) unless mode is non-zero.
  'CEILING.MATH': def((v, sig, mode) => {
    const n = num1(v);
    const s = sig === undefined ? 1 : Math.abs(num1(sig));
    const m = mode === undefined ? 0 : num1(mode);
    const e = firstError([n, s, m]);
    if (e) return e;
    if (s === 0) return 0;
    if (n < 0 && m !== 0) return -Math.ceil(-n / s) * s;
    return Math.ceil(n / s) * s;
  }),
  'FLOOR.MATH': def((v, sig, mode) => {
    const n = num1(v);
    const s = sig === undefined ? 1 : Math.abs(num1(sig));
    const m = mode === undefined ? 0 : num1(mode);
    const e = firstError([n, s, m]);
    if (e) return e;
    if (s === 0) return 0;
    if (n < 0 && m !== 0) return -Math.floor(-n / s) * s;
    return Math.floor(n / s) * s;
  }),
  MROUND: def((v, mult) => {

    const n = num1(v);
    const m = num1(mult);
    const e = firstError([n, m]);
    if (e) return e;
    if (m === 0) return 0;
    if ((n > 0 && m < 0) || (n < 0 && m > 0)) return ERR.NUM('MROUND with opposite signs');
    return roundHalfAwayFromZero(n / m, 0) * m;
  }),
  EVEN: def((v) => {
    const n = num1(v);
    if (isError(n)) return n;
    const r = Math.ceil(Math.abs(n) / 2) * 2;
    return n < 0 ? -r : r;
  }),
  ODD: def((v) => {
    const n = num1(v);
    if (isError(n)) return n;
    const a = Math.abs(n);
    const r = a <= 1 ? 1 : Math.ceil((a - 1) / 2) * 2 + 1;
    return n < 0 ? -r : r;
  }),
  SUMSQ: def((...args) => {
    const nums = numericArgs(args);
    return isError(nums) ? nums : nums.reduce((a, b) => a + b * b, 0);
  }),
  // SUMPRODUCT multiplies its arrays position by position. Text and blanks
  // count as 0 — the convention that makes it usable as a conditional sum.
  SUMPRODUCT: def((...args) => {
    const e = firstError(args);
    if (e) return e;
    const vectors = args.map(vec);
    const len = vectors[0].length;
    if (vectors.some((v) => v.length !== len)) return ERR.VALUE('SUMPRODUCT arrays must be the same size');
    let total = 0;
    for (let i = 0; i < len; i++) {
      let product = 1;
      for (const v of vectors) product *= typeof v[i] === 'number' ? v[i] : 0;
      total += product;
    }
    return total;
  }),
  RAND: def(() => Math.random(), { volatile: true }),
  RANDBETWEEN: def((lo, hi) => {
    const a = num1(lo);
    const b = num1(hi);
    const e = firstError([a, b]);
    if (e) return e;
    if (b < a) return ERR.NUM('RANDBETWEEN needs bottom <= top');
    const low = Math.ceil(a);
    const high = Math.floor(b);
    return low + Math.floor(Math.random() * (high - low + 1));
  }, { volatile: true }),

  // ---- statistics ---------------------------------------------------------
  MEDIAN: def((...args) => {
    const nums = numericArgs(args);
    if (isError(nums)) return nums;
    if (!nums.length) return ERR.NUM('MEDIAN of nothing');
    const s = [...nums].sort((a, b) => a - b);
    const mid = s.length >> 1;
    return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  }),
  // The most frequent number, first encountered winning a tie; all-distinct is #N/A.
  MODE: def((...args) => {
    const nums = numericArgs(args);
    if (isError(nums)) return nums;
    const counts = new Map();
    for (const n of nums) counts.set(n, (counts.get(n) ?? 0) + 1);
    let best = null;
    let bestCount = 1;
    for (const [n, count] of counts) {
      if (count > bestCount) { best = n; bestCount = count; }
    }
    return best === null ? ERR.NA('no value repeats') : best;
  }),
  'MODE.SNGL': def((...args) => FUNCTIONS.MODE.fn(...args)),
  'STDEV.S': def((...args) => {
    const v = variance(args, { population: false });
    return isError(v) ? v : Math.sqrt(v);
  }),
  'STDEV.P': def((...args) => {
    const v = variance(args, { population: true });
    return isError(v) ? v : Math.sqrt(v);
  }),
  STDEV: def((...args) => FUNCTIONS['STDEV.S'].fn(...args)),
  STDEVP: def((...args) => FUNCTIONS['STDEV.P'].fn(...args)),
  'VAR.S': def((...args) => variance(args, { population: false })),
  'VAR.P': def((...args) => variance(args, { population: true })),
  VAR: def((...args) => variance(args, { population: false })),
  VARP: def((...args) => variance(args, { population: true })),
  LARGE: def((range, k) => {
    const nums = aggregateNumbers([range]);
    if (isError(nums)) return nums;
    const n = num1(k);
    if (isError(n)) return n;
    if (n < 1 || n > nums.length) return ERR.NUM('LARGE k is outside the data');
    return [...nums].sort((a, b) => b - a)[n - 1];
  }),
  SMALL: def((range, k) => {
    const nums = aggregateNumbers([range]);
    if (isError(nums)) return nums;
    const n = num1(k);
    if (isError(n)) return n;
    if (n < 1 || n > nums.length) return ERR.NUM('SMALL k is outside the data');
    return [...nums].sort((a, b) => a - b)[n - 1];
  }),
  // Descending unless order is non-zero, ties share the top rank — Excel's RANK.
  RANK: def((value, range, order) => {
    const n = num1(value);
    if (isError(n)) return n;
    const nums = aggregateNumbers([range]);
    if (isError(nums)) return nums;
    const ord = order === undefined ? 0 : num1(order);
    if (isError(ord)) return ord;
    const asc = ord !== 0;
    if (!nums.includes(n)) return ERR.NA('RANK value is not in the data');
    return 1 + nums.filter((x) => (asc ? x < n : x > n)).length;
  }),
  'RANK.EQ': def((value, range, order) => FUNCTIONS.RANK.fn(value, range, order)),

  // ---- conditional aggregation, many criteria -----------------------------
  SUMIFS: def((sumRange, ...pairs) => {
    const targets = vec(sumRange);
    const hits = ifsIndices(targets.length, pairs);
    if (isError(hits)) return hits;
    let total = 0;
    for (const i of hits) {
      const v = targets[i];
      if (isError(v)) return v;
      if (typeof v === 'number') total += v;
    }
    return total;
  }),
  COUNTIFS: def((...pairs) => {
    if (!pairs.length) return ERR.VALUE('COUNTIFS needs at least one range and criteria');
    const hits = ifsIndices(vec(pairs[0]).length, pairs);
    return isError(hits) ? hits : hits.length;
  }),
  AVERAGEIFS: def((avgRange, ...pairs) => {
    const targets = vec(avgRange);
    const hits = ifsIndices(targets.length, pairs);
    if (isError(hits)) return hits;
    const picked = hits.map((i) => targets[i]).filter((v) => typeof v === 'number');
    const e = firstError(hits.map((i) => targets[i]));
    if (e) return e;
    return picked.length ? picked.reduce((a, b) => a + b, 0) / picked.length : ERR.DIV0();
  }),
  MAXIFS: def((range, ...pairs) => {
    const targets = vec(range);
    const hits = ifsIndices(targets.length, pairs);
    if (isError(hits)) return hits;
    const picked = hits.map((i) => targets[i]).filter((v) => typeof v === 'number');
    return picked.length ? Math.max(...picked) : 0;
  }),
  MINIFS: def((range, ...pairs) => {
    const targets = vec(range);
    const hits = ifsIndices(targets.length, pairs);
    if (isError(hits)) return hits;
    const picked = hits.map((i) => targets[i]).filter((v) => typeof v === 'number');
    return picked.length ? Math.min(...picked) : 0;
  }),

  // ---- logic, extended ----------------------------------------------------
  IFS: def((...thunks) => {
    if (thunks.length % 2 !== 0) return ERR.VALUE('IFS takes condition/value pairs');
    for (let i = 0; i < thunks.length; i += 2) {
      const c = toBoolean(firstOf(thunks[i]()));
      if (isError(c)) return c;
      if (c) return thunks[i + 1]();
    }
    return ERR.NA('no IFS condition was true');
  }, { lazy: true }),
  SWITCH: def((...thunks) => {
    if (thunks.length < 3) return ERR.VALUE('SWITCH needs a value and at least one case/result pair');
    const value = thunks[0]();
    if (isError(value)) return value;
    // After the value come case/result pairs; a leftover odd argument is the default.
    const hasDefault = (thunks.length - 1) % 2 === 1;
    const pairsEnd = hasDefault ? thunks.length - 1 : thunks.length;
    for (let i = 1; i + 1 < pairsEnd; i += 2) {
      const candidate = thunks[i]();
      if (isError(candidate)) return candidate;
      const cmp = compareValues(value, candidate);
      if (!isError(cmp) && cmp === 0) return thunks[i + 1]();
    }
    return hasDefault ? thunks[thunks.length - 1]() : ERR.NA('no SWITCH case matched');
  }, { lazy: true }),
  XOR: def((...args) => {
    let odd = false;
    for (const v of flatten(args)) {
      if (isBlank(v)) continue;
      const b = toBoolean(v);
      if (isError(b)) return b;
      if (b) odd = !odd;
    }
    return odd;
  }),

  // ---- text, extended ------------------------------------------------------
  // FIND is case-sensitive and literal; SEARCH is case-insensitive and takes
  // wildcards. Both are 1-based and #VALUE! when nothing is found.
  FIND: def((what, where, start) => {
    const f = text1(what);
    const s = text1(where);
    const at = start === undefined ? 1 : num1(start);
    const e = firstError([f, s, at]);
    if (e) return e;
    if (at < 1 || at > s.length + 1) return ERR.VALUE('FIND start is outside the text');
    const idx = s.indexOf(f, at - 1);
    return idx < 0 ? ERR.VALUE('FIND found nothing') : idx + 1;
  }),
  SEARCH: def((what, where, start) => {
    const f = text1(what);
    const s = text1(where);
    const at = start === undefined ? 1 : num1(start);
    const e = firstError([f, s, at]);
    if (e) return e;
    if (at < 1 || at > s.length + 1) return ERR.VALUE('SEARCH start is outside the text');
    if (/[*?]/.test(f)) {
      // Wildcards match a SUBSTRING starting at each candidate position.
      const re = new RegExp(wildcardRegex(f).source.slice(1, -1), 'i');
      const idx = s.slice(at - 1).search(re);
      return idx < 0 ? ERR.VALUE('SEARCH found nothing') : at + idx;
    }
    const idx = s.toLowerCase().indexOf(f.toLowerCase(), at - 1);
    return idx < 0 ? ERR.VALUE('SEARCH found nothing') : idx + 1;
  }),
  REPLACE: def((old, start, count, replacement) => {
    const s = text1(old);
    const at = num1(start);
    const n = num1(count);
    const r = text1(replacement);
    const e = firstError([s, at, n, r]);
    if (e) return e;
    if (at < 1 || n < 0) return ERR.VALUE();
    return s.slice(0, at - 1) + r + s.slice(at - 1 + n);
  }),
  REPT: def((v, times) => {
    const s = text1(v);
    const n = times === undefined ? 0 : num1(times);
    const e = firstError([s, n]);
    if (e) return e;
    if (n < 0) return ERR.VALUE('REPT count must not be negative');
    const count = Math.floor(n);
    if (s.length * count > 32767) return ERR.VALUE('REPT result is too long');
    return s.repeat(count);
  }),
  PROPER: def((v) => {
    const s = text1(v);
    if (isError(s)) return s;
    return s.toLowerCase().replace(/(^|[^A-Za-z])([a-z])/g, (_, before, ch) => before + ch.toUpperCase());
  }),
  EXACT: def((a, b) => {
    const x = text1(a);
    const y = text1(b);
    return firstError([x, y]) || x === y;
  }),
  CHAR: def((v) => {
    const n = num1(v);
    if (isError(n)) return n;
    const code = Math.trunc(n);
    return code >= 1 && code <= 255 ? String.fromCharCode(code) : ERR.VALUE('CHAR takes 1-255');
  }),
  CODE: def((v) => {
    const s = text1(v);
    if (isError(s)) return s;
    return s.length ? s.charCodeAt(0) : ERR.VALUE('CODE of empty text');
  }),
  TEXTJOIN: def((delim, ignoreEmpty, ...args) => {
    const d = text1(delim);
    if (isError(d)) return d;
    const skip = toBoolean(Array.isArray(ignoreEmpty) ? ignoreEmpty.flat(Infinity)[0] : ignoreEmpty);
    if (isError(skip)) return skip;
    const e = firstError(args);
    if (e) return e;
    const parts = flatten(args).filter((v) => !(skip && isBlank(v))).map(toText);
    return parts.join(d);
  }),

  // ---- lookup, extended ----------------------------------------------------
  HLOOKUP: def((needle, table, rowIndex, approximate) => {
    const key = Array.isArray(needle) ? needle.flat(Infinity)[0] : needle;
    if (isError(key)) return key;
    const idx = num1(rowIndex);
    if (isError(idx)) return idx;
    if (!Array.isArray(table) || !Array.isArray(table[0])) return ERR.VALUE('HLOOKUP needs a range');
    if (idx < 1 || idx > table.length) return ERR.REF('HLOOKUP row ' + idx + ' is outside the range');
    const exact = approximate === undefined ? false : !toBoolean(Array.isArray(approximate) ? approximate.flat(Infinity)[0] : approximate);
    const headers = table[0];

    if (exact) {
      for (let c = 0; c < headers.length; c++) {
        const cmp = compareValues(headers[c], key);
        if (isError(cmp)) return cmp;
        if (cmp === 0) return table[idx - 1][c] ?? '';
      }
      return ERR.NA('HLOOKUP found no match');
    }
    let best = -1;
    for (let c = 0; c < headers.length; c++) {
      const cmp = compareValues(headers[c], key);
      if (isError(cmp)) return cmp;
      if (cmp <= 0) best = c;
      else break;
    }
    return best >= 0 ? (table[idx - 1][best] ?? '') : ERR.NA('HLOOKUP found no match');
  }),
  /**
   * XLOOKUP(lookup, lookupArray, returnArray, [ifNotFound], [matchMode], [searchMode]).
   * Match modes: 0 exact (default), -1 exact or next smaller, 1 exact or next
   * larger, 2 wildcard. Search mode 1 is first-to-last, -1 last-to-first; the
   * binary modes behave as linear, which is a performance shape, not a result.
   */
  XLOOKUP: def((needle, lookupArray, returnArray, ifNotFound, matchMode, searchMode) => {
    const key = Array.isArray(needle) ? needle.flat(Infinity)[0] : needle;
    if (isError(key)) return key;
    const looks = vec(lookupArray);
    const returns = vec(returnArray);
    if (looks.length !== returns.length) return ERR.VALUE('XLOOKUP arrays must be the same size');
    const mode = matchMode === undefined ? 0 : num1(matchMode);
    if (isError(mode)) return mode;
    const dir = searchMode === undefined ? 1 : num1(searchMode);
    if (isError(dir)) return dir;

    const order = [...looks.keys()];
    if (dir < 0) order.reverse();

    if (mode === 2) {
      const re = wildcardRegex(toText(key));
      for (const i of order) {
        if (re.test(toText(looks[i]))) return returns[i];
      }
    } else {
      let bestIdx = -1;
      let bestValue = null;
      for (const i of order) {
        const cmp = compareValues(looks[i], key);
        if (isError(cmp)) continue;
        if (cmp === 0) return returns[i];
        if (mode === -1 && cmp < 0) {
          if (bestIdx < 0 || compareValues(looks[i], bestValue) > 0) { bestIdx = i; bestValue = looks[i]; }
        } else if (mode === 1 && cmp > 0) {
          if (bestIdx < 0 || compareValues(looks[i], bestValue) < 0) { bestIdx = i; bestValue = looks[i]; }
        }
      }
      if (bestIdx >= 0) return returns[bestIdx];
    }
    if (ifNotFound !== undefined) return Array.isArray(ifNotFound) ? ifNotFound.flat(Infinity)[0] : ifNotFound;
    return ERR.NA('XLOOKUP found no match');
  }),
  CHOOSE: def((...thunks) => {
    if (thunks.length < 2) return ERR.VALUE('CHOOSE needs an index and at least one value');
    const raw = thunks[0]();
    const idx = toNumber(Array.isArray(raw) ? raw.flat(Infinity)[0] : raw);
    if (isError(idx)) return idx;
    const n = Math.trunc(idx);
    if (n < 1 || n > thunks.length - 1) return ERR.VALUE('CHOOSE index ' + n + ' is outside the values');
    return thunks[n]();
  }, { lazy: true }),

  // ---- dates and times, extended ------------------------------------------
  EDATE: def((start, months) => {
    const s = num1(start);
    const m = num1(months);
    const e = firstError([s, m]);
    if (e) return e;
    const p = partsOf(s);
    return serialOfClamped(p.y, p.m + Math.trunc(m), p.day);
  }),
  EOMONTH: def((start, months) => {
    const s = num1(start);
    const m = num1(months);
    const e = firstError([s, m]);
    if (e) return e;
    const p = partsOf(s);
    return serialOfClamped(p.y, p.m + Math.trunc(m), 31);
  }),
  DAYS: def((end, start) => {
    const a = num1(end);
    const b = num1(start);
    return firstError([a, b]) || Math.trunc(a) - Math.trunc(b);
  }),
  WEEKDAY: def((serial, type) => {
    const s = num1(serial);
    if (isError(s)) return s;
    const t = type === undefined ? 1 : num1(type);
    if (isError(t)) return t;
    const sunday0 = serialToDate(s).getUTCDay(); // 0 = Sunday
    if (t === 1) return sunday0 + 1;
    if (t === 2) return sunday0 === 0 ? 7 : sunday0;
    if (t === 3) return sunday0 === 0 ? 6 : sunday0 - 1;
    return ERR.NUM('WEEKDAY type must be 1, 2 or 3');
  }),
  HOUR: def((v) => {
    const n = num1(v);
    if (isError(n)) return n;
    return Math.floor(Math.round((n - Math.floor(n)) * 86400) / 3600) % 24;
  }),
  MINUTE: def((v) => {
    const n = num1(v);
    if (isError(n)) return n;
    return Math.floor((Math.round((n - Math.floor(n)) * 86400) % 3600) / 60);
  }),
  SECOND: def((v) => {
    const n = num1(v);
    if (isError(n)) return n;
    return Math.round((n - Math.floor(n)) * 86400) % 60;
  }),
  TIME: def((h, m, s) => {
    const hh = num1(h);
    const mm = num1(m);
    const ss = num1(s);
    const e = firstError([hh, mm, ss]);
    if (e) return e;
    const seconds = hh * 3600 + mm * 60 + ss;
    if (seconds < 0) return ERR.NUM('TIME before midnight');
    return (seconds % 86400) / 86400;
  }),

  // ---- reference geometry --------------------------------------------------
  // ROW(ref)/COLUMN(ref) are answered by the evaluator from the AST, where a
  // reference still has a position; these entries carry the argument-less
  // form and keep the catalog honest.
  ROW: def((v, ctx) => ((ctx.row ?? 0) + 1), { wantsContext: true }),
  COLUMN: def((v, ctx) => ((ctx.col ?? 0) + 1), { wantsContext: true }),
  ROWS: def((v) => (Array.isArray(v) ? v.length : 1)),
  COLUMNS: def((v) => (Array.isArray(v) ? (Array.isArray(v[0]) ? v[0].length : v.length) : 1)),
  // OFFSET and INDIRECT are answered by the evaluator too — they build a new
  // reference at run time, which needs the AST and the resolver. These entries
  // carry the volatile flag the engine reads and keep the catalog honest; the
  // bodies are only reachable if the evaluator's interception is broken.
  OFFSET: def(() => ERR.VALUE('OFFSET needs a reference'), { volatile: true }),
  INDIRECT: def(() => ERR.REF('INDIRECT needs a reference'), { volatile: true }),
  ADDRESS: def((row, col, abs, a1, sheetName) => {
    const r = num1(row);
    const c = num1(col);
    const mode = abs === undefined || abs === null || abs === '' ? 1 : num1(abs);
    const e = firstError([r, c, mode]);
    if (e) return e;
    if (r < 1 || c < 1 || mode < 1 || mode > 4) return ERR.VALUE();
    if (a1 !== undefined && a1 !== null && a1 !== '' && toBoolean(a1) === false) {
      return ERR.VALUE('R1C1-style addresses are not supported');
    }
    // abs: 1 = $A$1, 2 = A$1, 3 = $A1, 4 = A1 — Excel's numbering.
    const colTxt = (mode === 1 || mode === 3 ? '$' : '') + indexToCol(c - 1);
    const rowTxt = (mode === 1 || mode === 2 ? '$' : '') + r;
    if (sheetName === undefined || sheetName === null || sheetName === '') return colTxt + rowTxt;
    const s = toText(sheetName);
    if (isError(s)) return s;
    const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(s) ? s : "'" + s.replace(/'/g, "''") + "'";
    return quoted + '!' + colTxt + rowTxt;
  }),

  // ---- dynamic arrays ------------------------------------------------------
  // These RETURN arrays, and the engine SPILLS what comes back into the
  // cells below and beside the formula. In a legacy scalar context the
  // evaluator's top-level collapse still applies, so nothing old breaks.
  SEQUENCE: def((rows, cols, start, step) => {
    const r = num1(rows);
    const c = cols === undefined || cols === null || cols === '' ? 1 : num1(cols);
    const s = start === undefined || start === null || start === '' ? 1 : num1(start);
    const d = step === undefined || step === null || step === '' ? 1 : num1(step);
    const e = firstError([r, c, s, d]);
    if (e) return e;
    if (r < 1 || c < 1 || r * c > 1048576) return ERR.VALUE('SEQUENCE needs positive dimensions');
    const out = [];
    let v = s;
    for (let i = 0; i < r; i++) {
      const line = [];
      for (let j = 0; j < c; j++) {
        line.push(v);
        v += d;
      }
      out.push(line);
    }
    return out;
  }),
  TRANSPOSE: def((v) => {
    const g = gridOf(v);
    const out = [];
    for (let j = 0; j < (g[0]?.length ?? 0); j++) out.push(g.map((line) => line[j] ?? ''));
    return out.length ? out : [['']];
  }),
  UNIQUE: def((v, byCol, exactlyOnce) => {
    // by_col is the row machinery run sideways: transpose, dedupe, transpose.
    if (toBoolean(byCol ?? false) === true) {
      const t = FUNCTIONS.UNIQUE.fn(transposeGrid(gridOf(v)), false, exactlyOnce);
      return isError(t) ? t : transposeGrid(t);
    }
    const rows = gridOf(v);
    // Text compares case-insensitively, as every comparison here does.
    const keyOf = (line) => line
      .map((x) => (typeof x === 'string' ? 't:' + x.toLowerCase() : typeof x + ':' + String(x)))
      .join('\u001f');
    const counts = new Map();
    for (const line of rows) counts.set(keyOf(line), (counts.get(keyOf(line)) ?? 0) + 1);
    const once = toBoolean(exactlyOnce ?? false) === true;
    const seen = new Set();
    const out = [];
    for (const line of rows) {
      const k = keyOf(line);
      if (once) {
        if (counts.get(k) === 1) out.push(line);
        continue;
      }
      if (!seen.has(k)) {
        seen.add(k);
        out.push(line);
      }
    }
    return out.length ? out : ERR.CALC('UNIQUE found nothing');
  }),
  SORT: def((v, index, order, byCol) => {
    // by_col sorts COLUMNS by a row's values — the row machinery, sideways.
    if (toBoolean(byCol ?? false) === true) {
      const t = FUNCTIONS.SORT.fn(transposeGrid(gridOf(v)), index, order, false);
      return isError(t) ? t : transposeGrid(t);
    }
    const rows = gridOf(v);
    const col = index === undefined || index === null || index === '' ? 1 : num1(index);
    const dir = order === undefined || order === null || order === '' ? 1 : num1(order);
    const e = firstError([col, dir]);
    if (e) return e;
    if (col < 1 || (rows[0] && col > rows[0].length)) return ERR.VALUE('SORT index is outside the array');
    if (dir !== 1 && dir !== -1) return ERR.VALUE('SORT order is 1 or -1');
    return [...rows].sort((a, b) => {
      const c = compareValues(a[col - 1], b[col - 1]);
      return (isError(c) ? 0 : c) * dir;
    });
  }),
  FILTER: def((v, include, ifEmpty) => {
    const rows = gridOf(v);
    const inc = gridOf(include).flat(Infinity);
    if (inc.length !== rows.length) {
      return ERR.VALUE('FILTER include must have one entry per row');
    }
    const out = [];
    for (let i = 0; i < rows.length; i++) {
      const b = toBoolean(inc[i]);
      if (isError(b)) return b;
      if (b) out.push(rows[i]);
    }
    if (out.length) return out;
    return ifEmpty === undefined ? ERR.CALC('FILTER kept nothing') : ifEmpty;
  }),

  // ---- information ---------------------------------------------------------
  ISEVEN: def((v) => {
    const n = num1(v);
    return isError(n) ? n : Math.trunc(n) % 2 === 0;
  }),
  ISODD: def((v) => {
    const n = num1(v);
    return isError(n) ? n : Math.abs(Math.trunc(n)) % 2 === 1;
  }),
  N: def((v) => {
    const raw = Array.isArray(v) ? v.flat(Infinity)[0] : v;
    if (isError(raw)) return raw;
    if (typeof raw === 'number') return raw;
    if (typeof raw === 'boolean') return raw ? 1 : 0;
    return 0;
  }),

  // ---- finance --------------------------------------------------------------
  // The standard annuity arithmetic, with `type` 1 meaning payments at the
  // start of the period. Signs follow the spreadsheet convention: money out is
  // negative, so a loan payment comes back negative.
  PMT: def((rate, nper, pv, fv, type) => {
    const r = num1(rate);
    const n = num1(nper);
    const p = num1(pv);
    const f = fv === undefined ? 0 : num1(fv);
    const t = type === undefined ? 0 : num1(type);
    const e = firstError([r, n, p, f, t]);
    if (e) return e;
    if (n === 0) return ERR.DIV0();
    if (r === 0) return -(p + f) / n;
    const growth = (1 + r) ** n;
    return -(p * growth + f) * r / ((growth - 1) * (1 + r * (t ? 1 : 0)));
  }),
  FV: def((rate, nper, pmt, pv, type) => {
    const r = num1(rate);
    const n = num1(nper);
    const m = num1(pmt);
    const p = pv === undefined ? 0 : num1(pv);
    const t = type === undefined ? 0 : num1(type);
    const e = firstError([r, n, m, p, t]);
    if (e) return e;
    if (r === 0) return -(p + m * n);
    const growth = (1 + r) ** n;
    return -(p * growth + m * (1 + r * (t ? 1 : 0)) * (growth - 1) / r);
  }),
  PV: def((rate, nper, pmt, fv, type) => {
    const r = num1(rate);
    const n = num1(nper);
    const m = num1(pmt);
    const f = fv === undefined ? 0 : num1(fv);
    const t = type === undefined ? 0 : num1(type);
    const e = firstError([r, n, m, f, t]);
    if (e) return e;
    if (r === 0) return -(f + m * n);
    const growth = (1 + r) ** n;
    return -(f + m * (1 + r * (t ? 1 : 0)) * (growth - 1) / r) / growth;
  }),
  NPER: def((rate, pmt, pv, fv, type) => {
    const r = num1(rate);
    const m = num1(pmt);
    const p = num1(pv);
    const f = fv === undefined ? 0 : num1(fv);
    const t = type === undefined ? 0 : num1(type);
    const e = firstError([r, m, p, f, t]);
    if (e) return e;
    if (r === 0) return m === 0 ? ERR.DIV0() : -(p + f) / m;
    const adj = m * (1 + r * (t ? 1 : 0)) / r;
    const inside = (adj - f) / (p + adj);
    return inside <= 0 ? ERR.NUM('NPER has no solution for these values') : Math.log(inside) / Math.log(1 + r);
  }),
  NPV: def((rate, ...args) => {
    const r = num1(rate);
    if (isError(r)) return r;
    if (r === -1) return ERR.DIV0();
    const nums = aggregateNumbers(args);
    if (isError(nums)) return nums;
    return nums.reduce((total, v, i) => total + v / (1 + r) ** (i + 1), 0);
  }),
  /**
   * Internal rate of return: the rate at which the cash flows' net present
   * value is zero, found by Newton's method from the guess (10% unless
   * given), as Excel does — and #NUM! when twenty steps do not converge or
   * the flows never change sign, which is when no rate exists.
   */
  IRR: def((values, guess) => {
    const flows = aggregateNumbers([values]);
    if (isError(flows)) return flows;
    const g = guess === undefined ? 0.1 : num1(guess);
    if (isError(g)) return g;
    if (!flows.some((v) => v > 0) || !flows.some((v) => v < 0)) return ERR.NUM('IRR needs both an outflow and an inflow');
    let r = g;
    for (let i = 0; i < 100; i++) {
      let f = 0;
      let df = 0;
      for (let k = 0; k < flows.length; k++) {
        const d = (1 + r) ** k;
        f += flows[k] / d;
        df -= (k * flows[k]) / (d * (1 + r));
      }
      if (Math.abs(f) < 1e-10) return r;
      if (df === 0 || !Number.isFinite(df)) break;
      const next = r - f / df;
      if (!Number.isFinite(next) || next <= -1) break;
      if (Math.abs(next - r) < 1e-12) return next;
      r = next;
    }
    return ERR.NUM('IRR did not converge');
  }),
  /**
   * The interest rate per period of an annuity — PMT solved for the rate,
   * by Newton's method from the guess (10% unless given), like Excel.
   */
  RATE: def((nper, pmt, pv, fv, type, guess) => {
    const n = num1(nper);
    const m = num1(pmt);
    const p = num1(pv);
    const f = fv === undefined ? 0 : num1(fv);
    const t = type === undefined ? 0 : num1(type);
    const g = guess === undefined ? 0.1 : num1(guess);
    const e = firstError([n, m, p, f, t, g]);
    if (e) return e;
    if (n <= 0) return ERR.NUM('RATE needs a positive number of periods');
    const balance = (r) => {
      if (Math.abs(r) < 1e-12) return p + m * n + f;
      const growth = (1 + r) ** n;
      return p * growth + m * (1 + r * (t ? 1 : 0)) * (growth - 1) / r + f;
    };
    let r = g;
    for (let i = 0; i < 100; i++) {
      const y = balance(r);
      if (Math.abs(y) < 1e-9) return r;
      const h = 1e-6;
      const dy = (balance(r + h) - balance(r - h)) / (2 * h);
      if (dy === 0 || !Number.isFinite(dy)) break;
      const next = r - y / dy;
      if (!Number.isFinite(next) || next <= -1) break;
      if (Math.abs(next - r) < 1e-12) return next;
      r = next;
    }
    return ERR.NUM('RATE did not converge');
  }),
  /**
   * Excel's DATEDIF, units as Excel spells them: "Y" complete years, "M"
   * complete months, "D" days, "MD" days ignoring months and years, "YM"
   * months ignoring years, "YD" days ignoring years. #NUM! when the start is
   * after the end, as in Excel.
   */
  DATEDIF: def((start, end, unit) => {
    const s = num1(start);
    const e = num1(end);
    const err = firstError([s, e]);
    if (err) return err;
    if (s > e) return ERR.NUM('DATEDIF start is after its end');
    const u = String(unit ?? '').toUpperCase().trim();
    const a = partsOf(Math.trunc(s));
    const b = partsOf(Math.trunc(e));
    const months = (b.y - a.y) * 12 + (b.m - a.m) - (b.day < a.day ? 1 : 0);
    if (u === 'D') return Math.trunc(e) - Math.trunc(s);
    if (u === 'M') return months;
    if (u === 'Y') return Math.trunc(months / 12);
    if (u === 'YM') return months % 12;
    if (u === 'MD') {
      if (b.day >= a.day) return b.day - a.day;
      // Days left in the month before the end's month, plus the end's day.
      const prevLast = new Date(Date.UTC(b.y, b.m, 0)).getUTCDate();
      return prevLast - a.day + b.day;
    }
    if (u === 'YD') {
      const sameYearStart = serialOfClamped(b.y - (new Date(Date.UTC(b.y, a.m, a.day)) > new Date(Date.UTC(b.y, b.m, b.day)) ? 1 : 0), a.m, a.day);
      return Math.trunc(e) - sameYearStart;
    }
    return ERR.NUM('DATEDIF unit must be Y, M, D, MD, YM or YD');
  }),
};

export const FUNCTION_NAMES = Object.keys(FUNCTIONS).sort();
export const isVolatile = (name) => Boolean(FUNCTIONS[name]?.volatile);
export { FormulaError, makeMatcher };
