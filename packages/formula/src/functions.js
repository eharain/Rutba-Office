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
  ERR, ERROR_TYPES, FormulaError, isError, isBlank, isEmpty, firstError,
  toNumber, toText, toBoolean, compareValues, roundHalfAwayFromZero,
  dateToSerial, serialToDate, serialToParts, partsToSerial, numericText, textAsNumber, raised,
} from './values.js';
import { moreFunctions } from './functions-more.js';
import { indexToCol } from './parser.js';
import { formatValue } from './numfmt.js';
import {
  etsQuestion, forecastAt, confidenceAt, prepareSeries, seasonFor, linearForecast, ForecastError,
} from './forecast.js';

/**
 * A forecasting function's body, with errors in its arguments passed on and
 * its own refusals (a timeline without a steady step, a date before the end
 * of it) answered as the Excel error each is.
 */
const forecastFn = (body) => (...args) => {
  const e = firstError(args);
  if (e) return e;
  try {
    return body(...args);
  } catch (err) {
    if (err instanceof ForecastError) return new FormulaError(err.type, err.message);
    throw err;
  }
};

/** An optional number argument: its default when left out or blank. */
const optionalNumber = (v, fallback) => {
  if (v === undefined) return fallback;
  const one = Array.isArray(v) ? v.flat(Infinity)[0] : v;
  if (isBlank(one)) return fallback;
  const n = toNumber(one);
  if (isError(n)) throw new ForecastError('#VALUE!', 'an argument is not a number');
  return n;
};

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

/** A row padded out to `width` with `pad` — #N/A unless said otherwise, as Excel pads a stack. */
const padTo = (line, width, pad = ERR.NA('padded to the widest')) => {
  const out = [...line];
  while (out.length < width) out.push(pad);
  return out;
};

/** TAKE and DROP: rows (and columns) from the start, or from the end when negative. */
function cutGrid(array, rows, cols, how) {
  const g = gridOf(array);
  const counts = [rows, cols].map((v) => (v === undefined || isBlank(Array.isArray(v) ? v.flat(Infinity)[0] : v) ? null : num1(v)));
  const e = firstError(counts.filter((n) => n !== null));
  if (e) return e;
  const cut = (list, n) => {
    if (n === null) return list;
    const k = Math.trunc(n);
    if (how === 'take') return k >= 0 ? list.slice(0, k) : list.slice(Math.max(0, list.length + k));
    return k >= 0 ? list.slice(k) : list.slice(0, Math.max(0, list.length + k));
  };
  const out = cut(g, counts[0]).map((line) => cut(line, counts[1]));
  if (!out.length || !out[0].length) return ERR.CALC((how === 'take' ? 'TAKE' : 'DROP') + ' left nothing');
  return out;
}

/** CHOOSECOLS and CHOOSEROWS: the 1-based picks (negative from the end) as indices, or #VALUE!. */
function pickIndices(picks, size, name) {
  if (!picks.length) return ERR.VALUE(name + ' needs at least one to choose');
  const out = [];
  for (const raw of picks.flatMap(vec)) {
    const n = toNumber(raw);
    if (isError(n)) return n;
    const k = Math.trunc(n);
    if (k === 0 || Math.abs(k) > size) return ERR.VALUE(name + ' asked for one the array does not have');
    out.push(k > 0 ? k - 1 : size + k);
  }
  return out;
}

/** TOCOL and TOROW: the values in reading order (or down each column), blanks and errors left out as asked. */
function lineUp(array, ignore, byColumn) {
  const g = gridOf(array);
  const mode = ignore === undefined || isBlank(Array.isArray(ignore) ? ignore.flat(Infinity)[0] : ignore) ? 0 : num1(ignore);
  if (isError(mode)) return mode;
  if (![0, 1, 2, 3].includes(mode)) return ERR.VALUE('ignore is 0, 1, 2 or 3');
  const down = byColumn !== undefined && toBoolean(Array.isArray(byColumn) ? byColumn.flat(Infinity)[0] : byColumn) === true;
  const order = down ? transposeGrid(g) : g;
  return order.flat().filter((v) => !((mode & 1) && isBlank(v)) && !((mode & 2) && isError(v)));
}

/** WRAPROWS and WRAPCOLS: a line of values folded into rows (or columns) of `count`. */
function wrapLine(vector, count, padWith, into) {
  const g = gridOf(vector);
  if (g.length > 1 && (g[0]?.length ?? 0) > 1) return ERR.VALUE('wrapping takes one row or one column');
  const n = num1(count);
  if (isError(n)) return n;
  const k = Math.trunc(n);
  if (k < 1) return ERR.NUM('the wrap count must be at least 1');
  const flat = g.flat();
  const pad = padWith === undefined ? ERR.NA('wrapping padded the last line') : (Array.isArray(padWith) ? padWith.flat(Infinity)[0] : padWith);
  const lines = [];
  for (let i = 0; i < flat.length; i += k) lines.push(padTo(flat.slice(i, i + k), k, pad));
  return into === 'rows' ? lines : transposeGrid(lines);
}

/** VALUETOTEXT and ARRAYTOTEXT: a value as text — strict quoting text, as Excel's format 1 does. */
function valueText(v, strict) {
  if (isError(v)) return v.type ?? String(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return String(v);
  const s = String(v ?? '');
  return strict ? '"' + s.replace(/"/g, '""') + '"' : s;
}

/**
 * TEXTBEFORE and TEXTAFTER: the words before (or after) the nth delimiter —
 * counted from the end when n is negative — case aside when asked, the
 * text's end standing for a delimiter when asked, `ifNotFound` (or #N/A)
 * when there are not that many.
 */
function textAround(side, value, delimiter, instance, matchMode, matchEnd, ifNotFound) {
  const name = side === 'before' ? 'TEXTBEFORE' : 'TEXTAFTER';
  const text = toText(Array.isArray(value) ? value.flat(Infinity)[0] : value);
  if (isError(text)) return text;
  const delims = vec(delimiter).map((d) => toText(d));
  const bad = delims.find(isError);
  if (bad) return bad;
  const n = instance === undefined || isBlank(Array.isArray(instance) ? instance.flat(Infinity)[0] : instance) ? 1 : Math.trunc(num1(instance));
  if (isError(n)) return n;
  if (n === 0 || Math.abs(n) > text.length + 1) return ERR.VALUE(name + ' instance is outside the text');
  const fold = matchMode !== undefined && num1(matchMode) === 1;
  const toEnd = matchEnd !== undefined && num1(matchEnd) === 1;
  const hay = fold ? text.toLowerCase() : text;
  // Every place a delimiter stands, in order: [start, end).
  const hits = [];
  for (let i = 0; i <= hay.length; i++) {
    for (const d of delims) {
      const needle = fold ? d.toLowerCase() : d;
      if (needle === '' ? true : hay.startsWith(needle, i)) { hits.push([i, i + needle.length]); break; }
    }
  }
  if (delims.every((d) => d === '')) {
    // An empty delimiter stands at the start (or, counting back, the end).
    return side === 'before' ? (n > 0 ? '' : text) : (n > 0 ? text : '');
  }
  const found = hits.filter(([a, b], k) => k === 0 || a >= hits[k - 1][1]);
  if (toEnd) { if (n > 0) found.push([text.length, text.length]); else found.unshift([0, 0]); }
  const hit = n > 0 ? found[n - 1] : found[found.length + n];
  if (!hit) return ifNotFound === undefined ? ERR.NA(name + ' found no such delimiter') : (Array.isArray(ifNotFound) ? ifNotFound.flat(Infinity)[0] : ifNotFound);
  return side === 'before' ? text.slice(0, hit[0]) : text.slice(hit[1]);
}

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
      if (isEmpty(arg)) continue;
      const n = toNumber(arg);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
}

/** A row or column number as a spreadsheet takes it: the part after the point is dropped. */
const truncated = (n) => (typeof n === 'number' ? Math.trunc(n) : n);

/** How many cells past the used part of whole columns every COUNTIFS criterion accepts: those are empty in all of them. */
function emptyRowsCounted(pairs) {
  let cells = null;
  for (let i = 0; i + 1 < pairs.length; i += 2) {
    const full = pairs[i]?.full;
    if (!full || !makeMatcher(pairs[i + 1])(null)) return 0;
    const total = full.rows * full.cols;
    if (cells !== null && cells !== total) return 0;
    cells = total;
  }
  return cells === null ? 0 : cells - vec(pairs[0]).length;
}

/** A value's date and time as the day-fractions the clock on this computer shows. */
const localSerial = (d) => {
  const day = dateToSerial(new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())));
  return { day, part: (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400 };
};

// Not Math.max(...list): a long column is more arguments than a call can take.
const highest = (list) => list.reduce((a, b) => (b > a ? b : a));
const lowest = (list) => list.reduce((a, b) => (b < a ? b : a));

const numericArgs = (args) => {
  const nums = aggregateNumbers(args);
  return isError(nums) ? nums : nums;
};

/**
 * What AND, OR and XOR weigh. In a range or a cell pointed at, text and blanks
 * do not count; typed into the formula, text that is not TRUE or FALSE is
 * #VALUE!. With nothing left to weigh the answer is #VALUE! too.
 */
function logicalsIn(args) {
  const out = [];
  for (const arg of args) {
    if (Array.isArray(arg)) {
      for (const v of arg.flat(Infinity)) {
        if (isError(v)) return v;
        if (typeof v === 'number') out.push(v !== 0);
        else if (typeof v === 'boolean') out.push(v);
      }
    } else if (!isEmpty(arg)) {
      const b = toBoolean(arg);
      if (isError(b)) return b;
      out.push(b);
    }
  }
  return out.length ? out : ERR.VALUE('there is nothing to weigh');
}

/** One required scalar, coerced to number. */
const num1 = (v) => toNumber(Array.isArray(v) ? v.flat(Infinity)[0] : v);
const text1 = (v) => toText(Array.isArray(v) ? v.flat(Infinity)[0] : v);

const def = (fn, opts = {}) => ({ fn, ...opts });

/** A date's number: nothing before the first day has a date. */
const onDate = (v) => { const n = num1(v); return !isError(n) && n < 0 ? ERR.NUM('a date cannot be negative') : n; };

/** Whether the part of a number format a negative number uses writes a date or a time. */
function dateSection(pattern) {
  const sections = String(pattern).replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '').split(';');
  return /[ymdhs]/i.test(sections.length > 1 ? sections[1] : sections[0]);
}

/** A text pattern with * and ?, ~ escaping them, as a whole-text test ignoring case. */
function patternOf(text) {
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '~' && /[*?~]/.test(text[i + 1] ?? '')) { i += 1; out += text[i].replace(/[*?]/, '\\$&'); }
    else if (c === '*') out += '[\\s\\S]*';
    else if (c === '?') out += '[\\s\\S]';
    else out += c.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp('^' + out + '$', 'iu');
}

const numberIn = (v) => (typeof v === 'number' ? v : typeof v === 'string' ? numericText(v) : null);

/**
 * Criteria matching for SUMIF, COUNTIF and their kin, as Excel has it: an
 * error matches that error and a boolean a boolean; a number matches numbers
 * and text that reads as one; < and > compare numbers with numbers and text
 * with text, and leave blanks out; "" finds blanks and empty text, "=" only
 * blanks, "<>" anything that is not blank; * and ? work on text only. A
 * reference to a blank cell asks for 0.
 */
function makeMatcher(criteria) {
  const raw = Array.isArray(criteria) ? criteria.flat(Infinity)[0] : criteria;
  if (isError(raw)) return (v) => isError(v) && v.type === raw.type;
  if (isEmpty(raw)) return (v) => numberIn(v) === 0;
  if (typeof raw === 'number') return (v) => numberIn(v) === raw;
  if (typeof raw === 'boolean') return (v) => v === raw;
  const m = /^(<=|>=|<>|=|<|>)([\s\S]*)$/.exec(raw);
  const op = m ? m[1] : '=';
  const text = m ? m[2] : raw;
  const yes = op !== '<>';

  if (text === '') {
    if (!m) return (v) => isEmpty(v) || v === '';
    if (op === '=') return (v) => isEmpty(v);
    if (op === '<>') return (v) => !isEmpty(v);
    return () => false;
  }
  const upper = text.trim().toUpperCase();
  if (ERROR_TYPES.includes(upper)) return (v) => (isError(v) && v.type === upper) === yes;
  if (upper === 'TRUE' || upper === 'FALSE') {
    const want = upper === 'TRUE';
    if (op === '=' || op === '<>') return (v) => (v === want) === yes;
  }
  const asNumber = textAsNumber(text);
  if (asNumber !== null) {
    if (op === '=') return (v) => numberIn(v) === asNumber;
    // Text that reads as the number is found by = and is still "not" it to <>.
    if (op === '<>') return (v) => v !== asNumber;
    return (v) => typeof v === 'number' && compareBy(op, v, asNumber);
  }
  if (op === '=' || op === '<>') {
    const test = /[*?~]/.test(text) ? patternOf(text) : null;
    const fold = text.toUpperCase();
    return (v) => (typeof v === 'string' && (test ? test.test(v) : v.toUpperCase() === fold)) === yes;
  }
  const fold = text.toUpperCase();
  return (v) => typeof v === 'string' && compareBy(op, v.toUpperCase(), fold);
}

function compareBy(op, a, b) {
  if (op === '<') return a < b;
  if (op === '<=') return a <= b;
  if (op === '>') return a > b;
  return a >= b;
}

/** One argument as a flat vector of scalars, preserving order. */
const vec = (arg) => (Array.isArray(arg) ? arg.flat(Infinity) : [arg]);

/**
 * Which cells an exact lookup takes for its key: values of the key's own
 * kind, text ignoring case, and with `wild` the * and ? of VLOOKUP and MATCH.
 * Blank cells and errors are never found; a blank key asks for 0.
 */
function exactTest(key, wild) {
  if (typeof key === 'string') {
    if (wild && /[*?~]/.test(key)) {
      const re = patternOf(key);
      return (v) => typeof v === 'string' && re.test(v);
    }
    const fold = key.toUpperCase();
    return (v) => typeof v === 'string' && v.toUpperCase() === fold;
  }
  const k = isEmpty(key) ? 0 : key;
  return (v) => v === k;
}

/** Whether an approximate lookup can place this cell against the key at all. */
const sameKind = (v, key) => !isEmpty(v) && !isError(v) && typeof v === typeof (isEmpty(key) ? 0 : key);

/**
 * A long table's first column, indexed for exact lookups: a range read whole
 * is one grid for a whole recalculation, and forty thousand VLOOKUPs into
 * seven thousand rows each walked it from the top. A column holding anything
 * other than numbers, text, booleans, blanks and errors is not indexed.
 */
const exactIndexes = new WeakMap();
const LONG_TABLE = 256;
const exactKey = (v) => (typeof v === 'number' ? 'n' + v : typeof v === 'string' ? 's' + v.toUpperCase() : 'b' + v);

function exactIndexOf(table) {
  let index = exactIndexes.get(table);
  if (index !== undefined) return index;
  index = new Map();
  for (let r = 0; r < table.length; r++) {
    const v = table[r][0];
    if (isError(v) || isEmpty(v)) continue;
    const t = typeof v;
    if (t !== 'number' && t !== 'string' && t !== 'boolean') { index = null; break; }
    const k = exactKey(v);
    if (!index.has(k)) index.set(k, r);
  }
  exactIndexes.set(table, index);
  return index;
}

/** The row an exact lookup stops at, or -1 for none. */
function exactRow(table, key) {
  const wild = typeof key === 'string' && /[*?~]/.test(key);
  const index = !wild && table.length > LONG_TABLE ? exactIndexOf(table) : null;
  if (!index) {
    const hit = exactTest(key, true);
    for (let r = 0; r < table.length; r++) if (hit(table[r][0])) return r;
    return -1;
  }
  return index.get(exactKey(isEmpty(key) ? 0 : key)) ?? -1;
}

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
      if (!t.match(t.cells[i])) { all = false; break; }
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

/** Excel's date parts from a serial, via the shared 1900-epoch arithmetic — 29 February 1900 included. */
const partsOf = (serial) => {
  const p = serialToParts(serial);
  return { y: p.y, m: p.m, day: p.d };
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
    return nums.length ? lowest(nums) : 0;
  }),
  MAX: def((...args) => {
    const nums = numericArgs(args);
    if (isError(nums)) return nums;
    return nums.length ? highest(nums) : 0;
  }),
  COUNT: def((...args) => {
    let n = 0;
    for (const arg of args) {
      if (Array.isArray(arg)) n += arg.flat(Infinity).filter((v) => typeof v === 'number').length;
      else if (!isEmpty(arg) && !isError(toNumber(arg))) n += 1;
    }
    return n;
  }),
  COUNTA: def((...args) => flatten(args).filter((v) => !isEmpty(v)).length),
  COUNTBLANK: def((range) => {
    const inUse = flatten([range]);
    const full = range?.full ? range.full.rows * range.full.cols - inUse.length : 0;
    return full + inUse.filter((v) => isBlank(v)).length;
  }),
  // SUBTOTAL(function_num, ref1, …): the aggregate a table's total row uses.
  // 1–11 are AVERAGE … VARP over what a filter has not hidden; 101–111 the
  // same over the rows nobody has hidden. The evaluator blanks the hidden
  // rows before they reach here.
  SUBTOTAL: def((fn, ...refs) => {
    const n = num1(fn);
    if (isError(n)) return n;
    const which = Math.trunc(n) > 100 ? Math.trunc(n) - 100 : Math.trunc(n);
    const name = ['AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV', 'STDEVP', 'SUM', 'VAR', 'VARP'][which - 1];
    if (!name) return ERR.VALUE(`SUBTOTAL has no function ${n}`);
    return FUNCTIONS[name].fn(...refs);
  }),
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
    return raised(x, y);
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
    // 4.35 is a hair under itself in binary, and =TRUNC(4.35,2) is still 4.35.
    return Math.trunc(Number((n * f).toPrecision(15))) / f;
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
      if (!match(cells[i])) continue;
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
    for (const c of cells) if (match(c)) n += 1;
    // A whole column is read as far as it is used; the rest of it is empty cells, which a criterion can count.
    if (range?.full && match(null)) n += range.full.rows * range.full.cols - cells.length;
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
    const asked = condition();
    // Several conditions, several answers, each from its own place in the two branches.
    if (Array.isArray(asked) && asked.flat(Infinity).length > 1) {
      const lines = Array.isArray(asked[0]) ? asked : [asked];
      const gridOf = (v) => (Array.isArray(v) ? (Array.isArray(v[0]) ? v : [v]) : [[v]]);
      const yes = gridOf(whenTrue ? whenTrue() : true);
      const no = gridOf(whenFalse ? whenFalse() : false);
      const pick = (g, i, j) => {
        const line = g.length === 1 ? g[0] : g[i];
        if (!line) return ERR.NA();
        const v = line.length === 1 ? line[0] : line[j];
        return v === undefined ? ERR.NA() : v;
      };
      const rows = Math.max(lines.length, yes.length, no.length);
      const cols = Math.max(lines[0]?.length ?? 0, yes[0]?.length ?? 0, no[0]?.length ?? 0);
      return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (__, j) => {
        const b = toBoolean(pick(lines, i, j));
        return isError(b) ? b : pick(b ? yes : no, i, j);
      }));
    }
    const c = toBoolean(firstOf(asked));
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
    const all = logicalsIn(args);
    return isError(all) ? all : all.every(Boolean);
  }),
  OR: def((...args) => {
    const all = logicalsIn(args);
    return isError(all) ? all : all.some(Boolean);
  }),
  NOT: def((v) => {
    const b = toBoolean(Array.isArray(v) ? v.flat(Infinity)[0] : v);
    return isError(b) ? b : !b;
  }),
  TRUE: def(() => true),
  FALSE: def(() => false),

  // ---- type tests ---------------------------------------------------------
  ISBLANK: def((v) => isEmpty(Array.isArray(v) ? v.flat(Infinity)[0] : v)),
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
  SUBSTITUTE: def((v, find, replace, instance) => {
    const s = text1(v);
    const f = text1(find);
    const r = text1(replace);
    const which = instance === undefined ? null : num1(instance);
    const e = firstError([s, f, r, which]);
    if (e) return e;
    if (f === '') return s;
    if (which === null) return s.split(f).join(r);
    if (which < 1) return ERR.VALUE('the instance starts at 1');
    const parts = s.split(f);
    const n = Math.trunc(which);
    if (n >= parts.length) return s;
    return parts.slice(0, n).join(f) + r + parts.slice(n).join(f);
  }),
  VALUE: def((v) => {
    if (v === null || v === undefined) return 0;
    const s = text1(v);
    if (isError(s)) return s;
    const n = textAsNumber(s);
    return n === null ? ERR.VALUE('"' + s + '" is not a number') : n;
  }),
  TEXT: def((v, fmt) => {
    const raw = Array.isArray(v) ? v.flat(Infinity)[0] : v;
    if (isError(raw)) return raw;
    const pattern = text1(fmt);
    if (isError(pattern)) return pattern;
    // The grid's own formatter, so a date code, a percentage, a fraction, a
    // currency or a colour section answers as the cell shows it — including
    // 29 February 1900. Text that is not a number stays text, as in Excel.
    const n = typeof raw === 'number' ? raw : toNumber(raw);
    // A date code has nothing to show for a negative number.
    if (typeof n === 'number' && n < 0 && dateSection(pattern)) return ERR.VALUE('a negative number is not a date');
    if (isError(n)) return toText(raw);
    return formatValue(n, pattern).text;
  }),

  // ---- lookup -------------------------------------------------------------
  VLOOKUP: def((needle, table, colIndex, approximate) => {
    const key = Array.isArray(needle) ? needle.flat(Infinity)[0] : needle;
    if (isError(key)) return key;
    const idx = truncated(num1(colIndex));
    if (isError(idx)) return idx;
    if (!Array.isArray(table) || !Array.isArray(table[0])) return ERR.VALUE('VLOOKUP needs a range');
    if (idx < 1) return ERR.VALUE('VLOOKUP columns start at 1');
    if (idx > table[0].length) return ERR.REF('VLOOKUP column ' + idx + ' is outside the range');
    const exact = approximate === undefined ? false : !toBoolean(Array.isArray(approximate) ? approximate.flat(Infinity)[0] : approximate);

    if (exact) {
      const row = exactRow(table, key);
      return row < 0 ? ERR.NA('VLOOKUP found no match') : (table[row][idx - 1] ?? null);
    }
    // approximate: last row whose first column is <= key, assuming sorted
    let best = null;
    for (const row of table) {
      if (!sameKind(row[0], key)) continue;
      if (compareValues(row[0], key) <= 0) best = row;
      else break;
    }
    return best ? (best[idx - 1] ?? null) : ERR.NA('VLOOKUP found no match');
  }),
  INDEX: def((table, rowNum, colNum) => {
    if (!Array.isArray(table)) return ERR.VALUE('INDEX needs a range');
    const grid = Array.isArray(table[0]) ? table : [table];
    const r = truncated(num1(rowNum));
    if (isError(r)) return r;
    const c = colNum === undefined ? 1 : truncated(num1(colNum));
    if (isError(c)) return c;
    // One index into a single row counts along it; a row or column of 0 means all of them.
    const [down, across] = grid.length === 1 && colNum === undefined ? [1, r] : [r, c];
    if (down === 0 && across === 0) return grid;
    if (down === 0) {
      if (across < 0 || across > (grid[0]?.length ?? 0)) return ERR.REF('INDEX column ' + across + ' is outside the range');
      const column = grid.map((line) => [line[across - 1]]);
      if (table.origin) {
        column.origin = { ...table.origin, col: table.origin.col + across - 1, left: table.origin.left + across - 1 };
        if (table.full) column.full = { rows: table.full.rows, cols: 1 };
      }
      return column;
    }
    const row = grid[down - 1];
    // Past the cells in use but inside a whole column: an empty cell, not an error.
    const inside = table.full && down >= 1 && down <= table.full.rows && across >= 1 && across <= table.full.cols;
    if (!row) return inside ? null : ERR.REF('INDEX row ' + down + ' is outside the range');
    if (across === 0) {
      const line = [row];
      if (table.origin) {
        line.origin = { ...table.origin, row: table.origin.row + down - 1, top: table.origin.row + down - 1 };
        if (table.full) line.full = { rows: 1, cols: table.full.cols };
      }
      return line;
    }
    const v = row[across - 1];
    if (v === undefined) return inside ? null : ERR.REF('INDEX column ' + across + ' is outside the range');
    return v;
  }),
  MATCH: def((needle, range, matchType) => {
    const key = Array.isArray(needle) ? needle.flat(Infinity)[0] : needle;
    if (isError(key)) return key;
    const cells = Array.isArray(range) ? range.flat(Infinity) : [range];
    const type = matchType === undefined ? 1 : num1(matchType);
    if (isError(type)) return type;
    if (type === 0) {
      const hit = exactTest(key, true);
      for (let i = 0; i < cells.length; i++) if (hit(cells[i])) return i + 1;
      return ERR.NA('MATCH found no match');
    }
    let best = null;
    for (let i = 0; i < cells.length; i++) {
      if (!sameKind(cells[i], key)) continue;
      const cmp = compareValues(cells[i], key);
      if (type > 0 ? cmp <= 0 : cmp >= 0) best = i + 1;
    }
    return best ?? ERR.NA('MATCH found no match');
  }),

  // ---- dates --------------------------------------------------------------
  // The date on this computer's clock: at 03:00 in Karachi it is already today.
  TODAY: def((_, ctx) => localSerial(ctx.now()).day, { wantsContext: true, volatile: true }),
  NOW: def((_, ctx) => {
    const { day, part } = localSerial(ctx.now());
    return day + part;
  }, { wantsContext: true, volatile: true }),
  DATE: def((y, m, d) => {
    const yy = num1(y);
    const mm = num1(m);
    const dd = num1(d);
    const e = firstError([yy, mm, dd]);
    if (e) return e;
    const year = Math.trunc(yy);
    if (year < 0 || year >= 10000) return ERR.NUM();
    // A two- or three-digit year counts from 1900, on either date system.
    const serial = partsToSerial(year < 1900 ? year + 1900 : year, Math.trunc(mm) - 1, Math.trunc(dd));
    return serial < 0 ? ERR.NUM() : serial;
  }),
  YEAR: def((v) => { const n = onDate(v); return isError(n) ? n : serialToParts(n).y; }),
  MONTH: def((v) => { const n = onDate(v); return isError(n) ? n : serialToParts(n).m + 1; }),
  DAY: def((v) => { const n = onDate(v); return isError(n) ? n : serialToParts(n).d; }),

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
    return Math.ceil(Number((n / s).toPrecision(15))) * s;
  }),
  FLOOR: def((v, sig) => {
    const n = num1(v);
    const s = sig === undefined ? 1 : num1(sig);
    const e = firstError([n, s]);
    if (e) return e;
    if (s === 0) return 0;
    if (n > 0 && s < 0) return ERR.NUM('FLOOR with opposite signs');
    return Math.floor(Number((n / s).toPrecision(15))) * s;
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
  // ---- forecasting (see forecast.js for the method) -----------------------
  'FORECAST.LINEAR': def(forecastFn((x, ys, xs) => {
    const at = toNumber(Array.isArray(x) ? x.flat(Infinity)[0] : x);
    if (isError(at)) return at;
    return linearForecast(at, vec(ys), vec(xs));
  })),
  FORECAST: def((...args) => FUNCTIONS['FORECAST.LINEAR'].fn(...args)),
  'FORECAST.ETS': def(forecastFn((target, values, timeline, seasonality, completion, aggregation) => {
    const q = etsQuestion({
      target: toNumber(Array.isArray(target) ? target.flat(Infinity)[0] : target),
      values: vec(values),
      timeline: vec(timeline),
      seasonality: optionalNumber(seasonality, 1),
      completion: optionalNumber(completion, 1),
      aggregation: optionalNumber(aggregation, 1),
    });
    return forecastAt(q.fit, q.h);
  })),
  'FORECAST.ETS.CONFINT': def(forecastFn((target, values, timeline, confidence, seasonality, completion, aggregation) => {
    const level = optionalNumber(confidence, 0.95);
    if (!(level > 0 && level < 1)) return ERR.NUM('the confidence level is between 0 and 1');
    const q = etsQuestion({
      target: toNumber(Array.isArray(target) ? target.flat(Infinity)[0] : target),
      values: vec(values),
      timeline: vec(timeline),
      seasonality: optionalNumber(seasonality, 1),
      completion: optionalNumber(completion, 1),
      aggregation: optionalNumber(aggregation, 1),
    });
    return confidenceAt(q.fit, q.h, level);
  })),
  'FORECAST.ETS.SEASONALITY': def(forecastFn((values, timeline, completion, aggregation) => {
    const series = prepareSeries(vec(values), vec(timeline), {
      completion: optionalNumber(completion, 1), aggregation: optionalNumber(aggregation, 1),
    });
    return seasonFor(series.ys, 1);
  })),
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
    if (isError(hits)) return hits;
    return hits.length + emptyRowsCounted(pairs);
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
    return picked.length ? highest(picked) : 0;
  }),
  MINIFS: def((range, ...pairs) => {
    const targets = vec(range);
    const hits = ifsIndices(targets.length, pairs);
    if (isError(hits)) return hits;
    const picked = hits.map((i) => targets[i]).filter((v) => typeof v === 'number');
    return picked.length ? lowest(picked) : 0;
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
    const all = logicalsIn(args);
    return isError(all) ? all : all.filter(Boolean).length % 2 === 1;
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
    const idx = truncated(num1(rowIndex));
    if (isError(idx)) return idx;
    if (!Array.isArray(table) || !Array.isArray(table[0])) return ERR.VALUE('HLOOKUP needs a range');
    if (idx < 1) return ERR.VALUE('HLOOKUP rows start at 1');
    if (idx > table.length) return ERR.REF('HLOOKUP row ' + idx + ' is outside the range');
    const exact = approximate === undefined ? false : !toBoolean(Array.isArray(approximate) ? approximate.flat(Infinity)[0] : approximate);
    const headers = table[0];

    if (exact) {
      const hit = exactTest(key, true);
      for (let c = 0; c < headers.length; c++) if (hit(headers[c])) return table[idx - 1][c] ?? null;
      return ERR.NA('HLOOKUP found no match');
    }
    let best = -1;
    for (let c = 0; c < headers.length; c++) {
      if (!sameKind(headers[c], key)) continue;
      if (compareValues(headers[c], key) <= 0) best = c;
      else break;
    }
    return best >= 0 ? (table[idx - 1][best] ?? null) : ERR.NA('HLOOKUP found no match');
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
    const s = onDate(serial);
    if (isError(s)) return s;
    const t = type === undefined ? 1 : num1(type);
    if (isError(t)) return t;
    const sunday0 = serialToParts(s).weekday; // 0 = Sunday, by Excel's arithmetic
    if (t === 1) return sunday0 + 1;
    if (t === 2) return sunday0 === 0 ? 7 : sunday0;
    if (t === 3) return sunday0 === 0 ? 6 : sunday0 - 1;
    return ERR.NUM('WEEKDAY type must be 1, 2 or 3');
  }),
  HOUR: def((v) => {
    const n = onDate(v);
    if (isError(n)) return n;
    return Math.floor(Math.round((n - Math.floor(n)) * 86400) / 3600) % 24;
  }),
  MINUTE: def((v) => {
    const n = onDate(v);
    if (isError(n)) return n;
    return Math.floor((Math.round((n - Math.floor(n)) * 86400) % 3600) / 60);
  }),
  SECOND: def((v) => {
    const n = onDate(v);
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
  ROWS: def((v) => (Array.isArray(v) ? (v.full?.rows ?? v.length) : 1)),
  COLUMNS: def((v) => (Array.isArray(v) ? v.full?.cols ?? (Array.isArray(v[0]) ? v[0].length : v.length) : 1)),
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
    const r1c1 = a1 !== undefined && a1 !== null && a1 !== '' && toBoolean(a1) === false;
    // abs: 1 = $A$1, 2 = A$1, 3 = $A1, 4 = A1, Excel's numbering. In R1C1 a relative part is in brackets: R2C[3].
    const colTxt = r1c1 ? (mode === 1 || mode === 3 ? `C${c}` : `C[${c}]`) : (mode === 1 || mode === 3 ? '$' : '') + indexToCol(c - 1);
    const rowTxt = r1c1 ? (mode === 1 || mode === 2 ? `R${r}` : `R[${r}]`) : (mode === 1 || mode === 2 ? '$' : '') + r;
    if (r1c1 && (sheetName === undefined || sheetName === null || sheetName === '')) return rowTxt + colTxt;
    if (sheetName === undefined || sheetName === null || sheetName === '') return colTxt + rowTxt;
    const s = toText(sheetName);
    if (isError(s)) return s;
    const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(s) ? s : "'" + s.replace(/'/g, "''") + "'";
    return quoted + '!' + (r1c1 ? rowTxt + colTxt : colTxt + rowTxt);
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

  // ---- Excel 365's arrays: shaping, stacking, picking ----------------------
  // These pass an error inside an array on as an element, as Excel does,
  // rather than letting it stand for the whole answer (`keepsErrors`); an
  // error in a number they are given is still the answer.
  XMATCH: def((needle, lookupArray, matchMode, searchMode) => {
    const looks = vec(lookupArray);
    const mode = matchMode === undefined || isBlank(firstOf(matchMode)) ? 0 : num1(matchMode);
    const dir = searchMode === undefined || isBlank(firstOf(searchMode)) ? 1 : num1(searchMode);
    const e = firstError([mode, dir]);
    if (e) return e;
    if (![0, -1, 1, 2].includes(mode)) return ERR.VALUE('XMATCH match mode is 0, -1, 1 or 2');
    if (![1, -1, 2, -2].includes(dir)) return ERR.VALUE('XMATCH search mode is 1, -1, 2 or -2');
    const order = [...looks.keys()];
    if (dir < 0) order.reverse();
    const one = (key) => {
      if (isError(key)) return key;
      if (mode === 2) {
        const re = wildcardRegex(toText(key));
        for (const i of order) if (re.test(toText(looks[i]))) return i + 1;
        return ERR.NA('XMATCH found no match');
      }
      let best = -1;
      for (const i of order) {
        const cmp = compareValues(looks[i], key);
        if (isError(cmp)) continue;
        if (cmp === 0) return i + 1;
        if (mode === -1 && cmp < 0 && (best < 0 || compareValues(looks[i], looks[best]) > 0)) best = i;
        if (mode === 1 && cmp > 0 && (best < 0 || compareValues(looks[i], looks[best]) < 0)) best = i;
      }
      return best >= 0 ? best + 1 : ERR.NA('XMATCH found no match');
    };
    return Array.isArray(needle) ? gridOf(needle).map((line) => line.map(one)) : one(needle);
  }),
  TEXTSPLIT: def((value, colDelim, rowDelim, ignoreEmpty, matchMode, padWith) => {
    const text = text1(value);
    if (isError(text)) return text;
    const list = (d) => (d === undefined || isBlank(firstOf(d)) ? [] : vec(d).map((x) => toText(x)).filter((x) => x !== ''));
    const cols = list(colDelim);
    const rows = list(rowDelim);
    if (!cols.length && !rows.length) return ERR.VALUE('TEXTSPLIT needs a delimiter');
    const skip = ignoreEmpty !== undefined && toBoolean(firstOf(ignoreEmpty)) === true;
    const fold = matchMode !== undefined && num1(matchMode) === 1;
    const pad = padWith === undefined ? ERR.NA('TEXTSPLIT padded a short row') : firstOf(padWith);
    const splitter = (delims) => {
      if (!delims.length) return null;
      const source = [...delims].sort((a, b) => b.length - a.length).map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|');
      return new RegExp(source, fold ? 'i' : '');
    };
    const byRow = splitter(rows);
    const byCol = splitter(cols);
    const lines = (byRow ? text.split(byRow) : [text]).filter((line) => !skip || line !== '');
    const grid = lines.map((line) => (byCol ? line.split(byCol) : [line]).filter((cell) => !skip || cell !== ''));
    const width = Math.max(1, ...grid.map((line) => line.length));
    const out = grid.map((line) => { const l = [...line]; while (l.length < width) l.push(pad); return l; });
    return out.length ? out : [['']];
  }),
  TEXTBEFORE: def((value, delimiter, instance, matchMode, matchEnd, ifNotFound) => textAround('before', value, delimiter, instance, matchMode, matchEnd, ifNotFound)),
  TEXTAFTER: def((value, delimiter, instance, matchMode, matchEnd, ifNotFound) => textAround('after', value, delimiter, instance, matchMode, matchEnd, ifNotFound)),
  VSTACK: def((...arrays) => {
    const grids = arrays.map(gridOf);
    const width = Math.max(...grids.map((g) => g[0]?.length ?? 0));
    return grids.flatMap((g) => g.map((line) => padTo(line, width)));
  }, { keepsErrors: true }),
  HSTACK: def((...arrays) => {
    const grids = arrays.map(gridOf);
    const height = Math.max(...grids.map((g) => g.length));
    const out = [];
    for (let i = 0; i < height; i++) out.push(grids.flatMap((g) => (i < g.length ? g[i] : padTo([], g[0]?.length ?? 1))));
    return out;
  }, { keepsErrors: true }),
  TAKE: def((array, rows, cols) => cutGrid(array, rows, cols, 'take'), { keepsErrors: true }),
  DROP: def((array, rows, cols) => cutGrid(array, rows, cols, 'drop'), { keepsErrors: true }),
  CHOOSECOLS: def((array, ...picks) => {
    const g = gridOf(array);
    const at = pickIndices(picks, g[0]?.length ?? 0, 'CHOOSECOLS');
    return isError(at) ? at : g.map((line) => at.map((j) => line[j]));
  }, { keepsErrors: true }),
  CHOOSEROWS: def((array, ...picks) => {
    const g = gridOf(array);
    const at = pickIndices(picks, g.length, 'CHOOSEROWS');
    return isError(at) ? at : at.map((i) => [...g[i]]);
  }, { keepsErrors: true }),
  TOCOL: def((array, ignore, byColumn) => {
    const flat = lineUp(array, ignore, byColumn);
    return isError(flat) ? flat : flat.length ? flat.map((v) => [v]) : ERR.CALC('TOCOL kept nothing');
  }, { keepsErrors: true }),
  TOROW: def((array, ignore, byColumn) => {
    const flat = lineUp(array, ignore, byColumn);
    return isError(flat) ? flat : flat.length ? [flat] : ERR.CALC('TOROW kept nothing');
  }, { keepsErrors: true }),
  WRAPROWS: def((vector, count, padWith) => wrapLine(vector, count, padWith, 'rows'), { keepsErrors: true }),
  WRAPCOLS: def((vector, count, padWith) => wrapLine(vector, count, padWith, 'cols'), { keepsErrors: true }),
  EXPAND: def((array, rows, cols, padWith) => {
    const g = gridOf(array);
    const h = rows === undefined || isBlank(firstOf(rows)) ? g.length : num1(rows);
    const w = cols === undefined || isBlank(firstOf(cols)) ? (g[0]?.length ?? 0) : num1(cols);
    const e = firstError([h, w]);
    if (e) return e;
    if (h < g.length || w < (g[0]?.length ?? 0)) return ERR.VALUE('EXPAND cannot make an array smaller');
    const pad = padWith === undefined ? ERR.NA('EXPAND padded the array') : firstOf(padWith);
    const out = [];
    for (let i = 0; i < Math.trunc(h); i++) out.push(i < g.length ? padTo(g[i], Math.trunc(w), pad) : padTo([], Math.trunc(w), pad));
    return out;
  }, { keepsErrors: true }),
  SORTBY: def((array, ...keys) => {
    const g = gridOf(array);
    if (!keys.length) return ERR.VALUE('SORTBY needs something to sort by');
    const sorts = [];
    for (let k = 0; k < keys.length; k += 2) {
      const by = gridOf(keys[k]);
      const order = keys[k + 1] === undefined || isBlank(firstOf(keys[k + 1])) ? 1 : num1(keys[k + 1]);
      if (isError(order)) return order;
      if (order !== 1 && order !== -1) return ERR.VALUE('SORTBY order is 1 or -1');
      sorts.push({ by, order });
    }
    // A key one column tall sorts the rows; one row wide, the columns.
    const byCols = sorts.every(({ by }) => by.length === 1 && (by[0]?.length ?? 0) === (g[0]?.length ?? 0) && g.length !== by[0].length);
    const lines = byCols ? transposeGrid(g) : g;
    const keyOf = ({ by }) => (byCols ? by[0] : by.map((line) => line[0]));
    if (sorts.some((s) => keyOf(s).length !== lines.length)) return ERR.VALUE('SORTBY needs a key the size of the array');
    const order = lines.map((_, i) => i).sort((a, b) => {
      for (const s of sorts) {
        const key = keyOf(s);
        const c = compareValues(key[a], key[b]);
        if (!isError(c) && c !== 0) return c * s.order;
      }
      return a - b;
    });
    const sorted = order.map((i) => lines[i]);
    return byCols ? transposeGrid(sorted) : sorted;
  }),
  RANDARRAY: def((rows, cols, min, max, whole) => {
    const opt = (v, d) => (v === undefined || isBlank(firstOf(v)) ? d : num1(v));
    const r = opt(rows, 1);
    const c = opt(cols, 1);
    const lo = opt(min, 0);
    const hi = opt(max, 1);
    const e = firstError([r, c, lo, hi]);
    if (e) return e;
    const integer = whole !== undefined && toBoolean(firstOf(whole)) === true;
    if (r < 1 || c < 1 || r * c > 1048576) return ERR.VALUE('RANDARRAY needs positive dimensions');
    if (hi < lo) return ERR.VALUE('RANDARRAY maximum is below its minimum');
    if (integer && (lo !== Math.trunc(lo) || hi !== Math.trunc(hi))) return ERR.VALUE('RANDARRAY whole numbers need a whole minimum and maximum');
    const draw = () => (integer ? lo + Math.floor(Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo));
    return Array.from({ length: Math.trunc(r) }, () => Array.from({ length: Math.trunc(c) }, draw));
  }, { volatile: true }),
  VALUETOTEXT: def((value, format) => {
    const strict = format !== undefined && num1(format) === 1;
    return Array.isArray(value) ? gridOf(value).map((line) => line.map((v) => valueText(v, strict))) : valueText(value, strict);
  }, { keepsErrors: true }),
  ARRAYTOTEXT: def((array, format) => {
    const strict = format !== undefined && num1(format) === 1;
    const g = gridOf(array);
    if (!strict) return g.flat().map((v) => valueText(v, false)).join(', ');
    return '{' + g.map((line) => line.map((v) => valueText(v, true)).join(',')).join(';') + '}';
  }, { keepsErrors: true }),
  // `@` written out: the one value an array stands for where one is wanted.
  SINGLE: def((v) => firstOf(v)),

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

Object.assign(FUNCTIONS, moreFunctions(FUNCTIONS));

// A number too big to hold is #NUM! wherever it arises, and goes no further as a number.
for (const entry of Object.values(FUNCTIONS)) {
  const inner = entry.fn;
  entry.fn = (...args) => {
    const answer = inner(...args);
    return typeof answer === 'number' && !Number.isFinite(answer) ? ERR.NUM() : answer;
  };
}

// These answer for an error themselves: a count or a lookup is not spoiled by
// one bad cell it does not use, and ISNUMBER exists to look at one.
for (const name of ['ISNUMBER', 'ISTEXT', 'ISBLANK', 'ISLOGICAL', 'ISNONTEXT', 'COUNT', 'COUNTA', 'COUNTBLANK',
  'COUNTIF', 'COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS', 'MAXIFS', 'MINIFS',
  'MATCH', 'VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'XMATCH', 'INDEX', 'ROWS', 'COLUMNS']) {
  if (FUNCTIONS[name]) FUNCTIONS[name].keepsErrors = true;
}

// A cell these are pointed at counts as a range of one.
for (const name of ['SUM', 'AVERAGE', 'MAX', 'MIN', 'PRODUCT', 'COUNT', 'MEDIAN', 'STDEV', 'STDEV.S', 'STDEV.P', 'STDEVP',
  'VAR', 'VAR.S', 'VAR.P', 'VARP', 'MAXA', 'MINA', 'AVERAGEA', 'AND', 'OR', 'XOR']) {
  if (FUNCTIONS[name]) FUNCTIONS[name].refsAsRanges = true;
}

export const FUNCTION_NAMES = Object.keys(FUNCTIONS).sort();
export const isVolatile = (name) => Boolean(FUNCTIONS[name]?.volatile);
export { FormulaError, makeMatcher };
