/**
 * Spreadsheet value semantics.
 *
 * The single most important thing to get right, and the thing that makes a
 * formula engine feel wrong when it is wrong: **errors are values, not
 * exceptions.** `=A1+1` where A1 is `#DIV/0!` yields `#DIV/0!` — it does not
 * throw, and it does not yield NaN. Every operator and nearly every function
 * propagates the first error it meets, left to right. `IFERROR` and the `IS*`
 * family are the only things that look at an error without spreading it.
 *
 * The second thing is coercion. A spreadsheet is not JavaScript:
 *   - an empty cell is 0 in arithmetic and "" in concatenation
 *   - TRUE is 1 in arithmetic, "TRUE" in concatenation
 *   - a text value that looks like a number IS a number in arithmetic
 *     ("5"+1 = 6), but a text value that does not is #VALUE!
 *   - comparison is case-insensitive for text, and numbers sort before text
 *
 * Getting these wrong produces a sheet that is subtly, unreproducibly different
 * from the one the customer built in Excel, which is worse than one that
 * obviously does not work.
 */

export const ERROR_TYPES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#CIRCULAR!', '#SPILL!', '#CALC!'];

export class FormulaError {
  constructor(type, detail) {
    if (!ERROR_TYPES.includes(type)) throw new Error('unknown error type: ' + type);
    this.type = type;
    this.detail = detail ?? null;
  }
  toString() { return this.type; }
  toJSON() { return this.type; }
}

export const ERR = {
  NULL: () => new FormulaError('#NULL!'),
  DIV0: () => new FormulaError('#DIV/0!'),
  VALUE: (d) => new FormulaError('#VALUE!', d),
  REF: (d) => new FormulaError('#REF!', d),
  NAME: (d) => new FormulaError('#NAME?', d),
  NUM: (d) => new FormulaError('#NUM!', d),
  NA: (d) => new FormulaError('#N/A', d),
  CIRCULAR: (d) => new FormulaError('#CIRCULAR!', d),
  SPILL: (d) => new FormulaError('#SPILL!', d),
  CALC: (d) => new FormulaError('#CALC!', d),
};

export const isError = (v) => v instanceof FormulaError;
/** Nothing, or empty text: what clears a cell, and what leaves out an optional argument. */
export const isBlank = (v) => v === null || v === undefined || v === '';
/** A cell with nothing in it. Empty text from a formula is text: =""+1 is #VALUE!. */
export const isEmpty = (v) => v === null || v === undefined;

/**
 * A long array's first error, once worked out: a range read whole is one
 * grid for a whole recalculation, and twenty thousand lookups into it each
 * flattened all hundred thousand of its cells to look for one.
 */
const errorsIn = new WeakMap();
const LONG = 256;

/** First error in a list, or null. Operators use this before doing any work. */
export function firstError(values) {
  for (const v of values) {
    if (isError(v)) return v;
    if (Array.isArray(v)) {
      const long = v.length > LONG;
      let nested = long ? errorsIn.get(v) : undefined;
      if (nested === undefined) {
        nested = firstError(v.flat(Infinity));
        if (long) errorsIn.set(v, nested);
      }
      if (nested) return nested;
    }
  }
  return null;
}

/** Text that is entirely a number, the way a spreadsheet decides it. */
export function numericText(s) {
  const t = String(s).trim();
  if (t === '') return null;
  // Reject things JS would happily accept but a spreadsheet would not.
  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?%?$/.test(t)) return null;
  const pct = t.endsWith('%');
  const n = Number(pct ? t.slice(0, -1) : t);
  return Number.isFinite(n) ? (pct ? n / 100 : n) : null;
}

const MONTH_NAMES = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

let dayFirstHere = null;
/** Whether this computer writes the day before the month, as Excel reads 3/4/2026 by it. */
function dayFirst() {
  if (dayFirstHere === null) {
    try {
      const order = new Intl.DateTimeFormat(undefined, { year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date(Date.UTC(2001, 10, 22, 12))).map((p) => p.type);
      dayFirstHere = order.indexOf('day') < order.indexOf('month');
    } catch { dayFirstHere = true; }
  }
  return dayFirstHere;
}

let yearHere = null;

/**
 * How dates written short are read from here on, as on another computer on
 * another day: `dayFirst` true reads 3/4/2026 as 3 April, false as 4 March;
 * `year` is the year a date written with none (1/2, 1 Jan) falls in. Null,
 * the default for each: this computer's own order, and this year by its
 * clock. For answers made elsewhere, held to how they were made.
 */
export function readDatesAs({ dayFirst = null, year = null } = {}) {
  dayFirstHere = dayFirst === null ? null : Boolean(dayFirst);
  yearHere = year === null ? null : Number(year);
}

/** A number raised to a power, with the answers Excel gives where arithmetic has none. */
export function raised(a, b) {
  if (a === 0 && b === 0) return ERR.NUM('0^0');
  if (a === 0 && b < 0) return ERR.DIV0();
  let r = a ** b;
  // A root of a negative number is real when the root is odd: (-8)^(1/3).
  if (a < 0 && !Number.isInteger(b)) {
    const root = Number((1 / b).toPrecision(15));
    if (Number.isInteger(root) && root % 2 !== 0) r = -((-a) ** b);
  }
  return Number.isFinite(r) ? r : ERR.NUM();
}

const fullYear = (y) => (y.length <= 2 ? (Number(y) < 30 ? 2000 : 1900) + Number(y) : Number(y));

function dateSerial(y, m, d) {
  if (y < 1900 || y > 9999 || m < 0 || m > 11 || d < 1) return null;
  const at = new Date(Date.UTC(y, m, d));
  return at.getUTCMonth() === m ? partsToSerial(y, m, d) : null;
}

/** A time written as text, as a part of a day, or null: 12:30, 7:05:09 PM. */
export function parseTimeText(text) {
  const m = /(?:^|\s)(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?\s*(am|pm|a|p)?\s*$/i.exec(String(text).trim());
  if (!m) return null;
  let h = Number(m[1]);
  const min = Number(m[2]);
  const sec = Number(m[3] ?? 0);
  if (min > 59 || sec >= 60) return null;
  if (m[4]) {
    if (h < 1 || h > 12) return null;
    h = (h % 12) + (/p/i.test(m[4]) ? 12 : 0);
  }
  if (h > 9999) return null;
  return (h * 3600 + min * 60 + sec) / 86400;
}

/**
 * A date written as text, as a serial, or null. The year must be there: 1/2
 * alone is left as text. With `dropTime` a time after the date is not added.
 */
export function parseDateText(text, { dropTime = false, guessYear = false } = {}) {
  const t = String(text).trim();
  const time = /\s\d{1,2}:\d/.test(t) ? t.slice(t.search(/\s\d{1,2}:\d/)) : '';
  const date = (time ? t.slice(0, t.length - time.length) : t).trim();
  let serial = null;
  let m;
  const monthOf = (name) => MONTH_NAMES.indexOf(name.slice(0, 3).toLowerCase());
  if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(date))) serial = dateSerial(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  else if ((m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})$/.exec(date))) {
    const [d, mon] = dayFirst() ? [m[1], m[2]] : [m[2], m[1]];
    serial = dateSerial(fullYear(m[3]), Number(mon) - 1, Number(d));
  } else if ((m = /^(\d{1,2})[- ]([A-Za-z]{3,9})[-, ]+(\d{4}|\d{2})$/.exec(date))) serial = monthOf(m[2]) < 0 ? null : dateSerial(fullYear(m[3]), monthOf(m[2]), Number(m[1]));
  else if ((m = /^([A-Za-z]{3,9})[- ](\d{1,2})[-, ]+(\d{4}|\d{2})$/.exec(date))) serial = monthOf(m[1]) < 0 ? null : dateSerial(fullYear(m[3]), monthOf(m[1]), Number(m[2]));
  // Worked with as a number, 1/2 is this year's and Jan 2026 the first of the month, as Excel takes them.
  if (serial === null && guessYear) {
    const year = yearHere ?? new Date().getFullYear();
    if ((m = /^(\d{1,2})[-/](\d{1,2})$/.exec(date))) {
      const [d, mon] = dayFirst() ? [m[1], m[2]] : [m[2], m[1]];
      serial = dateSerial(year, Number(mon) - 1, Number(d));
    } else if ((m = /^(\d{1,2})[- ]([A-Za-z]{3,9})$/.exec(date))) serial = monthOf(m[2]) < 0 ? null : dateSerial(year, monthOf(m[2]), Number(m[1]));
    else if ((m = /^([A-Za-z]{3,9})[-, ]+(\d{1,4})$/.exec(date))) serial = monthOf(m[1]) < 0 ? null : dateSerial(fullYear(m[2]), monthOf(m[1]), 1);
  }
  if (serial === null) return null;
  if (dropTime || !time) return serial;
  const part = parseTimeText(time);
  return part === null ? null : serial + part;
}

/** What arithmetic reads a text as: a number, 1,234.5, a date or a time; else null. */
export function textAsNumber(s) {
  const plain = numericText(s);
  if (plain !== null) return plain;
  const t = String(s).trim();
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d*)?%?$/.test(t)) return numericText(t.replace(/,/g, ''));
  // Brackets are how accounts write a negative.
  const bracketed = /^\(([\d.,]+%?)\)$/.exec(t);
  if (bracketed) {
    const inner = numericText(bracketed[1].replace(/,/g, ''));
    return inner === null ? null : -inner;
  }
  const date = parseDateText(t, { guessYear: true });
  if (date !== null) return date;
  return /^\d{1,2}:\d/.test(t) ? parseTimeText(t) : null;
}

/** Coerce to a number for arithmetic. Returns a FormulaError if impossible. */
export function toNumber(v) {
  if (isError(v)) return v;
  if (isEmpty(v)) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : ERR.NUM();
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    const n = textAsNumber(v);
    return n === null ? ERR.VALUE('"' + v + '" is not a number') : n;
  }
  if (v instanceof Date) return dateToSerial(v);
  return ERR.VALUE();
}

/** Coerce to text for concatenation. */
export function toText(v) {
  if (isError(v)) return v;
  if (isBlank(v)) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return formatNumber(v);
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v);
}

/** Coerce to boolean for logical context. */
export function toBoolean(v) {
  if (isError(v)) return v;
  if (isEmpty(v)) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') {
    const u = v.trim().toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    const n = numericText(v);
    if (n !== null) return n !== 0;
  }
  return ERR.VALUE('cannot read "' + v + '" as TRUE or FALSE');
}

/**
 * Default number rendering. Spreadsheets show up to 15 significant digits and
 * hide binary-floating-point noise; `0.1+0.2` must display as `0.3`, not
 * `0.30000000000000004`, or every arithmetic sheet looks broken.
 */
export function formatNumber(n) {
  if (!Number.isFinite(n)) return String(n);
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const rounded = Number(n.toPrecision(15));
  return String(rounded);
}

/** Round the way a spreadsheet does: half away from zero, not banker's. */
export function roundHalfAwayFromZero(n, digits = 0) {
  const f = 10 ** digits;
  const scaled = n * f;
  // Nudge past floating-point error before deciding the half case, or
  // ROUND(2.675, 2) returns 2.67 and a finance sheet is a penny out.
  const corrected = Number(scaled.toPrecision(15));
  const r = corrected < 0 ? -Math.round(-corrected) : Math.round(corrected);
  return r / f;
}

// ---- dates ---------------------------------------------------------------
// Serial numbers with the 1900 epoch, including the leap-year bug that is part
// of the format: day 60 is "29 Feb 1900", a date that never existed. Files
// carry serials, so we reproduce the arithmetic rather than the calendar.

// Serial 1 is 1900-01-01, so the zero point is 1899-12-31. Serial 60 is the
// date that never existed ("29 Feb 1900"), which is why everything from
// 1900-03-01 onward is shifted by one.
const EPOCH_UTC = Date.UTC(1899, 11, 31);
const MS_PER_DAY = 86400000;

// A workbook on the 1904 date system counts from 1 January 1904, which is day
// 1462 of the other. The file's serials stay as they are; only the calendar
// they mean moves, so this is applied where a serial meets a calendar.
const DAYS_TO_1904 = 1462;
let dayShift = 0;

/** Run `fn` with serials read on the 1904 system when `on`. */
export function withDate1904(on, fn) {
  if (!on) return fn();
  const was = dayShift;
  dayShift = DAYS_TO_1904;
  try { return fn(); } finally { dayShift = was; }
}

export function dateToSerial(date) {
  const utc = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  const days = (utc - EPOCH_UTC) / MS_PER_DAY;
  return (days >= 60 ? days + 1 : days) - dayShift;
}

/**
 * A Date for a serial — for anything that needs a calendar. Serial 60 has no
 * Date: it is 29 February 1900, which never happened, so it answers as
 * 1 March. Use `serialToParts` where the day matters.
 */
export function serialToDate(serial) {
  const n = Math.floor(serial) + dayShift;
  const days = n > 60 ? n - 1 : n;
  return new Date(EPOCH_UTC + days * MS_PER_DAY);
}

/**
 * The calendar parts of a serial, as Excel has them.
 *
 * Serial 60 is 29 February 1900 in Excel and in every file Excel wrote; no
 * Date can hold it, so the parts are given directly. The weekday is Excel's
 * arithmetic, not the calendar's: serial 1 was a Sunday to Excel, whatever
 * the almanac says of 1 January 1900, and from 1 March 1900 the two agree —
 * the phantom day cancels the error. A workbook's WEEKDAY() must agree with
 * Excel's, which is why it is not read off a Date.
 *
 * @returns {{ y: number, m: number, d: number, weekday: number }} m is the
 *   month index (January 0); weekday is 0 for Sunday.
 */
export function serialToParts(serial) {
  const n = Math.floor(serial) + dayShift;
  const weekday = (((n - 1) % 7) + 7) % 7;
  if (n === 60) return { y: 1900, m: 1, d: 29, weekday };
  // Serial zero is the day before the first, which Excel writes as 0 January 1900.
  if (n === 0) return { y: 1900, m: 0, d: 0, weekday };
  const date = serialToDate(Math.floor(serial));
  return { y: date.getUTCFullYear(), m: date.getUTCMonth(), d: date.getUTCDate(), weekday };
}

/**
 * The serial of a year, month index and day, with Excel's overflow —
 * (2023, 12, 1) is January 2024, (2023, 0, 0) is the last of December 2022 —
 * and the day that never existed: (1900, 1, 29) and (1900, 2, 0) are 60.
 */
export function partsToSerial(y, m, d) {
  if (y === 1900 && ((m === 1 && d === 29) || (m === 2 && d === 0))) return 60 - dayShift;
  return dateToSerial(new Date(Date.UTC(y, m, d)));
}

/**
 * Compare two values the way a spreadsheet does.
 * Text comparison is case-insensitive; numbers sort before text. A blank takes
 * the other side's kind, so it equals 0, "" and FALSE; empty text is text.
 * @returns {-1|0|1|FormulaError}
 */
export function compareValues(a, b) {
  const err = firstError([a, b]);
  if (err) return err;

  if (isEmpty(a) && isEmpty(b)) return 0;
  const zeroOf = (v) => (typeof v === 'number' ? 0 : typeof v === 'boolean' ? false : '');
  if (isEmpty(a)) return compareValues(zeroOf(b), b);
  if (isEmpty(b)) return compareValues(a, zeroOf(a));

  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2); // number < text < boolean
  if (rank(a) !== rank(b)) return rank(a) < rank(b) ? -1 : 1;

  // Numbers that agree to fifteen significant digits are the same number, as a spreadsheet compares them: 0.1+0.2 is 0.3.
  if (typeof a === 'number') return a === b || a.toPrecision(15) === b.toPrecision(15) ? 0 : a < b ? -1 : 1;
  if (typeof a === 'boolean') return a === b ? 0 : a ? 1 : -1;
  const x = String(a).toUpperCase();
  const y = String(b).toUpperCase();
  return x < y ? -1 : x > y ? 1 : 0;
}
