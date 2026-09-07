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
export const isBlank = (v) => v === null || v === undefined || v === '';

/** First error in a list, or null. Operators use this before doing any work. */
export function firstError(values) {
  for (const v of values) {
    if (isError(v)) return v;
    if (Array.isArray(v)) {
      const nested = firstError(v.flat(Infinity));
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

/** Coerce to a number for arithmetic. Returns a FormulaError if impossible. */
export function toNumber(v) {
  if (isError(v)) return v;
  if (isBlank(v)) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : ERR.NUM();
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    const n = numericText(v);
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
  if (isBlank(v)) return false;
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

export function dateToSerial(date) {
  const utc = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  const days = (utc - EPOCH_UTC) / MS_PER_DAY;
  return days >= 60 ? days + 1 : days;
}

export function serialToDate(serial) {
  const n = Math.floor(serial);
  const days = n > 60 ? n - 1 : n;
  return new Date(EPOCH_UTC + days * MS_PER_DAY);
}

/**
 * Compare two values the way a spreadsheet does.
 * Text comparison is case-insensitive; numbers sort before text; blanks equal
 * both 0 and "".
 * @returns {-1|0|1|FormulaError}
 */
export function compareValues(a, b) {
  const err = firstError([a, b]);
  if (err) return err;

  const aBlank = isBlank(a);
  const bBlank = isBlank(b);
  if (aBlank && bBlank) return 0;
  if (aBlank) return typeof b === 'number' ? (0 < b ? -1 : 0 > b ? 1 : 0) : (b === '' ? 0 : -1);
  if (bBlank) return typeof a === 'number' ? (a < 0 ? -1 : a > 0 ? 1 : 0) : (a === '' ? 0 : 1);

  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2); // number < text < boolean
  if (rank(a) !== rank(b)) return rank(a) < rank(b) ? -1 : 1;

  if (typeof a === 'number') return a < b ? -1 : a > b ? 1 : 0;
  if (typeof a === 'boolean') return a === b ? 0 : a ? 1 : -1;
  const x = String(a).toUpperCase();
  const y = String(b).toUpperCase();
  return x < y ? -1 : x > y ? 1 : 0;
}
