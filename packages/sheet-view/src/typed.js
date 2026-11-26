// What a typed entry is, as Excel reads one: a date, a time, a percentage,
// an amount of money or a number written with commas becomes a number, and
// brings the number format that shows it the way it was typed.
//
// Pure: text in, `{ value, format, kind }` out, or null when the entry is
// none of these (a plain number, words, a formula). The view decides what to
// do with the format: a cell that already wears one keeps its own.

import { parseDateText, parseTimeText, readsDayFirst, formatNumber, serialToParts, withDate1904 } from '@rutba/formula';
import { isDateFormat, splitSections } from '@rutba/formula/numfmt';

const MONEY = '£$€¥₹₨';
const PLAIN = String.raw`(?:\d+\.?\d*|\.\d+)`;
const GROUPED = String.raw`\d{1,3}(?:,\d{3})+(?:\.\d*)?`;
const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];

const decimalsOf = (digits) => (digits.includes('.') && !digits.endsWith('.') ? digits.length - digits.indexOf('.') - 1 : 0);
const numberOf = (digits) => {
  const n = Number(digits.replace(/,/g, ''));
  return Number.isFinite(n) ? n : null;
};

// Functions whose answer is always a day.
const DATE_MAKERS = new Set(['DATE', 'TODAY', 'EDATE', 'EOMONTH', 'WORKDAY', 'WORKDAY.INTL', 'DATEVALUE']);

/** The function a formula is one call of (its closing bracket ends the formula), in capitals; else null. */
function outerCall(body) {
  const m = /^([A-Z][A-Z0-9.]*)\(/i.exec(body);
  if (!m) return null;
  let depth = 0;
  let quoted = false;
  for (let i = m[0].length - 1; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') quoted = !quoted;
    else if (quoted) continue;
    else if (ch === '(') depth++;
    else if (ch === ')') {
      depth--;
      if (depth === 0) return i === body.length - 1 ? m[1].toUpperCase() : null;
    }
  }
  return null;
}

/**
 * The number format a typed formula brings to a cell that has none of its
 * own, as Excel gives one: `=DATE(2026,11,30)` and `=TODAY()` are shown as
 * dates, `=NOW()` as a date and a time, `=TIME(9,30,0)` as a time, and a
 * formula that is one cell, or one cell with days added or taken away
 * (`=A5`, `=A5+7`), takes that cell's date format. Null for everything else.
 *
 * `formatOf(ref)` gives the format of a cell on the same sheet by its
 * reference (A5), for the last kind.
 */
export function formulaFormat(input, formatOf = null) {
  if (typeof input !== 'string' || !input.startsWith('=')) return null;
  const body = input.slice(1).trim();
  const shortDate = readsDayFirst() ? 'dd/mm/yyyy' : 'mm/dd/yyyy';
  const call = outerCall(body);
  if (call) {
    if (DATE_MAKERS.has(call)) return shortDate;
    if (call === 'NOW') return shortDate + ' hh:mm';
    if (call === 'TIME' || call === 'TIMEVALUE') return 'hh:mm';
    return null;
  }
  const m = /^\$?([A-Z]{1,3})\$?(\d{1,7})(?:\s*[+-]\s*\d+(?:\.\d+)?)?$/i.exec(body);
  if (!m || !formatOf) return null;
  const code = formatOf(m[1].toUpperCase() + m[2]);
  return code && code !== 'General' && isDateFormat(code) ? code : null;
}

/** A month's name, whole or cut to three letters or more (Sept too). */
function isMonthName(word) {
  const w = word.toLowerCase();
  return w.length >= 3 && MONTHS.some((m) => m.startsWith(w));
}

function readPercent(t) {
  const m = new RegExp(`^([+-]?)(${GROUPED}|${PLAIN})\\s?%$`).exec(t);
  if (!m) return null;
  const n = numberOf(m[2]);
  if (n === null) return null;
  const value = Number(((m[1] === '-' ? -n : n) / 100).toPrecision(15));
  return { value, format: decimalsOf(m[2]) ? '0.00%' : '0%', kind: 'percent' };
}

function readMoney(t) {
  // The sign before the symbol or after it, or the amount in brackets as accounts write a negative.
  const bracketed = /^\((.*)\)$/.exec(t);
  const inner = bracketed ? bracketed[1].trim() : t;
  const m = new RegExp(`^([+-]?)([${MONEY}])\\s?([+-]?)(${GROUPED}|${PLAIN})$`).exec(inner);
  if (!m || (m[1] && m[3])) return null;
  const n = numberOf(m[4]);
  if (n === null) return null;
  const negative = Boolean(bracketed) !== ((m[1] || m[3]) === '-');
  return { value: negative ? -n : n, format: `"${m[2]}"#,##0${decimalsOf(m[4]) ? '.00' : ''}`, kind: 'money' };
}

function readGrouped(t) {
  const m = new RegExp(`^([+-]?)(${GROUPED})$`).exec(t);
  if (!m) return null;
  const n = numberOf(m[2]);
  if (n === null) return null;
  return { value: m[1] === '-' ? -n : n, format: decimalsOf(m[2]) ? '#,##0.00' : '#,##0', kind: 'number' };
}

const TIME = /^(\d{1,4}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?(?:\s*(am|pm|a|p))?$/i;

function timeFormat(m, value) {
  const seconds = m[3] !== undefined;
  if (m[4]) return seconds ? 'h:mm:ss AM/PM' : 'h:mm AM/PM';
  if (value >= 1) return seconds ? '[h]:mm:ss' : '[h]:mm';
  return seconds ? 'h:mm:ss' : 'h:mm';
}

function readTime(t) {
  const m = TIME.exec(t);
  if (!m) return null;
  const value = parseTimeText(t);
  if (value === null) return null;
  return { value, format: timeFormat(m, value), kind: 'time' };
}

function readDate(t, { date1904 }) {
  // A time after the date comes with it: 9/11/2026 14:30.
  const at = t.search(/\s\d{1,2}:\d/);
  const date = (at < 0 ? t : t.slice(0, at)).trim();
  const time = at < 0 ? '' : t.slice(at).trim();
  const shortDate = readsDayFirst() ? 'dd/mm/yyyy' : 'mm/dd/yyyy';
  let format = null;
  let guessYear = false;
  let m;
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(date) || /^\d{1,2}[-/]\d{1,2}[-/](\d{4}|\d{2})$/.test(date)) format = shortDate;
  else if ((m = /^\d{1,2}[- ]([A-Za-z]{3,9})[-, ]+(?:\d{4}|\d{2})$/.exec(date))) format = isMonthName(m[1]) ? 'd mmm yyyy' : null;
  else if ((m = /^([A-Za-z]{3,9})[- ]\d{1,2}[-, ]+(?:\d{4}|\d{2})$/.exec(date))) format = isMonthName(m[1]) ? 'd mmm yyyy' : null;
  else if ((m = /^\d{1,2}[- ]([A-Za-z]{3,9})$/.exec(date))) { format = isMonthName(m[1]) ? 'd mmm' : null; guessYear = true; }
  else if ((m = /^([A-Za-z]{3,9})[-, ]+\d{4}$/.exec(date))) { format = isMonthName(m[1]) ? 'mmm yyyy' : null; guessYear = true; }
  if (!format) return null;
  const value = withDate1904(date1904, () => parseDateText(time ? `${date} ${time}` : date, { guessYear }));
  if (value === null || value < 0) return null;
  return { value, format: time ? `${format} hh:mm${/:\d{1,2}:\d/.test(time) ? ':ss' : ''}` : format, kind: 'date' };
}

/**
 * A typed entry read as Excel reads it. `date1904` is the workbook's date
 * system. Null for anything that is not a date, a time, a percentage, money
 * or a number with commas.
 */
export function readTyped(text, { date1904 = false } = {}) {
  if (typeof text !== 'string') return null;
  const t = text.trim();
  if (t === '' || t.startsWith('=') || t.startsWith("'") || !/\d/.test(t)) return null;
  return readPercent(t) ?? readMoney(t) ?? readGrouped(t) ?? readTime(t) ?? readDate(t, { date1904 });
}

/**
 * What the formula bar and the cell editor show for a number: a date as a
 * short date, a time as a time and a percentage with its sign, as Excel
 * does, so the entry can be amended in the form it was typed in. Null for a
 * number that is edited as the plain number it is.
 */
export function typedText(value, code, { date1904 = false } = {}) {
  if (typeof value !== 'number' || !Number.isFinite(value) || !code || code === 'General' || code === '@') return null;
  const body = splitSections(code)[0]?.body ?? code;
  if (isDateFormat(body)) {
    if (value < 0) return null;
    const bare = body.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, (b) => (/^\[(h+|m+|s+)\]$/i.test(b) ? b : ''));
    const days = Math.floor(value + 1e-9);
    const part = Math.max(0, value - days);
    const two = (n) => String(n).padStart(2, '0');
    const clock = () => {
      const total = Math.round(part * 86400);
      return `${Math.floor(total / 3600)}:${two(Math.floor(total / 60) % 60)}:${two(total % 60)}`;
    };
    const hasDate = /[yd]/i.test(bare) || /m{3,}/i.test(bare) || (/m/i.test(bare) && !/[hs]/i.test(bare));
    if (!hasDate) {
      // A time alone, or hours counted past a day.
      if (/\[h+\]/i.test(bare) || value >= 1) {
        const total = Math.round(value * 86400);
        return `${Math.floor(total / 3600)}:${two(Math.floor(total / 60) % 60)}:${two(total % 60)}`;
      }
      return clock();
    }
    const parts = withDate1904(date1904, () => serialToParts(days));
    const d = two(parts.d);
    const mo = two(parts.m + 1);
    const y = parts.y;
    const dateText = readsDayFirst() ? `${d}/${mo}/${y}` : `${mo}/${d}/${y}`;
    return part * 86400 >= 0.5 ? `${dateText} ${clock()}` : dateText;
  }
  if (/%/.test(body.replace(/"[^"]*"/g, '').replace(/\\./g, ''))) return formatNumber(Number((value * 100).toPrecision(15))) + '%';
  return null;
}
