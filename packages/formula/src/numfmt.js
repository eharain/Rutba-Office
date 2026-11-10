/**
 * Number formats — turning a stored value into the string a user reads.
 *
 * This is the difference between a bound ERP sheet that looks like a finance
 * report and one that looks like a debug dump. The value `1661.4000000000001`
 * is correct and unreadable; `1,661.40 PKR` is the same number and is the
 * product.
 *
 * Format codes are specified in ECMA-376 §18.8.30, which also fixes the built-in
 * numFmtId table reproduced below — ids 0-49 are reserved and every consumer
 * agrees on them, which is why a file can say `numFmtId="4"` and mean
 * `#,##0.00` without spelling it out.
 *
 * The subtleties that make this worth writing carefully:
 *
 *   - A format has up to FOUR sections: positive;negative;zero;text. With two,
 *     the second covers negatives AND the first covers zero. The negative
 *     section formats the ABSOLUTE value — the minus sign, or the brackets, come
 *     from the format itself. Getting this wrong renders -500 as "-(500)".
 *   - `m` means month or minute depending on neighbours. `mm:ss` is minutes;
 *     `mm-dd` is months. The rule is: a token adjacent to an hour or second
 *     token is minutes. Nothing else disambiguates it.
 *   - `_x` inserts a space the width of x, and `*x` repeats x to fill. Both are
 *     alignment instructions we consume and skip rather than render literally.
 *   - `[Red]` and friends are colour, not text. They come back as a hint so the
 *     view can apply it, rather than being printed.
 *   - `[$-hhhhhhhh]` carries a locale, and above it a calendar and a set of
 *     digits: `[$-2000000]` writes a number in Arabic-Indic digits (١٢٣),
 *     `[$-3000000]` in Urdu's and Persian's (۱۲۳), `[$-60000]` a date in the
 *     Hijri calendar. A section that starts `B2` is Hijri too, `B1` Gregorian,
 *     as Excel writes them.
 */
import { serialToParts, formatNumber, withDate1904 } from './values.js';

/**
 * The short date this computer writes, as a format code: `dd/mm/yyyy` on a
 * British machine, `mm/dd/yyyy` on an American one, `yyyy-mm-dd` on a
 * Swedish one. Excel's built-in format 14 is documented as `mm-dd-yy` but
 * shows the system's short date, and so does every file whose cells wear
 * it — a British workbook's dates would read as American here otherwise.
 */
export function localeShortDate(locale = undefined) {
  try {
    const parts = new Intl.DateTimeFormat(locale, { year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date(Date.UTC(2001, 10, 22, 12)));
    const code = parts
      .map((p) => (p.type === 'day' ? 'dd' : p.type === 'month' ? 'mm' : p.type === 'year' ? 'yyyy' : p.type === 'literal' ? p.value.replace(/[^-/. ]/g, '') : ''))
      .join('');
    return /dd|mm|yyyy/.test(code) ? code.trim() : 'dd/mm/yyyy';
  } catch {
    return 'dd/mm/yyyy';
  }
}

/** ECMA-376 §18.8.30 built-in formats; 14 and 22 follow the system's short date. */
export const BUILTIN_FORMATS = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: localeShortDate(),
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: `${localeShortDate()} h:mm`,
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Split on `;` while respecting quoted text and bracketed directives. */
export function splitSections(code) {
  const sections = [];
  let current = '';
  let inQuote = false;
  let inBracket = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '\\') { current += ch + (code[i + 1] ?? ''); i += 1; continue; }
    if (ch === '"') { inQuote = !inQuote; current += ch; continue; }
    if (!inQuote && ch === '[') { inBracket = true; current += ch; continue; }
    if (!inQuote && ch === ']') { inBracket = false; current += ch; continue; }
    if (ch === ';' && !inQuote && !inBracket) { sections.push(current); current = ''; continue; }
    current += ch;
  }
  sections.push(current);
  return sections;
}

const COLOURS = ['black', 'blue', 'cyan', 'green', 'magenta', 'red', 'white', 'yellow'];

/** The digits a `[$-…]` code's numeral-system byte names: their zero, from which 0 to 9 follow. */
const DIGIT_ZERO = { 2: 0x660, 3: 0x6f0 };

/** A text's 0 to 9 in another set of digits; Western digits are left as they are. */
function inDigits(text, system) {
  const zero = DIGIT_ZERO[system];
  return zero ? String(text).replace(/[0-9]/g, (d) => String.fromCharCode(zero + Number(d))) : text;
}

/** Pull `[Red]`, `[$-409]`, `[h]` etc out of a section. */
function extractDirectives(section) {
  let colour = null;
  let elapsed = null;
  let condition = null;
  let digits = 0;
  let calendar = 0;
  let lcid = 0;
  const raw = section.replace(/"[^"]*"|\\.|\[([^\]]*)\]/g, (whole, inner) => {
    if (inner === undefined) return whole;
    const lower = inner.toLowerCase();
    if (COLOURS.includes(lower)) { colour = lower; return ''; }
    if (/^color\s*\d+$/i.test(inner)) { colour = 'color' + inner.replace(/\D/g, ''); return ''; }
    if (/^(h+|m+|s+)$/i.test(inner)) { elapsed = lower[0]; return inner; } // [h] elapsed time
    const cond = /^([<>=]{1,2})(-?[\d.]+)$/.exec(inner);
    if (cond) { condition = { op: cond[1], value: Number(cond[2]) }; return ''; }
    if (inner.startsWith('$')) {
      // A locale, a currency or both: [$-409], [$€-407], [$-2060401]. Above the
      // locale's sixteen bits, a calendar (a byte) and a set of digits (a byte).
      const code = /-([0-9a-f]{1,8})$/i.exec(inner);
      if (code) {
        const n = parseInt(code[1], 16);
        lcid = n & 0xffff;
        calendar = (n >>> 16) & 0xff;
        digits = (n >>> 24) & 0xff;
      }
      // What stands before the hyphen is a currency symbol to print: [$£-809].
      const symbol = inner.slice(1).split('-')[0];
      return symbol ? '"' + symbol.replace(/"/g, '') + '"' : '';
    }
    return '';
  });
  // Excel's calendar marks at the start of a section: B1 Gregorian, B2 Hijri.
  const mark = /^(\s*)B([12])/i.exec(raw);
  if (mark) calendar = mark[2] === '2' ? 6 : 1;
  const body = mark ? raw.slice(mark[0].length) : raw;
  return { body, colour, elapsed, condition, digits, hijri: calendar === 6 || calendar === 23, lcid };
}

/** Hijri (Umm al-Qura) dates and month names, made once per language. */
const hijriFormats = new Map();
function hijriFormat(lang) {
  if (!hijriFormats.has(lang)) {
    let parts = null;
    let months = null;
    try {
      parts = new Intl.DateTimeFormat(`${lang}-u-ca-islamic-umalqura-nu-latn`, { timeZone: 'UTC', year: 'numeric', month: 'numeric', day: 'numeric' });
      months = new Intl.DateTimeFormat(`${lang}-u-ca-islamic-umalqura-nu-latn`, { timeZone: 'UTC', month: 'long' });
    } catch { /* no Hijri calendar here: dates stay Gregorian */ }
    hijriFormats.set(lang, { parts, months });
  }
  return hijriFormats.get(lang);
}

/**
 * A Gregorian day in the Hijri calendar: { y, m (from 0), d, month }, the
 * month's name in Arabic for an Arabic locale, else as English writes it.
 */
function toHijri(y, m, d, lcid) {
  const lang = (lcid & 0x3ff) === 0x01 ? 'ar' : 'en';
  const { parts, months } = hijriFormat(lang);
  if (!parts) return null;
  const at = new Date(Date.UTC(y, m, d, 12));
  const got = Object.fromEntries(parts.formatToParts(at).map((p) => [p.type, p.value]));
  const year = parseInt(String(got.year ?? got.relatedYear).replace(/\D/g, ''), 10);
  const month = parseInt(got.month, 10);
  const day = parseInt(got.day, 10);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
  return { y: year, m: month - 1, d: day, month: months.format(at) };
}

const DATE_TOKEN = /(am\/pm|a\/p|yyyy|yyy|yy|y|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|\.0+)/i;

/** Does this section format a date/time rather than a number? */
export function isDateFormat(body) {
  // strip quoted literals and escapes before deciding
  const bare = body.replace(/"[^"]*"/g, '').replace(/\\./g, '');
  // Month tokens count: `mmmm` alone is a date format and carries no y/d/h/s.
  // The `#?` guard is what keeps a numeric pattern out of this branch.
  return /(y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2}|am\/pm|a\/p)/i.test(bare) && !/[#?]/.test(bare);
}

/**
 * Decide, for each `m` run, whether it means month or minute.
 * A run adjacent (ignoring separators) to an hour or second token is minutes.
 */
function classifyMonthMinute(tokens) {
  const out = tokens.map((t) => ({ ...t }));
  for (let i = 0; i < out.length; i++) {
    const t = out[i];
    if (t.kind !== 'date' || !/^m+$/i.test(t.text)) continue;
    let minutes = false;
    for (let j = i - 1; j >= 0; j--) {
      const p = out[j];
      if (p.kind === 'literal' && /^[\s:.\-/]*$/.test(p.text)) continue;
      if (p.kind === 'date' && /^h+$/i.test(p.text)) minutes = true;
      break;
    }
    for (let j = i + 1; j < out.length && !minutes; j++) {
      const n = out[j];
      if (n.kind === 'literal' && /^[\s:.\-/]*$/.test(n.text)) continue;
      if (n.kind === 'date' && /^s+$/i.test(n.text)) minutes = true;
      break;
    }
    t.meaning = minutes ? 'minute' : 'month';
  }
  return out;
}

/** Tokenize one section body. */
export function tokenize(body) {
  const tokens = [];
  const dates = isDateFormat(body);
  let i = 0;
  while (i < body.length) {
    const ch = body[i];
    if (ch === '\\') { tokens.push({ kind: 'literal', text: body[i + 1] ?? '' }); i += 2; continue; }
    if (ch === '"') {
      const end = body.indexOf('"', i + 1);
      tokens.push({ kind: 'literal', text: body.slice(i + 1, end < 0 ? body.length : end) });
      i = end < 0 ? body.length : end + 1;
      continue;
    }
    if (ch === '_') { tokens.push({ kind: 'literal', text: ' ' }); i += 2; continue; } // width of next char
    if (ch === '*') { i += 2; continue; } // fill — an alignment hint, not content
    if (ch === '@') { tokens.push({ kind: 'text' }); i += 1; continue; }
    if (ch === '%') { tokens.push({ kind: 'percent' }); i += 1; continue; }
    if (!dates && /^general/i.test(body.slice(i, i + 7))) { tokens.push({ kind: 'general' }); i += 7; continue; }
    if (!dates && (('#0?'.includes(ch)) || (ch === '.' && /[#0?]/.test(body[i + 1] ?? '')))) {
      let run = '';
      while (i < body.length && '#0?,.'.includes(body[i])) { run += body[i]; i += 1; }
      tokens.push({ kind: 'number', text: run });
      continue;
    }

    const dateMatch = DATE_TOKEN.exec(body.slice(i));
    if (dateMatch && dateMatch.index === 0) {
      tokens.push({ kind: 'date', text: dateMatch[0] });
      i += dateMatch[0].length;
      continue;
    }
    if ('#0?'.includes(ch)) {
      let run = '';
      while (i < body.length && '#0?,.'.includes(body[i])) { run += body[i]; i += 1; }
      tokens.push({ kind: 'number', text: run });
      continue;
    }
    tokens.push({ kind: 'literal', text: ch });
    i += 1;
  }
  return tokens;
}

/**
 * A number as digits, rounded in decimal: fifteen significant digits and then
 * half up, so 2.675 at two places is 2.68 and 123456789012345678 ends in
 * zeros, as Excel shows them. Answers [whole part, fraction of `places` digits].
 */
function decimalDigits(x, places) {
  if (x === 0) return ['0', '0'.repeat(places)];
  const [mantissa, e] = x.toExponential(14).split('e');
  const digits = mantissa.replace('.', '');
  const point = Number(e) + 1;
  let ints;
  let frac;
  if (point <= 0) { ints = '0'; frac = '0'.repeat(-point) + digits; }
  else if (point >= digits.length) { ints = digits + '0'.repeat(point - digits.length); frac = ''; }
  else { ints = digits.slice(0, point); frac = digits.slice(point); }
  if (frac.length > places) {
    const up = frac[places] >= '5';
    frac = frac.slice(0, places);
    if (up) {
      const width = ints.length + frac.length;
      const raised = (BigInt(ints + frac) + 1n).toString().padStart(width, '0');
      const keep = raised.length - frac.length;
      ints = raised.slice(0, keep);
      frac = raised.slice(keep);
    }
  } else frac = frac.padEnd(places, '0');
  return [ints.replace(/^0+(?=\d)/, ''), frac];
}

/**
 * Number placeholders, however literals split them: `#,##0.00` is one piece
 * and `000-00-0000` three, which between them hold one number. Each slot is
 * `0` (a digit, zero if none), `#` (nothing if none) or `?` (a space if none).
 * Answers the text for each piece, in order.
 */
function renderPieces(pieces, value, { percent }) {
  const joined = pieces.join('');
  const dot = joined.indexOf('.');
  let intCode = dot < 0 ? joined : joined.slice(0, dot);
  let fracCode = dot < 0 ? '' : joined.slice(dot + 1);
  // Commas straight after the last digit place divide by a thousand each.
  const trailing = (code) => (/,+$/.exec(code) ?? [''])[0].length;
  let scale = trailing(dot < 0 ? intCode : fracCode);
  if (dot >= 0 && !fracCode.replace(/,/g, '')) scale += trailing(intCode);
  else if (dot >= 0) scale += trailing(intCode);
  const grouped = /[#0?],+[#0?]/.test(intCode);

  let n = value;
  if (percent) n *= 100;
  if (scale) n /= 1000 ** scale;

  const intSlots = intCode.replace(/,/g, '').split('');
  const fracSlots = fracCode.replace(/,/g, '').split('');
  const [ints, frac] = decimalDigits(Math.abs(n), fracSlots.length);
  const allZero = /^0*$/.test(ints + frac);

  // Whole part, from the right; digits with no slot left go before the first.
  const number = ints === '0' ? '' : ints;
  const out = new Array(intSlots.length).fill('');
  let d = number.length - 1;
  for (let i = intSlots.length - 1; i >= 0; i--) {
    if (d >= 0) out[i] = number[d--];
    else out[i] = intSlots[i] === '0' ? '0' : intSlots[i] === '?' ? ' ' : '';
  }
  const overflow = d >= 0 ? number.slice(0, d + 1).split('') : [];
  const line = [...overflow, ...out];
  if (grouped) {
    let seen = 0;
    for (let i = line.length - 1; i >= 0; i--) {
      if (!/\d/.test(line[i])) continue;
      if (seen > 0 && seen % 3 === 0) line[i] += ',';
      seen += 1;
    }
  }
  const whole = line.slice(overflow.length);
  if (overflow.length) {
    if (whole.length) whole[0] = line.slice(0, overflow.length).join('') + whole[0];
    else whole.push(line.join(''));
  }

  // Fraction, from the left; a place past the last digit that matters is left out.
  let last = frac.length;
  while (last > 0 && frac[last - 1] === '0' && fracSlots[last - 1] !== '0') last -= 1;
  const part = fracSlots.map((slot, i) => (i < last ? frac[i] : slot === '0' ? '0' : slot === '?' ? ' ' : ''));

  let wi = 0;
  let fi = 0;
  let afterDot = false;
  const rendered = pieces.map((piece) => {
    let text = '';
    for (const c of piece) {
      if (c === ',') continue;
      if (c === '.') { if (!afterDot) { text += '.'; afterDot = true; } continue; }
      if (afterDot) text += part[fi++] ?? '';
      else text += whole[wi++] ?? '';
    }
    return text;
  });
  return { rendered, allZero };
}

/** Renders the MAGNITUDE. The sign is the caller's decision: see formatValue. */
function renderNumber(value, pattern, opts) {
  return renderPieces([pattern], value, opts).rendered[0];
}

const pad = (n, width) => String(n).padStart(width, '0');

function renderDate(value, tokens, { elapsed, hijri = false, lcid = 0 }) {
  // The parts, not a Date: serial 60 is 29 February 1900 here, as in Excel.
  const places = Math.max(0, ...tokens.filter((t) => t.kind === 'date' && /^\.0+$/.test(t.text)).map((t) => t.text.length - 1));
  // Seconds are rounded to what is shown; a time that rounds to midnight is the next day.
  const unit = 10 ** places;
  let wholeDays = Math.floor(value);
  let clock = Math.round((value - wholeDays) * 86400 * unit) / unit;
  if (clock >= 86400) { clock -= 86400; wholeDays += 1; }
  const gregorian = serialToParts(wholeDays);
  // In the Hijri calendar the year, month and day are Hijri's; the weekday is the same day's.
  const h = hijri ? toHijri(gregorian.y, gregorian.m, gregorian.d, lcid) : null;
  const date = h ? { ...gregorian, y: h.y, m: h.m, d: h.d } : gregorian;
  const monthName = (m) => (h ? h.month : MONTHS[m]);
  const totalSeconds = Math.floor(clock);
  const hours24 = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  const usesAmPm = tokens.some((t) => t.kind === 'date' && /^(am\/pm|a\/p)$/i.test(t.text));
  const hours12 = hours24 % 12 === 0 ? 12 : hours24 % 12;
  let elapsedShown = false;

  return tokens.map((t) => {
    if (t.kind !== 'date') return t.kind === 'literal' ? t.text : '';
    const raw = t.text;
    const lower = raw.toLowerCase();
    if (lower === 'am/pm') return hours24 < 12 ? 'AM' : 'PM';
    if (lower === 'a/p') return hours24 < 12 ? 'A' : 'P';
    if (/^y{3,4}$/.test(lower)) return String(date.y);
    if (/^y{1,2}$/.test(lower)) return pad(date.y % 100, 2);
    if (/^d{4}$/.test(lower)) return DAYS[date.weekday];
    if (/^d{3}$/.test(lower)) return DAYS[date.weekday].slice(0, 3);
    if (/^d{2}$/.test(lower)) return pad(date.d, 2);
    if (/^d$/.test(lower)) return String(date.d);
    if (elapsed && !elapsedShown && lower[0] === elapsed) {
      elapsedShown = true;
      const perDay = { h: 24, m: 1440, s: 86400 }[elapsed];
      const total = Math.floor((wholeDays * 86400 + totalSeconds) / (86400 / perDay));
      return raw.length > 1 ? pad(total, raw.length) : String(total);
    }
    if (/^h+$/.test(lower)) {
      const h = usesAmPm ? hours12 : hours24;
      return raw.length > 1 ? pad(h, 2) : String(h);
    }
    if (/^s+$/.test(lower)) return raw.length > 1 ? pad(seconds, 2) : String(seconds);
    if (/^\.0+$/.test(lower)) {
      const shown = Math.round((clock - totalSeconds) * 10 ** (raw.length - 1));
      return '.' + pad(shown, raw.length - 1);
    }
    if (/^m+$/.test(lower)) {
      if (t.meaning === 'minute') return raw.length > 1 ? pad(minutes, 2) : String(minutes);
      if (raw.length === 5) return monthName(date.m)[0];
      if (raw.length === 4) return monthName(date.m);
      if (raw.length === 3) return h ? monthName(date.m) : MONTHS[date.m].slice(0, 3);
      if (raw.length === 2) return pad(date.m + 1, 2);
      return String(date.m + 1);
    }
    return raw;
  }).join('');
}

/**
 * A scientific code: mantissa placeholders, `E`, a sign rule and exponent
 * placeholders — `0.00E+00`, `##0.0E+0`. `E+` always writes the sign, `E-`
 * only when it is negative; the exponent is padded to its placeholders. More
 * than one integer placeholder is Excel's engineering notation: the exponent
 * is a multiple of their count, so `##0.0E+0` shows 1234567 as 1.2E+6 and
 * 12345678 as 12.3E+6.
 */
const SCIENTIFIC = /^([#0,]*)(?:\.([0#]*))?E([+-])(0+)$/i;

function renderScientific(magnitude, body) {
  const [, intCode, fracCode = '', signRule, expCode] = SCIENTIFIC.exec(body.trim());
  const k = Math.max(1, (intCode.match(/[0#]/g) || []).length);
  let exponent = 0;
  let mantissa = 0;
  if (magnitude > 0) {
    const order = Math.floor(Math.log10(magnitude));
    exponent = k > 1 ? Math.floor(order / k) * k : order;
    mantissa = magnitude / Math.pow(10, exponent);
    // Rounding can carry the mantissa to the next power: 9.995 at two places is 10.00.
    const places = fracCode.length;
    if (Number(mantissa.toFixed(places)) >= Math.pow(10, k)) {
      exponent += k > 1 ? k : 1;
      mantissa = magnitude / Math.pow(10, exponent);
    }
  }
  const mantissaText = renderNumber(mantissa, `${intCode || '0'}${fracCode ? '.' + fracCode : ''}`, {});
  const sign = exponent < 0 ? '-' : signRule === '+' ? '+' : '';
  return `${mantissaText}E${sign}${String(Math.abs(exponent)).padStart(expCode.length, '0')}`;
}

/**
 * A fraction code: an optional whole part, then numerator placeholders, a
 * slash, and denominator placeholders or a fixed denominator — `# ?/?`,
 * `# ??/??`, `?/8`, `0 ?/16`. Excel's rules: the placeholders' count bounds
 * the denominator (one `?` allows up to 9, two up to 99); `?` pads with a
 * space so a column of fractions lines up; a whole part of `#` disappears
 * when it is zero and leaves its space behind, which is why a half in
 * `# ?/?` is " 1/2"; a fraction that rounds away leaves spaces its width.
 */
const FRACTION = /^([#0]*)(\s*)([?#0]+)\/([?#0]+|[1-9][0-9]*)(.*)$/;

function renderFraction(magnitude, body) {
  const m = FRACTION.exec(body.trim());
  const [, wholeCode, gap, numCode, denCode, tail] = m;
  const mixed = wholeCode.length > 0;
  let whole = mixed ? Math.floor(magnitude) : 0;
  const rest = mixed ? magnitude - whole : magnitude;
  let num;
  let den;
  if (/^[0-9]+$/.test(denCode)) {
    den = Number(denCode);
    num = Math.round(rest * den);
  } else {
    // The closest fraction with a denominator the placeholders allow.
    const maxDen = Math.pow(10, denCode.length) - 1;
    let best = { num: Math.round(rest), den: 1, err: Math.abs(rest - Math.round(rest)) };
    for (let d = 2; d <= maxDen; d++) {
      const n = Math.round(rest * d);
      const err = Math.abs(rest - n / d);
      if (err + 1e-12 < best.err) best = { num: n, den: d, err };
      if (err === 0) break;
    }
    num = best.num;
    den = best.den;
  }
  if (num >= den && den > 0 && mixed) {
    whole += Math.floor(num / den);
    num -= Math.floor(num / den) * den;
  }
  const padLeft = (s, code) => (code.includes('?') ? s.padStart(code.length, ' ') : s);
  const padRight = (s, code) => (code.includes('?') ? s.padEnd(code.length, ' ') : s);
  let out = '';
  if (mixed) out += whole === 0 && !wholeCode.includes('0') ? '' : String(whole);
  out += gap;
  if (num === 0 && mixed) {
    // Nothing to show past the whole number: spaces where the fraction would be.
    out += ' '.repeat(numCode.length + 1 + denCode.length);
  } else {
    out += padLeft(String(num), numCode) + '/' + padRight(String(den), denCode);
  }
  return out + tail;
}

/**
 * Format a value for display.
 *
 * @param {*} value        number, string, boolean, or a FormulaError-like object
 * @param {string} code    a format code, or 'General'
 * @returns {{text: string, colour: string|null, align: 'left'|'right'}}
 */
export function formatValue(value, code = 'General', { date1904 = false } = {}) {
  if (date1904) return withDate1904(true, () => formatValue(value, code));
  // Errors and booleans ignore the format entirely, as they do in a spreadsheet.
  if (value && typeof value === 'object' && typeof value.type === 'string' && value.type.startsWith('#')) {
    return { text: value.type, colour: null, align: 'center' };
  }
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', colour: null, align: 'center' };
  if (value === null || value === undefined || value === '') return { text: '', colour: null, align: 'left' };

  const sections = splitSections(String(code || 'General'));
  const isNumber = typeof value === 'number';

  if (!isNumber) {
    // Text uses the fourth section when present, otherwise renders as-is.
    // or a last section that says where the text goes.
    const lastOne = sections[sections.length - 1];
    const textSection = sections.length >= 4 ? sections[3] : tokenize(extractDirectives(lastOne).body).some((t) => t.kind === 'text') ? lastOne : null;
    if (!textSection) return { text: String(value), colour: null, align: 'left' };
    const { body, colour } = extractDirectives(textSection);
    const rendered = tokenize(body)
      .map((t) => (t.kind === 'text' ? String(value) : t.kind === 'literal' ? t.text : ''))
      .join('');
    return { text: rendered, colour, align: 'left' };
  }

  const { section, implied, matched } = chooseSection(sections, value);
  const { body, colour, elapsed, condition, digits, hijri, lcid } = extractDirectives(section);
  // Whatever the section writes, in the digits its code names.
  const out = (text) => ({ text: inDigits(text, digits), colour, align: 'right' });
  // A section with only a condition that does not hold shows the plain number.
  if (!matched) return out(formatNumber(value));
  // A negative in the negative section has its sign in the code: -500 in
  // `(#,##0)` is (500). So does a condition that only negatives can meet.
  const formatSuppliesSign = implied === 'negative' || Boolean(condition && (condition.op === '<' || condition.op === '<=') && condition.value <= 0);

  const trimmed = body.trim();
  if (trimmed === '' ) return { text: '', colour, align: 'right' };

  if (isDateFormat(body)) {
    const tokens = classifyMonthMinute(tokenize(body));
    return out(renderDate(value, tokens, { elapsed, hijri, lcid }));
  }

  if (SCIENTIFIC.test(body.trim())) {
    const sign = formatSuppliesSign || value >= 0 ? '' : '-';
    return out(sign + renderScientific(Math.abs(value), body));
  }

  if (FRACTION.test(body.trim())) {
    const sign = formatSuppliesSign || value >= 0 ? '' : '-';
    return out(sign + renderFraction(Math.abs(value), body));
  }

  const tokens = tokenize(body);
  const percent = tokens.some((t) => t.kind === 'percent');
  if (/^general$/i.test(trimmed) && value < 0 && formatSuppliesSign) return out(formatNumber(-value));
  if (/^general$/i.test(trimmed)) return out(formatNumber(value));
  const magnitude = Math.abs(value);
  const pieces = tokens.filter((t) => t.kind === 'number').map((t) => t.text);
  const { rendered, allZero } = pieces.length ? renderPieces(pieces, magnitude, { percent }) : { rendered: [], allZero: false };
  // A value that rounds to nothing has no sign: -0.001 in 0.00 is 0.00.
  const sign = !formatSuppliesSign && value < 0 && !allZero ? '-' : '';
  let piece = 0;
  const text = tokens.map((t) => {
    if (t.kind === 'number') return rendered[piece++];
    if (t.kind === 'percent') return '%';
    if (t.kind === 'literal') return t.text;
    if (t.kind === 'text' || t.kind === 'general') return formatNumber(magnitude);
    return '';
  }).join('');
  return out(sign + text);
}

/**
 * Which section of a code a number takes. Without conditions: positive and
 * zero; negative; or positive, negative, zero. With them, each in turn, and a
 * last section with none takes whatever is left.
 */
function chooseSection(sections, value) {
  const numeric = sections.slice(0, 3);
  const conditions = numeric.map((x) => extractDirectives(x).condition);
  if (!conditions.some(Boolean)) {
    if (numeric.length === 1) return { section: numeric[0], matched: true };
    if (value < 0) return { section: numeric[1], implied: 'negative', matched: true };
    if (numeric.length === 2 || value > 0) return { section: numeric[0], matched: true };
    return { section: numeric[2], matched: true };
  }
  for (let i = 0; i < numeric.length; i++) {
    if (!conditions[i]) {
      if (i === numeric.length - 1) return { section: numeric[i], matched: true };
      if (i === 1 && value < 0) return { section: numeric[i], implied: 'negative', matched: true };
      if (i === 0 && value > 0) return { section: numeric[i], matched: true };
      continue;
    }
    if (testCondition(value, conditions[i])) return { section: numeric[i], matched: true };
  }
  return { section: numeric[numeric.length - 1], matched: false };
}

function testCondition(value, { op, value: threshold }) {
  switch (op) {
    case '<': return value < threshold;
    case '<=': return value <= threshold;
    case '>': return value > threshold;
    case '>=': return value >= threshold;
    case '=': return value === threshold;
    case '<>': return value !== threshold;
    default: return true;
  }
}

/**
 * Read the style table out of a workbook package: cellXfs -> numFmt code.
 *
 * We only need the number format, so this deliberately ignores fonts, fills and
 * borders. Those stay preserved-but-unrendered until there is a reason to read
 * them, and the capability register says so.
 */
export function readNumberFormats(pkg) {
  if (!pkg.has('xl/styles.xml')) return { byStyleIndex: [], custom: {} };
  const xml = pkg.text('xl/styles.xml');

  const custom = {};
  for (const m of xml.matchAll(/<numFmt\b([^>]*)\/>/g)) {
    const id = /numFmtId="(\d+)"/.exec(m[1]);
    const code = /formatCode="([^"]*)"/.exec(m[1]);
    if (id && code) custom[Number(id[1])] = unescXml(code[1]);
  }

  const cellXfsBlock = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  const byStyleIndex = [];
  if (cellXfsBlock) {
    for (const m of cellXfsBlock[1].matchAll(/<xf\b([^>]*?)(?:\/>|>[\s\S]*?<\/xf>)/g)) {
      const id = /numFmtId="(\d+)"/.exec(m[1]);
      const numFmtId = id ? Number(id[1]) : 0;
      byStyleIndex.push(custom[numFmtId] ?? BUILTIN_FORMATS[numFmtId] ?? 'General');
    }
  }
  return { byStyleIndex, custom };
}

const unescXml = (s) =>
  String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
