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
 */
import { serialToDate, formatNumber } from '@rutba/formula';

/** ECMA-376 §18.8.30 built-in formats. */
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
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
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

/** Pull `[Red]`, `[$-409]`, `[h]` etc out of a section. */
function extractDirectives(section) {
  let colour = null;
  let elapsed = false;
  let condition = null;
  const body = section.replace(/\[([^\]]*)\]/g, (_, inner) => {
    const lower = inner.toLowerCase();
    if (COLOURS.includes(lower)) { colour = lower; return ''; }
    if (/^color\s*\d+$/i.test(inner)) { colour = 'color' + inner.replace(/\D/g, ''); return ''; }
    if (/^(h+|m+|s+)$/i.test(inner)) { elapsed = true; return inner; } // [h] elapsed time
    const cond = /^([<>=]{1,2})(-?[\d.]+)$/.exec(inner);
    if (cond) { condition = { op: cond[1], value: Number(cond[2]) }; return ''; }
    if (inner.startsWith('$')) return ''; // locale/currency hint, e.g. [$-409]
    return '';
  });
  return { body, colour, elapsed, condition };
}

const DATE_TOKEN = /(am\/pm|a\/p|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|\.0+)/i;

/** Does this section format a date/time rather than a number? */
export function isDateFormat(body) {
  // strip quoted literals and escapes before deciding
  const bare = body.replace(/"[^"]*"/g, '').replace(/\\./g, '');
  // Month tokens count: `mmmm` alone is a date format and carries no y/d/h/s.
  // The `#?` guard is what keeps a numeric pattern out of this branch.
  return /(y{2,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2}|am\/pm|a\/p)/i.test(bare) && !/[#?]/.test(bare);
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

/** Parse a numeric pattern like `#,##0.00` into its parts. */
function parseNumericPattern(pattern) {
  const [intPart, fracPart = ''] = pattern.split('.');
  const grouped = intPart.includes(',');
  const cleanInt = intPart.replace(/,/g, '');
  // trailing commas scale by thousands: `#,##0,` shows thousands
  const trailingCommas = (/,+$/.exec(intPart) ?? [''])[0].length;
  return {
    grouped,
    minInt: (cleanInt.match(/0/g) ?? []).length,
    minFrac: (fracPart.match(/0/g) ?? []).length,
    maxFrac: (fracPart.match(/[0#?]/g) ?? []).length,
    scale: trailingCommas,
  };
}

/** Renders the MAGNITUDE. The sign is the caller's decision — see formatValue. */
function renderNumber(value, pattern, { percent }) {
  const p = parseNumericPattern(pattern);
  let n = value;
  if (percent) n *= 100;
  if (p.scale) n /= 1000 ** p.scale;

  const fixed = Math.abs(n).toFixed(p.maxFrac);
  let [ints, frac = ''] = fixed.split('.');

  // trim optional trailing zeros down to the minimum the pattern demands
  while (frac.length > p.minFrac && frac.endsWith('0')) frac = frac.slice(0, -1);
  while (ints.length < p.minInt) ints = '0' + ints;
  if (p.minInt === 0 && ints === '0' && p.maxFrac > 0) ints = '';

  if (p.grouped) ints = ints.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return ints + (frac ? '.' + frac : '');
}

const pad = (n, width) => String(n).padStart(width, '0');

function renderDate(value, tokens, { elapsed }) {
  const date = serialToDate(value);
  const dayFraction = value - Math.floor(value);
  const totalSeconds = Math.round(dayFraction * 86400);
  const hours24 = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  const usesAmPm = tokens.some((t) => t.kind === 'date' && /^(am\/pm|a\/p)$/i.test(t.text));
  const hours12 = hours24 % 12 === 0 ? 12 : hours24 % 12;

  return tokens.map((t) => {
    if (t.kind !== 'date') return t.kind === 'literal' ? t.text : '';
    const raw = t.text;
    const lower = raw.toLowerCase();
    if (lower === 'am/pm') return hours24 < 12 ? 'AM' : 'PM';
    if (lower === 'a/p') return hours24 < 12 ? 'A' : 'P';
    if (/^y{3,4}$/.test(lower)) return String(date.getUTCFullYear());
    if (/^y{1,2}$/.test(lower)) return pad(date.getUTCFullYear() % 100, 2);
    if (/^d{4}$/.test(lower)) return DAYS[date.getUTCDay()];
    if (/^d{3}$/.test(lower)) return DAYS[date.getUTCDay()].slice(0, 3);
    if (/^d{2}$/.test(lower)) return pad(date.getUTCDate(), 2);
    if (/^d$/.test(lower)) return String(date.getUTCDate());
    if (/^h+$/.test(lower)) {
      if (elapsed) return String(Math.floor(value * 24));
      const h = usesAmPm ? hours12 : hours24;
      return raw.length > 1 ? pad(h, 2) : String(h);
    }
    if (/^s+$/.test(lower)) return raw.length > 1 ? pad(seconds, 2) : String(seconds);
    if (/^\.0+$/.test(lower)) return '.' + pad(0, raw.length - 1);
    if (/^m+$/.test(lower)) {
      if (t.meaning === 'minute') return raw.length > 1 ? pad(minutes, 2) : String(minutes);
      if (raw.length === 5) return MONTHS[date.getUTCMonth()][0];
      if (raw.length === 4) return MONTHS[date.getUTCMonth()];
      if (raw.length === 3) return MONTHS[date.getUTCMonth()].slice(0, 3);
      if (raw.length === 2) return pad(date.getUTCMonth() + 1, 2);
      return String(date.getUTCMonth() + 1);
    }
    return raw;
  }).join('');
}

/**
 * Format a value for display.
 *
 * @param {*} value        number, string, boolean, or a FormulaError-like object
 * @param {string} code    a format code, or 'General'
 * @returns {{text: string, colour: string|null, align: 'left'|'right'}}
 */
export function formatValue(value, code = 'General') {
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
    const textSection = sections.length >= 4 ? sections[3] : null;
    if (!textSection) return { text: String(value), colour: null, align: 'left' };
    const { body, colour } = extractDirectives(textSection);
    const rendered = tokenize(body)
      .map((t) => (t.kind === 'text' ? String(value) : t.kind === 'literal' ? t.text : ''))
      .join('');
    return { text: rendered, colour, align: 'left' };
  }

  // Which section applies: positive / negative / zero.
  let section;
  if (sections.length === 1) section = sections[0];
  else if (sections.length === 2) section = value < 0 ? sections[1] : sections[0];
  else section = value > 0 ? sections[0] : value < 0 ? sections[1] : sections[2];

  const { body, colour, elapsed, condition } = extractDirectives(section);
  if (condition && !testCondition(value, condition)) {
    // a conditional section that does not apply falls back to General
    return { text: formatNumber(value), colour, align: 'right' };
  }

  const trimmed = body.trim();
  if (trimmed === '' ) return { text: '', colour, align: 'right' };
  if (/^general$/i.test(trimmed)) return { text: formatNumber(value), colour, align: 'right' };

  if (isDateFormat(body)) {
    const tokens = classifyMonthMinute(tokenize(body));
    return { text: renderDate(value, tokens, { elapsed }), colour, align: 'right' };
  }

  const tokens = tokenize(body);
  const percent = tokens.some((t) => t.kind === 'percent');
  // The negative section formats the absolute value; its sign comes from the
  // literal text in the format, so -500 with `(#,##0)` is "(500)" not "(-500)".
  const formatSuppliesSign = sections.length > 1;
  const magnitude = Math.abs(value);
  const sign = !formatSuppliesSign && value < 0 ? '-' : '';
  let numberRendered = false;

  const text = tokens.map((t) => {
    if (t.kind === 'number') {
      numberRendered = true;
      return renderNumber(magnitude, t.text, { percent });
    }
    if (t.kind === 'percent') return '%';
    if (t.kind === 'literal') return t.text;
    if (t.kind === 'text') return formatNumber(magnitude);
    return '';
  }).join('');

  // A section with no numeric placeholder renders only its literals — that is
  // how `#,##0;(#,##0);"-"` shows a dash for zero rather than "-0".
  void numberRendered;
  return { text: sign + text, colour, align: 'right' };
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
