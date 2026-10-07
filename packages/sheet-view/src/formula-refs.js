// The references in a formula as it is being typed, and where a click may
// put another — what Excel colours while a formula is entered.
//
// A formula half typed is not a formula: "=SUM(A1:B" has an unclosed
// bracket and a range with no end. The parser refuses it, rightly; this
// does not parse, it reads — each cell, range, whole column or row, with
// its sheet if it names one, where it starts and ends in the text, and the
// colour Excel would give it (one per distinct reference, in the order they
// first appear). Text inside quotes is text.
//
// Pure: a string in, positions out.

/** The colours references take, in turn — Excel's order, legible on white and on the dark theme. */
export const REF_COLOURS = ['#2f6fd6', '#d2433a', '#8a3fc7', '#1e8a46', '#e0750b', '#00838f', '#c2185b', '#7a6b00'];

const MAX_ROW = 1048575;
const MAX_COL = 16383;

// A sheet prefix: 'Quoted name'! or Plain_Name!
const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!`;
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})`;
const RANGE_RE = new RegExp(String.raw`^(?:${SHEET})?${CELL}(?::${CELL})?(?![\w(.])`);
const COLS_RE = new RegExp(String.raw`^(?:${SHEET})?(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(.])`);
const ROWS_RE = new RegExp(String.raw`^(?:${SHEET})?(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\w(.])`);

/** Column letters as a 0-based index. */
export function colIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** A 0-based column index as letters. */
export function colLetters(index) {
  let s = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** A range, top-left first, as Excel writes it: "B3", or "B3:D7". */
export function rangeText({ top, left, bottom, right }) {
  const a = colLetters(left) + (top + 1);
  return top === bottom && left === right ? a : `${a}:${colLetters(right) + (bottom + 1)}`;
}

/** Two corners — where a drag began and where it is — as a range. */
export function rangeOf(a, b = a) {
  return { top: Math.min(a.row, b.row), left: Math.min(a.col, b.col), bottom: Math.max(a.row, b.row), right: Math.max(a.col, b.col) };
}

const sheetOf = (m) => (m[1] != null ? m[1].replace(/''/g, "'") : m[2] ?? null);
const valid = (r) => r.top >= 0 && r.left >= 0 && r.bottom <= MAX_ROW && r.right <= MAX_COL && r.top <= r.bottom && r.left <= r.right;

/**
 * Every reference in a formula: [{ start, end, text, sheet, range, colour }],
 * `start` and `end` in the formula's own text (the "=" included), `range`
 * 0-based and inclusive, `colour` an index into REF_COLOURS. Nothing for
 * text that is not a formula.
 */
export function formulaReferences(formula) {
  const text = String(formula ?? '');
  if (!text.startsWith('=')) return [];
  const out = [];
  const colours = new Map();
  let i = 1;
  while (i < text.length) {
    const ch = text[i];
    if (ch === '"') { // a string: skipped whole, closed or not
      const close = text.indexOf('"', i + 1);
      i = close < 0 ? text.length : close + 1;
      while (text[i] === '"') { const next = text.indexOf('"', i + 1); i = next < 0 ? text.length : next + 1; }
      continue;
    }
    // A reference starts where a name could not be going on.
    const before = text[i - 1];
    if (/[A-Za-z0-9_.$]/.test(before) && i > 1) { i += 1; continue; }
    const rest = text.slice(i);
    let found = null;
    let m = RANGE_RE.exec(rest);
    if (m) {
      const a = { row: Number(m[6]) - 1, col: colIndex(m[4]) };
      const b = m[8] != null ? { row: Number(m[10]) - 1, col: colIndex(m[8]) } : a;
      found = { length: m[0].length, sheet: sheetOf(m), range: rangeOf(a, b) };
    } else if ((m = COLS_RE.exec(rest))) {
      const l = colIndex(m[4]);
      const r = colIndex(m[6]);
      found = { length: m[0].length, sheet: sheetOf(m), range: { top: 0, bottom: MAX_ROW, left: Math.min(l, r), right: Math.max(l, r) } };
    } else if ((m = ROWS_RE.exec(rest))) {
      const t = Number(m[4]) - 1;
      const b = Number(m[6]) - 1;
      found = { length: m[0].length, sheet: sheetOf(m), range: { top: Math.min(t, b), bottom: Math.max(t, b), left: 0, right: MAX_COL } };
    }
    if (found && valid(found.range)) {
      const refText = text.slice(i, i + found.length);
      const key = refText.replace(/\$/g, '').toUpperCase();
      if (!colours.has(key)) colours.set(key, colours.size % REF_COLOURS.length);
      out.push({ start: i, end: i + found.length, text: refText, sheet: found.sheet, range: found.range, colour: colours.get(key) });
      i += found.length;
      continue;
    }
    // A name or a function: past it whole, so "LOG10" is not read as LOG10.
    const word = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest);
    i += word ? word[0].length : 1;
  }
  return out;
}

/** Whether `caret` sits inside a string in `formula`. */
function inString(formula, caret) {
  let open = false;
  for (let i = 0; i < caret; i++) if (formula[i] === '"') open = !open;
  return open;
}

/**
 * Where a click on a cell would put its reference, with the caret at
 * `caret`: { start, end } — the span to replace (empty to insert) — or null
 * when a click there means leaving the cell, as it does in Excel. A
 * reference goes in after "=", an opening bracket, a comma or an operator;
 * and a reference the caret has just been put after (one pointed at, or
 * typed) is replaced by the next one.
 */
export function pointSpan(formula, caret) {
  const text = String(formula ?? '');
  if (!text.startsWith('=') || caret == null || caret < 1 || caret > text.length) return null;
  if (inString(text, caret)) return null;
  const after = formulaReferences(text).find((r) => r.end === caret);
  if (after) return { start: after.start, end: after.end };
  let k = caret - 1;
  while (k > 0 && text[k] === ' ') k -= 1;
  // The rest of the formula must not run straight on into a word or a number.
  if (/^[A-Za-z0-9_$"]/.test(text.slice(caret))) return null;
  return /[=(,;+\-*/^&<>:%]/.test(text[k]) ? { start: caret, end: caret } : null;
}

/**
 * F4: a reference made absolute, then row-only, then column-only, then
 * relative again — A1, $A$1, A$1, $A1 — for each cell in it.
 */
export function cycleAbsolute(refText) {
  const m = /^((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(.*)$/.exec(refText);
  const sheet = m[1] || '';
  const step = (cell) => {
    const c = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(cell);
    if (!c) return cell;
    const [, dc, col, dr, row] = c;
    // relative → both → row → column → relative
    const next = !dc && !dr ? ['$', '$'] : dc && dr ? ['', '$'] : !dc && dr ? ['$', ''] : ['', ''];
    return next[0] + col + next[1] + row;
  };
  return sheet + m[2].split(':').map(step).join(':');
}
