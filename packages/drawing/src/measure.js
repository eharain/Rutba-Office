/**
 * Text metrics without a DOM.
 *
 * A chart has to reserve space for its axis labels BEFORE it can lay out the
 * plot area, and on a server there is no element to measure. So this estimates
 * advance widths from a per-character table.
 *
 * It is an APPROXIMATION and says so. The widths below are proportional-sans
 * averages; the browser will lay the text out for real and may differ by a few
 * percent. That is fine for reserving margins and for deciding whether a label
 * fits, which is all it is used for — it is never used to position glyphs.
 *
 * Where exactness matters (Studio, eventually), the same interface can be backed
 * by real font metrics; nothing above this file assumes the estimate.
 */

/** Advance width as a fraction of font size, for a proportional sans. */
const WIDTHS = {
  ' ': 0.26, '!': 0.26, '"': 0.36, '#': 0.55, $: 0.55, '%': 0.87, '&': 0.67, "'": 0.19,
  '(': 0.33, ')': 0.33, '*': 0.39, '+': 0.58, ',': 0.28, '-': 0.33, '.': 0.28, '/': 0.28,
  0: 0.55, 1: 0.55, 2: 0.55, 3: 0.55, 4: 0.55, 5: 0.55, 6: 0.55, 7: 0.55, 8: 0.55, 9: 0.55,
  ':': 0.28, ';': 0.28, '<': 0.58, '=': 0.58, '>': 0.58, '?': 0.55, '@': 1.0,
  A: 0.67, B: 0.67, C: 0.72, D: 0.72, E: 0.67, F: 0.61, G: 0.78, H: 0.72, I: 0.28, J: 0.5,
  K: 0.67, L: 0.56, M: 0.83, N: 0.72, O: 0.78, P: 0.67, Q: 0.78, R: 0.72, S: 0.67, T: 0.61,
  U: 0.72, V: 0.67, W: 0.94, X: 0.67, Y: 0.67, Z: 0.61,
  '[': 0.28, '\\': 0.28, ']': 0.28, '^': 0.47, _: 0.55, '`': 0.33,
  a: 0.55, b: 0.55, c: 0.5, d: 0.55, e: 0.55, f: 0.28, g: 0.55, h: 0.55, i: 0.22, j: 0.22,
  k: 0.5, l: 0.22, m: 0.83, n: 0.55, o: 0.55, p: 0.55, q: 0.55, r: 0.33, s: 0.5, t: 0.28,
  u: 0.55, v: 0.5, w: 0.72, x: 0.5, y: 0.5, z: 0.5,
  '{': 0.33, '|': 0.26, '}': 0.33, '~': 0.58,
};
const DEFAULT_WIDTH = 0.55;
const BOLD_FACTOR = 1.06;

/** Estimated width of a string, in the same units as `size`. */
export function measureText(value, { size = 11, weight = 'normal' } = {}) {
  if (value === null || value === undefined) return 0;
  const s = String(value);
  let total = 0;
  for (const ch of s) {
    // A character we have no entry for is most likely wider than Latin lowercase
    // (CJK, emoji), so guess generously rather than under-reserve and clip.
    const w = WIDTHS[ch] ?? (ch.codePointAt(0) > 0x2e80 ? 1.0 : DEFAULT_WIDTH);
    total += w;
  }
  const bold = weight === 'bold' || Number(weight) >= 600;
  return total * size * (bold ? BOLD_FACTOR : 1);
}

/** Widest of a list — what an axis needs to reserve. */
export const widestText = (values, opts) =>
  values.reduce((max, v) => Math.max(max, measureText(v, opts)), 0);

/** Rough line height. Chart labels are single-line; this is for stacked text. */
export const lineHeight = (size = 11) => Math.round(size * 1.35);

/** Ascent above the baseline, for vertically centring a label on a mark. */
export const capHeight = (size = 11) => size * 0.71;

/** Greedy word wrap to a pixel width. Returns the lines. */
export function wrapText(value, width, opts = {}) {
  const words = String(value ?? '').split(/\s+/).filter(Boolean);
  if (!words.length) return [];
  const lines = [];
  let current = words[0];
  for (const word of words.slice(1)) {
    const candidate = current + ' ' + word;
    if (measureText(candidate, opts) <= width) current = candidate;
    else { lines.push(current); current = word; }
  }
  lines.push(current);
  return lines;
}

/**
 * Shorten to fit, with an ellipsis. Returns the original when it already fits,
 * so a caller can tell whether truncation happened by comparing.
 */
export function truncateText(value, width, opts = {}) {
  const s = String(value ?? '');
  if (measureText(s, opts) <= width) return s;
  const ellipsis = '…';
  const budget = width - measureText(ellipsis, opts);
  if (budget <= 0) return '';
  let out = '';
  for (const ch of s) {
    if (measureText(out + ch, opts) > budget) break;
    out += ch;
  }
  return out + ellipsis;
}
