// The colours a colour palette offers, as Office lays them out: the theme's
// ten colours across the top, each with the five lighter and darker steps
// Office works out for it beneath, then the ten standard colours, then the
// colours chosen lately. Pure: no DOM, no React, so the tests can hold it to
// Office's own numbers.

/** Office's theme, for a document that names no other. */
export const OFFICE_COLOURS = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
  accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
};

/** The theme's columns, in Office's order: Background 1, Text 1, Background 2, Text 2, Accent 1 to 6. */
export const THEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

/** Office's ten standard colours, by the names Office gives them. */
export const STANDARD = [
  ['C00000', 'darkRed'], ['FF0000', 'red'], ['FFC000', 'orange'], ['FFFF00', 'yellow'], ['92D050', 'lightGreen'],
  ['00B050', 'green'], ['00B0F0', 'lightBlue'], ['0070C0', 'blue'], ['002060', 'darkBlue'], ['7030A0', 'purple'],
];

/**
 * Word's fifteen highlight colours, five to a row as Word draws them: the
 * value is the `w:highlight` name the file stores, the hex what it shows.
 */
export const HIGHLIGHT = [
  ['yellow', 'FFFF00'], ['green', '00FF00'], ['cyan', '00FFFF'], ['magenta', 'FF00FF'], ['blue', '0000FF'],
  ['red', 'FF0000'], ['darkBlue', '000080'], ['darkCyan', '008080'], ['darkGreen', '008000'], ['darkMagenta', '800080'],
  ['darkRed', '800000'], ['darkYellow', '808000'], ['darkGray', '808080'], ['lightGray', 'C0C0C0'], ['black', '000000'],
];

/**
 * `#abc`, `abc`, `#AABBCC`, `aabbcc` → `AABBCC`; Excel's `FFAABBCC` (alpha
 * first, as a sheet reports a colour) → `AABBCC`; anything else → null.
 */
export function hexOf(value) {
  const s = String(value ?? '').trim().replace(/^#/, '');
  if (/^[0-9a-f]{3}$/i.test(s)) return s.split('').map((c) => c + c).join('').toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(s)) return s.toUpperCase();
  if (/^[0-9a-f]{8}$/i.test(s)) return s.slice(2).toUpperCase();
  return null;
}

const channels = (hex) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
const toHex = (rgb) => rgb.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase();

/** RGB 0 to 255 → HSL, each 0 to 1. */
export function hslOf(hex) {
  const [r, g, b] = channels(hex).map((v) => v / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

/** HSL, each 0 to 1 → hex. */
export function hexFromHsl([h, s, l]) {
  if (s === 0) return toHex([l * 255, l * 255, l * 255]);
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const part = (t) => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return toHex([part(h + 1 / 3) * 255, part(h) * 255, part(h - 1 / 3) * 255]);
}

/** HSV, each 0 to 1 → hex: the custom picker's square and strip. */
export function hexFromHsv([h, s, v]) {
  const i = Math.floor(h * 6) % 6;
  const f = h * 6 - Math.floor(h * 6);
  const p = v * (1 - s);
  const q = v * (1 - f * s);
  const u = v * (1 - (1 - f) * s);
  const rgb = [[v, u, p], [q, v, p], [p, v, u], [p, q, v], [u, p, v], [v, p, q]][i];
  return toHex(rgb.map((c) => c * 255));
}

/** Hex → HSV, each 0 to 1. */
export function hsvOf(hex) {
  const [r, g, b] = channels(hex).map((c) => c / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const d = max - min;
  let h = 0;
  if (d) h = max === r ? ((g - b) / d + (g < b ? 6 : 0)) / 6 : max === g ? ((b - r) / d + 2) / 6 : ((r - g) / d + 4) / 6;
  return [h, max ? d / max : 0, max];
}

/** Hex → [r, g, b], 0 to 255. */
export const rgbOf = (hex) => channels(hexOf(hex) || '000000');
/** [r, g, b] → hex. */
export const hexFromRgb = (rgb) => toHex(rgb);

/**
 * The five steps under a theme colour, as Office chooses them by how light
 * the colour is: `{ hex, lighter }` or `{ hex, darker }`, a percentage.
 * Lighter by p moves the lightness p of the way to white; darker by p takes
 * p of it away. Black and white step in small amounts, as Office's do.
 */
export function steps(hex) {
  const base = hexOf(hex) || '000000';
  const [h, s, l] = hslOf(base);
  const lighter = (p) => ({ hex: hexFromHsl([h, s, l + (1 - l) * p / 100]), lighter: p });
  const darker = (p) => ({ hex: hexFromHsl([h, s, l * (1 - p / 100)]), darker: p });
  if (l <= 0.0001) return [50, 35, 25, 15, 5].map(lighter);
  if (l >= 0.9999) return [5, 15, 25, 35, 50].map(darker);
  if (l < 0.2) return [90, 75, 50, 25, 10].map(lighter);
  if (l > 0.8) return [10, 25, 50, 75, 90].map(darker);
  return [lighter(80), lighter(60), lighter(40), darker(25), darker(50)];
}

/**
 * The theme grid: ten columns, each `{ slot, hex, steps }`, from a theme's
 * colours (bare or `#` hex; missing slots are Office's).
 */
export function themeGrid(colours = null) {
  return THEME_SLOTS.map((slot) => {
    const hex = hexOf(colours?.[slot]) || OFFICE_COLOURS[slot];
    return { slot, hex, steps: steps(hex) };
  });
}

/** The recent colours: newest first, ten at most, each once. */
export function remember(list, hex, max = 10) {
  const h = hexOf(hex);
  if (!h) return list.slice(0, max);
  return [h, ...list.filter((x) => x !== h)].slice(0, max);
}

/** Whether two colours are the same colour, whichever way they are written. */
export const sameColour = (a, b) => {
  const x = hexOf(a);
  return Boolean(x) && x === hexOf(b);
};
