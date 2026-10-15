/**
 * Readable ink on a document's own colours.
 *
 * A file brings colours of its own: a cell's fill, a font's colour, a
 * border's. The suite draws its chrome in the light or the dark theme, but a
 * document's colours are the author's and stay as they are, the way Excel
 * keeps a cell's fill in its dark mode. Where the two meet, words must still
 * be read:
 *
 * - a cell with a fill keeps that fill, and a font colour the file set keeps
 *   its colour too (the author put them together);
 * - automatic text (no colour in the file) takes the suite's dark or light
 *   ink, whichever reads better on the fill under it;
 * - with no fill, the words sit on the suite's own surface, which is dark at
 *   night: a colour the file set that would vanish there (navy, dark teal) is
 *   lifted, keeping its hue, until it reads.
 *
 * Plain functions, no DOM: the painters call them per cell and the tests
 * measure them.
 */

/** The suite's two inks and its two cell surfaces (theme.css). */
export const INK = { light: '#16181c', dark: '#f4f6f8' };
export const SURFACE = { light: '#ffffff', dark: '#1e2127' };

/** WCAG's bar for body text, and the one for large or bold words. */
export const AA = 4.5;
export const AA_LARGE = 3;

/** '#rgb', '#rrggbb', '#aarrggbb' (Excel's ARGB) or 'rgb(r, g, b)' as [r, g, b], or null. */
export function parseColour(value) {
  if (typeof value !== 'string') return null;
  const s = value.trim().toLowerCase();
  let m = /^#?([0-9a-f]{3})$/.exec(s);
  if (m) return [...m[1]].map((h) => parseInt(h + h, 16));
  m = /^#?(?:[0-9a-f]{2})?([0-9a-f]{6})$/.exec(s);
  if (m) return [0, 2, 4].map((i) => parseInt(m[1].slice(i, i + 2), 16));
  m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?\s*\)$/.exec(s);
  if (m) {
    // A colour more than half see-through says little about what is behind it.
    if (m[4] != null) {
      const a = m[4].endsWith('%') ? parseFloat(m[4]) / 100 : parseFloat(m[4]);
      if (a < 0.5) return null;
    }
    return [m[1], m[2], m[3]].map((n) => Math.min(255, Number(n)));
  }
  if (s === 'white') return [255, 255, 255];
  if (s === 'black') return [0, 0, 0];
  return null;
}

const hex = (rgb) => '#' + rgb.map((n) => Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, '0')).join('');

/** Relative luminance (WCAG 2), 0 for black to 1 for white. */
export function luminance(colour) {
  const rgb = Array.isArray(colour) ? colour : parseColour(colour);
  if (!rgb) return null;
  const [r, g, b] = rgb.map((n) => {
    const c = n / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** The contrast ratio of two colours, 1 to 21; null when either cannot be read. */
export function contrast(a, b) {
  const la = luminance(a);
  const lb = luminance(b);
  if (la == null || lb == null) return null;
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

/** Whichever of the suite's two inks reads better on a background. */
export function inkOn(background) {
  const dark = contrast(INK.light, background);
  const light = contrast(INK.dark, background);
  if (dark == null || light == null) return null;
  if (Math.max(dark, light) >= AA) return dark >= light ? INK.light : INK.dark;
  // A middling fill (a mid grey) that neither ink quite reads on: pure black or white.
  return contrast('#000000', background) >= contrast('#ffffff', background) ? '#000000' : '#ffffff';
}

const toHsl = ([r, g, b]) => {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
};
const fromHsl = ([h, s, l]) => {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const ch = (t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [ch(h + 1 / 3) * 255, ch(h) * 255, ch(h - 1 / 3) * 255];
};

/**
 * A colour made to read on a background at `ratio`, keeping its hue: the
 * colour itself when it already reads; otherwise its lightness mirrored (dark
 * navy becomes light blue, black becomes white, the way a dark mode maps a
 * document's colours), and moved further only if the mirror still does not
 * read.
 */
export function readableOn(colour, background, ratio = AA) {
  const rgb = parseColour(colour);
  const bg = parseColour(background);
  if (!rgb || !bg) return colour;
  if (contrast(rgb, bg) >= ratio) return colour;
  const [h, s, l] = toHsl(rgb);
  const up = luminance(bg) < 0.18;
  const reads = (light) => contrast(parseColour(hex(fromHsl([h, s, light]))), bg) >= ratio;
  const mirror = up ? Math.max(l, 1 - l) : Math.min(l, 1 - l);
  if (reads(mirror)) return hex(fromHsl([h, s, mirror]));
  const far = up ? 1 : 0;
  if (!reads(far)) return hex(fromHsl([h, s, far]));
  // Halve the way toward the far end until the ratio is just met.
  let lo = mirror;
  let hi = far;
  for (let i = 0; i < 20; i++) {
    const mid = (lo + hi) / 2;
    if (reads(mid)) hi = mid;
    else lo = mid;
  }
  return hex(fromHsl([h, s, hi]));
}

/**
 * A cell's words and edges, worked out against what is under them.
 *
 * `fill`: the cell's fill as the file (or a rule, or a table style) gave it,
 * or null. `font`: the file's font colour, or null for automatic. `dark`: the
 * suite's theme. Returns `{ ink, background }`, where `ink` is null when the
 * suite's own ink (the cell's CSS) is already right, and `background` is the
 * colour under the words.
 */
export function cellInk({ fill = null, font = null, dark = false } = {}) {
  const filled = parseColour(fill) ? fill : null;
  const background = filled || (dark ? SURFACE.dark : SURFACE.light);
  if (filled) {
    // The author's pair stays the author's.
    if (font && parseColour(font)) return { ink: font, background };
    return { ink: inkOn(filled), background };
  }
  if (!font || !parseColour(font)) return { ink: null, background };
  return { ink: dark ? readableOn(font, background, AA) : font, background };
}

/**
 * An edge's colour on what is beside it: an automatic edge (no colour) is the
 * file's black on a light page and the suite's ink at night; a set colour
 * keeps itself on a fill and is lifted on the dark surface the way words are.
 * Edges are lines, not words: the large-text bar is enough.
 */
export function edgeInk({ colour = null, fill = null, dark = false } = {}) {
  const filled = parseColour(fill) ? fill : null;
  if (!colour || !parseColour(colour)) {
    if (!dark) return '#000000';
    return filled ? (inkOn(filled) === INK.light ? '#000000' : INK.dark) : '#8d949f';
  }
  if (filled || !dark) return colour;
  return readableOn(colour, SURFACE.dark, AA_LARGE);
}
