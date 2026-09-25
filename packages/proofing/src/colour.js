// Colours for the contrast rule: WCAG 2.1's relative luminance and contrast
// ratio, and the few spellings a colour arrives in from the three formats.
//
// Word writes `2F5496` (no hash) or a highlight NAME (`yellow`, `darkBlue`);
// the sheet and deck readers hand over `#2F5496`; a theme colour has been
// resolved to hex by the time it reaches here. Anything this cannot read is
// null, and a rule with a null colour says nothing rather than guessing.

/** Word's highlighter names, in the colours Word paints them. */
export const HIGHLIGHT_HEX = {
  yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000',
  darkBlue: '#00008B', darkCyan: '#008B8B', darkGreen: '#006400', darkMagenta: '#8B008B', darkRed: '#8B0000',
  darkYellow: '#808000', darkGray: '#A9A9A9', lightGray: '#D3D3D3', black: '#000000', white: '#FFFFFF',
};

/** `#RRGGBB`, upper case, from any of the spellings above — or null. */
export function normaliseColour(value) {
  if (value == null) return null;
  const s = String(value).trim();
  if (!s || s === 'auto' || s === 'none') return null;
  if (HIGHLIGHT_HEX[s]) return HIGHLIGHT_HEX[s];
  const hex = /^#?([0-9a-f]{6})$/i.exec(s);
  if (hex) return '#' + hex[1].toUpperCase();
  const short = /^#?([0-9a-f]{3})$/i.exec(s);
  if (short) return '#' + short[1].split('').map((c) => c + c).join('').toUpperCase();
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(s);
  if (rgb) return '#' + [rgb[1], rgb[2], rgb[3]].map((n) => Math.max(0, Math.min(255, Number(n))).toString(16).padStart(2, '0')).join('').toUpperCase();
  if (s.toLowerCase() === 'white') return '#FFFFFF';
  if (s.toLowerCase() === 'black') return '#000000';
  return null;
}

/** WCAG relative luminance of `#RRGGBB`, 0 (black) to 1 (white). */
export function luminance(colour) {
  const hex = normaliseColour(colour);
  if (!hex) return null;
  const channel = (i) => {
    const c = parseInt(hex.slice(1 + i * 2, 3 + i * 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(0) + 0.7152 * channel(1) + 0.0722 * channel(2);
}

/** The contrast ratio of two colours, 1 to 21 — or null when either is unreadable. */
export function contrastRatio(a, b) {
  const la = luminance(a);
  const lb = luminance(b);
  if (la == null || lb == null) return null;
  const [hi, lo] = la >= lb ? [la, lb] : [lb, la];
  return (hi + 0.05) / (lo + 0.05);
}

/**
 * Large text by WCAG's measure: 18 pt, or 14 pt bold. Large text needs 3:1,
 * everything else 4.5:1 (level AA) — the bar Office's checker holds text to.
 */
export function isLargeText(sizePt, bold) {
  const size = Number(sizePt) || 0;
  return size >= 18 || (bold && size >= 14);
}

export function requiredRatio(sizePt, bold) {
  return isLargeText(sizePt, bold) ? 3 : 4.5;
}

/** Black or white, whichever reads better on `background` — the fix's colour. */
export function readableOn(background) {
  const lb = luminance(background);
  if (lb == null) return '#000000';
  return (1.05) / (lb + 0.05) >= (lb + 0.05) / 0.05 ? '#FFFFFF' : '#000000';
}
