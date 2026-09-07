// The unit and colour arithmetic every part of a slide needs.
//
// PowerPoint measures in English Metric Units — 914400 to the inch, 12700 to
// the point — chosen so that both inches and centimetres divide exactly. Text
// sizes are in hundredths of a point, angles in sixtieths of a degree, and
// percentages in thousandths. Getting one of these wrong moves a shape by an
// inch, so they live here and nowhere else.

export const EMU_PER_INCH = 914400;
export const EMU_PER_POINT = 12700;
export const EMU_PER_CM = 360000;

/** EMU to CSS pixels at 96 dpi — the unit the renderer draws in. */
export const emuToPx = (emu) => (Number(emu) || 0) / EMU_PER_INCH * 96;
export const pxToEmu = (px) => Math.round((Number(px) || 0) / 96 * EMU_PER_INCH);
export const emuToPt = (emu) => (Number(emu) || 0) / EMU_PER_POINT;
export const ptToEmu = (pt) => Math.round((Number(pt) || 0) * EMU_PER_POINT);

/** Hundredths of a point, as written in `sz="1800"`. */
export const szToPt = (sz) => (Number(sz) || 0) / 100;
export const ptToSz = (pt) => Math.round((Number(pt) || 0) * 100);

/** Sixtieths of a degree, as written in `rot="5400000"`. */
export const rotToDeg = (rot) => ((Number(rot) || 0) / 60000) % 360;
export const degToRot = (deg) => Math.round((Number(deg) || 0) * 60000);

/** Thousandths of a percent, as written in `lumMod val="60000"`. */
export const pctOf = (v, fallback = 1) => {
  if (v == null || v === '') return fallback;
  const s = String(v);
  if (s.endsWith('%')) return parseFloat(s) / 100;
  return Number(s) / 100000;
};

const HEX = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');

export function rgbToHex({ r, g, b }) {
  return `#${HEX(r)}${HEX(g)}${HEX(b)}`;
}

export function hexToRgb(hex) {
  const h = String(hex || '').replace('#', '');
  const full = h.length === 3 ? h.split('').map((c) => c + c).join('') : h.padEnd(6, '0');
  return {
    r: parseInt(full.slice(0, 2), 16) || 0,
    g: parseInt(full.slice(2, 4), 16) || 0,
    b: parseInt(full.slice(4, 6), 16) || 0,
  };
}

function toHsl({ r, g, b }) {
  const R = r / 255;
  const G = g / 255;
  const B = b / 255;
  const max = Math.max(R, G, B);
  const min = Math.min(R, G, B);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === R) h = ((G - B) / d + (G < B ? 6 : 0)) / 6;
  else if (max === G) h = ((B - R) / d + 2) / 6;
  else h = ((R - G) / d + 4) / 6;
  return { h, s, l };
}

function fromHsl({ h, s, l }) {
  if (s === 0) {
    const v = l * 255;
    return { r: v, g: v, b: v };
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t) => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return { r: hue(h + 1 / 3) * 255, g: hue(h) * 255, b: hue(h - 1 / 3) * 255 };
}

/**
 * Apply the colour transforms a theme colour carries — the reason "accent1"
 * appears in six visibly different shades in one deck.
 * @param {string} hex
 * @param {Array<{ name: string, val: string }>} transforms
 */
export function applyColorTransforms(hex, transforms = []) {
  let rgb = hexToRgb(hex);
  let alpha = 1;
  for (const t of transforms) {
    const v = pctOf(t.val, 1);
    switch (t.name) {
      case 'lumMod': {
        const hsl = toHsl(rgb);
        rgb = fromHsl({ ...hsl, l: Math.min(1, hsl.l * v) });
        break;
      }
      case 'lumOff': {
        const hsl = toHsl(rgb);
        rgb = fromHsl({ ...hsl, l: Math.min(1, Math.max(0, hsl.l + v)) });
        break;
      }
      case 'shade': {
        rgb = { r: rgb.r * v, g: rgb.g * v, b: rgb.b * v };
        break;
      }
      case 'tint': {
        rgb = { r: rgb.r * v + 255 * (1 - v), g: rgb.g * v + 255 * (1 - v), b: rgb.b * v + 255 * (1 - v) };
        break;
      }
      case 'satMod': {
        const hsl = toHsl(rgb);
        rgb = fromHsl({ ...hsl, s: Math.min(1, hsl.s * v) });
        break;
      }
      case 'alpha':
        alpha = v;
        break;
      default:
        break;
    }
  }
  const hexOut = rgbToHex(rgb);
  return alpha >= 1 ? hexOut : { hex: hexOut, alpha };
}

/** The 140 or so colour names OOXML allows in `prstClr`. The common ones. */
export const PRESET_COLORS = {
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff',
  yellow: '#ffff00', cyan: '#00ffff', magenta: '#ff00ff', gray: '#808080', grey: '#808080',
  darkGray: '#a9a9a9', lightGray: '#d3d3d3', orange: '#ffa500', purple: '#800080',
  brown: '#a52a2a', pink: '#ffc0cb', navy: '#000080', teal: '#008080', olive: '#808000',
  maroon: '#800000', lime: '#00ff00', silver: '#c0c0c0', gold: '#ffd700', transparent: 'none',
};
