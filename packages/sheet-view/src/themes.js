/**
 * Page Layout → Themes, Colours, Fonts and Effects for a workbook.
 *
 * A workbook's look beyond its own cells is one part, `xl/theme/theme1.xml`:
 * twelve colours, a heading and a body face, and a format scheme. A cell
 * whose fill or font says `theme="4"`, a table style, a chart and a shape
 * styled from the theme, and the Normal style's font (`<scheme val="minor"/>`)
 * all take their look from there — so swapping the part restyles the
 * workbook without touching a cell. The themes are the suite's own, the same
 * eleven Presentation offers, so a deck and a workbook can match.
 */
import {
  THEMES, PALETTES, FONT_PAIRS, EFFECT_PRESETS, COLOUR_SLOTS,
  themePartXml, clrSchemeXml, fontSchemeXml, fmtSchemeXml, themeById,
} from '@rutba/office-formats/themes';

export { THEMES, PALETTES, FONT_PAIRS, EFFECT_PRESETS, COLOUR_SLOTS };

export const THEME_PART = 'xl/theme/theme1.xml';
export const THEME_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';
export const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';

/** Office's own theme, which Excel gives a workbook that names no other: what a file without a theme part looks like. */
export const OFFICE = {
  name: 'Office Theme',
  colorName: 'Office',
  colors: {
    dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
    accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
    hlink: '0563C1', folHlink: '954F72',
  },
  fontName: 'Office',
  fonts: { major: 'Calibri Light', minor: 'Calibri' },
};

const attr = (tag, name) => new RegExp('\\b' + name + '="([^"]*)"').exec(tag ?? '')?.[1];
const unesc = (s) => String(s ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/**
 * What the Page Layout tab shows as current: the theme's name and which of
 * the suite's it is, its twelve colours as hex and their scheme's name, its
 * two faces and their scheme's name, its format scheme.
 */
export function readWorkbookDesign(pkg) {
  if (!pkg.has(THEME_PART)) {
    return { exists: false, name: OFFICE.name, builtIn: null, colors: { ...OFFICE.colors }, colorName: OFFICE.colorName, fonts: { ...OFFICE.fonts }, fontName: OFFICE.fontName, effects: null, effectName: 'Office' };
  }
  const xml = pkg.text(THEME_PART);
  const name = unesc(attr(/<a:theme\b[^>]*>/.exec(xml)?.[0], 'name') ?? 'Theme');
  const clr = /<a:clrScheme\b([^>]*)>([\s\S]*?)<\/a:clrScheme>/.exec(xml);
  const colors = {};
  for (const k of COLOUR_SLOTS) {
    const m = clr ? new RegExp('<a:' + k + '>([\\s\\S]*?)</a:' + k + '>').exec(clr[2]) : null;
    const v = m ? (attr(/<a:srgbClr\b[^>]*>/.exec(m[1])?.[0], 'val') ?? attr(/<a:sysClr\b[^>]*>/.exec(m[1])?.[0], 'lastClr')) : null;
    colors[k] = v ? v.toUpperCase() : OFFICE.colors[k];
  }
  const fontBlock = /<a:fontScheme\b([^>]*)>([\s\S]*?)<\/a:fontScheme>/.exec(xml);
  const face = (tag) => {
    const m = fontBlock ? new RegExp('<a:' + tag + '>[\\s\\S]*?<a:latin\\b([^>]*)/?>').exec(fontBlock[2]) : null;
    return m ? unesc(attr(m[1], 'typeface') ?? '') : '';
  };
  const effectName = unesc(attr(/<a:fmtScheme\b[^>]*>/.exec(xml)?.[0], 'name') ?? '');
  return {
    exists: true,
    name,
    builtIn: THEMES.find((t) => t.name === name)?.id ?? null,
    colors,
    colorName: unesc(clr ? attr(clr[1], 'name') ?? '' : ''),
    fonts: { major: face('majorFont') || OFFICE.fonts.major, minor: face('minorFont') || OFFICE.fonts.minor },
    fontName: unesc(fontBlock ? attr(fontBlock[1], 'name') ?? '' : ''),
    effects: EFFECT_PRESETS.find((p) => p.name === effectName)?.id ?? null,
    effectName,
  };
}

/** One of `themeElements`' three schemes replaced (or put in its place when missing). */
export function withThemeElement(xml, tag, element) {
  const re = new RegExp('<a:' + tag + '\\b[^>]*?(?:/>|>[\\s\\S]*?</a:' + tag + '>)');
  if (re.test(xml)) return xml.replace(re, () => element);
  const order = ['clrScheme', 'fontScheme', 'fmtScheme'];
  const after = order.slice(0, order.indexOf(tag)).reverse().map((t) => new RegExp('</a:' + t + '>')).find((r) => r.test(xml));
  if (after) return xml.replace(after, (m) => m + element);
  return xml.replace(/<a:themeElements>/, (m) => m + element);
}

/**
 * The theme part as a design choice makes it: `{ theme }` one of the suite's
 * themes whole; `{ colors }` a palette id or twelve slots (with `name`);
 * `{ fonts }` a pair id or `{ major, minor }` (with `name`); `{ effects }` a
 * format scheme. The part the workbook has is changed in its one element;
 * a workbook with none gets Office's with the change.
 */
export function designedThemeXml(pkg, spec) {
  const base = pkg.has(THEME_PART)
    ? pkg.text(THEME_PART)
    : themePartXml({ name: OFFICE.name, colors: OFFICE.colors, colorName: OFFICE.colorName, fonts: OFFICE.fonts, fontName: OFFICE.fontName, effects: 'flat' });
  if (spec.theme) {
    const t = themeById(spec.theme);
    if (!t) throw new Error('no theme "' + spec.theme + '"');
    // A dark theme's palette is for slides on a dark ground; a sheet is
    // white paper, so its text and background colours keep the page light.
    const colors = t.dark ? { ...t.palette, dk1: t.palette.dk1, lt1: 'FFFFFF' } : t.palette;
    return themePartXml({ name: t.name, colors, colorName: t.name, fonts: t.fonts, fontName: t.name, effects: t.effects });
  }
  if (spec.colors) {
    const builtIn = typeof spec.colors === 'string' ? PALETTES.find((p) => p.id === spec.colors) : null;
    if (typeof spec.colors === 'string' && !builtIn) throw new Error('no palette "' + spec.colors + '"');
    const colors = builtIn ? builtIn.colors : spec.colors;
    for (const k of COLOUR_SLOTS) {
      if (!/^#?[0-9a-f]{6}$/i.test(String(colors?.[k] ?? ''))) throw new Error(k + ' needs a colour like 1F6FB2');
    }
    return withThemeElement(base, 'clrScheme', clrSchemeXml(String(spec.name || builtIn?.name || 'Custom').trim() || 'Custom', colors));
  }
  if (spec.fonts) {
    const pair = typeof spec.fonts === 'string' ? FONT_PAIRS.find((p) => p.id === spec.fonts) : null;
    if (typeof spec.fonts === 'string' && !pair) throw new Error('no font pair "' + spec.fonts + '"');
    const major = String(pair ? pair.major : spec.fonts.major ?? '').trim();
    const minor = String(pair ? pair.minor : spec.fonts.minor ?? '').trim();
    if (!major || !minor) throw new Error('a heading font and a body font are both needed');
    return withThemeElement(base, 'fontScheme', fontSchemeXml(String(spec.name || pair?.name || 'Custom').trim() || 'Custom', { major, minor }));
  }
  if (spec.effects) {
    if (!EFFECT_PRESETS.some((p) => p.id === spec.effects)) throw new Error('no effects "' + spec.effects + '"');
    return withThemeElement(base, 'fmtScheme', fmtSchemeXml(spec.effects));
  }
  throw new Error('say which theme, colours, fonts or effects');
}

/**
 * The style table's fonts made to follow the theme's faces: a font that
 * names its scheme (`<scheme val="minor"/>`, `"major"`) takes the new face,
 * as Excel rewrites them; and the Normal style's font — the workbook's
 * default — joins the body face when it wore the old one.
 */
export function stylesFollowingFonts(stylesXml, { major, minor }, was = OFFICE.fonts) {
  const block = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(stylesXml);
  if (!block) return stylesXml;
  const normalFont = Number(attr(/<cellStyleXfs\b[^>]*>\s*<xf\b[^>]*>/.exec(stylesXml)?.[0]?.replace(/^<cellStyleXfs\b[^>]*>\s*/, ''), 'fontId') ?? 0);
  let i = -1;
  const fonts = block[1].replace(/<font\b[^>]*?(?:\/>|>[\s\S]*?<\/font>)/g, (font) => {
    i += 1;
    const scheme = attr(/<scheme\b[^>]*\/>/.exec(font)?.[0], 'val');
    const name = attr(/<name\b[^>]*\/>/.exec(font)?.[0], 'val');
    let kind = scheme === 'major' || scheme === 'minor' ? scheme : null;
    if (!kind && i === normalFont && (!name || name === was.minor || /^(Calibri|Aptos|Aptos Narrow)$/.test(name))) kind = 'minor';
    if (!kind) return font;
    const face = kind === 'major' ? major : minor;
    let next = font;
    if (/\/>$/.test(next) && !/<\/font>$/.test(next)) next = next.replace(/\s*\/>$/, '></font>');
    next = /<name\b[^>]*\/>/.test(next)
      ? next.replace(/<name\b[^>]*\/>/, '<name val="' + face.replace(/"/g, '&quot;') + '"/>')
      : next.replace('</font>', '<name val="' + face.replace(/"/g, '&quot;') + '"/></font>');
    // Schema order: name, charset, family, … scheme last.
    if (!/<scheme\b/.test(next)) next = next.replace('</font>', '<scheme val="' + kind + '"/></font>');
    return next;
  });
  return stylesXml.replace(block[1], () => fonts);
}
