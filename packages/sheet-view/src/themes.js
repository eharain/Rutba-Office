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
  OFFICE_THEME, readThemeDesign, withThemeElement, designedThemePart,
} from '@rutba/office-formats/themes';

export { THEMES, PALETTES, FONT_PAIRS, EFFECT_PRESETS, COLOUR_SLOTS };

export const THEME_PART = 'xl/theme/theme1.xml';
export const THEME_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';
export const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';

/** Office's own theme, which Excel gives a workbook that names no other: what a file without a theme part looks like. */
export const OFFICE = OFFICE_THEME;

const attr = (tag, name) => new RegExp('\\b' + name + '="([^"]*)"').exec(tag ?? '')?.[1];

/**
 * What the Page Layout tab shows as current: the theme's name and which of
 * the suite's it is, its twelve colours as hex and their scheme's name, its
 * two faces and their scheme's name, its format scheme.
 */
export function readWorkbookDesign(pkg) {
  return readThemeDesign(pkg.has(THEME_PART) ? pkg.text(THEME_PART) : null);
}

export { withThemeElement };

/**
 * The theme part as a design choice makes it — see `designedThemePart`. The
 * part the workbook has is changed in its one element; a workbook with none
 * gets Office's with the change.
 */
export function designedThemeXml(pkg, spec) {
  return designedThemePart(pkg.has(THEME_PART) ? pkg.text(THEME_PART) : null, spec);
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
      ? next.replace(/<name\b[^>]*\/>/, () => '<name val="' + face.replace(/"/g, '&quot;') + '"/>')
      : next.replace('</font>', () => '<name val="' + face.replace(/"/g, '&quot;') + '"/></font>');
    // Schema order: name, charset, family, … scheme last.
    if (!/<scheme\b/.test(next)) next = next.replace('</font>', () => '<scheme val="' + kind + '"/></font>');
    return next;
  });
  return stylesXml.replace(block[1], () => fonts);
}
