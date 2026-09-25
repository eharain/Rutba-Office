// The suite's own themes — Design → Themes, Variants, Colours, Fonts, Effects.
//
// A theme is one part, `ppt/theme/themeN.xml`: twelve colours, a heading and
// a body font, and the "format scheme" of fills, lines, effects and
// backgrounds that shapes and backgrounds refer to by index. Everything a
// slide does not state for itself comes from there, through the master, so
// swapping those three elements restyles a deck without touching one slide.
//
// The themes here are our own: names, palettes and pairings chosen for this
// suite, each written as a complete theme part PowerPoint opens as its own.
// Fonts are ones Windows ships (most also on macOS); where a machine lacks
// one, the renderer falls back to a face of the same kind — a serif for a
// serif — rather than to whatever comes first.

import { escapeXml } from '@rutba/office-formats/xml';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** The twelve colour slots in the order the schema puts them. */
export const COLOUR_SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** What PowerPoint's Customise Colours dialog calls each slot. */
export const COLOUR_SLOT_NAMES = {
  dk1: 'Text/Background – Dark 1',
  lt1: 'Text/Background – Light 1',
  dk2: 'Text/Background – Dark 2',
  lt2: 'Text/Background – Light 2',
  accent1: 'Accent 1',
  accent2: 'Accent 2',
  accent3: 'Accent 3',
  accent4: 'Accent 4',
  accent5: 'Accent 5',
  accent6: 'Accent 6',
  hlink: 'Hyperlink',
  folHlink: 'Followed Hyperlink',
};

const palette = (dk1, lt1, dk2, lt2, accents, hlink, folHlink) => ({
  dk1, lt1, dk2, lt2,
  accent1: accents[0], accent2: accents[1], accent3: accents[2],
  accent4: accents[3], accent5: accents[4], accent6: accents[5],
  hlink, folHlink,
});

/**
 * The built-in themes. `palette` is the first variant; `dark` puts the
 * master's colour map the other way round (backgrounds from dk1, text from
 * lt1), which is how a dark theme stays a theme rather than a pile of
 * hard-coded white text. `background` is the master's treatment.
 */
export const THEMES = [
  {
    id: 'rutba', name: 'Rutba', effects: 'soft', background: 'plain',
    fonts: { major: 'Segoe UI Semibold', minor: 'Segoe UI' },
    palette: palette('000000', 'FFFFFF', '1F3864', 'EEF2F8', ['2B5FD9', '0F9D58', 'E08B2B', 'C2408F', '7B5CD6', '3AAFA9'], '0563C1', '954F72'),
  },
  {
    id: 'harbour', name: 'Harbour', effects: 'soft', background: 'band',
    fonts: { major: 'Segoe UI Semibold', minor: 'Segoe UI' },
    palette: palette('1C2B39', 'FFFFFF', '1F4E79', 'E6EEF6', ['1F6FB2', '2AA6A0', 'F07F5E', 'F4B942', '6C63B8', '8AA1B4'], '1F6FB2', '6C63B8'),
  },
  {
    id: 'meadow', name: 'Meadow', effects: 'soft', background: 'glow',
    fonts: { major: 'Trebuchet MS', minor: 'Trebuchet MS' },
    palette: palette('22301F', 'FFFFFF', '2F5233', 'EEF4E6', ['4C9A2A', 'A4C639', 'F2A541', '2E86AB', 'C8553D', '7D6B91'], '2E86AB', '7D6B91'),
  },
  {
    id: 'ember', name: 'Ember', effects: 'lifted', background: 'side',
    fonts: { major: 'Georgia', minor: 'Segoe UI' },
    palette: palette('2B2522', 'FFFDF9', '5A2E1F', 'F6EDE4', ['C8553D', 'E58F3A', 'D9B44A', '7A9E7E', '5B7DB1', '8C5E8F'], 'B5452F', '8C5E8F'),
  },
  {
    id: 'nightfall', name: 'Nightfall', effects: 'soft', background: 'deep', dark: true,
    fonts: { major: 'Segoe UI Semibold', minor: 'Segoe UI' },
    palette: palette('141A2E', 'F4F6FB', '24345C', 'C9D3EA', ['5AA9FF', '45D0B5', 'FFB454', 'FF7A8A', 'B18CFF', '8FD14F'], '7FBCFF', 'C3A6FF'),
  },
  {
    id: 'paperkite', name: 'Paper Kite', effects: 'flat', background: 'rule',
    fonts: { major: 'Palatino Linotype', minor: 'Segoe UI' },
    palette: palette('111111', 'FFFFFF', '333333', 'F3F1EC', ['D7263D', '1B998B', '2E294E', 'F46036', '5C8001', '7F7F7F'], 'D7263D', '2E294E'),
  },
  {
    id: 'basalt', name: 'Basalt', effects: 'flat', background: 'cap',
    fonts: { major: 'Bahnschrift', minor: 'Segoe UI' },
    palette: palette('1E2326', 'FFFFFF', '37474F', 'ECEFF1', ['00897B', '546E7A', 'FFA000', '5C6BC0', 'EF5350', '8D6E63'], '00796B', '5C6BC0'),
  },
  {
    id: 'saffron', name: 'Saffron', effects: 'lifted', background: 'glow',
    fonts: { major: 'Trebuchet MS', minor: 'Segoe UI' },
    palette: palette('2A1F3D', 'FFFFFF', '4A2C6D', 'FBF3DC', ['E8A317', '7B3F98', 'D1495B', '00A6A6', '3B8EA5', '9BC53D'], '7B3F98', 'D1495B'),
  },
  {
    id: 'tidewater', name: 'Tidewater', effects: 'soft', background: 'deep', dark: true,
    fonts: { major: 'Georgia', minor: 'Segoe UI' },
    palette: palette('0F2A2E', 'F1F7F6', '14474E', 'BFD9D6', ['4FD1C5', 'F6AD55', 'B79CF2', 'F687B3', '68D391', '63B3ED'], '63B3ED', 'B79CF2'),
  },
  {
    id: 'orchard', name: 'Orchard', effects: 'soft', background: 'wash',
    fonts: { major: 'Constantia', minor: 'Corbel' },
    palette: palette('3B2A33', 'FFFFFF', '6B2F4A', 'F9EEF1', ['B83B5E', 'F08A5D', 'E9B949', '6A2C70', '3F8F8B', '8E9AAF'], 'B83B5E', '6A2C70'),
  },
  {
    id: 'studio', name: 'Studio', effects: 'outline', background: 'capAccent',
    fonts: { major: 'Franklin Gothic Medium', minor: 'Franklin Gothic Book' },
    palette: palette('202124', 'FFFFFF', '2B3A55', 'F1F3F5', ['2563EB', 'F43F5E', '10B981', 'F59E0B', '8B5CF6', '06B6D4'], '2563EB', '8B5CF6'),
  },
];

/** Colours → the built-in palettes: each theme's own, then a few that belong to no theme. */
export const PALETTES = [
  ...THEMES.map((t) => ({ id: t.id, name: t.name, colors: t.palette })),
  { id: 'graphite', name: 'Graphite', colors: palette('000000', 'FFFFFF', '404040', 'EDEDED', ['595959', '7F7F7F', 'A5A5A5', '3F3F3F', '8C8C8C', 'BFBFBF'], '404040', '7F7F7F') },
  { id: 'lagoon', name: 'Lagoon', colors: palette('0B2530', 'FFFFFF', '0E4C5E', 'E3F4F6', ['0096C7', '48CAE4', '00B4A0', 'F9844A', 'F9C74F', '577590'], '0077B6', '577590') },
  { id: 'citrus', name: 'Citrus', colors: palette('263238', 'FFFFFF', '33691E', 'FFFDE7', ['F9A825', '7CB342', 'FB8C00', '00ACC1', 'E53935', '8E24AA'], '00838F', '8E24AA') },
  { id: 'berry', name: 'Berry', colors: palette('2D1B2E', 'FFFFFF', '4E1F4E', 'F6EAF4', ['8E3B8E', 'D64161', 'F28C28', '3A7CA5', '2A9D8F', 'B5838D'], '3A7CA5', '8E3B8E') },
];

/** Fonts → the built-in pairs: a heading face and a body face. */
export const FONT_PAIRS = [
  { id: 'rutba', name: 'Rutba', major: 'Segoe UI Semibold', minor: 'Segoe UI' },
  { id: 'classic', name: 'Classic', major: 'Georgia', minor: 'Georgia' },
  { id: 'editorial', name: 'Editorial', major: 'Palatino Linotype', minor: 'Segoe UI' },
  { id: 'humanist', name: 'Humanist', major: 'Trebuchet MS', minor: 'Trebuchet MS' },
  { id: 'friendly', name: 'Friendly', major: 'Candara', minor: 'Candara' },
  { id: 'technical', name: 'Technical', major: 'Bahnschrift', minor: 'Segoe UI' },
  { id: 'bookish', name: 'Bookish', major: 'Constantia', minor: 'Corbel' },
  { id: 'grotesque', name: 'Grotesque', major: 'Franklin Gothic Medium', minor: 'Franklin Gothic Book' },
  { id: 'newsprint', name: 'Newsprint', major: 'Cambria', minor: 'Calibri' },
  { id: 'plain', name: 'Plain', major: 'Arial', minor: 'Arial' },
  { id: 'readable', name: 'Readable', major: 'Verdana', minor: 'Verdana' },
];

/**
 * Effects → the format schemes: what a shape that takes its look from the
 * theme (a `p:style` reference, as PowerPoint's own shape styles are) is
 * filled, outlined and lifted with. Each list holds the three levels the
 * schema asks for, subtle to intense.
 */
export const EFFECT_PRESETS = [
  { id: 'flat', name: 'Flat', description: 'Solid fills, fine lines, no shadows' },
  { id: 'soft', name: 'Soft', description: 'Gentle gradients and a soft shadow on intense shapes' },
  { id: 'lifted', name: 'Lifted', description: 'Rich gradients and shapes lifted off the slide' },
  { id: 'outline', name: 'Outline', description: 'Flat fills with strong outlines' },
];

const scheme = (val, mods = '') => (mods ? `<a:schemeClr val="${val}">${mods}</a:schemeClr>` : `<a:schemeClr val="${val}"/>`);
const solid = (mods = '') => `<a:solidFill>${scheme('phClr', mods)}</a:solidFill>`;
const grad = (a, b, ang = 5400000) =>
  `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr">${a}</a:schemeClr></a:gs>` +
  `<a:gs pos="100000"><a:schemeClr val="phClr">${b}</a:schemeClr></a:gs></a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;
const line = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid()}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
const shadow = (blur, dist, alpha) =>
  `<a:effectStyle><a:effectLst><a:outerShdw blurRad="${blur}" dist="${dist}" dir="5400000" algn="ctr" rotWithShape="0">` +
  `<a:srgbClr val="000000"><a:alpha val="${alpha}"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>`;
const noEffect = '<a:effectStyle><a:effectLst/></a:effectStyle>';

/** The fmtScheme's four lists for one preset. */
function formatLists(id) {
  const bgFills = solid() + solid('<a:tint val="95000"/><a:satMod val="170000"/>') + grad('<a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/>', '<a:shade val="63000"/><a:satMod val="120000"/>');
  switch (id) {
    case 'flat':
      return {
        fills: solid() + solid('<a:tint val="85000"/>') + solid('<a:shade val="90000"/>'),
        lines: line(6350) + line(12700) + line(19050),
        effects: noEffect + noEffect + noEffect,
        bgFills,
      };
    case 'lifted':
      return {
        fills: solid() + grad('<a:tint val="72000"/><a:satMod val="110000"/>', '<a:shade val="94000"/>') + grad('<a:tint val="92000"/>', '<a:shade val="68000"/><a:satMod val="120000"/>'),
        lines: line(9525) + line(19050) + line(28575),
        effects: noEffect + shadow(50800, 25400, 30000) + shadow(76200, 38100, 42000),
        bgFills,
      };
    case 'outline':
      return {
        fills: solid() + solid('<a:tint val="60000"/>') + solid('<a:tint val="35000"/>'),
        lines: line(19050) + line(28575) + line(38100),
        effects: noEffect + noEffect + noEffect,
        bgFills,
      };
    case 'soft':
    default:
      return {
        fills: solid() + grad('<a:tint val="67000"/><a:satMod val="105000"/>', '<a:tint val="95000"/>') + grad('<a:tint val="94000"/>', '<a:shade val="80000"/>'),
        lines: line(6350) + line(12700) + line(19050),
        effects: noEffect + noEffect + shadow(57150, 19050, 63000),
        bgFills,
      };
  }
}

const hex6 = (v) => String(v || '').replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6);

/** `<a:clrScheme>` with its twelve slots. Black and white text/backgrounds are written as the system colours, the way PowerPoint writes them. */
export function clrSchemeXml(name, colors) {
  const slot = (k) => {
    const v = hex6(colors[k]);
    if (k === 'dk1' && v === '000000') return `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>`;
    if (k === 'lt1' && v === 'FFFFFF') return `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>`;
    return `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`;
  };
  return `<a:clrScheme name="${escapeXml(name)}">${COLOUR_SLOTS.map(slot).join('')}</a:clrScheme>`;
}

/** `<a:fontScheme>`: a heading and a body face; the East Asian and complex-script faces left to the system, as the schema allows. */
export function fontSchemeXml(name, { major, minor }) {
  const face = (tag, typeface) => `<a:${tag}><a:latin typeface="${escapeXml(typeface)}"/><a:ea typeface=""/><a:cs typeface=""/></a:${tag}>`;
  return `<a:fontScheme name="${escapeXml(name)}">${face('majorFont', major)}${face('minorFont', minor)}</a:fontScheme>`;
}

/** `<a:fmtScheme>` for one of the presets. */
export function fmtSchemeXml(id) {
  const preset = EFFECT_PRESETS.find((p) => p.id === id) || EFFECT_PRESETS[1];
  const l = formatLists(preset.id);
  return `<a:fmtScheme name="${escapeXml(preset.name)}"><a:fillStyleLst>${l.fills}</a:fillStyleLst><a:lnStyleLst>${l.lines}</a:lnStyleLst>` +
    `<a:effectStyleLst>${l.effects}</a:effectStyleLst><a:bgFillStyleLst>${l.bgFills}</a:bgFillStyleLst></a:fmtScheme>`;
}

/** A whole theme part. */
export function themePartXml({ name, colors, colorName = name, fonts, fontName = name, effects = 'soft' }) {
  return `${DECL}<a:theme xmlns:a="${A_NS}" name="${escapeXml(name)}"><a:themeElements>` +
    clrSchemeXml(colorName, colors) + fontSchemeXml(fontName, fonts) + fmtSchemeXml(effects) +
    `</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

/**
 * The master's background for a treatment, in theme colours only — so a
 * variant or a new palette carries it along: `plain` the background colour,
 * `wash` a fall from it into Background 2, `glow` a corner warmed by the
 * first accent, `deep` a dark fall for a dark theme.
 */
export function masterBackgroundXml(kind) {
  const bgPr = (fill) => `<p:bg><p:bgPr>${fill}<a:effectLst/></p:bgPr></p:bg>`;
  const stop = (pos, clr, mods = '') => `<a:gs pos="${pos}">${scheme(clr, mods)}</a:gs>`;
  const grad = (stops, ang) => bgPr(`<a:gradFill flip="none" rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${ang}" scaled="1"/></a:gradFill>`);
  // A band of colour across the background: hard stops either side of it,
  // so a theme has a mark of its own without a shape on the master.
  const band = (from, to, clr, ang, mods = '') =>
    grad(`${from > 0 ? stop(0, 'bg1') + stop(from, 'bg1') : ''}${stop(from, clr, mods)}${stop(to, clr, mods)}${to < 100000 ? stop(to, 'bg1') + stop(100000, 'bg1') : ''}`, ang);
  switch (kind) {
    case 'wash':
      return grad(`${stop(0, 'bg1')}${stop(62000, 'bg1')}${stop(100000, 'bg2')}`, 5400000);
    case 'glow':
      return grad(`${stop(0, 'bg1')}${stop(58000, 'bg1')}${stop(100000, 'accent1', '<a:lumMod val="20000"/><a:lumOff val="80000"/>')}`, 2700000);
    case 'deep':
      return grad(`${stop(0, 'bg1')}${stop(55000, 'bg1')}${stop(100000, 'bg2')}`, 5400000);
    case 'band':
      return band(92500, 100000, 'accent1', 5400000);
    case 'side':
      return band(0, 2400, 'accent1', 0);
    case 'rule':
      return band(5200, 6500, 'accent1', 5400000);
    case 'cap':
      return band(0, 4200, 'tx2', 5400000);
    case 'capAccent':
      return band(0, 4200, 'accent1', 5400000);
    case 'plain':
    default:
      return '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';
  }
}

/** A master's colour map: the usual one, or the dark one where Background 1 is the theme's dark colour. */
export function clrMapAttrs(dark) {
  return dark
    ? 'bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"'
    : 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';
}

// ---- variants ---------------------------------------------------------------

const rgb = (h) => [0, 2, 4].map((i) => parseInt(hex6(h).slice(i, i + 2), 16));
const toHex = (c) => c.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase();
const mix = (a, b, t) => toHex(rgb(a).map((v, i) => v + (rgb(b)[i] - v) * t));
const luminance = (h) => {
  const [r, g, b] = rgb(h).map((v) => {
    const s = v / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
};

/** The accents moved along, so a variant leads with another of them. */
function rotated(p, by) {
  const acc = [p.accent1, p.accent2, p.accent3, p.accent4, p.accent5, p.accent6];
  const r = acc.slice(by).concat(acc.slice(0, by));
  return { ...p, accent1: r[0], accent2: r[1], accent3: r[2], accent4: r[3], accent5: r[4], accent6: r[5] };
}

/**
 * Design → Variants: four colourings of one theme, the first its own. The
 * second and third lead with another accent, their second light colour
 * tinted towards it; the fourth turns a light theme dark (or a dark one
 * light), through the colour map as PowerPoint's own dark variants do.
 * Works for any palette, so a deck whose theme came from elsewhere has
 * variants too.
 */
export function variantsOf(base, dark = false) {
  const tinted = (p) => ({ ...p, lt2: dark ? p.lt2 : mix(p.accent1, 'FFFFFF', 0.9), dk2: dark ? mix(p.accent1, p.dk1, 0.72) : p.dk2 });
  const flipped = dark
    ? { ...base, dk1: mix(base.dk1, '000000', 0.2), lt1: 'FFFFFF', dk2: mix(base.dk2, base.dk1, 0.2), lt2: mix(base.accent1, 'FFFFFF', 0.9) }
    : { ...base, dk1: mix(base.dk2, '000000', 0.55), lt1: 'F7F8FA', dk2: mix(base.dk2, '000000', 0.25), lt2: mix(base.lt2, base.dk2, 0.25) };
  // A light palette turned dark needs its first accent to hold its own on the dark ground.
  if (!dark && luminance(flipped.accent1) < 0.12) Object.assign(flipped, rotated(flipped, 1));
  return [
    { name: 'Original', colors: base, dark },
    { name: 'Second accent', colors: tinted(rotated(base, 1)), dark },
    { name: 'Third accent', colors: tinted(rotated(base, 2)), dark },
    { name: dark ? 'Light' : 'Dark', colors: flipped, dark: !dark },
  ];
}

export function themeById(id) {
  return THEMES.find((t) => t.id === id) || null;
}
