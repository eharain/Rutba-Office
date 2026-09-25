// The suite's own themes, for Design → Themes, Variants, Colours, Fonts and
// Effects. The themes themselves — palettes, font pairs, format schemes and
// the theme part they make — are shared with Worksheets and live in
// @rutba/office-formats/themes; what is here is PresentationML's own: the
// master's background treatments and its colour map.

export {
  COLOUR_SLOTS,
  COLOUR_SLOT_NAMES,
  THEMES,
  PALETTES,
  FONT_PAIRS,
  EFFECT_PRESETS,
  clrSchemeXml,
  fontSchemeXml,
  fmtSchemeXml,
  themePartXml,
  variantsOf,
  themeById,
} from '@rutba/office-formats/themes';

const scheme = (val, mods = '') => (mods ? `<a:schemeClr val="${val}">${mods}</a:schemeClr>` : `<a:schemeClr val="${val}"/>`);

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
