/**
 * The shared visual language for anything Rutba draws.
 *
 * One palette across Workspace, Studio and Mail, so a chart pasted from a sheet
 * into an email into a deck stays the same chart. Colour in a product is not a
 * taste question — it is an accessibility question with a pass/fail answer, so
 * these values were validated rather than chosen.
 *
 * VALIDATION RESULTS (recorded so a future change can be compared, not guessed):
 *
 *   light, surface #fcfcfb — lightness band PASS, chroma floor PASS,
 *     worst adjacent CVD ΔE 9.1 (#eda100↔#1baf7a, protan), worst adjacent
 *     normal-vision ΔE 19.6. Contrast WARN: aqua 2.74, yellow 2.11,
 *     magenta 2.62 sit below 3:1 — see RELIEF_REQUIRED below.
 *   dark, surface #1a1a19 — every check PASS, all eight ≥ 3:1.
 *   all-pairs (scatter, bubble — where any two series can touch): only the
 *     FIRST THREE slots clear the floors. Hence SCATTER_SERIES_CAP.
 *
 * If you change a hex, re-run the validator before shipping:
 *   node scripts/validate_palette.js "<hex,…>" --mode light  --surface "#fcfcfb"
 *   node scripts/validate_palette.js "<hex,…>" --mode dark   --surface "#1a1a19"
 *
 * The slot ORDER is the colour-blind-safety mechanism, not decoration. Hues are
 * assigned in fixed order and never cycled: a ninth series folds into "Other" or
 * becomes small multiples. Colour follows the entity, never its rank, so
 * filtering a series must not repaint the survivors.
 */

/** Categorical slots, in fixed assignment order. */
export const CATEGORICAL = {
  light: ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7', '#e34948'],
  dark: ['#3987e5', '#d95926', '#199e70', '#c98500', '#d55181', '#008300', '#9085e9', '#e66767'],
};
export const CATEGORICAL_HUES = ['blue', 'orange', 'aqua', 'yellow', 'magenta', 'green', 'violet', 'red'];

/**
 * Slots that fall below 3:1 against the light surface. A chart using these must
 * carry visible direct labels or offer a table view — colour alone is not enough
 * to find the series. Empty in dark mode, where all eight clear 3:1.
 */
export const RELIEF_REQUIRED = { light: [2, 3, 4], dark: [] };

/**
 * Forms where any two series can end up adjacent (scatter, bubble, choropleth)
 * cannot use all eight — only the first three clear the all-pairs floors.
 */
export const SCATTER_SERIES_CAP = 3;

/** Sequential: one hue, light to dark. Never a rainbow. */
export const SEQUENTIAL_BLUE = [
  '#cde2fb', '#b7d3f6', '#9ec5f4', '#86b6ef', '#6da7ec', '#5598e7',
  '#3987e5', '#2a78d6', '#256abf', '#1c5cab', '#184f95', '#104281', '#0d366b',
];
/** An ORDINAL ramp (discrete ordered marks) must stay clear of the surface. */
export const ORDINAL_START = { light: 3, dark: 10 };

/** Diverging: two poles that read as opposite, neutral grey between them. */
export const DIVERGING = {
  low: '#2a78d6',
  high: '#e34948',
  midpoint: { light: '#f0efec', dark: '#383835' },
};

/**
 * Status is reserved. Never reuse one of these for "series 4" — and never let a
 * status colour carry meaning alone; it ships with an icon and a label.
 */
export const STATUS = {
  good: '#0ca30c',
  warning: '#fab219',
  serious: '#ec835a',
  critical: '#d03b3b',
};

/** Chrome and ink. Text always wears an ink token, never the series colour. */
export const INK = {
  light: {
    surface: '#fcfcfb',
    plane: '#f9f9f7',
    primary: '#0b0b0b',
    secondary: '#52514e',
    muted: '#898781',
    gridline: '#e1e0d9',
    axis: '#c3c2b7',
    border: 'rgba(11,11,11,0.10)',
  },
  dark: {
    surface: '#1a1a19',
    plane: '#0d0d0d',
    primary: '#ffffff',
    secondary: '#c3c2b7',
    muted: '#898781',
    gridline: '#2c2c2a',
    axis: '#383835',
    border: 'rgba(255,255,255,0.10)',
  },
};

/**
 * Mark geometry. Thin marks, recessive chrome — the data should be the darkest
 * thing on the surface.
 */
export const MARKS = {
  lineWidth: 2,
  markerSize: 8,
  barRadius: 4, // rounded only at the data end, anchored to the baseline
  fillGap: 2, // surface-coloured gap between adjacent or stacked fills
  gridWidth: 1,
  axisWidth: 1,
};

export const FONT = {
  family: 'system-ui, -apple-system, "Segoe UI", sans-serif',
  title: 14,
  label: 11,
  legend: 11,
};

export const theme = (mode = 'light') => ({
  mode,
  categorical: CATEGORICAL[mode] ?? CATEGORICAL.light,
  ink: INK[mode] ?? INK.light,
  reliefRequired: RELIEF_REQUIRED[mode] ?? [],
  marks: MARKS,
  font: FONT,
});

/**
 * Colour for series `i`. Fixed order, never cycled — past the eighth slot the
 * caller is expected to fold into "Other" rather than invent a hue, so this
 * returns the muted ink instead of wrapping around.
 */
export function seriesColour(index, mode = 'light') {
  const slots = CATEGORICAL[mode] ?? CATEGORICAL.light;
  return index < slots.length ? slots[index] : (INK[mode] ?? INK.light).muted;
}

/** Does this chart need direct labels or a table view to be readable? */
export function needsRelief(seriesCount, mode = 'light') {
  const relief = RELIEF_REQUIRED[mode] ?? [];
  return relief.some((slot) => slot < seriesCount);
}
