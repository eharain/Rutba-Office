// WordArt's Transform: words drawn along a path rather than in straight
// lines — the `a:prstTxWarp` preset on a text body, which Word, PowerPoint
// and Excel share. The arcs, the circle and the button are drawn here, each
// line of the body along its own path of the preset, in SVG with
// <textPath>, so the same drawing serves the page, the slide, the sheet, the
// thumbnail and print.

import { measureText } from './measure.js';

/** The presets offered, in the order Office's Transform gallery shows them. */
export const WARP_PRESETS = [
  { id: 'textNoShape', label: 'No Transform' },
  { id: 'textArchUp', label: 'Arch Up' },
  { id: 'textArchDown', label: 'Arch Down' },
  { id: 'textCircle', label: 'Circle' },
  { id: 'textButton', label: 'Button' },
];

/** Whether a preset is one this drawing follows (the rest are drawn straight). */
export const drawnWarp = (preset) => ['textArchUp', 'textArchDown', 'textCircle', 'textButton'].includes(preset);

/** Half an ellipse's length, by Ramanujan's second approximation. */
function halfEllipse(a, b) {
  const h = ((a - b) / (a + b)) ** 2;
  return (Math.PI * (a + b) * (1 + (3 * h) / (10 + Math.sqrt(4 - 3 * h)))) / 2;
}

const f = (n) => Number(n).toFixed(2);

/**
 * The paths a preset lays its lines along, inside `box` (the text box less
 * its insets): each `{ d, length, at, outside, cap, share, spread }` — `at`
 * where along the path the line's middle sits (a share of its length);
 * `outside` whether the letters stand outside the curve, so the curve is
 * drawn in by their height (`inset`) to keep them in the box; `cap` the
 * tallest the letters may be, as a share of the box's height; `share` how
 * much of the path the line fills; `spread` whether the letters are spaced
 * out to go the whole way round.
 */
export function warpPaths(preset, { x, y, w, h }, { inset = [] } = {}) {
  const cx = x + w / 2;
  const cy = y + h / 2;
  const radii = (i, by = 1) => {
    const d = Math.max(0, Math.min((inset[i] || 0) * by, w / 2 - 2, h / 2 - 2));
    return { rx: w / 2 - d, ry: h / 2 - d };
  };
  const arc = (sweep, i, by) => {
    const { rx, ry } = radii(i, by);
    return { d: `M ${f(cx - rx)} ${f(cy)} A ${f(rx)} ${f(ry)} 0 0 ${sweep} ${f(cx + rx)} ${f(cy)}`, length: halfEllipse(rx, ry), at: 0.5 };
  };
  // Over the top, left to right: the letters stand outside the curve.
  if (preset === 'textArchUp') return [{ ...arc(1, 0, 1), outside: true, cap: 0.42, share: 0.92 }];
  // Along the bottom, left to right: they stand inside it, their feet to the
  // edge; drawn in a little and not run to the steep ends, where they would
  // stand out past the box.
  if (preset === 'textArchDown') return [{ ...arc(0, 0, 0.3), outside: false, cap: 0.34, share: 0.78 }];
  if (preset === 'textCircle') {
    // Round from the foot, by the left, over the top and back: the letters
    // spaced out to go the whole way round, the words' middle at the top.
    const { rx, ry } = radii(0);
    return [{ d: `M ${f(cx)} ${f(cy + ry)} A ${f(rx)} ${f(ry)} 0 1 1 ${f(cx)} ${f(cy - ry)} A ${f(rx)} ${f(ry)} 0 1 1 ${f(cx)} ${f(cy + ry)}`, length: 2 * halfEllipse(rx, ry), at: 0.5, outside: true, cap: 0.28, share: 0.97, spread: true }];
  }
  if (preset === 'textButton') {
    // The top arc, a straight middle, and the bottom arc: one line each.
    const mid = cy + (inset[1] || 0) * 0.45;
    return [
      { ...arc(1, 0, 1), outside: true, cap: 0.22, share: 0.9 },
      { d: `M ${f(x)} ${f(mid)} L ${f(x + w)} ${f(mid)}`, length: w, at: 0.5, outside: false, cap: 0.22, share: 0.95 },
      { ...arc(0, 2, 0.3), outside: false, cap: 0.22, share: 0.78 },
    ];
  }
  return [];
}

/** A short stable name for a path, from what it draws: the same words in the same place are the same path. */
function pathId(seed) {
  let hash = 2166136261;
  for (let i = 0; i < seed.length; i++) hash = Math.imul(hash ^ seed.charCodeAt(i), 16777619);
  return `wp${(hash >>> 0).toString(36)}`;
}

/**
 * A text body's lines along its preset's paths, as SVG. `lines` are lists of
 * runs `{ text, size, bold, italic, color, font, noFill, outline }`; `span(run)`
 * gives a run's own attributes, the caller's, so a run looks as it does
 * straight. Each line is set at the size that fills its share of its path,
 * as Office's WordArt fills its Transform, no taller than the preset allows
 * in the box; round a circle the letters are spaced out to go all the way.
 */
export function warpedTextSvg({ preset, box, lines, span, family = 'sans-serif', color = '#1a1a1a' }) {
  // A family named in quotes ("Segoe UI") goes in an attribute that is itself in double quotes.
  const face = String(family).replace(/"/g, "'");
  const first = warpPaths(preset, box);
  if (!first.length) return '';
  // More lines than paths: the rest run on along the last path, as one line.
  const laid = first.map((_, i) => (i < first.length - 1 ? lines[i] || [] : lines.slice(i).flatMap((l, k) => (k ? [{ ...(l[0] || {}), text: ' ' }, ...l] : l))));
  const widthOf = (runs) => runs.reduce((n, r) => n + measureText(r.text, { size: r.size || 18, weight: r.bold ? 'bold' : 'normal' }), 0);
  const scaleFor = (runs, p) => {
    const width = widthOf(runs);
    if (!width) return 1;
    const tallest = (box.h * p.cap) / (runs[0].size || 18);
    // Spread round a circle, a line need only not be longer than the way round.
    const fill = (p.length * p.share) / width;
    return Math.max(0.05, p.spread ? Math.min(tallest, fill) : Math.min(fill, tallest));
  };
  const fits = first.map((p, i) => { const runs = (laid[i] || []).filter((r) => r.text); return runs.length ? scaleFor(runs, p) : 1; });
  // A curve the letters stand outside is drawn in by their height, then the
  // size worked out again on the shorter curve.
  const inset = first.map((p, i) => ((laid[i] || []).find((r) => r.text) ? (laid[i].find((r) => r.text).size || 18) * fits[i] * 0.78 : 0));
  const paths = warpPaths(preset, box, { inset });
  const out = [];
  paths.forEach((p, i) => {
    const runs = (laid[i] || []).filter((r) => r.text);
    if (!runs.length) return;
    const fit = p.outside ? Math.min(fits[i], scaleFor(runs, p)) : fits[i];
    // Held back by the box's height, a line fills its share of the path all
    // the same, its letters drawn wider, as Office's WordArt stretches them.
    const natural = widthOf(runs) * fit;
    const stretch = !p.spread && natural < p.length * p.share * 0.9;
    const id = pathId(`${preset}|${p.d}|${i}|${runs.map((r) => r.text).join('')}`);
    const size = (runs[0].size || 18) * fit;
    const spans = runs.map((r) => {
      const attrs = span ? span(r) : '';
      const sized = r.size && Math.abs(r.size * fit - size) > 0.01 ? ` font-size="${f(r.size * fit)}"` : '';
      return `<tspan${attrs ? ` ${attrs}` : ''}${sized}>${escape(r.text)}</tspan>`;
    }).join('');
    const along = p.spread || stretch
      ? `startOffset="${(((1 - p.share) / 2) * 100).toFixed(1)}%" textLength="${f(p.length * p.share)}" lengthAdjust="${p.spread ? 'spacing' : 'spacingAndGlyphs'}"`
      : `startOffset="${(p.at * 100).toFixed(1)}%" text-anchor="middle"`;
    out.push(`<defs><path id="${id}" d="${p.d}" fill="none"/></defs>`
      + `<text font-size="${f(size)}" font-family="${face}" fill="${runs[0].color || color}" xml:space="preserve">`
      + `<textPath href="#${id}" ${along}>${spans}</textPath></text>`);
  });
  return out.join('');
}

const escape = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
