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

const TAU = Math.PI * 2;

/**
 * The warps: the words stretched to fill the box between a top and a bottom
 * curve, as Office's Transform → Warp gallery has them. Each curve gives the
 * height in the box (0 the top, 1 the bottom) at a place across it (0 the
 * left, 1 the right), at Office's default handle.
 */
const ENVELOPES = {
  textTriangle: { label: 'Triangle: Up', top: (u) => 0.6 * Math.abs(2 * u - 1), bottom: () => 1 },
  textTriangleInverted: { label: 'Triangle: Down', top: () => 0, bottom: (u) => 1 - 0.6 * Math.abs(2 * u - 1) },
  textChevron: { label: 'Chevron: Up', top: (u) => 0.25 * Math.abs(2 * u - 1), bottom: (u) => 0.75 + 0.25 * Math.abs(2 * u - 1) },
  textChevronInverted: { label: 'Chevron: Down', top: (u) => 0.25 * (1 - Math.abs(2 * u - 1)), bottom: (u) => 1 - 0.25 * Math.abs(2 * u - 1) },
  textCurveUp: { label: 'Curve: Up', top: (u) => 0.4 * Math.cos((Math.PI * u) / 2), bottom: (u) => 0.6 + 0.4 * Math.cos((Math.PI * u) / 2) },
  textCurveDown: { label: 'Curve: Down', top: (u) => 0.4 * Math.sin((Math.PI * u) / 2), bottom: (u) => 0.6 + 0.4 * Math.sin((Math.PI * u) / 2) },
  textWave1: { label: 'Wave: Down', top: (u) => 0.125 * (1 + Math.sin(TAU * u)), bottom: (u) => 0.75 + 0.125 * (1 + Math.sin(TAU * u)) },
  textWave2: { label: 'Wave: Up', top: (u) => 0.125 * (1 - Math.sin(TAU * u)), bottom: (u) => 0.75 + 0.125 * (1 - Math.sin(TAU * u)) },
  textDoubleWave1: { label: 'Double Wave: Down', top: (u) => 0.0625 * (1 + Math.sin(2 * TAU * u)), bottom: (u) => 0.875 + 0.0625 * (1 + Math.sin(2 * TAU * u)) },
  textWave4: { label: 'Double Wave: Up', top: (u) => 0.0625 * (1 - Math.sin(2 * TAU * u)), bottom: (u) => 0.875 + 0.0625 * (1 - Math.sin(2 * TAU * u)) },
  textInflate: { label: 'Inflate', top: (u) => 0.18 * (1 - Math.sin(Math.PI * u)), bottom: (u) => 1 - 0.18 * (1 - Math.sin(Math.PI * u)) },
  textDeflate: { label: 'Deflate', top: (u) => 0.18 * Math.sin(Math.PI * u), bottom: (u) => 1 - 0.18 * Math.sin(Math.PI * u) },
  textInflateBottom: { label: 'Inflate: Bottom', top: () => 0, bottom: (u) => 1 - 0.3 * (1 - Math.sin(Math.PI * u)) },
  textDeflateBottom: { label: 'Deflate: Bottom', top: () => 0, bottom: (u) => 1 - 0.3 * Math.sin(Math.PI * u) },
  textInflateTop: { label: 'Inflate: Top', top: (u) => 0.3 * (1 - Math.sin(Math.PI * u)), bottom: () => 1 },
  textDeflateTop: { label: 'Deflate: Top', top: (u) => 0.3 * Math.sin(Math.PI * u), bottom: () => 1 },
  textFadeRight: { label: 'Fade: Right', top: (u) => 0.33 * u, bottom: (u) => 1 - 0.33 * u },
  textFadeLeft: { label: 'Fade: Left', top: (u) => 0.33 * (1 - u), bottom: (u) => 1 - 0.33 * (1 - u) },
  textSlantUp: { label: 'Slant: Up', top: (u) => 0.3 * (1 - u), bottom: (u) => 1 - 0.3 * u },
  textSlantDown: { label: 'Slant: Down', top: (u) => 0.3 * u, bottom: (u) => 1 - 0.3 * (1 - u) },
};

/** The rest of the gallery, after the five on the ribbon: the warps. */
export const WARP_MORE = Object.entries(ENVELOPES).map(([id, e]) => ({ id, label: e.label }));

/** Whether a preset is one this drawing follows (the rest are drawn straight). */
export const drawnWarp = (preset) => ['textArchUp', 'textArchDown', 'textCircle', 'textButton'].includes(preset) || Object.hasOwn(ENVELOPES, String(preset));

/** A preset's name as the gallery gives it. */
export const warpLabel = (preset) => WARP_PRESETS.find((p) => p.id === preset)?.label || ENVELOPES[preset]?.label || null;

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
  if (ENVELOPES[preset]) return envelopeSvg(ENVELOPES[preset], { box, lines, span, face, color });
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

// Where a letter stands in its em: its top a little over the capitals, its
// foot under the descenders, as a share of the size.
const ASCENT = 0.76;
const DESCENT = 0.22;

/**
 * A warp's words: each line stretched across the box and between its share
 * of the envelope's two curves, a letter at a time. A letter is stood at its
 * place across, as tall as the curves are apart there and sheared to their
 * slope, so its uprights stay upright, as Office's warps keep them. Lines
 * share the height between the curves, one band each.
 */
function envelopeSvg(env, { box, lines, span, face, color }) {
  const laid = lines.map((l) => (l || []).filter((r) => r.text));
  const count = laid.length;
  if (!laid.some((l) => l.length)) return '';
  const BASE = 100;
  const out = [];
  laid.forEach((runs, k) => {
    // Each letter with its run and how far it goes at the base size.
    const glyphs = [];
    for (const r of runs) {
      const scale = BASE / (r.size || 18);
      for (const ch of Array.from(String(r.text))) glyphs.push({ ch, run: r, adv: measureText(ch, { size: r.size || 18, weight: r.bold ? 'bold' : 'normal' }) * scale });
    }
    const width = glyphs.reduce((n, g) => n + g.adv, 0);
    if (!width) return;
    const sx = box.w / width;
    // The band this line has between the curves, in the box's units.
    const at = (u, t) => box.y + box.h * (env.top(u) + (env.bottom(u) - env.top(u)) * t);
    const top = (u) => at(u, k / count);
    const bottom = (u) => at(u, (k + 1) / count);
    let along = 0;
    for (const g of glyphs) {
      const u = (along + g.adv / 2) / width;
      along += g.adv;
      if (!g.ch.trim()) continue;
      const yt = top(u);
      const yb = bottom(u);
      const sy = Math.max(0.01, (yb - yt) / ((ASCENT + DESCENT) * BASE));
      const foot = yt + ASCENT * BASE * sy;
      // The slope at the letter: the band's middle a little either side of it.
      const e = 0.002;
      const slope = (((top(u + e) + bottom(u + e)) - (top(u - e) + bottom(u - e))) / 2) / (2 * e * box.w);
      const cx = box.x + u * box.w;
      const attrs = span ? span(g.run) : '';
      out.push(`<text font-size="${BASE}" text-anchor="middle" transform="matrix(${sx.toFixed(4)} ${(slope * sx).toFixed(4)} 0 ${sy.toFixed(4)} ${f(cx)} ${f(foot)})"${attrs ? ` ${attrs}` : ''}>${escape(g.ch)}</text>`);
    }
  });
  return out.length ? `<g font-family="${face}" fill="${laid.flat()[0]?.color || color}" xml:space="preserve">${out.join('')}</g>` : '';
}
