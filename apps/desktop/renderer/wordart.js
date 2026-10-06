// Insert → WordArt, for Presentations and Worksheets: PowerPoint's and
// Excel's gallery, as styles the engines write — a fill (or none), an
// outline round the letters, and a shadow or glow round the words — each
// shown on an "A" in the menu, as Office shows them.

import React from 'react';

/** PowerPoint's WordArt shadow: 3 pt down and to the right, 3 pt soft, 43% black. */
const SHADOW = { blurPx: 4, distPx: 4, dir: 45, color: '#000000', alpha: 0.43 };

/** The gallery: each style's name, as Office names them, and the look its words take. */
export const WORDART_STYLES = [
  { label: 'Fill: Black, Text color 1; Shadow', run: { color: '#262626', textEffects: { shadow: SHADOW } } },
  { label: 'Fill: Blue, Accent color 1; Shadow', run: { color: '#4472C4', textEffects: { shadow: SHADOW } } },
  { label: 'Fill: White; Outline: Blue, Accent color 1; Glow: Blue', run: { color: '#FFFFFF', outline: { width: 1, color: '#4472C4' }, textEffects: { glow: { radiusPt: 5, color: '#4472C4', alpha: 0.4 } } } },
  { label: 'Fill: Gold, Accent color 4; Glow: Gold', run: { color: '#FFC000', textEffects: { glow: { radiusPt: 5, color: '#FFC000', alpha: 0.4 } } } },
  { label: 'Fill: Orange, Accent color 2; Outline: White; Shadow', run: { color: '#ED7D31', outline: { width: 1, color: '#FFFFFF' }, textEffects: { shadow: SHADOW } } },
  { label: 'Fill: None; Outline: Orange, Accent color 2', run: { noFill: true, outline: { width: 1.5, color: '#ED7D31' } } },
  { label: 'Fill: Gray, Background color 2; Inner Shadow', run: { color: '#7F7F7F', textEffects: { shadow: { ...SHADOW, distPx: 2, blurPx: 2 } } } },
];

/** What a run's look draws as in CSS — for the menu's samples. */
export function wordArtCss(run, scale = 1) {
  const shadows = [];
  const fx = run.textEffects || {};
  if (fx.glow) shadows.push(`0 0 ${Math.round((fx.glow.radiusPt || 5) * scale)}px ${fx.glow.color}`);
  if (fx.shadow) {
    const rad = ((fx.shadow.dir || 0) * Math.PI) / 180;
    const d = (fx.shadow.distPx || 0) * scale * 0.5;
    shadows.push(`${(d * Math.cos(rad)).toFixed(1)}px ${(d * Math.sin(rad)).toFixed(1)}px ${((fx.shadow.blurPx || 0) * scale * 0.5).toFixed(1)}px rgba(0,0,0,${fx.shadow.alpha ?? 0.4})`);
  }
  return {
    color: run.noFill ? 'transparent' : run.color,
    WebkitTextStroke: run.outline ? `${Math.max(0.6, (run.outline.width || 1) * 0.8 * scale).toFixed(1)}px ${run.outline.color}` : undefined,
    textShadow: shadows.length ? shadows.join(', ') : undefined,
  };
}

/** One style on an "A", as the gallery shows it. */
export function WordArtSample({ run }) {
  return <span style={{ fontWeight: 700, fontSize: 19, lineHeight: 1, fontFamily: 'Calibri, Carlito, sans-serif', ...wordArtCss(run) }}>A</span>;
}

/** The menu's items: each style with its sample, `pick` given the style. */
export const wordArtMenu = (pick) => WORDART_STYLES.map((s) => ({ label: s.label, preview: <WordArtSample run={s.run} />, run: () => pick(s) }));
