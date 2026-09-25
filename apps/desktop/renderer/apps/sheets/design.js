// Worksheets: Page Layout → Themes, Colours, Fonts and Effects.
//
// The same galleries Presentation's Design tab has, over the same eleven
// themes of the suite's own, each choice drawn as a little sheet in it — a
// table with its header in the first accent, banded rows, a chart in the
// accents, "Aa" in the heading face — so a theme is picked by looking at a
// workbook in it. A pick is one operation on the workbook (one undo step)
// that rewrites its theme part; theme-coloured cells, tables, charts and the
// default font follow.

import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { Icon } from '@rutba/office-ui';
import { THEMES, PALETTES, FONT_PAIRS, EFFECT_PRESETS } from '@rutba/office-formats/themes';

const hex = (v) => `#${String(v || '000000').replace('#', '')}`;
/** A face and the kind of face to fall back to, for the window's own previews. */
const stack = (font) => `"${font}", ${/georgia|times|palatino|cambria|constantia|garamond|book antiqua/i.test(font) ? 'Georgia, serif' : '"Segoe UI", system-ui, sans-serif'}`;

/** Mix a colour towards white: `t` 0 is the colour, 1 is white. */
function tint(c, t) {
  const n = parseInt(String(c).replace('#', ''), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(v + (255 - v) * t));
  return `rgb(${ch.join(',')})`;
}

/** A little sheet in a theme's colours and faces: a titled table, banded rows, a chart. */
export function SheetPreview({ colors, fonts, width = 176 }) {
  const c = (k) => hex(colors?.[k]);
  const bars = [30, 44, 22, 37];
  return (
    <svg viewBox="0 0 176 110" width={width} height={Math.round((width * 110) / 176)} className="sh-dg-sheet" aria-hidden="true">
      <rect x="0" y="0" width="176" height="110" fill={c('lt1')} />
      <text x="10" y="24" fontSize="17" fontFamily={stack(fonts?.major || 'Segoe UI')} fill={c('dk2')}>Aa</text>
      <text x="40" y="23" fontSize="8" fontFamily={stack(fonts?.minor || 'Segoe UI')} fill={c('dk1')} opacity=".7">Quarterly sales</text>
      <rect x="10" y="33" width="86" height="11" fill={c('accent1')} rx="1" />
      <text x="14" y="41" fontSize="7" fontWeight="600" fontFamily={stack(fonts?.minor || 'Segoe UI')} fill="#fff">Region</text>
      <text x="92" y="41" fontSize="7" fontWeight="600" fontFamily={stack(fonts?.minor || 'Segoe UI')} fill="#fff" textAnchor="end">Units</text>
      {[0, 1, 2, 3].map((i) => (
        <g key={i}>
          <rect x="10" y={44 + i * 10} width="86" height="10" fill={i % 2 ? tint(c('accent1'), 0.82) : c('lt1')} />
          <rect x="14" y={47.5 + i * 10} width={[26, 20, 30, 18][i]} height="3" rx="1.5" fill={c('dk1')} opacity=".55" />
          <rect x={92 - [14, 10, 16, 12][i]} y={47.5 + i * 10} width={[14, 10, 16, 12][i]} height="3" rx="1.5" fill={c('dk1')} opacity=".55" />
        </g>
      ))}
      <line x1="10" y1="84" x2="96" y2="84" stroke={c('accent1')} strokeWidth="1" />
      {bars.map((h, i) => <rect key={i} x={108 + i * 15} y={88 - h} width="11" height={h} rx="1.5" fill={c(`accent${i + 1}`)} />)}
      <line x1="104" y1="88.5" x2="168" y2="88.5" stroke={c('dk1')} strokeOpacity=".25" strokeWidth="1" />
      <rect x="10" y="96" width="56" height="4" rx="2" fill={c('accent5')} opacity=".9" />
      <rect x="70" y="96" width="30" height="4" rx="2" fill={c('accent6')} opacity=".9" />
    </svg>
  );
}

/** Effects' own picture: three shapes in the first three accents, subtle to intense, as the scheme styles them. */
export function EffectsPreview({ id, colors }) {
  const c = (k) => hex(colors?.[k]);
  const shape = (n, x, y) => {
    const col = c(`accent${n}`);
    const g = `fx-${id}-${n}`;
    const grad = id === 'soft' || id === 'lifted';
    const shadow = (id === 'lifted' && n > 1) || (id === 'soft' && n === 3);
    const outline = id === 'outline';
    const fill = outline ? tint(col, n === 1 ? 0 : n === 2 ? 0.4 : 0.65) : grad && n > 1 ? `url(#${g})` : col;
    const stroke = outline ? col : 'none';
    return (
      <g key={n}>
        <defs>
          <linearGradient id={g} x1="0" y1="0" x2="0" y2="1">
            <stop offset="0" stopColor={tint(col, 0.35)} />
            <stop offset="1" stopColor={col} />
          </linearGradient>
          <filter id={`${g}-s`} x="-20%" y="-20%" width="150%" height="170%"><feDropShadow dx="0" dy="3" stdDeviation={id === 'lifted' ? 3 : 2.2} floodOpacity={id === 'lifted' ? 0.35 : 0.28} /></filter>
        </defs>
        <rect x={x} y={y} width="44" height="30" rx="6" fill={fill} stroke={stroke} strokeWidth={outline ? n + 0.5 : 0} filter={shadow ? `url(#${g}-s)` : undefined} />
      </g>
    );
  };
  return (
    <svg viewBox="0 0 176 84" width="176" height="84" className="sh-dg-sheet" aria-hidden="true">
      <rect x="0" y="0" width="176" height="84" fill="#fff" />
      {shape(1, 10, 12)}{shape(2, 66, 28)}{shape(3, 122, 44)}
    </svg>
  );
}

function Accents({ colors }) {
  return (
    <span className="sl-accents" aria-hidden="true">
      {[1, 2, 3, 4, 5, 6].map((n) => <span key={n} style={{ background: hex(colors?.[`accent${n}`]) }} />)}
    </span>
  );
}

function Swatches({ colors }) {
  return (
    <span className="sl-dg-swatches">
      {['dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((k) => (
        <span key={k} className="sl-dg-swatch" style={{ width: 16, height: 16, background: hex(colors?.[k]) }} />
      ))}
    </span>
  );
}

/**
 * A popover gallery under the ribbon button: the suite's themes (or
 * palettes, font pairs, format schemes), the workbook's own ringed. Escape
 * or a click elsewhere closes it; a pick closes it and hands the choice on.
 */
export function WorkbookGallery({ kind, anchor, design, onPick, onCustomise, onClose }) {
  const ref = useRef(null);
  const [pos, setPos] = useState({ left: anchor?.left ?? 100, top: anchor?.bottom ?? 100 });
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onClose(); };
    const onDown = (e) => { if (ref.current && !ref.current.contains(e.target)) onClose(); };
    window.addEventListener('keydown', onKey);
    window.addEventListener('mousedown', onDown, true);
    return () => {
      window.removeEventListener('keydown', onKey);
      window.removeEventListener('mousedown', onDown, true);
    };
  }, [onClose]);
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const left = Math.max(8, Math.min(anchor?.left ?? 100, window.innerWidth - r.width - 8));
    const top = Math.min(anchor?.bottom ?? 100, Math.max(8, window.innerHeight - r.height - 8));
    if (left !== pos.left || top !== pos.top) setPos({ left, top });
  });
  const d = design || {};
  const same = (a, b) => ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].every((k) => String(a?.[k] || '').toUpperCase() === String(b?.[k] || '').toUpperCase());
  const title = { themes: 'Themes', colours: 'Colours', fonts: 'Fonts', effects: 'Effects' }[kind];
  return (
    <div className={`sl-dg sl-dg-${kind} sh-dg`} ref={ref} style={{ left: pos.left, top: pos.top }} role="dialog" aria-label={title}>
      <div className="sl-dg-head">
        <span>{title}</span>
        {kind === 'themes' ? <span className="sl-dg-sub">This workbook: {d.name || 'Office Theme'}</span> : null}
        {kind === 'colours' ? <span className="sl-dg-sub">Now: {d.colorName || 'Office'}</span> : null}
        {kind === 'fonts' && d.fonts ? <span className="sl-dg-sub">Now: {d.fonts.major} / {d.fonts.minor}</span> : null}
      </div>
      {kind === 'themes' ? (
        <div className="sl-dg-grid">
          {THEMES.map((t) => {
            const current = d.builtIn === t.id;
            return (
              <button key={t.id} type="button" className={`sl-dg-card${current ? ' current' : ''}`} data-theme={t.id}
                data-tip={`${t.name} — ${t.fonts.major} and ${t.fonts.minor}`} onClick={() => onPick({ theme: t.id })}>
                <span className="sl-dg-pic"><SheetPreview colors={t.palette} fonts={t.fonts} /></span>
                <span className="sl-dg-name">{current ? <Icon name="check" size={12} /> : null}<span className="sl-dg-label">{t.name}</span><Accents colors={t.palette} /></span>
              </button>
            );
          })}
        </div>
      ) : null}
      {kind === 'colours' ? (
        <div className="sl-dg-list">
          {PALETTES.map((p) => {
            const current = same(p.colors, d.colors);
            return (
              <button key={p.id} type="button" className={`sl-dg-row${current ? ' current' : ''}`} data-palette={p.id} onClick={() => onPick({ colors: p.id })}>
                <Swatches colors={p.colors} />
                <span className="sl-dg-rowname">{p.name}</span>
                {current ? <Icon name="check" size={13} /> : null}
              </button>
            );
          })}
        </div>
      ) : null}
      {kind === 'fonts' ? (
        <div className="sl-dg-list">
          {FONT_PAIRS.map((p) => {
            const current = d.fonts?.major === p.major && d.fonts?.minor === p.minor;
            return (
              <button key={p.id} type="button" className={`sl-dg-row sl-dg-fontrow${current ? ' current' : ''}`} data-pair={p.id} onClick={() => onPick({ fonts: p.id })}>
                <span className="sl-dg-aa" style={{ fontFamily: stack(p.major) }}>Aa</span>
                <span className="sl-dg-faces">
                  <span className="sl-dg-rowname">{p.name}</span>
                  <span className="sl-dg-major" style={{ fontFamily: stack(p.major) }}>{p.major}</span>
                  <span className="sl-dg-minor" style={{ fontFamily: stack(p.minor) }}>{p.minor}</span>
                </span>
                {current ? <Icon name="check" size={13} /> : null}
              </button>
            );
          })}
        </div>
      ) : null}
      {kind === 'effects' ? (
        <div className="sl-dg-grid sl-dg-effects">
          {EFFECT_PRESETS.map((p) => {
            const current = d.effects === p.id;
            return (
              <button key={p.id} type="button" className={`sl-dg-card${current ? ' current' : ''}`} data-effects={p.id} data-tip={`${p.name} — ${p.description}`} onClick={() => onPick({ effects: p.id })}>
                <span className="sl-dg-pic"><EffectsPreview id={p.id} colors={d.colors} /></span>
                <span className="sl-dg-name">{current ? <Icon name="check" size={12} /> : null}{p.name}</span>
              </button>
            );
          })}
        </div>
      ) : null}
      {kind === 'colours' || kind === 'fonts' ? (
        <div className="sl-dg-foot">
          <button type="button" className="sl-dg-custom" onClick={() => onCustomise(kind)}>
            <Icon name="settings" size={14} /> {kind === 'colours' ? 'Customise Colours…' : 'Customise Fonts…'}
          </button>
        </div>
      ) : null}
      {kind === 'effects' ? <div className="sl-dg-foot sl-dg-note">Effects reach shapes styled from the theme, as Excel's own shape styles are.</div> : null}
    </div>
  );
}

export const SHEET_DESIGN_CSS = `
.sh-dg.sl-dg-themes { width: 836px; }
.sh-dg .sl-dg-pic { background: #fff; }
.sh-dg .sh-dg-sheet { display: block; width: 100%; height: auto; }
`;
