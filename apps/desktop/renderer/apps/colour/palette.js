// One colour palette for Word, Sheets and Slides, laid out as Office lays
// its out: Automatic (or No colour) at the top, the theme's ten colours with
// five lighter and darker steps under each, the ten standard colours, the
// colours chosen lately, and More Colours for any colour at all, by picking
// or by its hex. A split button opens it: the button's face puts on the
// colour last chosen, its arrow opens the palette, as Word's Font Color does.
//
// The palette drops from the arrow into a box of its own drawn over the
// window. It carries the menu's class, so a folded ribbon group it opened
// from stays open under it, and a press anywhere else, Escape or the window
// losing the keyboard put it away. The keyboard goes back where it was.

import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Button, Dialog, Field, Input, Icon } from '@rutba/office-ui';
import { t } from '../colour-words.js';
import { slotName, standardName, highlightName, words } from './names.js';
import { themeGrid, STANDARD, HIGHLIGHT, hexOf, hsvOf, hexFromHsv, rgbOf, hexFromRgb, remember, sameColour } from './grid.js';

/* ── what is remembered ─────────────────────────────────────────────────── */

// The colours chosen lately, shared by every palette in the window and kept
// for the next time it opens; the colour each split button last put on.
const RECENT_KEY = 'rutba.office.recentColours';
let recent = (() => {
  try {
    const list = JSON.parse(globalThis.localStorage?.getItem(RECENT_KEY) || '[]');
    return Array.isArray(list) ? list.map(hexOf).filter(Boolean).slice(0, 10) : [];
  } catch {
    return [];
  }
})();
const keepRecent = (hex) => {
  recent = remember(recent, hex);
  try { globalThis.localStorage?.setItem(RECENT_KEY, JSON.stringify(recent)); } catch { /* a window without storage keeps them until it closes */ }
};
const last = new Map();
const listeners = new Set();
const setLast = (kind, value) => {
  last.set(kind, value);
  for (const fn of listeners) fn();
};

/** For the tests: forget the recent colours and the last ones put on. */
export function forgetColours() {
  recent = [];
  last.clear();
  try { globalThis.localStorage?.removeItem(RECENT_KEY); } catch { /* nothing kept */ }
}

/* ── names ──────────────────────────────────────────────────────────────── */

const stepName = (base, step) => (step.lighter
  ? t('{colour}, lighter {percent}%', { colour: base, percent: step.lighter })
  : t('{colour}, darker {percent}%', { colour: base, percent: step.darker }));

/** What a colour is called in a tip: its palette name where it has one, else its hex. */
export function colourName(value, { theme = null, highlight = false } = {}) {
  if (value === null || value === undefined) return null;
  if (highlight) {
    const h = HIGHLIGHT.find(([k]) => k === value);
    return h ? highlightName(h[0]) : String(value);
  }
  const hex = hexOf(value);
  if (!hex) return String(value);
  const std = STANDARD.find(([h]) => h === hex);
  if (std) return standardName(std[1]);
  for (const col of themeGrid(theme)) {
    if (col.hex === hex) return slotName(col.slot);
    const s = col.steps.find((x) => x.hex === hex);
    if (s) return stepName(slotName(col.slot), s);
  }
  return `#${hex}`;
}

/* ── the palette ────────────────────────────────────────────────────────── */

/**
 * The palette's box. `highlight` lays out Word's fifteen highlight colours
 * instead (the value is the highlight's name); otherwise a value is a bare
 * six-digit hex. `empty` is the top row's `{ label, value }` (Automatic, No
 * fill), or none.
 */
export function ColourPalette({ x, y, rtl = false, theme = null, current, empty = null, highlight = false, more = true, keyboard = false, onPick, onMore, onClose, label }) {
  const ref = useRef(null);
  const [pos, setPos] = useState({ x, y });
  const [from] = useState(() => (typeof document === 'undefined' ? null : document.activeElement));

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const left = rtl ? x - r.width : x;
    setPos({ x: Math.max(8, Math.min(left, window.innerWidth - r.width - 8)), y: Math.max(8, Math.min(y, window.innerHeight - r.height - 8)) });
  }, [x, y, rtl]);

  useEffect(() => {
    const box = ref.current;
    if (keyboard) (box?.querySelector('button.on') || box?.querySelector('button'))?.focus({ preventScroll: true });
    const off = () => onClose?.();
    const key = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); onClose?.(); } };
    window.addEventListener('mousedown', off);
    window.addEventListener('blur', off);
    window.addEventListener('resize', off);
    window.addEventListener('keydown', key, true);
    return () => {
      window.removeEventListener('mousedown', off);
      window.removeEventListener('blur', off);
      window.removeEventListener('resize', off);
      window.removeEventListener('keydown', key, true);
      const now = document.activeElement;
      if ((!now || now === document.body || box?.contains(now)) && from && from.isConnected && from !== document.body) from.focus({ preventScroll: true });
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // The arrows move between the squares by where they are drawn: across a
  // row, and up and down a column, from one section to the next.
  const keys = (e) => {
    const all = [...(ref.current?.querySelectorAll('button:not(:disabled)') || [])];
    const at = all.indexOf(document.activeElement);
    if (at < 0) return;
    if (e.key === 'Tab') { e.preventDefault(); all[(at + (e.shiftKey ? -1 : 1) + all.length) % all.length].focus(); return; }
    const dirs = { ArrowLeft: [rtl ? 1 : -1, 0], ArrowRight: [rtl ? -1 : 1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
    const d = dirs[e.key];
    if (!d) return;
    e.preventDefault();
    e.stopPropagation();
    const box = (b) => { const r = b.getBoundingClientRect(); return { cx: r.left + r.width / 2, cy: r.top + r.height / 2, r }; };
    const me = box(all[at]);
    let best = null;
    let bestScore = Infinity;
    for (const b of all) {
      if (b === all[at]) continue;
      const o = box(b);
      const dx = o.cx - me.cx;
      const dy = o.cy - me.cy;
      if (d[0] && (Math.sign(dx) !== d[0] || Math.abs(dy) > Math.max(me.r.height, o.r.height) / 2)) continue;
      if (d[1] && (Math.sign(dy) !== d[1] || Math.abs(dy) < 2)) continue;
      // Across: the nearest in the row. Up and down: the nearest row, then the nearest column in it.
      const score = d[0] ? Math.abs(dx) : Math.abs(dy) * 1000 + Math.abs(dx);
      if (score < bestScore) { bestScore = score; best = b; }
    }
    best?.focus();
  };

  // A press on a square keeps the keyboard where it is, as the ribbon's own
  // buttons do: a cell being typed in stays in its editor with its chosen
  // words still chosen, rather than putting its words in before the colour
  // arrives. Kept from the window too, which would put the palette away.
  const keepFocus = (e) => {
    e.stopPropagation();
    if (e.target instanceof Element && e.target.closest('button')) e.preventDefault();
  };
  const pick = (value) => { onPick?.(value); onClose?.(); };
  const cell = (value, hex, name, extra = '') => {
    const on = highlight ? current === value : sameColour(current, hex);
    return (
      <button key={`${extra}${value}`} type="button" className={`rw-cell${on ? ' on' : ''}${extra ? ` ${extra}` : ''}`} style={{ background: `#${hex}` }}
        data-colour={hex} data-value={value} data-tip={name} aria-label={name} aria-pressed={on} onClick={() => pick(value)} />
    );
  };

  const grid = themeGrid(theme);
  return (
    <div ref={ref} className={`rw-menu rw-colours${highlight ? ' rw-colours-hl' : ''}`} role="dialog" aria-label={label} style={{ left: pos.x, top: pos.y }}
      onMouseDown={keepFocus} onKeyDown={keys}>
      {empty ? (
        <button type="button" className={`rw-colours-row${current === null || current === undefined ? ' on' : ''}`} data-value="" onClick={() => pick(empty.value ?? null)}>
          {empty.swatch ? <span className="rw-colours-chip" style={{ background: empty.swatch }} /> : <Icon name="close" />}
          <span className="rw-menu-label">{empty.label}</span>
        </button>
      ) : null}
      {highlight ? (
        <div className="rw-colours-grid rw-colours-five">{HIGHLIGHT.map(([k, hex]) => cell(k, hex, highlightName(k)))}</div>
      ) : (
        <>
          <div className="rw-menu-head" role="presentation">{words().themeColours}</div>
          <div className="rw-colours-grid rw-colours-theme">{grid.map((c) => cell(c.hex, c.hex, slotName(c.slot), 'base'))}</div>
          <div className="rw-colours-grid rw-colours-steps">
            {[0, 1, 2, 3, 4].flatMap((i) => grid.map((c) => cell(c.steps[i].hex, c.steps[i].hex, stepName(slotName(c.slot), c.steps[i]), `step${i}`)))}
          </div>
          <div className="rw-menu-head" role="presentation">{t('Standard colours')}</div>
          <div className="rw-colours-grid">{STANDARD.map(([hex, k]) => cell(hex, hex, standardName(k), 'std'))}</div>
          {recent.length ? (
            <>
              <div className="rw-menu-head" role="presentation">{t('Recent colours')}</div>
              <div className="rw-colours-grid rw-colours-recent">{recent.map((hex) => cell(hex, hex, colourName(hex, { theme }), 'recent'))}</div>
            </>
          ) : null}
          {more ? (
            <>
              <hr />
              <button type="button" className="rw-colours-row rw-colours-more" onClick={() => { onClose?.(); onMore?.(); }}>
                <Icon name="contrast" />
                <span className="rw-menu-label">{t('More colours…')}</span>
              </button>
            </>
          ) : null}
        </>
      )}
    </div>
  );
}

/* ── More Colours ───────────────────────────────────────────────────────── */

/**
 * Any colour: a square of saturation and brightness over a strip of hues, or
 * its hex, or its red, green and blue, with the new colour shown beside the
 * one it replaces. OK answers a bare six-digit hex.
 */
export function MoreColoursDialog({ initial = null, onOk, onClose, title }) {
  const start = hexOf(initial) || 'FF0000';
  const [hsv, setHsv] = useState(() => hsvOf(start));
  const [hex, setHex] = useState(start);
  const [typed, setTyped] = useState(start);
  const square = useRef(null);
  const strip = useRef(null);

  const fromHsv = (next) => { setHsv(next); const h = hexFromHsv(next); setHex(h); setTyped(h); };
  const fromHex = (value) => {
    setTyped(value);
    const h = hexOf(value);
    if (h && h.length === 6) { setHex(h); const [hh, s, v] = hsvOf(h); setHsv(([oldH]) => [s === 0 ? oldH : hh, s, v]); }
  };
  const drag = (el, move) => (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    const at = (ev) => { const r = el.current.getBoundingClientRect(); move(Math.max(0, Math.min(1, (ev.clientX - r.left) / r.width)), Math.max(0, Math.min(1, (ev.clientY - r.top) / r.height))); };
    at(e);
    const up = () => { window.removeEventListener('pointermove', at); window.removeEventListener('pointerup', up); };
    window.addEventListener('pointermove', at);
    window.addEventListener('pointerup', up);
  };
  const nudge = (keyMove) => (e) => {
    const step = e.shiftKey ? 0.1 : 0.01;
    const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
    if (!d) return;
    e.preventDefault();
    keyMove(d);
  };
  const rgb = rgbOf(hex);
  const setChannel = (i, v) => {
    const n = Math.max(0, Math.min(255, Math.round(Number(v) || 0)));
    const next = rgb.slice();
    next[i] = n;
    fromHex(hexFromRgb(next));
  };
  const ok = () => { const h = hexOf(typed) || hex; onOk?.(h); };
  const pure = hexFromHsv([hsv[0], 1, 1]);

  return (
    <Dialog title={title || words().colours} width={420} onClose={onClose} className="rw-colours-dialog"
      actions={<><Button label={words().cancel} onClick={onClose} /><Button primary label={words().ok} className="rw-colours-ok" disabled={!hexOf(typed)} onClick={ok} /></>}>
      <div className="rw-colours-custom">
        <div className="rw-colours-pickers">
          <div ref={square} className="rw-colours-sv" role="slider" tabIndex={0} aria-label={t('Saturation and brightness')} aria-valuetext={`#${hex}`}
            style={{ background: `linear-gradient(to top, #000, transparent), linear-gradient(to right, #fff, #${pure})` }}
            onPointerDown={drag(square, (sx, sy) => fromHsv([hsv[0], sx, 1 - sy]))}
            onKeyDown={nudge(([dx, dy]) => fromHsv([hsv[0], Math.max(0, Math.min(1, hsv[1] + dx)), Math.max(0, Math.min(1, hsv[2] - dy))]))}>
            <span className="rw-colours-knob" style={{ left: `${hsv[1] * 100}%`, top: `${(1 - hsv[2]) * 100}%` }} />
          </div>
          <div ref={strip} className="rw-colours-hue" role="slider" tabIndex={0} aria-label={words().hue} aria-valuemin={0} aria-valuemax={360} aria-valuenow={Math.round(hsv[0] * 360)}
            onPointerDown={drag(strip, (sx, sy) => fromHsv([Math.min(0.999, sy), hsv[1], hsv[2]]))}
            onKeyDown={nudge(([, dy]) => fromHsv([Math.max(0, Math.min(0.999, hsv[0] + (dy || 0))), hsv[1], hsv[2]]))}>
            <span className="rw-colours-knob rw-colours-hue-knob" style={{ top: `${hsv[0] * 100}%` }} />
          </div>
        </div>
        <div className="rw-colours-fields">
          <Field label={t('Hex')}>
            <Input className="rw-input rw-colours-hex" value={`#${typed}`.replace(/^##/, '#')} maxLength={7} onChange={(e) => fromHex(e.target.value.replace(/^#/, ''))}
              onKeyDown={(e) => { if (e.key === 'Enter' && hexOf(typed)) { e.preventDefault(); ok(); } }} />
          </Field>
          <div className="rw-colours-rgb">
            {[words().red, words().green, words().blue].map((name, i) => (
              <Field key={i} label={name}>
                <Input type="number" min={0} max={255} value={rgb[i]} onChange={(e) => setChannel(i, e.target.value)} />
              </Field>
            ))}
          </div>
          <div className="rw-colours-compare" aria-hidden="true">
            <span style={{ background: `#${hex}` }}>{words().newer}</span>
            <span style={{ background: `#${start}` }}>{t('Current')}</span>
          </div>
        </div>
      </div>
    </Dialog>
  );
}

/* ── the split button ───────────────────────────────────────────────────── */

/**
 * The ribbon's colour button, in two parts as Office's: the face puts on the
 * colour last chosen (drawn in the bar under its picture), the arrow opens
 * the palette. `kind` names what it colours, so each button remembers its own
 * last colour ('font', 'highlight', 'fill'…); `first` is the colour it starts
 * with. `onPick(value)` gets a bare hex, the highlight's name, or the empty
 * row's value.
 */
export function ColourSplit({ kind, icon, name, theme = null, current, empty = null, highlight = false, first = 'FF0000', disabled = false, pressed, onPick, className = '' }) {
  const [, redraw] = useState(0);
  const [open, setOpen] = useState(null);
  const [asking, setAsking] = useState(false);
  useEffect(() => {
    const fn = () => redraw((n) => n + 1);
    listeners.add(fn);
    return () => listeners.delete(fn);
  }, []);
  const value = last.has(kind) ? last.get(kind) : first;
  const shown = highlight ? HIGHLIGHT.find(([k]) => k === value)?.[1] : hexOf(value);
  const named = value === null ? empty?.label : colourName(value, { theme, highlight });
  const put = (v) => {
    setLast(kind, v);
    if (!highlight && v) keepRecent(v);
    onPick?.(v);
  };
  const drop = (e) => {
    const r = e.currentTarget.closest('.rw-split')?.getBoundingClientRect() || e.currentTarget.getBoundingClientRect();
    const rtl = getComputedStyle(e.currentTarget).direction === 'rtl';
    setOpen({ x: rtl ? r.right : r.left, y: r.bottom + 2, rtl, keyboard: e.detail === 0 });
  };
  return (
    <span className={`rw-split${className ? ` ${className}` : ''}`} data-kind={kind}>
      <Button icon={icon} className="rw-split-face" disabled={disabled} pressed={pressed}
        title={t('{name}: apply {colour}', { name, colour: named || words().automatic })} onClick={() => put(value)}>
        <span className="rw-split-bar" aria-hidden="true" style={{ background: shown ? `#${shown}` : 'transparent', borderColor: shown ? 'transparent' : undefined }} />
      </Button>
      <Button className="rw-split-arrow" disabled={disabled} aria-haspopup="dialog" aria-expanded={Boolean(open)}
        title={t('{name}: more colours to choose from', { name })} onClick={(e) => (open ? setOpen(null) : drop(e))}>
        <Icon name="chevronDown" size={9} strokeWidth={2.2} />
      </Button>
      {open && typeof document !== 'undefined' ? createPortal(
        <ColourPalette {...open} theme={theme} current={current} empty={empty} highlight={highlight} label={name}
          onPick={put} onMore={() => setAsking(true)} onClose={() => setOpen(null)} />,
        document.body,
      ) : null}
      {asking ? (
        <MoreColoursDialog initial={hexOf(current) || hexOf(value)} title={name} onClose={() => setAsking(false)} onOk={(hex) => { setAsking(false); put(hex); }} />
      ) : null}
    </span>
  );
}

/* ── its look ───────────────────────────────────────────────────────────── */

export const COLOUR_CSS = `
.rw-split { display: inline-flex; align-items: stretch; border-radius: var(--r-1); }
.rw-split .rw-btn.rw-split-face { position: relative; border-start-end-radius: 0; border-end-end-radius: 0; padding-inline-end: 3px; }
.rw-split .rw-btn.rw-split-arrow { min-width: 0; width: 14px; padding: 0; border-start-start-radius: 0; border-end-start-radius: 0; justify-content: center; }
.rw-split:hover .rw-btn:not(:disabled) { background: var(--hover); }
.rw-split:hover .rw-btn:not(:disabled):hover { background: var(--active); }
.rw-split-bar { position: absolute; left: 50%; bottom: 3px; width: 16px; height: 4px; margin-left: -8px; border-radius: 1px; border: 1px solid var(--line-strong); box-sizing: border-box; pointer-events: none; }
.rw-split .rw-btn.rw-split-face svg { margin-bottom: 3px; }
.rw-menu.rw-colours { min-width: 0; width: max-content; padding: 6px 8px 8px; }
.rw-colours .rw-menu-head { padding: 6px 2px 4px; }
.rw-colours-grid { display: grid; grid-template-columns: repeat(10, 18px); gap: 4px; }
.rw-colours-five { grid-template-columns: repeat(5, 22px); gap: 5px; padding: 4px 2px 2px; }
.rw-colours-five .rw-cell { height: 22px !important; }
.rw-colours-theme { margin-bottom: 4px; }
.rw-colours-steps { gap: 0 4px; }
.rw-menu.rw-colours button.rw-cell { width: auto; min-height: 0; height: 18px; padding: 0; border-radius: 2px; border: 1px solid rgba(0, 0, 0, .18); gap: 0; }
.rw-menu.rw-colours button.rw-cell:hover, .rw-menu.rw-colours button.rw-cell:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; position: relative; z-index: 1; }
.rw-menu.rw-colours button.rw-cell.on { outline: 2px solid var(--ink); outline-offset: 1px; box-shadow: inset 0 0 0 1px #fff; }
.rw-menu.rw-colours button.rw-colours-row { min-height: 28px; padding: 4px 6px; margin-bottom: 2px; }
.rw-menu.rw-colours button.rw-colours-row.on { background: var(--selected); }
.rw-colours-chip { width: 16px; height: 16px; border-radius: 2px; border: 1px solid rgba(0, 0, 0, .25); flex: none; }
.rw-colours-custom { display: flex; gap: 16px; }
.rw-colours-pickers { display: flex; gap: 10px; flex: none; }
.rw-colours-sv { position: relative; width: 180px; height: 150px; border-radius: var(--r-1); cursor: crosshair; border: 1px solid var(--line); touch-action: none; }
.rw-colours-hue { position: relative; width: 16px; height: 150px; border-radius: var(--r-1); cursor: ns-resize; border: 1px solid var(--line); touch-action: none;
  background: linear-gradient(to bottom, #f00 0%, #ff0 16.66%, #0f0 33.33%, #0ff 50%, #00f 66.66%, #f0f 83.33%, #f00 100%); }
.rw-colours-sv:focus-visible, .rw-colours-hue:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.rw-colours-knob { position: absolute; width: 12px; height: 12px; margin: -6px 0 0 -6px; border-radius: 50%; border: 2px solid #fff; box-shadow: 0 0 0 1px rgba(0, 0, 0, .55); pointer-events: none; }
.rw-colours-hue-knob { left: 50%; width: 20px; height: 6px; margin: -3px 0 0 -10px; border-radius: 3px; }
.rw-colours-fields { display: flex; flex-direction: column; gap: 8px; min-width: 0; flex: 1; }
.rw-colours-rgb { display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; }
.rw-colours-rgb input { width: 100%; box-sizing: border-box; }
.rw-colours-compare { display: grid; grid-template-columns: 1fr 1fr; border: 1px solid var(--line); border-radius: var(--r-1); overflow: hidden; }
.rw-colours-compare span { height: 34px; display: flex; align-items: flex-end; padding: 2px 5px; font-size: 10.5px; color: #fff; text-shadow: 0 0 2px #000; }
`;

// The look goes on the page once, the first time any window brings the palette in.
if (typeof document !== 'undefined' && !document.getElementById('rw-colours-css')) {
  const style = document.createElement('style');
  style.id = 'rw-colours-css';
  style.textContent = COLOUR_CSS;
  document.head.appendChild(style);
}
