// Zoom by Ctrl and the wheel, and by Ctrl+Plus, Ctrl+Minus and Ctrl+0.
//
// Word, Excel and PowerPoint zoom the page, the sheet or the slide with Ctrl
// and the wheel, 10% a notch, and keep the point under the pointer where it
// was; the ribbon and the panes stay the size they are. Each app here keeps
// its own level (the same one its status bar's slider sets), so this never
// touches the window's zoom, which scales the ribbon too (owner, 2026-09-24).
//
// The arithmetic is plain functions, tested on their own; `useWheelZoom`
// binds them to a window: the wheel, the three keys as commands, and the
// application menu's Zoom In, Zoom Out and Actual Size.

import { useCallback, useEffect, useLayoutEffect, useRef } from 'react';
import { useCommands } from './commands.js';
import { t } from './messages.js';

/** One notch of the wheel, one press of a key: 10%. */
export const ZOOM_STEP = 0.1;

const round2 = (v) => Math.round(v * 100) / 100;
const clampTo = (v, min, max) => Math.max(min, Math.min(max, v));

/**
 * One step in or out from `level`, landing on the 10% grid as Word's do:
 * 100% goes to 110% or 90%, and 137% to 140% or 130%.
 */
export function stepZoom(level, direction, { min = 0.1, max = 5, step = ZOOM_STEP } = {}) {
  const unit = Math.round(step * 100);
  const pct = Math.round((Number(level) || 1) * 1000) / 10;
  const next = direction > 0
    ? Math.floor(pct / unit + 1e-6) * unit + unit
    : Math.ceil(pct / unit - 1e-6) * unit - unit;
  return round2(clampTo(next / 100, min, max));
}

/**
 * What one wheel event asks of the zoom.
 *
 * A mouse wheel turns in notches (a whole hundred pixels, or lines or pages):
 * each is one 10% step. A touchpad's pinch arrives as Ctrl and the wheel too,
 * in small fractional amounts; those scale smoothly, Chromium's way
 * (a pinch of `s` is a wheel of -100 ln s), and the exact level is carried
 * between events so a slow pinch is not lost to rounding.
 *
 * @param {number} level the level the app shows now
 * @param {number|null} exact the unrounded level the last pinch left, or null
 * @param {{ deltaY: number, deltaMode?: number }} wheel
 * @returns {{ level: number, exact: number }}
 */
export function wheelZoom(level, exact, { deltaY, deltaMode = 0 }, { min = 0.1, max = 5, step = ZOOM_STEP } = {}) {
  const now = Number(level) || 1;
  if (!deltaY) return { level: now, exact: exact ?? now };
  const notched = deltaMode !== 0 || (Number.isInteger(deltaY) && Math.abs(deltaY) >= 50);
  if (notched) {
    const size = deltaMode === 1 ? 3 : deltaMode === 2 ? 1 : 100;
    const notches = Math.max(1, Math.round(Math.abs(deltaY) / size));
    let next = now;
    for (let i = 0; i < notches; i++) next = stepZoom(next, deltaY < 0 ? 1 : -1, { min, max, step });
    return { level: next, exact: next };
  }
  // The carried level is only good while nothing else has moved the zoom.
  const from = exact != null && Math.abs(exact - now) < 0.006 ? exact : now;
  const scaled = clampTo(from * Math.exp(-deltaY / 100), min, max);
  return { level: round2(scaled), exact: scaled };
}

/** Where a point sits on the thing being zoomed, as fractions of its box. */
export function anchorAt(rect, x, y) {
  return {
    x,
    y,
    fx: rect && rect.width ? (x - rect.left) / rect.width : 0,
    fy: rect && rect.height ? (y - rect.top) / rect.height : 0,
  };
}

/**
 * How far, in screen pixels, the anchored point has moved once the thing is
 * drawn at its new size: scrolling by this much puts it back under the point.
 */
export function anchorShift(anchor, rect) {
  if (!anchor || !rect) return { dx: 0, dy: 0 };
  return {
    dx: rect.left + anchor.fx * rect.width - anchor.x,
    dy: rect.top + anchor.fy * rect.height - anchor.y,
  };
}

/**
 * Screen pixels per scrolled pixel, from one scroll that moved the area by
 * `moved` and the anchored point from `before` to `after` pixels away; 0
 * when that cannot be told (the area could not scroll any further).
 */
export function scrollRatio(before, after, moved) {
  if (Math.abs(moved) < 0.5) return 0;
  const per = (before - after) / moved;
  return per > 0.05 ? per : 0;
}

/**
 * Ctrl and the wheel, Ctrl+Plus (or Ctrl+=), Ctrl+Minus and Ctrl+0, and the
 * menu's Zoom In, Zoom Out and Actual Size, all on the app's own level.
 *
 * @param {object} o
 * @param {object} [o.shell] the window's bridge, for the menu's commands
 * @param {number} o.zoom the level the app shows now
 * @param {(level: number) => void} o.setZoom
 * @param {number} [o.min] @param {number} [o.max] the app's own limits
 * @param {string} o.area a selector for the scrolling area the wheel zooms over
 * @param {string} [o.target] a selector, inside the area, for the thing drawn
 *   at the level: the point under the pointer is kept on it
 * @param {number} [o.reset] the level Ctrl+0 goes to
 */
export function useWheelZoom({ shell, zoom, setZoom, min = 0.1, max = 5, area, target = null, reset = 1 }) {
  const live = useRef({});
  live.current = { zoom: Number(zoom) || 1, setZoom, min, max, area, target, reset };
  const exact = useRef(null);
  const pending = useRef(null);
  // When this page last took Ctrl and the wheel itself (see the menu below).
  const wheeled = useRef(-Infinity);

  /** Go to `level`, keeping the point (`x`, `y`) of the area `el` where it is. */
  const go = useCallback((level, el, point) => {
    const { zoom: now, setZoom: set, target: sel } = live.current;
    if (Math.abs(level - now) < 0.001) return;
    pending.current = null;
    const box = el || (live.current.area ? document.querySelector(live.current.area) : null);
    const thing = box && sel ? box.querySelector(sel) : null;
    if (box && thing) {
      const r = box.getBoundingClientRect();
      const at = point || { x: r.left + r.width / 2, y: r.top + r.height / 2 };
      pending.current = { level, box, thing, anchor: anchorAt(thing.getBoundingClientRect(), at.x, at.y) };
    }
    set?.(level);
  }, []);

  // Once the app has drawn the new level: scroll so the anchored point is
  // back where it was. Measured, then measured again: an area that is
  // itself zoomed (the sheet's grid) scrolls in its own pixels, not the
  // screen's, and the second pass finds out by how much.
  useLayoutEffect(() => {
    const p = pending.current;
    if (!p || Math.abs(p.level - (Number(zoom) || 1)) > 0.001) return;
    pending.current = null;
    if (!p.box.isConnected || !p.thing.isConnected) return;
    const { box, thing, anchor } = p;
    const first = anchorShift(anchor, thing.getBoundingClientRect());
    const from = { left: box.scrollLeft, top: box.scrollTop };
    box.scrollLeft = from.left + first.dx;
    box.scrollTop = from.top + first.dy;
    const second = anchorShift(anchor, thing.getBoundingClientRect());
    // How many screen pixels one scrolled pixel turned out to be, per axis.
    const perX = scrollRatio(first.dx, second.dx, box.scrollLeft - from.left);
    const perY = scrollRatio(first.dy, second.dy, box.scrollTop - from.top);
    if (Math.abs(second.dx) > 0.5 && perX) box.scrollLeft += second.dx / perX;
    if (Math.abs(second.dy) > 0.5 && perY) box.scrollTop += second.dy / perY;
  }, [zoom]);

  useEffect(() => {
    const onWheel = (e) => {
      if (!e.ctrlKey || e.defaultPrevented) return;
      // Never the browser's own zoom, wherever the wheel is turned.
      e.preventDefault();
      wheeled.current = performance.now();
      const { zoom: now, min: lo, max: hi, area: sel } = live.current;
      const el = sel && e.target?.closest ? e.target.closest(sel) : null;
      if (!el) return;
      const next = wheelZoom(now, exact.current, e, { min: lo, max: hi });
      exact.current = next.exact;
      go(next.level, el, { x: e.clientX, y: e.clientY });
    };
    window.addEventListener('wheel', onWheel, { passive: false });
    return () => window.removeEventListener('wheel', onWheel);
  }, [go]);

  const zoomIn = useCallback(() => { exact.current = null; const l = live.current; go(stepZoom(l.zoom, 1, l)); }, [go]);
  const zoomOut = useCallback(() => { exact.current = null; const l = live.current; go(stepZoom(l.zoom, -1, l)); }, [go]);
  const zoomReset = useCallback(() => { exact.current = null; const l = live.current; go(clampTo(l.reset, l.min, l.max)); }, [go]);

  useCommands({
    'view.zoomIn': { label: t('Zoom in'), icon: 'zoomIn', key: 'Mod+Plus', global: true, run: zoomIn },
    'view.zoomOut': { label: t('Zoom out'), icon: 'zoomOut', key: 'Mod+-', global: true, run: zoomOut },
    'view.zoomReset': { label: t('Back to 100%'), icon: 'check', key: 'Mod+0', global: true, run: zoomReset },
  }, []);

  // The application menu, and Ctrl and the wheel over something the page
  // cannot hear (a mail message's sandboxed frame), which the shell turns
  // into the same Zoom In and Zoom Out.
  useEffect(() => {
    if (!shell?.on) return undefined;
    return shell.on('app:command', ({ command, args } = {}) => {
      // The shell's Zoom In from the wheel follows a wheel this page took
      // already (Chromium asks even so, under a synthetic wheel at least):
      // one notch is one step, so that one is spent.
      if (args?.wheel && performance.now() - wheeled.current < 400) return;
      if (command === 'view.zoomIn') zoomIn();
      else if (command === 'view.zoomOut') zoomOut();
      else if (command === 'view.zoomReset') zoomReset();
    });
  }, [shell, zoomIn, zoomOut, zoomReset]);

  return { zoomIn, zoomOut, zoomReset };
}
