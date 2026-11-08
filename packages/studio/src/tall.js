// The tall picture (plan A11; owner, 8 Nov: "wide and tall again choice based
// enabling both sends both simultaneously"). TikTok and Instagram are made for
// 1080 × 1920, so a scene has a tall layout as well as its wide one.
//
// A scene's tall layout is made from its wide one the first time: the sources
// kept, stacked down the frame and cropped to fit, the titles and pictures
// put back over them, and it can then be changed as any layout is. A layout
// changed by hand is kept in the stack under the scene's key. A layer's box is
// in fractions of the tall frame, so it holds at either size.
//
// Nothing here touches a canvas, so the whole of it is tested in Node.

import { clampBox, clampCrop, normaliseItem } from './stack.js';

/** The tall output sizes, by the studio's size setting. */
export const TALL_SIZES = {
  720: { width: 720, height: 1280 },
  1080: { width: 1080, height: 1920 },
};

const TALL_RATIO = 9 / 16;

export const tallSize = (size) => TALL_SIZES[size] || TALL_SIZES[1080];

const r4 = (n) => Math.round(n * 10000) / 10000;

/** Whether an item is one of the pictures that get a share of the frame, not something laid over them. */
const isPicture = (it) => ['screen', 'window', 'camera', 'web'].includes(it.kind);

/** The part of a clip that fills a box of this shape, cut about its middle. `aspect` is the source's. */
export function cropToFill(aspect, crop, boxAspect) {
  const c = crop || { x: 0, y: 0, w: 1, h: 1 };
  const clip = (aspect * c.w) / c.h;
  if (Math.abs(clip - boxAspect) < 0.001) return clampCrop(c);
  if (clip > boxAspect) {
    const w = c.w * (boxAspect / clip);
    return clampCrop({ x: c.x + (c.w - w) / 2, y: c.y, w, h: c.h });
  }
  const h = c.h * (clip / boxAspect);
  return clampCrop({ x: c.x, y: c.y + (c.h - h) / 2, w: c.w, h });
}

/**
 * The tall layers made from the shown items, top first. Each picture gets a
 * band of the frame: one picture fills it; of several, a wide source (a screen,
 * a window, a page) is shown whole across the width and the others share what
 * is left, each cropped to fill its band. Titles sit across the foot and
 * pictures in the top corner, over the rest.
 */
export function makeTallLayers(items, { aspectOf = () => 16 / 9 } = {}) {
  const shown = items.filter((it) => it.on);
  const pictures = shown.filter(isPicture);
  const wideFirst = [...pictures].sort((a, b) => (b.kind === 'screen' || b.kind === 'window' || b.kind === 'web' ? 1 : 0) - (a.kind === 'screen' || a.kind === 'window' || a.kind === 'web' ? 1 : 0));
  const out = [];
  const aspect = (it) => { const a = aspectOf(it); return a > 0 ? a : 16 / 9; };
  // A band's height as a fraction of the frame, for a picture shown whole at the full width.
  const wholeHeight = (it) => Math.min(1, TALL_RATIO / ((aspect(it) * (it.crop?.w ?? 1)) / (it.crop?.h ?? 1)));
  const bands = [];
  if (wideFirst.length === 1) {
    bands.push({ it: wideFirst[0], h: 1, whole: false });
  } else if (wideFirst.length > 1) {
    const whole = wideFirst.filter((it) => it.kind !== 'camera');
    const rest = wideFirst.filter((it) => it.kind === 'camera');
    // Whole pictures may not take more than two thirds of the frame between them, when others need room.
    let wholeH = whole.reduce((n, it) => n + wholeHeight(it), 0);
    const limit = rest.length ? 2 / 3 : 1;
    const k = wholeH > limit ? limit / wholeH : 1;
    for (const it of whole) bands.push({ it, h: wholeHeight(it) * k, whole: k === 1 });
    wholeH *= k;
    const each = rest.length ? (1 - wholeH) / rest.length : 0;
    for (const it of rest) bands.push({ it, h: each, whole: false });
  }
  // Spread what is left between them, so the picture reaches the foot of the frame.
  const used = bands.reduce((n, b) => n + b.h, 0);
  const gap = bands.length > 1 ? Math.max(0, (1 - used) / (bands.length - 1)) : 0;
  let y = 0;
  const placed = new Map();
  for (const b of bands) {
    const box = clampBox({ x: 0, y, w: 1, h: Math.max(0.02, b.h) });
    const boxAspect = TALL_RATIO / box.h;
    placed.set(b.it.id, { id: b.it.id, box, crop: b.whole ? clampCrop(b.it.crop) : cropToFill(aspect(b.it), b.it.crop, boxAspect), shape: 'rect' });
    y += b.h + gap;
  }
  // Top first: overlays above the pictures, the pictures in the order they had.
  for (const it of shown) {
    if (it.kind === 'title') out.push({ id: it.id, box: clampBox({ x: 0.05, y: 0.84, w: 0.9, h: 0.09 }), crop: null, shape: it.shape });
    else if (it.kind === 'picture') {
      const a = aspect(it);
      const w = 0.24;
      const h = Math.min(0.3, (w * TALL_RATIO) / a);
      out.push({ id: it.id, box: clampBox({ x: 1 - w - 0.04, y: 0.03, w, h }), crop: clampCrop(it.crop), shape: it.shape });
    }
  }
  for (const it of shown) if (placed.has(it.id)) out.push(placed.get(it.id));
  return out;
}

/** A layout kept by hand: only well-formed layers. */
export function normaliseTallLayers(raw) {
  if (!Array.isArray(raw)) return null;
  return raw
    .filter((l) => l && typeof l === 'object' && typeof l.id === 'string')
    .map((l) => ({ id: l.id.slice(0, 60), box: clampBox(l.box), crop: clampCrop(l.crop), shape: ['rect', 'rounded', 'circle'].includes(l.shape) ? l.shape : 'rect' }));
}

/** The tall layouts kept in a stack, by scene key. */
export function normaliseTallLayouts(raw) {
  const out = {};
  if (!raw || typeof raw !== 'object') return out;
  for (const [k, v] of Object.entries(raw).slice(0, 40)) {
    const layers = normaliseTallLayers(v);
    if (layers) out[k.slice(0, 80)] = layers;
  }
  return out;
}

/** The stack with a scene's tall layout kept (null to go back to the one made from the wide). */
export function setTallLayers(stack, key, layers) {
  const tall = { ...(stack.tall || {}) };
  if (layers) tall[key] = normaliseTallLayers(layers);
  else delete tall[key];
  return { ...stack, tall };
}

/**
 * The tall layers a stack shows for the scene on air: the ones kept for that
 * scene when every shown item is among them, or ones made from what is shown.
 */
export function tallLayersFor(stack, key, opts) {
  const kept = key ? stack.tall?.[key] : null;
  const shown = stack.items.filter((it) => it.on);
  if (kept?.length && shown.every((it) => kept.some((l) => l.id === it.id))) return kept;
  return makeTallLayers(stack.items, opts);
}

/**
 * The stack the tall picture draws: the same items, the same switches, each
 * in its tall place, drawn in the tall layout's order. An item a web page
 * shows is marked `tall`, so it is drawn from the page laid out at the tall
 * size.
 */
export function tallStackOf(stack, key, opts = {}) {
  const layers = tallLayersFor(stack, key, opts);
  const byId = new Map(stack.items.map((it) => [it.id, it]));
  const shown = [];
  for (const l of layers) {
    const it = byId.get(l.id);
    if (!it) continue;
    const item = normaliseItem({ ...it, on: true, box: l.box, crop: l.crop, shape: l.shape });
    if (item) { shown.push(it.kind === 'web' ? { ...item, tall: true } : item); byId.delete(l.id); }
  }
  const hidden = stack.items.filter((it) => byId.has(it.id)).map((it) => ({ ...it, on: false }));
  return { ...stack, items: [...shown, ...hidden] };
}

/** How much of a tall layer's box a source fills, for a check: its aspect after the clip. */
export const clipAspect = (aspect, crop) => r4((aspect * (crop?.w ?? 1)) / (crop?.h ?? 1));
