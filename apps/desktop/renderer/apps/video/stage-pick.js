// What a click on Video's stage picks, kept apart from the stage's window so
// a test reads it.
//
// Everything that shows at the playhead and stands somewhere on the frame is
// picked there: a clip, a still, a title, a shape, an icon, word art, a
// paragraph, a chart, a QR code. What lies over all of the frame and is seen
// through (an adjustment, the shade, a progress bar, the fades) is picked on
// the timeline, so a click on the frame finds what is under it.
//
// Two are picked on the frame though the engine gives them no box:
//
//   An end card covers the frame and hides what is under it, so once it has
//   come in a click on the frame is a click on it. It wears the frame's own
//   outline and no grips: it is not moved or resized, its words are typed in
//   the properties.
//
//   A QR code with nothing typed draws nothing (there is nothing to scan), so
//   the stage shows where it stands and says what it waits for, and it is
//   picked, moved and resized there like one that is drawn.

import { layerBounds, withLayerStateAt } from '@rutba/studio/host';

/** The kinds that stand in a box of their own on the frame, which the grips on the preview move and resize. */
export const BOXED_KINDS = ['textblock', 'chart', 'qr'];

/** The kinds of clip that are picked on the frame. */
const PICTURE = new Set(['video', 'image', 'text', 'shape', 'icon', 'wordart', ...BOXED_KINDS, 'outro']);

/** An end card comes in over half a second (the engine's paintOutro); from half of that it is what a click on the frame meets. */
export const OUTRO_SEEN_AFTER = 0.25;

/** A clip or still that fills the frame rather than standing in a box of its own. */
export const fillsFrame = (l) => (l.type === 'video' && !l.fw) || (l.type === 'image' && Boolean(l.fit));

/** What covers the whole frame and is picked on it all the same: an end card. It is never moved or resized there. */
export const coversFrame = (l) => l.type === 'outro';

/** A QR code with nothing to scan: the engine draws none, and the stage says what it waits for. */
export const emptyQr = (l) => l.type === 'qr' && l.visible === false && !String(l.data || '').trim();

const showing = (l, time) => {
  if (l.timing && !(time >= l.timing.start && time <= l.timing.end)) return false;
  if (l.type === 'outro') return l.visible !== false && (!l.timing || time - l.timing.start >= OUTRO_SEEN_AFTER);
  return l.visible !== false || emptyQr(l);
};

/** What can be picked at `time`: drawn, showing, on screen at the playhead, and not locked (an introduction's background and screen, or a locked lane). */
export function liveLayers(plan, time, isLocked = () => false) {
  return plan.layers.filter((l) => PICTURE.has(l.type) && !l.locked && !isLocked(l.id) && showing(l, time));
}

/** Where a layer stands on the frame, in frame pixels: the engine's own measure, and the whole frame for an end card. */
export function boundsOf(ctx, plan, layer) {
  if (coversFrame(layer)) return { x: 0, y: 0, w: plan.W, h: plan.H };
  return layerBounds(ctx, plan, layer);
}

/** The topmost live layer under (x, y) in frame pixels, by its turned outline. */
export function pickAt(ctx, plan, layers, time, x, y) {
  const ordered = [...layers].sort((a, b) => (b.z || 0) - (a.z || 0));
  for (const layer of ordered) {
    const b = withLayerStateAt(plan, layer, time, () => boundsOf(ctx, plan, layer));
    if (!b) continue;
    let px = x;
    let py = y;
    if (layer.rot && !coversFrame(layer)) {
      const cx = b.x + b.w / 2;
      const cy = b.y + b.h / 2;
      const a = (-layer.rot * Math.PI) / 180;
      px = cx + (x - cx) * Math.cos(a) - (y - cy) * Math.sin(a);
      py = cy + (x - cx) * Math.sin(a) + (y - cy) * Math.cos(a);
    }
    if (px >= b.x && px <= b.x + b.w && py >= b.y && py <= b.y + b.h) return layer;
  }
  return null;
}

/** The boxes of the QR codes that wait for an address at `time`, in frame pixels: [{ id, x, y, w, h }]. */
export function emptyQrBoxes(ctx, plan, time, isLocked = () => false) {
  const out = [];
  for (const l of liveLayers(plan, time, isLocked)) {
    if (!emptyQr(l)) continue;
    const b = withLayerStateAt(plan, l, time, () => boundsOf(ctx, plan, l));
    if (b) out.push({ id: l.id, x: Math.round(b.x), y: Math.round(b.y), w: Math.round(b.w), h: Math.round(b.h) });
  }
  return out;
}
