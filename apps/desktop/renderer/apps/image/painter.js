// Painting on a layer's pixels or on its mask (docs/STUDIO.md 7.2, I10, I12
// and I13): a stroke of the brush, the eraser, the clone stamp or the
// healing brush laid down dab by dab as the pointer moves, and the one-shot
// fills (the bucket, the gradient, Fill and Delete inside a selection).
//
// Everything is worked in the canvas's own pixels first: the stroke's
// coverage (its dabs), the paint (a colour, or the picture copied from
// elsewhere, cut to that coverage and to the selection), and only then laid
// on the layer's bitmap through the map from the canvas to it, with the
// layer's own operation (over for paint, out for the eraser and for hiding
// part of a mask). The bitmap as it was is kept, so every frame of a stroke is
// drawn again from it, and its changed region is handed back for undo.

import { dabSprite, dabSpacing, DabWalker } from '@rutba/imaging/paint';
import { affineAt, mapPoint } from '@rutba/studio/layers';
import { makeCanvas, ctx2d, copyCanvas, clampRect, grab, selectionCanvas, editable, touched } from './bitmaps.js';
import { toolOff } from './work.js';

const sprites = new Map();

/** The brush's dab as a canvas, kept per size and hardness. */
function spriteFor(diameter, hardness) {
  const d = Math.max(1, Math.round(diameter));
  const key = `${d}|${Math.round(hardness * 100)}`;
  let c = sprites.get(key);
  if (!c) {
    const { size, data } = dabSprite(d, hardness);
    c = makeCanvas(size, size);
    const g = c.getContext('2d');
    const img = g.createImageData(size, size);
    img.data.set(data);
    g.putImageData(img, 0, 0);
    if (sprites.size > 40) sprites.clear();
    sprites.set(key, c);
  }
  return c;
}

/** The box `rect` (canvas pixels) covers once mapped onto a bitmap, padded a pixel, clamped. */
function mappedRect(map, rect, width, height) {
  const pts = [[rect.x, rect.y], [rect.x + rect.width, rect.y], [rect.x + rect.width, rect.y + rect.height], [rect.x, rect.y + rect.height]].map(([x, y]) => mapPoint(map, x, y));
  const xs = pts.map((p) => p.x);
  const ys = pts.map((p) => p.y);
  return clampRect({ x: Math.min(...xs) - 2, y: Math.min(...ys) - 2, width: Math.max(...xs) - Math.min(...xs) + 4, height: Math.max(...ys) - Math.min(...ys) + 4 }, width, height);
}

/**
 * Lay `paint` (a canvas the size of the picture, already cut to coverage and
 * selection) onto a bitmap through `map`, over its `before` pixels, within
 * `rect` (canvas pixels). `op` is 'over' or 'out'.
 */
function lay(target, before, paint, map, rect, { op = 'over', opacity = 1 } = {}) {
  const dest = target.width ? target : null;
  if (!dest) return null;
  const r = mappedRect(map, rect, dest.width, dest.height);
  if (!r) return null;
  const g = dest.getContext('2d');
  g.save();
  g.setTransform(1, 0, 0, 1, 0, 0);
  g.beginPath();
  g.rect(r.x, r.y, r.width, r.height);
  g.clip();
  g.clearRect(r.x, r.y, r.width, r.height);
  g.drawImage(before, r.x, r.y, r.width, r.height, r.x, r.y, r.width, r.height);
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const [a, b, c, d, e, f] = affineAt(map, cx, cy);
  g.setTransform(a, b, c, d, e, f);
  g.globalAlpha = Math.max(0, Math.min(1, opacity));
  g.globalCompositeOperation = op === 'out' ? 'destination-out' : 'source-over';
  g.imageSmoothingEnabled = true;
  // The whole paint, held to the region by the clip: the paint just outside
  // `rect` is from earlier frames and must come back with the rest.
  g.drawImage(paint, 0, 0);
  g.restore();
  return r;
}

/**
 * One stroke. `target` is { entry, map } (the bitmap entry and the 3 × 3 map
 * from the canvas to it); `size` the canvas's size; `tool` 'brush',
 * 'eraser', 'clone', 'heal', 'reveal' or 'hide'; `source` the picture as
 * it shows (for the clone stamp and the healing brush) with its `offset`.
 */
export class Stroke {
  constructor({ target, size, tool, colour = '#000000', diameter = 24, hardness = 0.8, opacity = 1, selection = null, source = null, offset = null, pressure = true }) {
    this.target = target;
    this.dest = editable(target.entry);
    this.before = copyCanvas(this.dest);
    this.W = size.width;
    this.H = size.height;
    this.tool = tool;
    this.colour = colour;
    this.diameter = Math.max(1, diameter);
    this.hardness = hardness;
    this.opacity = opacity;
    this.selection = selection;
    this.source = source;
    this.offset = offset || { dx: 0, dy: 0 };
    this.usePressure = pressure;
    this.coverage = makeCanvas(this.W, this.H);
    this.paint = makeCanvas(this.W, this.H);
    this.walker = new DabWalker(dabSpacing(this.diameter));
    this.box = null;
    this.dirty = null;
    this.sprite = spriteFor(this.diameter, hardness);
  }

  get op() {
    return this.tool === 'eraser' || this.tool === 'hide' ? 'out' : 'over';
  }

  /** The pointer reached `point` ({ x, y, pressure } in canvas pixels). */
  add(point) {
    const p = { x: point.x, y: point.y, pressure: this.usePressure ? point.pressure ?? 1 : 1 };
    const dabs = this.walker.add(p);
    if (!dabs.length) return;
    const g = this.coverage.getContext('2d');
    for (const dab of dabs) {
      // A pen pressed lightly draws a smaller dab, never under a quarter of the brush.
      const s = this.diameter * (0.25 + 0.75 * Math.max(0, Math.min(1, dab.pressure)));
      const half = (s * this.sprite.width) / this.diameter / 2;
      g.drawImage(this.sprite, dab.x - half, dab.y - half, half * 2, half * 2);
      const r = { x: dab.x - half - 1, y: dab.y - half - 1, width: half * 2 + 2, height: half * 2 + 2 };
      this.dirty = union(this.dirty, r);
      this.box = union(this.box, r);
    }
  }

  /** Make the paint over `rect` from the coverage: a colour, or the source moved by the offset; cut to the selection. */
  fillPaint(rect) {
    const g = this.paint.getContext('2d');
    g.save();
    // Held to the region: source-in and destination-in reach the whole canvas otherwise.
    g.beginPath();
    g.rect(rect.x, rect.y, rect.width, rect.height);
    g.clip();
    g.globalCompositeOperation = 'source-over';
    g.clearRect(rect.x, rect.y, rect.width, rect.height);
    g.drawImage(this.coverage, rect.x, rect.y, rect.width, rect.height, rect.x, rect.y, rect.width, rect.height);
    g.globalCompositeOperation = 'source-in';
    if ((this.tool === 'clone' || this.tool === 'heal') && this.source) {
      const sx = rect.x + this.offset.dx;
      const sy = rect.y + this.offset.dy;
      g.drawImage(this.source, sx, sy, rect.width, rect.height, rect.x, rect.y, rect.width, rect.height);
    } else {
      g.fillStyle = this.tool === 'reveal' || this.tool === 'hide' || this.tool === 'eraser' ? '#ffffff' : this.colour;
      g.fillRect(rect.x, rect.y, rect.width, rect.height);
    }
    if (this.selection) {
      g.globalCompositeOperation = 'destination-in';
      g.drawImage(selectionCanvas(this.selection), rect.x, rect.y, rect.width, rect.height, rect.x, rect.y, rect.width, rect.height);
    }
    g.restore();
  }

  /** Draw what has changed since the last frame onto the bitmap. */
  frame() {
    if (!this.dirty) return false;
    const r = clampRect(this.dirty, this.W, this.H);
    this.dirty = null;
    if (!r) return false;
    this.fillPaint(r);
    lay(this.dest, this.before, this.paint, this.target.map, r, { op: this.op, opacity: this.opacity });
    touched(this.target.entry);
    return true;
  }

  /**
   * The stroke finished: for the healing brush, the copied texture given the
   * light round it (in the worker). Answers the undo record { url, rect,
   * before, after } of the bitmap, or null when nothing was painted.
   */
  async finish(url) {
    this.frame();
    const box = this.box && clampRect(this.box, this.W, this.H);
    if (!box) return null;
    if (this.tool === 'heal' && this.source) await this.heal(box);
    const r = mappedRect(this.target.map, box, this.dest.width, this.dest.height);
    if (!r) return null;
    return { url, rect: r, before: grab(this.before, r), after: grab(this.dest, r) };
  }

  async heal(box) {
    const read = (canvas, x, y) => {
      const c = makeCanvas(box.width, box.height);
      c.getContext('2d').drawImage(canvas, x, y, box.width, box.height, 0, 0, box.width, box.height);
      return ctx2d(c).getImageData(0, 0, box.width, box.height).data;
    };
    const src = read(this.source, box.x + this.offset.dx, box.y + this.offset.dy);
    const dst = read(this.source, box.x, box.y);
    const cov = read(this.coverage, box.x, box.y);
    const coverage = new Uint8Array(box.width * box.height);
    for (let i = 0; i < coverage.length; i++) coverage[i] = cov[i * 4 + 3];
    const radius = Math.max(3, this.diameter * 0.6);
    const { rgba } = await toolOff({ op: 'heal', width: box.width, height: box.height, src: src.buffer, dst: dst.buffer, coverage: coverage.buffer, radius });
    const g = this.paint.getContext('2d');
    const img = g.createImageData(box.width, box.height);
    img.data.set(new Uint8ClampedArray(rgba));
    g.save();
    g.beginPath();
    g.rect(box.x, box.y, box.width, box.height);
    g.clip();
    g.clearRect(box.x, box.y, box.width, box.height);
    g.putImageData(img, box.x, box.y);
    g.globalCompositeOperation = 'destination-in';
    g.drawImage(this.coverage, box.x, box.y, box.width, box.height, box.x, box.y, box.width, box.height);
    if (this.selection) g.drawImage(selectionCanvas(this.selection), box.x, box.y, box.width, box.height, box.x, box.y, box.width, box.height);
    g.restore();
    lay(this.dest, this.before, this.paint, this.target.map, box, { op: 'over', opacity: this.opacity });
    touched(this.target.entry);
  }
}

function union(a, b) {
  if (!a) return { ...b };
  const x = Math.min(a.x, b.x);
  const y = Math.min(a.y, b.y);
  return { x, y, width: Math.max(a.x + a.width, b.x + b.width) - x, height: Math.max(a.y + a.height, b.y + b.height) - y };
}

/**
 * A one-shot paint laid on a bitmap: `paint` a canvas the size of the
 * picture (the bucket's colour, a gradient, a fill or the selection to
 * delete), `rect` where it lies. Answers the undo record, or null.
 */
export function layOnce(target, url, paint, rect, { op = 'over', opacity = 1 } = {}) {
  const dest = editable(target.entry);
  const before = copyCanvas(dest);
  const r = lay(dest, before, paint, target.map, rect, { op, opacity });
  if (!r) return null;
  touched(target.entry);
  return { url, rect: r, before: grab(before, r), after: grab(dest, r) };
}

/** A canvas the size of the picture filled with `colour` where `selection` is (all of it with none). */
export function colourPaint(size, colour, selection) {
  const c = makeCanvas(size.width, size.height);
  const g = c.getContext('2d');
  g.fillStyle = colour;
  g.fillRect(0, 0, c.width, c.height);
  if (selection) {
    g.globalCompositeOperation = 'destination-in';
    g.drawImage(selectionCanvas(selection), 0, 0);
  }
  return c;
}

/** Put an undo record's pixels back (`which`: 'before' or 'after'). */
export function restoreBits(entry, bit, which) {
  const dest = editable(entry);
  dest.getContext('2d').putImageData(bit[which], bit.rect.x, bit.rect.y);
  touched(entry);
}
