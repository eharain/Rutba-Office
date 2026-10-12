// Painting on pixels (the Image app's advanced editor, docs/STUDIO.md 7.2,
// I12 and I13): the brush's dab and how a stroke is laid out in dabs, the
// clone stamp's offset, the bucket fill and the gradient, and the healing
// brush's patch blend. A painted pixel layer or a layer's mask is an RGBA or
// alpha bitmap; these work on the bytes.
//
// The window lays a brush stroke down as dabs drawn on a canvas, live, from
// `dabSprite` and `DabWalker`; the fill, the gradient and the heal are worked
// here, in the image worker, on the whole region at once.
//
// Pure arithmetic over arrays: no DOM, no canvas. Runs in the worker and in
// Node for the tests.

import { blurPlane } from './select.js';

const clamp = (v, lo, hi) => (v < lo ? lo : v > hi ? hi : v);

/**
 * How much paint a round brush lays at `d`, the distance from its centre as
 * a fraction of its radius: all of it inside `hardness` of the radius, then
 * a smooth fall to none at the edge. Hardness 1 is a hard-edged disc (with
 * one pixel of smoothing, which the sprite gives).
 */
export function brushAlpha(d, hardness = 0.8) {
  if (d >= 1) return 0;
  const h = clamp(Number(hardness), 0, 0.999);
  if (d <= h) return 1;
  const t = (d - h) / (1 - h);
  return 1 - t * t * (3 - 2 * t);
}

/**
 * The brush as a picture: a white dab `diameter` pixels across, its alpha
 * the brush's fall-off, sampled four by four in each pixel so a small brush is
 * round. Answers { size, data } (RGBA, size × size), the dab centred.
 */
export function dabSprite(diameter, hardness = 0.8) {
  const dia = Math.max(1, Number(diameter) || 1);
  const size = Math.ceil(dia) + 2;
  const r = dia / 2;
  const c = size / 2;
  const data = new Uint8ClampedArray(size * size * 4);
  // A hard brush still gets a pixel of edge, or it is jagged.
  const h = Math.min(clamp(Number(hardness), 0, 1), Math.max(0, 1 - 1 / Math.max(1, r)));
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let a = 0;
      for (let j = 0; j < 4; j++) {
        for (let i = 0; i < 4; i++) {
          const dx = x + (i + 0.5) / 4 - c;
          const dy = y + (j + 0.5) / 4 - c;
          a += brushAlpha(Math.hypot(dx, dy) / r, h);
        }
      }
      const o = (y * size + x) * 4;
      data[o] = 255;
      data[o + 1] = 255;
      data[o + 2] = 255;
      data[o + 3] = Math.round((255 * a) / 16);
    }
  }
  return { size, data };
}

/** How far apart a stroke's dabs fall: an eighth of the brush, never under a pixel. */
export const dabSpacing = (diameter) => Math.max(1, Number(diameter) * 0.12);

/**
 * A stroke laid out as dabs: points come in as the pointer moves (`{ x, y,
 * pressure }`), and dabs come out evenly spaced along the path between them,
 * the pressure eased from one point to the next. The first point is a dab.
 */
export class DabWalker {
  constructor(spacing) {
    this.spacing = Math.max(0.5, Number(spacing) || 1);
    this.last = null;
    this.carry = 0;
  }

  /** The dabs to lay for the pointer reaching `point`. */
  add(point) {
    const p = { x: point.x, y: point.y, pressure: point.pressure ?? 1 };
    if (!this.last) {
      this.last = p;
      this.carry = 0;
      return [p];
    }
    const a = this.last;
    const len = Math.hypot(p.x - a.x, p.y - a.y);
    const out = [];
    let at = this.spacing - this.carry;
    while (at <= len) {
      const t = len ? at / len : 0;
      out.push({ x: a.x + (p.x - a.x) * t, y: a.y + (p.y - a.y) * t, pressure: a.pressure + (p.pressure - a.pressure) * t });
      at += this.spacing;
    }
    this.carry = len - (at - this.spacing);
    this.last = p;
    return out;
  }
}

/**
 * The clone stamp's link: Alt+click sets `source`, and the first dab of the
 * next stroke `start`; from then on every dab at a point copies from that
 * point moved by the same offset. Answers { dx, dy }.
 */
export function cloneOffset(source, start) {
  return { dx: source.x - start.x, dy: source.y - start.y };
}

/** Where a dab at `point` copies from, through an offset from `cloneOffset`. */
export function cloneSourceAt(point, offset) {
  return { x: point.x + offset.dx, y: point.y + offset.dy };
}

/**
 * Source-over: `colour` ([r, g, b, a] 0 to 255) laid on one RGBA pixel at
 * `i` with a strength of `k` (0 to 1).
 */
function over(out, i, r, g, b, a, k) {
  const sa = (a / 255) * k;
  if (sa <= 0) return;
  const da = out[i + 3] / 255;
  const oa = sa + da * (1 - sa);
  if (oa <= 0) return;
  out[i] = (r * sa + out[i] * da * (1 - sa)) / oa;
  out[i + 1] = (g * sa + out[i + 1] * da * (1 - sa)) / oa;
  out[i + 2] = (b * sa + out[i + 2] * da * (1 - sa)) / oa;
  out[i + 3] = oa * 255;
}

/**
 * The bucket: `colour` laid over every pixel of `region` (an alpha plane,
 * the magic wand's answer, already cut to any selection), at `opacity`.
 * Answers new RGBA.
 */
export function fillRegion(rgba, width, height, region, colour, { opacity = 1 } = {}) {
  const out = new Uint8ClampedArray(rgba);
  const [r, g, b, a = 255] = colour;
  for (let p = 0, n = width * height; p < n; p++) {
    const k = region ? region[p] / 255 : 1;
    if (k > 0) over(out, p * 4, r, g, b, a, k * opacity);
  }
  return out;
}

/**
 * Where a pixel at (x, y) is along a gradient dragged from `start` to
 * `end`, 0 to 1: along the line for `linear`, out from the start for
 * `radial`, and there and back for `reflected`.
 */
export function gradientAt(kind, start, end, x, y) {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const len2 = dx * dx + dy * dy;
  if (len2 < 1e-9) return 1;
  if (kind === 'radial') return clamp(Math.hypot(x - start.x, y - start.y) / Math.sqrt(len2), 0, 1);
  const t = ((x - start.x) * dx + (y - start.y) * dy) / len2;
  if (kind === 'reflected') return clamp(Math.abs(t), 0, 1);
  return clamp(t, 0, 1);
}

/**
 * A gradient from colour `from` to colour `to` ([r, g, b, a] each, `to`
 * may be clear) laid over an RGBA picture at `opacity`, only where `region`
 * (an alpha plane: a selection) lets it. Measured at each pixel's centre.
 * Answers new RGBA.
 */
export function fillGradient(rgba, width, height, { kind = 'linear', start, end, from, to, opacity = 1, region = null }) {
  const out = new Uint8ClampedArray(rgba);
  const f = [from[0], from[1], from[2], from[3] ?? 255];
  const g = [to[0], to[1], to[2], to[3] ?? 255];
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const p = y * width + x;
      const k = region ? region[p] / 255 : 1;
      if (k <= 0) continue;
      const t = gradientAt(kind, start, end, x + 0.5, y + 0.5);
      // Mixed in premultiplied alpha, so a colour fading to clear does not go grey on the way.
      const a = f[3] + (g[3] - f[3]) * t;
      if (a <= 0) continue;
      const mix = (c) => (f[c] * f[3] * (1 - t) + g[c] * g[3] * t) / a;
      over(out, p * 4, mix(0), mix(1), mix(2), a, k * opacity);
    }
  }
  return out;
}

/**
 * The healing brush's blend, over one patch: the texture of `src` (what the
 * clone stamp would put down) with the colour and light of `dst` (what is
 * there now), so a blemish is covered by skin that matches round it. The
 * low frequencies are swapped: out = src - blur(src) + blur(dst), with a
 * blur of about `radius` pixels. `coverage` (an alpha plane, the brush's
 * dabs) says how much of each pixel is healed; the rest stays `dst`. Both
 * patches are RGBA of `width` × `height`. Answers new RGBA.
 */
export function healPatch(src, dst, width, height, { radius = 8, coverage = null } = {}) {
  const n = width * height;
  const out = new Uint8ClampedArray(dst);
  const plane = (data, c) => {
    const p = new Float32Array(n);
    for (let i = 0; i < n; i++) p[i] = data[i * 4 + c];
    return p;
  };
  // The light round the place healed is read from round it, not from the
  // blemish itself: a blur weighted by what is NOT being healed (normalised
  // convolution), so a dark spot does not leave its colour in the patch.
  let weight = null;
  let weightBlur = null;
  if (coverage) {
    weight = new Float32Array(n);
    for (let i = 0; i < n; i++) weight[i] = 1 - coverage[i] / 255;
    weightBlur = blurPlane(weight, width, height, radius);
  }
  for (let c = 0; c < 3; c++) {
    const bs = blurPlane(plane(src, c), width, height, radius);
    let bd;
    if (weight) {
      const pd = plane(dst, c);
      for (let i = 0; i < n; i++) pd[i] *= weight[i];
      const num = blurPlane(pd, width, height, radius);
      const plain = blurPlane(plane(dst, c), width, height, radius);
      bd = new Float32Array(n);
      for (let i = 0; i < n; i++) bd[i] = weightBlur[i] > 0.02 ? num[i] / weightBlur[i] : plain[i];
    } else {
      bd = blurPlane(plane(dst, c), width, height, radius);
    }
    for (let i = 0; i < n; i++) {
      const k = coverage ? coverage[i] / 255 : 1;
      if (k <= 0) continue;
      const healed = clamp(src[i * 4 + c] - bs[i] + bd[i], 0, 255);
      out[i * 4 + c] = dst[i * 4 + c] * (1 - k) + healed * k;
    }
  }
  if (coverage) for (let i = 0; i < n; i++) if (coverage[i] && out[i * 4 + 3] < 255) out[i * 4 + 3] = Math.max(out[i * 4 + 3], (src[i * 4 + 3] * coverage[i]) / 255);
  return out;
}

/** A colour written `#rrggbb` as [r, g, b, 255]. */
export function hexColour(hex, alpha = 255) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || ''));
  if (!m) return [0, 0, 0, alpha];
  const v = parseInt(m[1], 16);
  return [(v >> 16) & 255, (v >> 8) & 255, v & 255, alpha];
}
