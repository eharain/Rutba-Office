// The edit pipeline.
//
// An edit is a list of operations, not a modified bitmap. That is what makes
// the image tool non-destructive: the original bytes are never touched, undo is
// dropping an operation, and "reset" is emptying the list. The pixels are
// produced only when something needs to see them — a preview, or an export.
//
// The operations are described here in a form with no DOM in it, so the same
// list can be planned and tested in Node and executed on a canvas in the
// window. `toCanvasPlan` turns the list into one matrix from the source's
// pixels to the result's, the canvas filter, the pixel settings and the
// marks drawn on top.
//
// The geometry, in the order it is applied to a source pixel:
//   1. straighten: turned by a small angle about the picture's centre, and
//      cut to the largest upright rectangle inside the turned picture;
//   2. crop: a rectangle of that, in the same (straightened) frame;
//   3. resize: scaled, by width and height;
//   4. flip, then a quarter turn: the picture as it is shown.
// A crop, a resize or a flip is given as the person sees the picture at the
// time, and is turned back into this order here, so flipping a picture that
// was turned on its side mirrors it the way it is on the screen.
//
// Marks (rectangle, ellipse, arrow, text, pen) and red-eye spots are kept in
// the SOURCE's pixels, so they stay on what they were drawn on whatever is
// cropped, turned or resized after them, as a mark in Paint does.

/** @typedef {{ op: string, [key: string]: any }} Operation */

export const OPS = {
  crop: 'crop',            // { x, y, width, height } as the picture is shown at the time
  rotate: 'rotate',        // { degrees: 90 | 180 | 270 }
  flip: 'flip',            // { axis: 'x' | 'y' } as shown: x mirrors left and right
  resize: 'resize',        // { width, height } as shown
  straighten: 'straighten',// { degrees } -45 to 45, clockwise; the last one counts
  adjust: 'adjust',        // { brightness, contrast, saturation, hue, blur, sepia, grayscale, invert }
  pixels: 'pixels',        // tone and detail, worked on the pixels (pixels.js)
  annotate: 'annotate',    // { shape: 'rect'|'ellipse'|'arrow'|'text'|'pen', ... } in source pixels
  redeye: 'redeye',        // { x, y, radius } in source pixels
};

export const DEFAULT_ADJUSTMENTS = {
  brightness: 100,
  contrast: 100,
  saturation: 100,
  hue: 0,
  blur: 0,
  sepia: 0,
  grayscale: 0,
  invert: 0,
};

/** The CSS/canvas filter string for a set of adjustments, or '' for none. */
export function filterString(adjust = {}) {
  const a = { ...DEFAULT_ADJUSTMENTS, ...adjust };
  const parts = [];
  if (a.brightness !== 100) parts.push(`brightness(${a.brightness}%)`);
  if (a.contrast !== 100) parts.push(`contrast(${a.contrast}%)`);
  if (a.saturation !== 100) parts.push(`saturate(${a.saturation}%)`);
  if (a.hue) parts.push(`hue-rotate(${a.hue}deg)`);
  if (a.blur) parts.push(`blur(${a.blur}px)`);
  if (a.sepia) parts.push(`sepia(${a.sepia}%)`);
  if (a.grayscale) parts.push(`grayscale(${a.grayscale}%)`);
  if (a.invert) parts.push(`invert(${a.invert}%)`);
  return parts.join(' ');
}

/** Pixel settings that change nothing: all zero, no levels moved, no curve bent. */
function pixelsNeutral(p) {
  if (!p) return true;
  for (const [k, v] of Object.entries(p)) {
    if (k === 'op') continue;
    if (k === 'levels') {
      if (v && ((v.inBlack || 0) !== 0 || (v.inWhite ?? 255) !== 255 || (v.gamma ?? 1) !== 1 || (v.outBlack || 0) !== 0 || (v.outWhite ?? 255) !== 255)) return false;
    } else if (k === 'curve') {
      if (v && Object.values(v).some((pts) => Array.isArray(pts) && pts.some(([x, y]) => x !== y))) return false;
    } else if (typeof v === 'number' && v !== 0) return false;
  }
  return true;
}

/** True when the list changes nothing, so Save can be offered honestly. */
export function isIdentity(operations) {
  return !operations.some((op) => {
    if (op.op === OPS.adjust) return filterString(op) !== '';
    if (op.op === OPS.pixels) return !pixelsNeutral(op);
    if (op.op === OPS.rotate) return op.degrees % 360 !== 0;
    if (op.op === OPS.straighten) return Boolean(op.degrees);
    return true;
  });
}

// ── matrices ────────────────────────────────────────────────────────────────
// [a, b, c, d, e, f], as a canvas takes them: x' = a x + c y + e, y' = b x + d y + f.

const IDENTITY = [1, 0, 0, 1, 0, 0];

/** m · n: n applied first. */
export function multiply(m, n) {
  return [
    m[0] * n[0] + m[2] * n[1],
    m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3],
    m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4],
    m[1] * n[4] + m[3] * n[5] + m[5],
  ];
}

export function invert(m) {
  const det = m[0] * m[3] - m[1] * m[2];
  if (!det) return [...IDENTITY];
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det, (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
}

export function apply(m, p) {
  return { x: m[0] * p.x + m[2] * p.y + m[4], y: m[1] * p.x + m[3] * p.y + m[5] };
}

const translate = (x, y) => [1, 0, 0, 1, x, y];
const scale = (x, y) => [x, 0, 0, y, 0, 0];
const rotation = (deg) => {
  const r = (deg * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  // Quarter turns exactly, so a picture turned four times is where it began.
  const snap = (v) => (Math.abs(v) < 1e-12 ? 0 : Math.abs(Math.abs(v) - 1) < 1e-12 ? Math.sign(v) : v);
  return [snap(c), snap(s), snap(-s), snap(c), 0, 0];
};

/**
 * The largest upright rectangle inside a w × h picture turned by `degrees`:
 * what Straighten keeps, so no corner of the frame is left empty.
 */
export function straightenRect(width, height, degrees) {
  const a = Math.abs((degrees * Math.PI) / 180);
  if (!a || width <= 0 || height <= 0) return { width, height };
  const sin = Math.abs(Math.sin(a));
  const cos = Math.abs(Math.cos(a));
  const wide = width >= height;
  const long = wide ? width : height;
  const short = wide ? height : width;
  let w;
  let h;
  if (short <= 2 * sin * cos * long || Math.abs(sin - cos) < 1e-10) {
    // Half-constrained: two corners of the rectangle touch the longer side.
    const x = 0.5 * short;
    [w, h] = wide ? [x / sin, x / cos] : [x / cos, x / sin];
  } else {
    // Fully constrained: all four corners touch the turned picture.
    const cos2 = cos * cos - sin * sin;
    w = (width * cos - height * sin) / cos2;
    h = (height * cos - width * sin) / cos2;
  }
  return { width: Math.max(1, Math.floor(w)), height: Math.max(1, Math.floor(h)) };
}

const clampRect = (r, bounds) => {
  const x = Math.max(bounds.x, Math.min(r.x, bounds.x + bounds.width - 1));
  const y = Math.max(bounds.y, Math.min(r.y, bounds.y + bounds.height - 1));
  const right = Math.min(r.x + r.width, bounds.x + bounds.width);
  const bottom = Math.min(r.y + r.height, bounds.y + bounds.height);
  return { x, y, width: Math.max(1, right - x), height: Math.max(1, bottom - y) };
};

/**
 * Turn the list into a plan a canvas can execute.
 *
 * The geometry is collapsed into one matrix from the source's pixels to the
 * result's, so the pixels are resampled once rather than once per operation,
 * which is both faster and visibly sharper.
 */
export function toCanvasPlan(source, operations) {
  const W0 = source.width;
  const H0 = source.height;
  let theta = 0;
  // The part of the (straightened) picture kept, in its own pixels.
  let rect = { x: 0, y: 0, width: W0, height: H0 };
  let sx = 1;
  let sy = 1;
  let rotate = 0;
  let flipX = false;
  let flipY = false;
  const adjust = { ...DEFAULT_ADJUSTMENTS };
  const pixels = {};
  const overlays = [];
  const spots = [];

  // The unturned size, and the part of the plan that shows it as it is shown.
  const unturned = () => ({ width: rect.width * sx, height: rect.height * sy });
  const shown = () => {
    const u = unturned();
    return rotate % 180 ? { width: u.height, height: u.width } : u;
  };
  // From the unturned, scaled picture (0..dw, 0..dh) to the result.
  const view = () => {
    const u = unturned();
    const t = shown();
    return multiply(translate(t.width / 2, t.height / 2), multiply(rotation(rotate), multiply(scale(flipX ? -1 : 1, flipY ? -1 : 1), translate(-u.width / 2, -u.height / 2))));
  };

  for (const op of operations) {
    switch (op.op) {
      case OPS.crop: {
        // Given as the picture is shown: back through the turn and the flip
        // to the unturned picture, then through the scale to its own pixels.
        const back = invert(view());
        const corners = [
          apply(back, { x: op.x, y: op.y }),
          apply(back, { x: op.x + op.width, y: op.y + op.height }),
        ];
        const ux = Math.min(corners[0].x, corners[1].x);
        const uy = Math.min(corners[0].y, corners[1].y);
        const uw = Math.abs(corners[1].x - corners[0].x);
        const uh = Math.abs(corners[1].y - corners[0].y);
        rect = clampRect({ x: rect.x + ux / sx, y: rect.y + uy / sy, width: uw / sx, height: uh / sy }, rect);
        break;
      }
      case OPS.rotate:
        rotate = (((rotate + (op.degrees || 0)) % 360) + 360) % 360;
        break;
      case OPS.flip: {
        // As shown: on a picture turned on its side, left and right on the
        // screen are the picture's own top and bottom.
        const side = rotate % 180 !== 0;
        const axis = side ? (op.axis === 'y' ? 'x' : 'y') : op.axis === 'y' ? 'y' : 'x';
        if (axis === 'y') flipY = !flipY;
        else flipX = !flipX;
        break;
      }
      case OPS.resize: {
        const side = rotate % 180 !== 0;
        const w = Math.max(1, Math.round(side ? op.height : op.width));
        const h = Math.max(1, Math.round(side ? op.width : op.height));
        sx = w / rect.width;
        sy = h / rect.height;
        break;
      }
      case OPS.straighten: {
        theta = Math.max(-45, Math.min(45, Number(op.degrees) || 0));
        const keep = straightenRect(W0, H0, theta);
        // What was kept, cut to what is left inside the turned picture; the
        // scale stays as it was, so a resized picture stays resized.
        rect = clampRect(rect, { x: (W0 - keep.width) / 2, y: (H0 - keep.height) / 2, width: keep.width, height: keep.height });
        break;
      }
      case OPS.adjust:
        Object.assign(adjust, op);
        break;
      case OPS.pixels:
        Object.assign(pixels, op);
        break;
      case OPS.annotate:
        overlays.push(op);
        break;
      case OPS.redeye:
        spots.push(op);
        break;
      default:
        break;
    }
  }
  delete adjust.op;
  delete pixels.op;

  // Source pixels → the straightened frame → kept → scaled → shown.
  const straighten = multiply(translate(W0 / 2, H0 / 2), multiply(rotation(theta), translate(-W0 / 2, -H0 / 2)));
  const matrix = multiply(view(), multiply(scale(sx, sy), multiply(translate(-rect.x, -rect.y), straighten)));
  const t = shown();
  const target = { width: Math.max(1, Math.round(t.width)), height: Math.max(1, Math.round(t.height)) };

  return {
    source: { width: W0, height: H0 },
    sourceRect: { x: rect.x, y: rect.y, width: rect.width, height: rect.height },
    straighten: theta,
    rotate,
    flipX,
    flipY,
    target,
    matrix,
    adjust,
    filter: filterString(adjust),
    pixels,
    overlays,
    spots,
  };
}

/** The size the image ends up, after crops, rotations, resizes and straightening. */
export function resultSize(source, operations) {
  return toCanvasPlan(source, operations).target;
}

/** A point in the result's pixels from one in the source's. */
export function toTarget(plan, point) {
  return apply(plan.matrix, point);
}

/** A point in the source's pixels from one in the result's. */
export function toSource(plan, point) {
  return apply(invert(plan.matrix), point);
}

/**
 * How large one source pixel is in the result, and the frame that undoes the
 * plan's turn and mirror: a mark made now is drawn upright, the way the
 * person sees the picture as they draw it, and turns with the picture after.
 */
export function markFrame(plan) {
  const [a, b, c, d] = plan.matrix;
  const k = Math.sqrt(Math.abs(a * d - b * c)) || 1;
  const inv = invert([a, b, c, d, 0, 0]);
  return { k, frame: [inv[0] * k, inv[1] * k, inv[2] * k, inv[3] * k] };
}

/**
 * A mark drawn on the result, in its pixels, as an operation kept in the
 * source's: `{ shape, x1, y1, x2, y2, points, text, colour, width, size }`.
 * A rectangle and an ellipse keep their centre, their size and the frame
 * that draws them upright; an arrow and a pen their points; a text its
 * corner, its size and the frame.
 */
export function markToOperation(plan, mark) {
  const { k, frame } = markFrame(plan);
  const at = (x, y) => {
    const p = toSource(plan, { x, y });
    return { x: round2(p.x), y: round2(p.y) };
  };
  const base = { op: OPS.annotate, shape: mark.shape, colour: mark.colour || '#e0393e', width: round2((mark.width || 4) / k) };
  if (mark.shape === 'rect' || mark.shape === 'ellipse') {
    const c = at((mark.x1 + mark.x2) / 2, (mark.y1 + mark.y2) / 2);
    return { ...base, cx: c.x, cy: c.y, w: round2(Math.abs(mark.x2 - mark.x1) / k), h: round2(Math.abs(mark.y2 - mark.y1) / k), frame: frame.map(round4), fill: mark.fill || null };
  }
  if (mark.shape === 'arrow') {
    return { ...base, points: [at(mark.x1, mark.y1), at(mark.x2, mark.y2)], head: round2(Math.max(12, (mark.width || 4) * 3.5) / k) };
  }
  if (mark.shape === 'pen') {
    return { ...base, points: (mark.points || []).map((p) => at(p.x, p.y)) };
  }
  if (mark.shape === 'text') {
    const c = at(mark.x1, mark.y1);
    return { ...base, x: c.x, y: c.y, size: round2((mark.size || 32) / k), text: String(mark.text || ''), frame: frame.map(round4) };
  }
  return base;
}

const round2 = (v) => Math.round(v * 100) / 100;
const round4 = (v) => Math.round(v * 10000) / 10000;

/**
 * Where a red-eye spot clicked on the result lies in the source, and how large.
 * `radius` is in the result's pixels.
 */
export function spotToOperation(plan, { x, y, radius }) {
  const { k } = markFrame(plan);
  const p = toSource(plan, { x, y });
  return { op: OPS.redeye, x: round2(p.x), y: round2(p.y), radius: round2(radius / k) };
}

/** A spot in the result's pixels, for the pixel pass that removes it. */
export function spotInTarget(plan, spot) {
  const { k } = markFrame(plan);
  const p = toTarget(plan, spot);
  return { x: p.x, y: p.y, radius: spot.radius * k };
}

/** Fit a box inside another, preserving the ratio — the viewer's zoom-to-fit. */
export function fit(inner, outer) {
  if (!inner.width || !inner.height) return { width: 0, height: 0, scale: 1 };
  const s = Math.min(outer.width / inner.width, outer.height / inner.height);
  return { width: Math.round(inner.width * s), height: Math.round(inner.height * s), scale: s };
}

/** A crop rectangle constrained to a ratio and kept inside the image. */
export function constrainCrop(rect, bounds, ratio) {
  let { x, y, width, height } = rect;
  if (ratio) {
    if (width / height > ratio) width = height * ratio;
    else height = width / ratio;
  }
  width = Math.max(8, Math.min(width, bounds.width));
  height = Math.max(8, Math.min(height, bounds.height));
  x = Math.max(0, Math.min(x, bounds.width - width));
  y = Math.max(0, Math.min(y, bounds.height - height));
  return { x: Math.round(x), y: Math.round(y), width: Math.round(width), height: Math.round(height) };
}

/**
 * A new size from a resize by pixels or by percent, the ratio kept when asked:
 * `{ by: 'pixels' | 'percent', width, height, keep }` against the size now.
 * Changing one side with the ratio kept works the other out.
 */
export function resizeTo(current, { by = 'pixels', width, height, keep = true, changed = 'width' }) {
  const ratio = current.width / current.height;
  let w;
  let h;
  if (by === 'percent') {
    const pw = Number(width) || 100;
    const ph = keep ? pw : Number(height) || 100;
    w = (current.width * pw) / 100;
    h = (current.height * ph) / 100;
  } else {
    w = Number(width) || current.width;
    h = Number(height) || current.height;
    if (keep) {
      if (changed === 'height') w = h * ratio;
      else h = w / ratio;
    }
  }
  return { width: Math.max(1, Math.min(30000, Math.round(w))), height: Math.max(1, Math.min(30000, Math.round(h))) };
}
