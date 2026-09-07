// The edit pipeline.
//
// An edit is a list of operations, not a modified bitmap. That is what makes
// the image tool non-destructive: the original bytes are never touched, undo is
// dropping an operation, and "reset" is emptying the list. The pixels are
// produced only when something needs to see them — a preview, or an export.
//
// The operations are described here in a form with no DOM in it, so the same
// list can be planned and tested in Node and executed on a canvas in the
// window. `toCanvasPlan` turns the list into the exact sequence of draws.

/** @typedef {{ op: string, [key: string]: any }} Operation */

export const OPS = {
  crop: 'crop',          // { x, y, width, height } in source pixels
  rotate: 'rotate',      // { degrees: 90 | 180 | 270 }
  flip: 'flip',          // { axis: 'x' | 'y' }
  resize: 'resize',      // { width, height }
  adjust: 'adjust',      // { brightness, contrast, saturation, hue, blur, sepia, grayscale, invert }
  annotate: 'annotate',  // { shape: 'rect'|'ellipse'|'arrow'|'text'|'pen', ... }
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

/** The size the image ends up, after crops, rotations and resizes. */
export function resultSize(source, operations) {
  let { width, height } = source;
  for (const op of operations) {
    if (op.op === OPS.crop) {
      width = Math.max(1, Math.round(op.width));
      height = Math.max(1, Math.round(op.height));
    } else if (op.op === OPS.rotate && (op.degrees === 90 || op.degrees === 270)) {
      [width, height] = [height, width];
    } else if (op.op === OPS.resize) {
      width = Math.max(1, Math.round(op.width));
      height = Math.max(1, Math.round(op.height));
    }
  }
  return { width, height };
}

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

/** True when the list changes nothing, so Save can be offered honestly. */
export function isIdentity(operations) {
  return !operations.some((op) => {
    if (op.op === OPS.adjust) return filterString(op) !== '';
    if (op.op === OPS.rotate) return op.degrees % 360 !== 0;
    return true;
  });
}

/**
 * Turn the list into a plan a canvas can execute in order.
 *
 * Crops and geometry are collapsed into one source rectangle, one rotation and
 * one destination size, so the pixels are resampled once rather than once per
 * operation — which is both faster and visibly sharper.
 */
export function toCanvasPlan(source, operations) {
  let sourceRect = { x: 0, y: 0, width: source.width, height: source.height };
  let rotate = 0;
  let flipX = false;
  let flipY = false;
  let target = { width: source.width, height: source.height };
  const adjust = { ...DEFAULT_ADJUSTMENTS };
  const overlays = [];

  for (const op of operations) {
    switch (op.op) {
      case OPS.crop:
        // A crop is expressed against what is currently shown, so it composes
        // with the crop before it rather than replacing it.
        sourceRect = {
          x: sourceRect.x + Math.max(0, Math.round(op.x)),
          y: sourceRect.y + Math.max(0, Math.round(op.y)),
          width: Math.max(1, Math.round(op.width)),
          height: Math.max(1, Math.round(op.height)),
        };
        target = { width: sourceRect.width, height: sourceRect.height };
        break;
      case OPS.rotate:
        rotate = (rotate + (op.degrees || 0)) % 360;
        if (op.degrees === 90 || op.degrees === 270) target = { width: target.height, height: target.width };
        break;
      case OPS.flip:
        if (op.axis === 'y') flipY = !flipY;
        else flipX = !flipX;
        break;
      case OPS.resize:
        target = { width: Math.max(1, Math.round(op.width)), height: Math.max(1, Math.round(op.height)) };
        break;
      case OPS.adjust:
        Object.assign(adjust, op);
        break;
      case OPS.annotate:
        overlays.push(op);
        break;
      default:
        break;
    }
  }

  return { sourceRect, rotate, flipX, flipY, target, adjust, filter: filterString(adjust), overlays };
}

/** Fit a box inside another, preserving the ratio — the viewer's zoom-to-fit. */
export function fit(inner, outer) {
  if (!inner.width || !inner.height) return { width: 0, height: 0, scale: 1 };
  const scale = Math.min(outer.width / inner.width, outer.height / inner.height);
  return { width: Math.round(inner.width * scale), height: Math.round(inner.height * scale), scale };
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
