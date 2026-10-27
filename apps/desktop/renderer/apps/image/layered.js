// The layered picture drawn: the project (an image document of
// @rutba/studio, layers.js) built into a plan at the size wanted and painted
// by the engine's own painter. The preview builds it at the window's scale and
// an export or a save at the canvas's own size; the geometry is fractional, so
// both show the same picture.

import { buildImagePlan, paintFrame, layerBounds } from '@rutba/studio/host';
import { imagePatches, canvasOf } from '@rutba/studio/layers';
import { decodePicture } from '@rutba/imaging/decode';
import { svgSize } from './formats.js';

const scratch = () => document.createElement('canvas');

/** A plan of the project at `scale` of its canvas, drawn on nothing yet. */
export function planAt(doc, assets, scale = 1) {
  const { width, height } = canvasOf(doc);
  const size = { width: Math.max(1, Math.round(width * scale)), height: Math.max(1, Math.round(height * scale)) };
  return buildImagePlan({ canvas: scratch(), size, options: doc.options, layerPatches: imagePatches(doc), assets });
}

/** Paint the project onto `canvas` at `scale`; the canvas is sized to it. Answers the plan. */
export function paintLayered(canvas, doc, assets, scale = 1) {
  const plan = planAt(doc, assets, scale);
  if (canvas.width !== plan.W) canvas.width = plan.W;
  if (canvas.height !== plan.H) canvas.height = plan.H;
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, plan.W, plan.H);
  paintFrame(ctx, plan, 0);
  return plan;
}

/** The project flattened at `scale`, on a canvas of its own; `ground` fills what is clear (a JPEG has no clear). */
export function renderLayered(doc, assets, scale = 1, { ground = null } = {}) {
  const canvas = scratch();
  paintLayered(canvas, doc, assets, scale);
  if (!ground) return canvas;
  const out = scratch();
  out.width = canvas.width;
  out.height = canvas.height;
  const ctx = out.getContext('2d');
  ctx.fillStyle = ground;
  ctx.fillRect(0, 0, out.width, out.height);
  ctx.drawImage(canvas, 0, 0);
  return out;
}

/** A layer's box in the plan's pixels, or null. */
export function boundsOf(plan, layerId) {
  const layer = plan.layers.find((l) => l.id === layerId);
  if (!layer) return null;
  const ctx = scratch().getContext('2d');
  const b = layerBounds(ctx, plan, layer);
  return b ? { ...b, rot: layer.rot || 0, layer } : null;
}

/**
 * An SVG put in as a layer: drawn at the size it gives itself, through an
 * <img> (a bitmap cannot be made from an SVG directly), kept as that PNG.
 */
async function decodeSvg(bytes) {
  const { width, height } = svgSize(new TextDecoder().decode(bytes));
  const url = URL.createObjectURL(new Blob([bytes], { type: 'image/svg+xml' }));
  try {
    const el = new Image();
    await new Promise((resolve, reject) => {
      el.onload = resolve;
      el.onerror = () => reject(new Error('This SVG picture could not be drawn: it may be damaged.'));
      el.src = url;
    });
    const canvas = scratch();
    canvas.width = width;
    canvas.height = height;
    const g = canvas.getContext('2d');
    g.clearRect(0, 0, width, height);
    g.drawImage(el, 0, 0, width, height);
    const img = await createImageBitmap(canvas);
    return { img, width, height, png: await pngBytes(canvas) };
  } finally {
    URL.revokeObjectURL(url);
  }
}

/**
 * Bytes of a picture, decoded: { img, width, height }. A format Chromium
 * does not draw (a TIFF, a QOI, a Targa) is read by decodePicture instead,
 * and answered with `png` too: the bytes to keep for it, since the project
 * and the canvas both want a picture Chromium can draw again.
 */
export async function decodeBitmap(bytes, type = 'image/png', name = '') {
  if (type === 'image/svg+xml') return decodeSvg(bytes);
  try {
    const img = await createImageBitmap(new Blob([bytes], { type }));
    return { img, width: img.width, height: img.height };
  } catch (err) {
    let pic;
    try {
      pic = decodePicture(bytes, name);
    } catch (refusal) {
      if (refusal?.picture) throw refusal;
      throw err;
    }
    const data = new ImageData(new Uint8ClampedArray(pic.rgba.buffer, pic.rgba.byteOffset, pic.rgba.byteLength), pic.width, pic.height);
    const canvas = scratch();
    canvas.width = pic.width;
    canvas.height = pic.height;
    canvas.getContext('2d').putImageData(data, 0, 0);
    const img = await createImageBitmap(canvas);
    return { img, width: img.width, height: img.height, png: await pngBytes(canvas) };
  }
}

/** A canvas as PNG bytes. */
export async function pngBytes(canvas) {
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}

/** The MIME type a picture's name says it is. */
export function mimeOf(name) {
  const ext = String(name || '').toLowerCase().split('.').pop();
  return { jpg: 'image/jpeg', jpeg: 'image/jpeg', webp: 'image/webp', gif: 'image/gif', bmp: 'image/bmp', avif: 'image/avif', ico: 'image/x-icon', svg: 'image/svg+xml' }[ext] || 'image/png';
}
