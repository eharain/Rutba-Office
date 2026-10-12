// The pixels the advanced editor paints (docs/STUDIO.md 7.2, I10 to I14):
// a layer's bitmap made a canvas that can be drawn on, a selection or a mask
// as a canvas whose alpha is its numbers, and a region of a bitmap read back
// for undo. Every bitmap a project names lives in the window's media map as
// `{ img, width, height, bytes, name, version }`; `version` goes up each time
// its pixels change in place, which is what the renderer's caches watch, and
// `bytes` is dropped so a save encodes the pixels again.

export function makeCanvas(width, height) {
  const c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(width));
  c.height = Math.max(1, Math.round(height));
  return c;
}

/** A 2D context that is read back often. */
export const ctx2d = (canvas) => canvas.getContext('2d', { willReadFrequently: true });

/** A bitmap entry's picture as a canvas it can be painted on (made so once). */
export function editable(entry) {
  if (!entry) return null;
  if (entry.img instanceof HTMLCanvasElement) return entry.img;
  const c = makeCanvas(entry.width, entry.height);
  ctx2d(c).drawImage(entry.img, 0, 0, c.width, c.height);
  entry.img = c;
  return c;
}

/** Say a bitmap's pixels changed: the renderer draws it again and a save encodes it again. */
export function touched(entry) {
  entry.version = (entry.version || 0) + 1;
  entry.bytes = null;
}

/** A new clear bitmap entry, `width` × `height`. */
export function blankEntry(width, height, name) {
  const img = makeCanvas(width, height);
  return { img, width: img.width, height: img.height, bytes: null, name, version: 1 };
}

/** A copy of a canvas. */
export function copyCanvas(src) {
  const c = makeCanvas(src.width, src.height);
  c.getContext('2d').drawImage(src, 0, 0);
  return c;
}

/** A rectangle clamped to a canvas, in whole pixels; null when nothing is left of it. */
export function clampRect(r, width, height) {
  const x0 = Math.max(0, Math.floor(r.x));
  const y0 = Math.max(0, Math.floor(r.y));
  const x1 = Math.min(width, Math.ceil(r.x + r.width));
  const y1 = Math.min(height, Math.ceil(r.y + r.height));
  return x1 > x0 && y1 > y0 ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null;
}

/** The pixels of `rect` of a canvas, as ImageData (for undo). */
export function grab(canvas, rect) {
  return ctx2d(canvas).getImageData(rect.x, rect.y, rect.width, rect.height);
}

/** An alpha plane (a selection's or a mask's numbers) as a white canvas with that alpha. */
export function alphaCanvas(plane, width, height) {
  const c = makeCanvas(width, height);
  const g = c.getContext('2d');
  const img = g.createImageData(c.width, c.height);
  const d = img.data;
  for (let i = 0, n = c.width * c.height; i < n; i++) {
    d[i * 4] = 255;
    d[i * 4 + 1] = 255;
    d[i * 4 + 2] = 255;
    d[i * 4 + 3] = plane[i];
  }
  g.putImageData(img, 0, 0);
  return c;
}

const selectionCanvases = new WeakMap();

/** A selection as a canvas the size of the picture, its alpha what is selected; kept per selection. */
export function selectionCanvas(sel) {
  if (!sel) return null;
  let c = selectionCanvases.get(sel);
  if (!c) {
    c = alphaCanvas(sel.data, sel.width, sel.height);
    selectionCanvases.set(sel, c);
  }
  return c;
}

/** The alpha of a canvas, as a plane of 0 to 255. */
export function alphaOf(canvas) {
  const d = ctx2d(canvas).getImageData(0, 0, canvas.width, canvas.height).data;
  const out = new Uint8Array(canvas.width * canvas.height);
  for (let i = 0; i < out.length; i++) out[i] = d[i * 4 + 3];
  return out;
}

/** RGBA bytes as a canvas. */
export function rgbaCanvas(rgba, width, height) {
  const c = makeCanvas(width, height);
  const g = c.getContext('2d');
  const img = g.createImageData(c.width, c.height);
  img.data.set(rgba instanceof Uint8ClampedArray ? rgba : new Uint8ClampedArray(rgba));
  g.putImageData(img, 0, 0);
  return c;
}

/** A whole canvas's RGBA, as an ArrayBuffer of its own (for a worker job). */
export function rgbaBuffer(canvas) {
  return ctx2d(canvas).getImageData(0, 0, canvas.width, canvas.height).data.buffer;
}
