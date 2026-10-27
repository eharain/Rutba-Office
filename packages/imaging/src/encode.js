// Writing the formats a canvas cannot: BMP, TIFF and ICO beside the PNG,
// JPEG and WebP Chromium writes itself. Pure JavaScript, so Image's window,
// the batch worker and the main process write the same bytes.

import { encodeBmp } from './codecs/bmp.js';
import { encodeTiff } from './codecs/tiff.js';
import { encodePng } from './codecs/png.js';

export { encodeBmp, encodeTiff, encodePng };

/** The sizes an icon carries, as Windows asks for them (largest first). */
export const ICON_SIZES = [256, 128, 64, 48, 32, 24, 16];

/**
 * The sizes an icon of a picture this large is written at: every standard
 * size no larger than the picture (an icon is not made by enlarging), and at
 * least the smallest one.
 */
export function iconSizes(width, height) {
  const long = Math.max(width, height);
  const fit = ICON_SIZES.filter((s) => s <= long);
  return fit.length ? fit : [Math.max(1, Math.min(256, long))];
}

/**
 * An .ico of PNG images, one a size (Windows Vista and later, every browser
 * and Office read PNG inside an icon).
 * @param {Array<{ width: number, height: number, png: Uint8Array }>} entries
 */
export function encodeIco(entries, { cursor = false, hotspots = [] } = {}) {
  const list = [...entries].sort((a, b) => b.width - a.width);
  const head = 6 + list.length * 16;
  const total = head + list.reduce((n, e) => n + e.png.length, 0);
  const out = new Uint8Array(total);
  const v = new DataView(out.buffer);
  v.setUint16(2, cursor ? 2 : 1, true);
  v.setUint16(4, list.length, true);
  let at = head;
  list.forEach((e, i) => {
    const d = 6 + i * 16;
    out[d] = e.width >= 256 ? 0 : e.width;
    out[d + 1] = e.height >= 256 ? 0 : e.height;
    if (cursor) {
      v.setUint16(d + 4, hotspots[i]?.x || 0, true);
      v.setUint16(d + 6, hotspots[i]?.y || 0, true);
    } else {
      v.setUint16(d + 4, 1, true); // planes
      v.setUint16(d + 6, 32, true); // bits a pixel
    }
    v.setUint32(d + 8, e.png.length, true);
    v.setUint32(d + 12, at, true);
    out.set(e.png, at);
    at += e.png.length;
  });
  return out;
}

/** The formats written here, by the name the windows give them. */
export const OWN_FORMATS = ['bmp', 'tiff', 'ico'];

/**
 * Straight RGBA in one of the formats written here; an icon from this one
 * picture at its own size (Image draws each size itself, for the quality a
 * canvas's smoothing gives, and calls encodeIco).
 */
export function encodePicture(rgba, width, height, format, opts = {}) {
  if (format === 'bmp') return encodeBmp(rgba, width, height);
  if (format === 'tiff') return encodeTiff([{ rgba, width, height }], opts);
  if (format === 'ico') return encodeIco([{ width, height, png: encodePng(rgba, width, height) }]);
  if (format === 'png') return encodePng(rgba, width, height, opts);
  throw new Error(`Rutba Office does not write ${String(format).toUpperCase()} pictures itself.`);
}
