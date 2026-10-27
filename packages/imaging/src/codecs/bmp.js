// Windows bitmaps, read and written: the BMP file, and the DIB inside an
// icon, a cursor, an old Word document or a clipboard. Every header Windows
// and OS/2 wrote (core, info, V2 to V5), 1 to 32 bits a pixel, RLE4, RLE8 and
// bit fields, top-down and bottom-up, with a PNG inside one too.

import { bad, reader, checkSize } from './bytes.js';
import { decodePng, isPng } from './png.js';

export function isBmp(b) {
  return b.length >= 26 && b[0] === 0x42 && b[1] === 0x4d;
}

export function bmpInfo(b) {
  const r = reader(b, 'BMP');
  if (!isBmp(b)) bad('This is not a BMP picture.');
  const size = r.u32(14);
  const core = size === 12;
  return { format: 'bmp', width: core ? r.u16(18) : Math.abs(r.i32(18)), height: core ? r.u16(20) : Math.abs(r.i32(22)), pages: 1 };
}

/** The lowest set bit and width of a channel mask. */
function maskOf(m) {
  if (!m) return { shift: 0, bits: 0, max: 0 };
  let shift = 0;
  while (!((m >>> shift) & 1)) shift++;
  let bits = 0;
  while ((m >>> (shift + bits)) & 1) bits++;
  return { shift, bits, max: 2 ** bits - 1, mask: m >>> 0 };
}

/**
 * A DIB (a BITMAPINFOHEADER and what follows) as RGBA.
 * @param {Uint8Array} b
 * @param {number} at where the DIB header starts
 * @param {{ pixelsAt?: number, icon?: boolean, limits?: object, kind?: string }} o
 *   pixelsAt: where the pixels are (a BMP file says; a DIB has them after its palette);
 *   icon: an icon's DIB, twice as tall for its AND mask
 */
export function decodeDib(b, at, { pixelsAt = -1, icon = false, limits, kind = 'BMP' } = {}) {
  const r = reader(b, kind);
  const size = r.u32(at);
  if (![12, 16, 40, 52, 56, 64, 108, 124].includes(size)) bad(`This ${kind} picture's header is one no Windows bitmap has: it is damaged.`);
  const core = size === 12;
  const width = core ? r.u16(at + 4) : r.i32(at + 4);
  let height = core ? r.i16(at + 6) : r.i32(at + 8);
  const bpp = core ? r.u16(at + 10) : r.u16(at + 14);
  const compression = core || size === 16 ? 0 : r.u32(at + 16);
  let colours = core || size === 16 ? 0 : r.u32(at + 32);
  const topDown = height < 0;
  height = Math.abs(height);
  if (icon) height = Math.floor(height / 2);
  if (width < 1) bad(`This ${kind} picture is ${width} pixels wide, which no picture is: it is damaged.`);
  checkSize(width, height, limits, kind);
  if (![1, 2, 4, 8, 16, 24, 32].includes(bpp) && !(bpp === 0 && (compression === 4 || compression === 5))) bad(`This ${kind} picture has ${bpp} bits a pixel, which no bitmap has: it is damaged.`);

  if (compression === 5 || compression === 4) {
    const start = pixelsAt >= 0 ? pixelsAt : at + size;
    const inner = b.subarray(start);
    if (compression === 5 && isPng(inner)) return { ...decodePng(inner, limits), topDown: true };
    bad(`This ${kind} picture holds a JPEG inside it, which Rutba Office does not read there yet.`);
  }

  // Bit fields: in the header (V2 and later), or after a 40-byte one.
  let masks = null;
  if (compression === 3 || compression === 6) {
    const m = size >= 52 ? at + 40 : at + size;
    masks = [r.u32(m), r.u32(m + 4), r.u32(m + 8), compression === 6 || size >= 56 ? r.u32(m + 12) : 0];
  } else if (bpp === 16) masks = [0x7c00, 0x03e0, 0x001f, 0];
  else if (bpp === 32) masks = [0xff0000, 0xff00, 0xff, 0];

  const paletteAt = at + size + (size === 40 && (compression === 3 || compression === 6) ? (compression === 6 ? 16 : 12) : 0);
  const entry = core ? 3 : 4;
  if (bpp <= 8) {
    if (!colours || colours > 1 << bpp) colours = 1 << bpp;
  } else colours = 0;
  const palette = new Uint8Array(colours * 3);
  for (let i = 0; i < colours; i++) {
    const p = paletteAt + i * entry;
    if (p + 3 > b.length) break;
    palette[i * 3] = b[p + 2];
    palette[i * 3 + 1] = b[p + 1];
    palette[i * 3 + 2] = b[p];
  }
  const start = pixelsAt >= 0 ? pixelsAt : paletteAt + colours * entry;
  const rgba = new Uint8Array(width * height * 4);
  const stride = Math.floor((width * bpp + 31) / 32) * 4;
  const rowOf = (y) => (topDown ? y : height - 1 - y);

  if (compression === 1 || compression === 2) {
    // Run-length: bottom-up always; what the runs do not reach stays clear.
    const four = compression === 2;
    let x = 0;
    let y = 0;
    let i = start;
    const put = (k) => {
      if (x >= width || y >= height) return;
      const o = (rowOf(y) * width + x) * 4;
      rgba[o] = palette[k * 3] ?? 0;
      rgba[o + 1] = palette[k * 3 + 1] ?? 0;
      rgba[o + 2] = palette[k * 3 + 2] ?? 0;
      rgba[o + 3] = 255;
    };
    while (i + 1 < b.length && y < height) {
      const n = b[i];
      const v = b[i + 1];
      i += 2;
      if (n) {
        for (let k = 0; k < n; k++, x++) put(four ? (k & 1 ? v & 15 : v >> 4) : v);
      } else if (v === 0) {
        x = 0;
        y++;
      } else if (v === 1) break;
      else if (v === 2) {
        if (i + 1 >= b.length) break;
        x += b[i];
        y += b[i + 1];
        i += 2;
      } else {
        for (let k = 0; k < v; k++, x++) {
          if (i >= b.length) break;
          put(four ? (k & 1 ? b[i] & 15 : b[i] >> 4) : b[i]);
          if (!four || k & 1) i++;
        }
        if (four && v & 1) i++;
        // runs are padded to a word
        const used = four ? Math.ceil(v / 2) : v;
        if (used & 1) i++;
      }
    }
    return { width, height, rgba };
  }

  const ch = masks ? masks.map(maskOf) : null;
  let anyAlpha = false;
  for (let y = 0; y < height; y++) {
    const row = start + rowOf(y) * stride;
    if (row >= b.length) break;
    for (let x = 0; x < width; x++) {
      const o = (y * width + x) * 4;
      if (bpp <= 8) {
        const bit = x * bpp;
        const p = row + (bit >> 3);
        if (p >= b.length) break;
        const k = (b[p] >> (8 - bpp - (bit & 7))) & ((1 << bpp) - 1);
        rgba[o] = palette[k * 3] ?? 0;
        rgba[o + 1] = palette[k * 3 + 1] ?? 0;
        rgba[o + 2] = palette[k * 3 + 2] ?? 0;
        rgba[o + 3] = 255;
      } else if (bpp === 24) {
        const p = row + x * 3;
        if (p + 2 >= b.length) break;
        rgba[o] = b[p + 2];
        rgba[o + 1] = b[p + 1];
        rgba[o + 2] = b[p];
        rgba[o + 3] = 255;
      } else {
        const p = row + x * (bpp >> 3);
        if (p + (bpp >> 3) > b.length) break;
        const v = bpp === 16 ? b[p] | (b[p + 1] << 8) : (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
        for (let c = 0; c < 3; c++) {
          const m = ch[c];
          rgba[o + c] = m.bits ? Math.round((((v & m.mask) >>> m.shift) * 255) / m.max) : 0;
        }
        const a = ch[3];
        if (a.bits) {
          rgba[o + 3] = Math.round((((v & a.mask) >>> a.shift) * 255) / a.max);
          if (rgba[o + 3]) anyAlpha = true;
        } else if (bpp === 32 && compression === 0) {
          // The byte Windows calls reserved, which icons and many programs use as alpha.
          rgba[o + 3] = b[p + 3];
          if (b[p + 3]) anyAlpha = true;
        } else rgba[o + 3] = 255;
      }
    }
  }
  // A 32-bit bitmap whose fourth byte is zero throughout has no alpha: it is opaque.
  if (bpp > 8 && !anyAlpha) for (let i = 3; i < rgba.length; i += 4) rgba[i] = 255;

  if (icon) {
    // The AND mask: one bit a pixel, set where the icon is see-through.
    const maskAt = start + stride * height;
    const mstride = Math.floor((width + 31) / 32) * 4;
    if (bpp < 32 || !anyAlpha) {
      for (let y = 0; y < height; y++) {
        const row = maskAt + rowOf(y) * mstride;
        for (let x = 0; x < width; x++) {
          const p = row + (x >> 3);
          if (p >= b.length) break;
          if ((b[p] >> (7 - (x & 7))) & 1) rgba[(y * width + x) * 4 + 3] = 0;
        }
      }
    }
  }
  return { width, height, rgba };
}

export function decodeBmp(b, { limits } = {}) {
  const r = reader(b, 'BMP');
  if (!isBmp(b)) bad('This is not a BMP picture.');
  const pixelsAt = r.u32(10);
  return decodeDib(b, 14, { pixelsAt: pixelsAt >= 26 && pixelsAt < b.length ? pixelsAt : -1, limits });
}

/**
 * Straight RGBA as a BMP: 24 bits a pixel, as Paint writes, or 32 with a
 * V5 header carrying alpha where the picture has see-through pixels (which
 * Windows, Office and the browsers all read).
 */
export function encodeBmp(rgba, width, height) {
  let alpha = false;
  for (let i = 3; i < rgba.length; i += 4) if (rgba[i] !== 255) { alpha = true; break; }
  const bpp = alpha ? 32 : 24;
  const header = alpha ? 124 : 40;
  const stride = Math.floor((width * bpp + 31) / 32) * 4;
  const pixelsAt = 14 + header;
  const size = pixelsAt + stride * height;
  const out = new Uint8Array(size);
  const v = new DataView(out.buffer);
  out[0] = 0x42;
  out[1] = 0x4d;
  v.setUint32(2, size, true);
  v.setUint32(10, pixelsAt, true);
  v.setUint32(14, header, true);
  v.setInt32(18, width, true);
  v.setInt32(22, height, true);
  v.setUint16(26, 1, true);
  v.setUint16(28, bpp, true);
  v.setUint32(30, alpha ? 3 : 0, true);
  v.setUint32(34, stride * height, true);
  v.setInt32(38, 2835, true); // 72 dpi
  v.setInt32(42, 2835, true);
  if (alpha) {
    v.setUint32(54, 0x00ff0000, true);
    v.setUint32(58, 0x0000ff00, true);
    v.setUint32(62, 0x000000ff, true);
    v.setUint32(66, 0xff000000, true);
    v.setUint32(70, 0x73524742, true); // 'sRGB'
    v.setUint32(122, 4, true); // LCS_GM_IMAGES
  }
  for (let y = 0; y < height; y++) {
    const row = pixelsAt + (height - 1 - y) * stride;
    for (let x = 0; x < width; x++) {
      const s = (y * width + x) * 4;
      const d = row + x * (bpp >> 3);
      out[d] = rgba[s + 2];
      out[d + 1] = rgba[s + 1];
      out[d + 2] = rgba[s];
      if (alpha) out[d + 3] = rgba[s + 3];
    }
  }
  return out;
}
