// Truevision TGA (Targa): an 18-byte header, an optional palette, then the
// pixels, plain or run-length packed; colour (15, 16, 24 or 32 bits), grey,
// or palette indices; stored from the bottom up unless the header says
// otherwise. TGA has no mark at its start, so it is known by its name, and
// its header is held to the few values a real one has.

import { bad, reader, checkSize, rgbaOf } from './bytes.js';

const TYPES = new Set([1, 2, 3, 9, 10, 11]);

function header(b) {
  const r = reader(b, 'TGA');
  r.need(0, 18);
  const h = {
    idLength: b[0],
    mapType: b[1],
    type: b[2],
    mapFirst: r.u16(3),
    mapLength: r.u16(5),
    mapBits: b[7],
    width: r.u16(12),
    height: r.u16(14),
    bits: b[16],
    descriptor: b[17],
  };
  if (!TYPES.has(h.type)) {
    if (h.type === 0) bad('This TGA file holds no picture.');
    bad('This is not a TGA picture Rutba Office knows: its header names a kind of picture TGA does not have.');
  }
  if (h.mapType > 1) bad('This TGA picture\'s header is damaged.');
  const colourMapped = h.type === 1 || h.type === 9;
  if (colourMapped && (h.mapType !== 1 || ![15, 16, 24, 32].includes(h.mapBits) || h.bits !== 8 && h.bits !== 16)) bad('This TGA picture\'s palette is damaged.');
  if ((h.type === 2 || h.type === 10) && ![15, 16, 24, 32].includes(h.bits)) bad(`This TGA picture says it has ${h.bits}-bit colour, which TGA does not: it is damaged.`);
  if ((h.type === 3 || h.type === 11) && h.bits !== 8 && h.bits !== 16) bad(`This grey TGA picture says it has ${h.bits}-bit pixels, which Rutba Office does not read.`);
  return h;
}

/** A name ending .tga and a header a real TGA has. */
export function looksTga(b) {
  try {
    header(b);
    return true;
  } catch {
    return false;
  }
}

export function tgaInfo(b) {
  const h = header(b);
  return { format: 'tga', width: h.width, height: h.height };
}

/** One pixel's colour, from 15, 16, 24 or 32 bits little-endian, into out at o. */
function colour(src, at, bits, out, o, alphaBits) {
  if (bits === 24 || bits === 32) {
    out[o] = src[at + 2];
    out[o + 1] = src[at + 1];
    out[o + 2] = src[at];
    out[o + 3] = bits === 32 ? src[at + 3] : 255;
  } else {
    const v = src[at] | (src[at + 1] << 8);
    out[o] = (((v >> 10) & 31) * 255 + 15) / 31 | 0;
    out[o + 1] = (((v >> 5) & 31) * 255 + 15) / 31 | 0;
    out[o + 2] = ((v & 31) * 255 + 15) / 31 | 0;
    out[o + 3] = bits === 16 && alphaBits ? (v & 0x8000 ? 255 : 0) : 255;
  }
}

export function decodeTga(b, { limits } = {}) {
  const h = header(b);
  const { width, height } = h;
  checkSize(width, height, limits, 'TGA');
  const alphaBits = h.descriptor & 15;
  const bpp = (h.bits + 7) >> 3;
  let p = 18 + h.idLength;
  // The palette, as RGBA.
  let palette = null;
  if (h.mapType === 1) {
    const eb = (h.mapBits + 7) >> 3;
    const need = h.mapLength * eb;
    if (p + need > b.length) bad('This TGA file ends early: it may be a download that stopped, or damaged.');
    if (h.type === 1 || h.type === 9) {
      palette = new Uint8Array((h.mapFirst + h.mapLength) * 4);
      for (let i = 0; i < h.mapLength; i++) colour(b, p + i * eb, h.mapBits, palette, (h.mapFirst + i) * 4, h.mapBits === 32 || alphaBits);
    }
    p += need;
  }
  // The pixels, unpacked into one run of bytes.
  const n = width * height;
  let data;
  if (h.type >= 9) {
    data = new Uint8Array(n * bpp);
    let o = 0;
    while (o < data.length && p < b.length) {
      const head = b[p++];
      const count = (head & 0x7f) + 1;
      if (head & 0x80) {
        if (p + bpp > b.length) break;
        for (let k = 0; k < count && o < data.length; k++) for (let c = 0; c < bpp; c++) data[o++] = b[p + c];
        p += bpp;
      } else {
        const len = Math.min(count * bpp, data.length - o, b.length - p);
        data.set(b.subarray(p, p + len), o);
        o += len;
        p += count * bpp;
      }
    }
  } else {
    data = b.subarray(p, p + n * bpp);
    if (data.length < bpp) bad('This TGA picture has a header and no pixels.');
  }
  const out = rgbaOf(width, height);
  const topDown = (h.descriptor & 0x20) !== 0;
  const rightToLeft = (h.descriptor & 0x10) !== 0;
  let anyAlpha = false;
  const have = Math.floor(data.length / bpp);
  for (let i = 0; i < n && i < have; i++) {
    const sx = i % width;
    const sy = (i / width) | 0;
    const x = rightToLeft ? width - 1 - sx : sx;
    const y = topDown ? sy : height - 1 - sy;
    const o = (y * width + x) * 4;
    const at = i * bpp;
    if (h.type === 1 || h.type === 9) {
      const k = bpp === 2 ? data[at] | (data[at + 1] << 8) : data[at];
      if (palette && k * 4 + 3 < palette.length) {
        out[o] = palette[k * 4]; out[o + 1] = palette[k * 4 + 1]; out[o + 2] = palette[k * 4 + 2]; out[o + 3] = palette[k * 4 + 3];
      }
    } else if (h.type === 3 || h.type === 11) {
      out[o] = out[o + 1] = out[o + 2] = data[at];
      if (bpp === 2) out[o + 3] = data[at + 1]; // grey and alpha
    } else {
      colour(data, at, h.bits, out, o, alphaBits);
    }
    if (out[o + 3] !== 0) anyAlpha = true;
  }
  // A 32-bit TGA whose alpha is zero everywhere was written with the byte unused: it is opaque.
  if (!anyAlpha) for (let i = 3; i < out.length; i += 4) out[i] = 255;
  return { width, height, rgba: out, pages: 1 };
}
