// QOI, the Quite OK Image format (qoiformat.org, specification 1.0): a
// header of 14 bytes, then one small operation a pixel or a run of pixels,
// then seven zero bytes and a one. Lossless, RGB or RGBA.

import { bad, reader, checkSize, rgbaOf } from './bytes.js';

export function isQoi(b) {
  return b.length >= 14 && b[0] === 0x71 && b[1] === 0x6f && b[2] === 0x69 && b[3] === 0x66; // qoif
}

export function qoiInfo(b) {
  const r = reader(b, 'QOI', false);
  if (!isQoi(b)) bad('This is not a QOI picture.');
  const width = r.u32(4);
  const height = r.u32(8);
  const channels = b[12];
  if (channels !== 3 && channels !== 4) bad(`This QOI picture says it has ${channels} channels; a QOI picture has 3 or 4: it is damaged.`);
  return { format: 'qoi', width, height, channels };
}

export function decodeQoi(b, { limits } = {}) {
  const { width, height } = qoiInfo(b);
  checkSize(width, height, limits, 'QOI');
  const out = rgbaOf(width, height);
  const index = new Uint8Array(64 * 4);
  let r = 0;
  let g = 0;
  let bl = 0;
  let a = 255;
  let run = 0;
  let p = 14;
  const end = b.length - 8; // the end mark is not pixels
  const total = width * height * 4;
  for (let o = 0; o < total; o += 4) {
    if (run > 0) {
      run--;
    } else if (p < end) {
      const op = b[p++];
      if (op === 0xfe) {
        r = b[p++]; g = b[p++]; bl = b[p++];
      } else if (op === 0xff) {
        r = b[p++]; g = b[p++]; bl = b[p++]; a = b[p++];
      } else if ((op & 0xc0) === 0x00) {
        const k = (op & 0x3f) * 4;
        r = index[k]; g = index[k + 1]; bl = index[k + 2]; a = index[k + 3];
      } else if ((op & 0xc0) === 0x40) {
        r = (r + ((op >> 4) & 3) - 2) & 255;
        g = (g + ((op >> 2) & 3) - 2) & 255;
        bl = (bl + (op & 3) - 2) & 255;
      } else if ((op & 0xc0) === 0x80) {
        const second = b[p++];
        const dg = (op & 0x3f) - 32;
        r = (r + dg + ((second >> 4) & 15) - 8) & 255;
        g = (g + dg) & 255;
        bl = (bl + dg + (second & 15) - 8) & 255;
      } else {
        run = op & 0x3f;
      }
      if (p > b.length) break; // an operation cut off by the file's end
      const k = ((r * 3 + g * 5 + bl * 7 + a * 11) % 64) * 4;
      index[k] = r; index[k + 1] = g; index[k + 2] = bl; index[k + 3] = a;
    } else {
      break; // the file ends early: the rows it reached are kept
    }
    out[o] = r; out[o + 1] = g; out[o + 2] = bl; out[o + 3] = a;
  }
  return { width, height, rgba: out, pages: 1 };
}

/** RGBA as QOI (RGBA channels when any pixel is see-through). */
export function encodeQoi(rgba, width, height) {
  let alpha = false;
  for (let i = 3; i < rgba.length; i += 4) if (rgba[i] !== 255) { alpha = true; break; }
  const out = new Uint8Array(14 + width * height * 5 + 8);
  const v = new DataView(out.buffer);
  out.set([0x71, 0x6f, 0x69, 0x66]);
  v.setUint32(4, width);
  v.setUint32(8, height);
  out[12] = alpha ? 4 : 3;
  out[13] = 0;
  const index = new Uint8Array(64 * 4);
  let p = 14;
  let pr = 0;
  let pg = 0;
  let pb = 0;
  let pa = 255;
  let run = 0;
  const total = width * height * 4;
  for (let o = 0; o < total; o += 4) {
    const r = rgba[o];
    const g = rgba[o + 1];
    const b = rgba[o + 2];
    const a = rgba[o + 3];
    if (r === pr && g === pg && b === pb && a === pa) {
      run++;
      if (run === 62 || o + 4 === total) { out[p++] = 0xc0 | (run - 1); run = 0; }
      continue;
    }
    if (run) { out[p++] = 0xc0 | (run - 1); run = 0; }
    const k = ((r * 3 + g * 5 + b * 7 + a * 11) % 64) * 4;
    if (index[k] === r && index[k + 1] === g && index[k + 2] === b && index[k + 3] === a) {
      out[p++] = k / 4;
    } else {
      index[k] = r; index[k + 1] = g; index[k + 2] = b; index[k + 3] = a;
      if (a === pa) {
        const dr = ((r - pr + 384) & 255) - 128;
        const dg = ((g - pg + 384) & 255) - 128;
        const db = ((b - pb + 384) & 255) - 128;
        const drg = dr - dg;
        const dbg = db - dg;
        if (dr >= -2 && dr <= 1 && dg >= -2 && dg <= 1 && db >= -2 && db <= 1) out[p++] = 0x40 | ((dr + 2) << 4) | ((dg + 2) << 2) | (db + 2);
        else if (dg >= -32 && dg <= 31 && drg >= -8 && drg <= 7 && dbg >= -8 && dbg <= 7) { out[p++] = 0x80 | (dg + 32); out[p++] = ((drg + 8) << 4) | (dbg + 8); }
        else { out[p++] = 0xfe; out[p++] = r; out[p++] = g; out[p++] = b; }
      } else {
        out[p++] = 0xff; out[p++] = r; out[p++] = g; out[p++] = b; out[p++] = a;
      }
    }
    pr = r; pg = g; pb = b; pa = a;
  }
  p += 7;
  out[p++] = 1;
  return out.slice(0, p);
}
