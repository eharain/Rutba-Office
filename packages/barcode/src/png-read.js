// A PNG's pixels, for reading a QR Code back off the picture a document
// keeps (the tests and the window checks). Eight bits a sample, not
// interlaced: what a canvas writes. Node only, for its inflate.

import zlib from 'node:zlib';

const CHANNELS = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

/** { width, height, rgba(x, y) -> [r, g, b, a] } of a PNG's bytes. */
export function readPng(bytes) {
  const buf = Buffer.from(bytes);
  if (buf.readUInt32BE(0) !== 0x89504e47) throw new Error('not a PNG');
  let at = 8;
  let width = 0; let height = 0; let depth = 0; let type = 0; let interlace = 0;
  let palette = null;
  const idat = [];
  while (at < buf.length) {
    const len = buf.readUInt32BE(at);
    const kind = buf.toString('latin1', at + 4, at + 8);
    const data = buf.subarray(at + 8, at + 8 + len);
    if (kind === 'IHDR') { width = data.readUInt32BE(0); height = data.readUInt32BE(4); depth = data[8]; type = data[9]; interlace = data[12]; }
    else if (kind === 'PLTE') palette = data;
    else if (kind === 'IDAT') idat.push(data);
    else if (kind === 'IEND') break;
    at += 12 + len;
  }
  if (depth !== 8 || interlace) throw new Error('only 8-bit, non-interlaced PNGs are read here');
  const ch = CHANNELS[type];
  const raw = zlib.inflateSync(Buffer.concat(idat));
  const stride = width * ch;
  const px = Buffer.alloc(stride * height);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)];
    const line = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1));
    for (let i = 0; i < stride; i++) {
      const a = i >= ch ? px[y * stride + i - ch] : 0;
      const b = y ? px[(y - 1) * stride + i] : 0;
      const c = y && i >= ch ? px[(y - 1) * stride + i - ch] : 0;
      let v = line[i];
      if (filter === 1) v += a;
      else if (filter === 2) v += b;
      else if (filter === 3) v += (a + b) >> 1;
      else if (filter === 4) {
        const p = a + b - c;
        const pa = Math.abs(p - a); const pb = Math.abs(p - b); const pc = Math.abs(p - c);
        v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
      }
      px[y * stride + i] = v & 255;
    }
  }
  const rgba = (x, y) => {
    const o = y * stride + x * ch;
    if (type === 0) return [px[o], px[o], px[o], 255];
    if (type === 4) return [px[o], px[o], px[o], px[o + 1]];
    if (type === 3) return [palette[px[o] * 3], palette[px[o] * 3 + 1], palette[px[o] * 3 + 2], 255];
    return [px[o], px[o + 1], px[o + 2], type === 6 ? px[o + 3] : 255];
  };
  return { width, height, rgba };
}

/** Whether a pixel reads as dark: little light in it, and not see-through. */
export const isDark = ([r, g, b, a]) => a > 127 && (r * 299 + g * 587 + b * 114) / 1000 < 128;
