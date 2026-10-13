// Reading a PNG into pixels, for pictures that arrive as files rather than
// as a window's canvas. Every colour type at every bit depth the format
// allows, non-interlaced or Adam7; the answer is always eight-bit RGBA.

import zlib from 'node:zlib';

const SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10];

export const isPng = (bytes) => bytes && bytes.length > 8 && SIGNATURE.every((b, i) => bytes[i] === b);

/** @returns {{ width: number, height: number, data: Uint8Array }} RGBA, eight bits a channel */
export function decodePng(bytes) {
  if (!isPng(bytes)) throw new Error('This is not a PNG picture.');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let at = 8;
  let header = null;
  let palette = null;
  let alpha = null;
  const idat = [];
  while (at + 8 <= bytes.length) {
    const length = view.getUint32(at);
    const type = String.fromCharCode(bytes[at + 4], bytes[at + 5], bytes[at + 6], bytes[at + 7]);
    const body = bytes.subarray(at + 8, at + 8 + length);
    if (type === 'IHDR') {
      header = { width: view.getUint32(at + 8), height: view.getUint32(at + 12), depth: body[8], colour: body[9], interlace: body[12] };
    } else if (type === 'PLTE') palette = body;
    else if (type === 'tRNS') alpha = body;
    else if (type === 'IDAT') idat.push(body);
    else if (type === 'IEND') break;
    at += 12 + length;
  }
  if (!header || !idat.length) throw new Error('The PNG picture is incomplete.');
  const { width, height, depth, colour, interlace } = header;
  const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colour];
  if (!channels) throw new Error('The PNG picture uses a colour type that does not exist.');
  const raw = zlib.inflateSync(Buffer.concat(idat.map((b) => Buffer.from(b.buffer, b.byteOffset, b.byteLength))));
  const out = new Uint8Array(width * height * 4);
  const bpp = Math.max(1, (channels * depth) >> 3);

  // Sample n of a row, scaled to eight bits (a palette index left as it is).
  const sample = (row, n) => {
    if (depth === 8) return row[n];
    if (depth === 16) return row[n * 2];
    const bit = n * depth;
    const v = (row[bit >> 3] >> (8 - depth - (bit & 7))) & ((1 << depth) - 1);
    return colour === 3 ? v : Math.round((v * 255) / ((1 << depth) - 1));
  };
  const put = (row, x, px, py) => {
    const o = (py * width + px) * 4;
    if (colour === 3) {
      const i = sample(row, x);
      out[o] = palette[i * 3]; out[o + 1] = palette[i * 3 + 1]; out[o + 2] = palette[i * 3 + 2];
      out[o + 3] = alpha && i < alpha.length ? alpha[i] : 255;
    } else if (colour === 0 || colour === 4) {
      const g = sample(row, x * channels);
      out[o] = out[o + 1] = out[o + 2] = g;
      out[o + 3] = colour === 4 ? sample(row, x * channels + 1) : 255;
    } else {
      out[o] = sample(row, x * channels); out[o + 1] = sample(row, x * channels + 1); out[o + 2] = sample(row, x * channels + 2);
      out[o + 3] = colour === 6 ? sample(row, x * channels + 3) : 255;
    }
  };

  let pos = 0;
  // One pass of rows (the whole picture, or one of Adam7's seven).
  const pass = (pw, ph, place) => {
    if (!pw || !ph) return;
    const stride = Math.ceil((pw * channels * depth) / 8);
    let prev = new Uint8Array(stride);
    for (let y = 0; y < ph; y++) {
      const filter = raw[pos++];
      const row = Uint8Array.from(raw.subarray(pos, pos + stride));
      pos += stride;
      for (let i = 0; i < stride; i++) {
        const a = i >= bpp ? row[i - bpp] : 0;
        const b = prev[i];
        const c = i >= bpp ? prev[i - bpp] : 0;
        if (filter === 1) row[i] = (row[i] + a) & 255;
        else if (filter === 2) row[i] = (row[i] + b) & 255;
        else if (filter === 3) row[i] = (row[i] + ((a + b) >> 1)) & 255;
        else if (filter === 4) {
          const p = a + b - c;
          const pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
          row[i] = (row[i] + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c)) & 255;
        }
      }
      for (let x = 0; x < pw; x++) {
        const [px, py] = place(x, y);
        put(row, x, px, py);
      }
      prev = row;
    }
  };
  if (interlace) {
    for (const [x0, y0, dx, dy] of [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]]) {
      pass(Math.ceil((width - x0) / dx), Math.ceil((height - y0) / dy), (x, y) => [x0 + x * dx, y0 + y * dy]);
    }
  } else pass(width, height, (x, y) => [x, y]);
  return { width, height, data: out };
}
