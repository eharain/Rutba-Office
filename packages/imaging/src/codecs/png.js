// PNG, read and written, in JavaScript.
//
// Written because every other reader here answers with pixels that a person
// must be able to see in any window, the main process included, which has
// no canvas; read because a PNG sits inside icons and cursors, and a picture
// handed to decodePicture as a PNG is answered like any other.

import { bad, reader, ascii, checkSize } from './bytes.js';
import { inflate, deflate, crc32 } from './zlib.js';

const SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

export function isPng(b) {
  return b.length >= 8 && SIGNATURE.every((v, i) => b[i] === v);
}

/** Width and height from the header, without unpacking anything. */
export function pngInfo(b) {
  const r = reader(b, 'PNG', false);
  if (!isPng(b) || ascii(b, 12, 4) !== 'IHDR') bad('This is not a PNG picture.');
  return { format: 'png', width: r.u32(16), height: r.u32(20), pages: 1 };
}

const ADAM7 = [
  [0, 0, 8, 8],
  [4, 0, 8, 8],
  [0, 4, 4, 8],
  [2, 0, 4, 4],
  [0, 2, 2, 4],
  [1, 0, 2, 2],
  [0, 1, 1, 2],
];

function paeth(a, b, c) {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
}

/** Undo a pass's row filters in place; rows the file did not reach stay zero. */
function unfilter(data, at, rowBytes, rows, bpp) {
  let prev = -1;
  for (let y = 0; y < rows; y++) {
    const start = at + y * (rowBytes + 1);
    const type = data[start];
    const row = start + 1;
    if (type > 4) bad('A row of this PNG picture is damaged.');
    for (let x = 0; x < rowBytes; x++) {
      const a = x >= bpp ? data[row + x - bpp] : 0;
      const b = prev >= 0 ? data[prev + x] : 0;
      const c = prev >= 0 && x >= bpp ? data[prev + x - bpp] : 0;
      let v = data[row + x];
      if (type === 1) v += a;
      else if (type === 2) v += b;
      else if (type === 3) v += (a + b) >> 1;
      else if (type === 4) v += paeth(a, b, c);
      data[row + x] = v;
    }
    prev = row;
  }
}

/**
 * A PNG decoded to straight RGBA, eight bits a channel.
 * @returns {{ width, height, rgba: Uint8Array }}
 */
export function decodePng(b, limits) {
  const r = reader(b, 'PNG', false);
  if (!isPng(b)) bad('This is not a PNG picture.');
  let at = 8;
  let ihdr = null;
  let palette = null;
  let trns = null;
  const idat = [];
  let idatLength = 0;
  while (at + 8 <= b.length) {
    const len = r.u32(at);
    const type = ascii(b, at + 4, 4);
    if (len > b.length - at - 8) {
      // A chunk cut short at the end of a file: what arrived of it is kept, as browsers keep it.
      if (type === 'IDAT') {
        idat.push(b.subarray(at + 8));
        idatLength += b.length - at - 8;
      }
      break;
    }
    const body = b.subarray(at + 8, at + 8 + len);
    if (type === 'IHDR') {
      if (len < 13) bad('This PNG picture\'s header is damaged.');
      ihdr = { width: r.u32(at + 8), height: r.u32(at + 12), depth: body[8], colour: body[9], interlace: body[12] };
    } else if (type === 'PLTE') palette = body;
    else if (type === 'tRNS') trns = body;
    else if (type === 'IDAT') {
      idat.push(body);
      idatLength += len;
    } else if (type === 'IEND') break;
    at += 12 + len;
  }
  if (!ihdr) bad('This PNG picture has no header: it is damaged.');
  const { width, height, depth, colour, interlace } = ihdr;
  checkSize(width, height, limits, 'PNG');
  const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colour];
  const depthOk = { 0: [1, 2, 4, 8, 16], 2: [8, 16], 3: [1, 2, 4, 8], 4: [8, 16], 6: [8, 16] }[colour];
  if (!channels || !depthOk.includes(depth)) bad('This PNG picture uses a colour type no PNG has: it is damaged.');
  if (colour === 3 && !palette) bad('This PNG picture has no palette for its colours: it is damaged.');
  if (!idat.length) bad('This PNG picture has no pixels in it.');

  const bitsPx = channels * depth;
  const bpp = Math.max(1, bitsPx >> 3);
  const passes = interlace
    ? ADAM7.map(([x0, y0, dx, dy]) => ({ x0, y0, dx, dy, w: Math.ceil((width - x0) / dx), h: Math.ceil((height - y0) / dy) })).filter((p) => p.w > 0 && p.h > 0)
    : [{ x0: 0, y0: 0, dx: 1, dy: 1, w: width, h: height }];
  let expected = 0;
  for (const p of passes) {
    p.rowBytes = Math.ceil((p.w * bitsPx) / 8);
    p.at = expected;
    expected += (p.rowBytes + 1) * p.h;
  }
  const joined = new Uint8Array(idatLength);
  let o = 0;
  for (const part of idat) {
    joined.set(part, o);
    o += part.length;
  }
  const unpacked = inflate(joined, { size: expected, max: expected });
  const data = new Uint8Array(expected);
  data.set(unpacked.subarray(0, expected));

  const rgba = new Uint8Array(width * height * 4);
  const tr = trns ? (depth === 16 ? null : trns) : null;
  const grayKey = trns && colour === 0 ? (trns[0] << 8) | trns[1] : -1;
  const rgbKey = trns && colour === 2 ? [(trns[0] << 8) | trns[1], (trns[2] << 8) | trns[3], (trns[4] << 8) | trns[5]] : null;
  const scale = depth < 8 ? 255 / ((1 << depth) - 1) : 1;
  for (const p of passes) {
    unfilter(data, p.at, p.rowBytes, p.h, bpp);
    for (let y = 0; y < p.h; y++) {
      const row = p.at + y * (p.rowBytes + 1) + 1;
      const oy = p.y0 + y * p.dy;
      for (let x = 0; x < p.w; x++) {
        const ox = p.x0 + x * p.dx;
        const d = (oy * width + ox) * 4;
        const sample = (i) => {
          if (depth === 8) return data[row + x * channels + i];
          if (depth === 16) return data[row + (x * channels + i) * 2];
          const bit = (x * channels + i) * depth;
          return (data[row + (bit >> 3)] >> (8 - depth - (bit & 7))) & ((1 << depth) - 1);
        };
        const wide = (i) => (data[row + (x * channels + i) * 2] << 8) | data[row + (x * channels + i) * 2 + 1];
        if (colour === 3) {
          const k = sample(0);
          rgba[d] = palette[k * 3] ?? 0;
          rgba[d + 1] = palette[k * 3 + 1] ?? 0;
          rgba[d + 2] = palette[k * 3 + 2] ?? 0;
          rgba[d + 3] = tr && k < tr.length ? tr[k] : 255;
        } else if (colour === 0 || colour === 4) {
          const raw = sample(0);
          const v = depth < 8 ? Math.round(raw * scale) : raw;
          rgba[d] = rgba[d + 1] = rgba[d + 2] = v;
          if (colour === 4) rgba[d + 3] = sample(1);
          else rgba[d + 3] = grayKey >= 0 && (depth === 16 ? wide(0) : raw) === grayKey ? 0 : 255;
        } else {
          rgba[d] = sample(0);
          rgba[d + 1] = sample(1);
          rgba[d + 2] = sample(2);
          if (colour === 6) rgba[d + 3] = sample(3);
          else if (rgbKey && (depth === 16 ? wide(0) === rgbKey[0] && wide(1) === rgbKey[1] && wide(2) === rgbKey[2] : rgba[d] === rgbKey[0] && rgba[d + 1] === rgbKey[1] && rgba[d + 2] === rgbKey[2])) rgba[d + 3] = 0;
          else rgba[d + 3] = 255;
        }
      }
    }
  }
  return { width, height, rgba };
}

function chunk(type, body) {
  const out = new Uint8Array(12 + body.length);
  const len = body.length;
  out[0] = len >>> 24;
  out[1] = (len >>> 16) & 0xff;
  out[2] = (len >>> 8) & 0xff;
  out[3] = len & 0xff;
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(body, 8);
  const crc = crc32(out.subarray(4, 8 + len));
  out[8 + len] = crc >>> 24;
  out[9 + len] = (crc >>> 16) & 0xff;
  out[10 + len] = (crc >>> 8) & 0xff;
  out[11 + len] = crc & 0xff;
  return out;
}

/**
 * Straight RGBA as a PNG: RGB when every pixel is opaque, each row given
 * the filter that leaves it smallest (the sum of its bytes as signed, which
 * is what libpng guesses with).
 * @param {{ level?: number, chunks?: Array<{ type: string, body: Uint8Array }> }} [o] extra chunks go before the pixels
 */
export function encodePng(rgba, width, height, { level = 6, chunks = [] } = {}) {
  let opaque = true;
  for (let i = 3; i < rgba.length; i += 4) if (rgba[i] !== 255) { opaque = false; break; }
  const ch = opaque ? 3 : 4;
  const rowBytes = width * ch;
  const raw = new Uint8Array((rowBytes + 1) * height);
  const line = new Uint8Array(rowBytes);
  const prev = new Uint8Array(rowBytes);
  const trial = new Uint8Array(rowBytes);
  const best = new Uint8Array(rowBytes);
  for (let y = 0; y < height; y++) {
    for (let x = 0, s = y * width * 4; x < width; x++, s += 4) {
      const d = x * ch;
      line[d] = rgba[s];
      line[d + 1] = rgba[s + 1];
      line[d + 2] = rgba[s + 2];
      if (ch === 4) line[d + 3] = rgba[s + 3];
    }
    let bestType = 0;
    let bestCost = Infinity;
    for (let type = 0; type < 5; type++) {
      let cost = 0;
      for (let x = 0; x < rowBytes; x++) {
        const a = x >= ch ? line[x - ch] : 0;
        const b = y ? prev[x] : 0;
        const c = y && x >= ch ? prev[x - ch] : 0;
        const pred = type === 0 ? 0 : type === 1 ? a : type === 2 ? b : type === 3 ? (a + b) >> 1 : paeth(a, b, c);
        const v = (line[x] - pred) & 0xff;
        trial[x] = v;
        cost += v < 128 ? v : 256 - v;
        if (cost >= bestCost) break;
      }
      if (cost < bestCost) {
        bestCost = cost;
        bestType = type;
        best.set(trial);
      }
    }
    raw[y * (rowBytes + 1)] = bestType;
    raw.set(best, y * (rowBytes + 1) + 1);
    prev.set(line);
  }
  const ihdr = new Uint8Array(13);
  const v = new DataView(ihdr.buffer);
  v.setUint32(0, width);
  v.setUint32(4, height);
  ihdr[8] = 8;
  ihdr[9] = opaque ? 2 : 6;
  const parts = [Uint8Array.from(SIGNATURE), chunk('IHDR', ihdr), ...chunks.map((c) => chunk(c.type, c.body)), chunk('IDAT', deflate(raw, { level })), chunk('IEND', new Uint8Array(0))];
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}
