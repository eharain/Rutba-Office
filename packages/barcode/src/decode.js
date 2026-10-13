// Reading a QR Code back: for the tests and the window checks, which hold
// what the encoder writes to what a reader reads. It reads a clean symbol,
// the kind the suite draws, rather than a photograph: the format information
// (the nearest of the thirty-two valid codes), the version (from the size, and
// its own information from version 7), the codewords out of the zigzag
// under the mask, the blocks put back together, each block's error
// correction checked against its data, and the segments parsed.

import { LEVELS, MASKS, ALPHANUMERIC, blockLayout, functionPatterns, zigzag, formatBits, versionBits, reedSolomon } from './qr.js';

const popcount = (n) => { let c = 0; while (n) { c += n & 1; n >>>= 1; } return c; };

/** The level and mask a format word is nearest to. */
function readFormat(word) {
  let best = null;
  for (const level of Object.keys(LEVELS)) {
    for (let mask = 0; mask < 8; mask++) {
      const d = popcount(formatBits(level, mask) ^ word);
      if (!best || d < best.d) best = { level, mask, d };
    }
  }
  if (best.d > 3) throw new Error('the format information cannot be read');
  return best;
}

/** UTF-8 bytes back to text. */
function fromUtf8(bytes) {
  let out = '';
  for (let i = 0; i < bytes.length;) {
    const b = bytes[i];
    let c;
    let n;
    if (b < 0x80) { c = b; n = 1; } else if (b >= 0xf0) { c = b & 7; n = 4; } else if (b >= 0xe0) { c = b & 15; n = 3; } else { c = b & 31; n = 2; }
    for (let k = 1; k < n; k++) c = (c << 6) | (bytes[i + k] & 63);
    out += String.fromCodePoint(c);
    i += n;
  }
  return out;
}

/**
 * The text of a symbol given as rows of booleans (dark true), and what it
 * was written with: { text, version, level, mask, mode }.
 */
export function decodeQr(modules) {
  const size = modules.length;
  const version = (size - 17) / 4;
  if (!Number.isInteger(version) || version < 1 || version > 40) throw new Error(`a symbol ${size} modules across is not a QR Code`);
  const at = (x, y) => (modules[y][x] ? 1 : 0);
  let word = 0;
  const bits = [];
  for (let i = 0; i <= 5; i++) bits.push(at(8, i));
  bits.push(at(8, 7), at(8, 8), at(7, 8));
  for (let i = 9; i < 15; i++) bits.push(at(14 - i, 8));
  bits.forEach((b, i) => { word |= b << i; });
  const { level, mask } = readFormat(word);
  if (version >= 7) {
    let vw = 0;
    for (let i = 0; i < 18; i++) vw |= at(size - 11 + (i % 3), Math.floor(i / 3)) << i;
    if (popcount(vw ^ versionBits(version)) > 3) throw new Error('the version information does not match the size');
  }
  const { fixed } = functionPatterns(version);
  const order = zigzag(size, fixed);
  const all = [];
  for (let i = 0; i + 8 <= order.length; i += 8) {
    let byte = 0;
    for (let k = 0; k < 8; k++) {
      const [x, y] = order[i + k];
      byte = (byte << 1) | (at(x, y) ^ (MASKS[mask](x, y) ? 1 : 0));
    }
    all.push(byte);
  }
  // The blocks, as the encoder interleaved them.
  const { eccPerBlock, blocks, total } = blockLayout(version, level);
  const short = blocks - (total % blocks);
  const shortLen = Math.floor(total / blocks);
  const parts = Array.from({ length: blocks }, (_, i) => ({ data: [], len: shortLen - eccPerBlock + (i < short ? 0 : 1), ecc: [] }));
  let k = 0;
  for (let i = 0; i < shortLen - eccPerBlock + 1; i++) for (const p of parts) if (i < p.len) p.data.push(all[k++]);
  for (let i = 0; i < eccPerBlock; i++) for (const p of parts) p.ecc.push(all[k++]);
  for (const [n, p] of parts.entries()) {
    const want = reedSolomon(p.data, eccPerBlock);
    if (want.some((b, i) => b !== p.ecc[i])) throw new Error(`block ${n + 1}'s error correction does not match its data`);
  }
  const data = parts.flatMap((p) => p.data);
  // The segments.
  let pos = 0;
  const read = (width) => {
    let v = 0;
    for (let i = 0; i < width; i++) { v = (v << 1) | ((data[pos >>> 3] >>> (7 - (pos & 7))) & 1); pos++; }
    return v;
  };
  const band = version <= 9 ? 0 : version <= 26 ? 1 : 2;
  let text = '';
  let mode = null;
  while (pos + 4 <= data.length * 8) {
    const indicator = read(4);
    if (indicator === 0) break;
    if (indicator === 1) {
      mode = mode || 'numeric';
      let count = read([10, 12, 14][band]);
      while (count >= 3) { text += String(read(10)).padStart(3, '0'); count -= 3; }
      if (count === 2) text += String(read(7)).padStart(2, '0');
      if (count === 1) text += String(read(4));
    } else if (indicator === 2) {
      mode = mode || 'alphanumeric';
      let count = read([9, 11, 13][band]);
      while (count >= 2) { const v = read(11); text += ALPHANUMERIC[Math.floor(v / 45)] + ALPHANUMERIC[v % 45]; count -= 2; }
      if (count === 1) text += ALPHANUMERIC[read(6)];
    } else if (indicator === 4) {
      mode = mode || 'byte';
      const count = read([8, 16, 16][band]);
      const bytes = [];
      for (let i = 0; i < count; i++) bytes.push(read(8));
      text += fromUtf8(bytes);
    } else if (indicator === 7) {
      // ECI: its designator is skipped; the bytes that follow are read as UTF-8.
      const first = read(8);
      if (first & 0x80) read(first & 0x40 ? 16 : 8);
    } else {
      throw new Error(`mode ${indicator} is not one this reader knows`);
    }
  }
  return { text, version, level, mask, mode };
}

/**
 * A symbol's modules read off a picture of it: `pixels(x, y)` true where the
 * picture is dark, `width` and `height` its size. The symbol is found by
 * its dark extent; the module's size by the top-left finder, seven modules
 * wide; each module read at its centre.
 */
export function modulesFromPicture({ width, height, dark }) {
  let left = width; let top = height; let right = -1; let bottom = -1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (!dark(x, y)) continue;
      if (x < left) left = x;
      if (x > right) right = x;
      if (y < top) top = y;
      if (y > bottom) bottom = y;
    }
  }
  if (right < 0) throw new Error('the picture has nothing dark in it');
  let run = 0;
  while (left + run <= right && dark(left + run, top + 1)) run++;
  const module = run / 7;
  const count = Math.round((right - left + 1) / module);
  const step = (right - left + 1) / count;
  const rows = [];
  for (let r = 0; r < count; r++) {
    const row = [];
    for (let c = 0; c < count; c++) row.push(Boolean(dark(Math.floor(left + (c + 0.5) * step), Math.floor(top + (r + 0.5) * step))));
    rows.push(row);
  }
  return rows;
}
