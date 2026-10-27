// The Netpbm family: PBM (black and white), PGM (grey) and PPM (colour), each
// as text (P1, P2, P3) or bytes (P4, P5, P6), and PAM (P7), which carries
// alpha. A header of whitespace-separated numbers, comments from # to the end
// of the line, then the samples; a depth over 255 is two bytes, high first.

import { bad, checkSize, rgbaOf } from './bytes.js';

export function isPnm(b) {
  return b.length >= 3 && b[0] === 0x50 && b[1] >= 0x31 && b[1] <= 0x37 && (b[2] === 0x0a || b[2] === 0x0d || b[2] === 0x20 || b[2] === 0x09 || b[2] === 0x23);
}

const space = (c) => c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d || c === 0x0b || c === 0x0c;

/** The header: kind, width, height, maxval, depth, and where the samples begin. */
function header(b) {
  if (!isPnm(b)) bad('This is not a Netpbm picture.');
  const kind = b[1] - 0x30;
  let p = 2;
  if (kind === 7) {
    // PAM: lines of KEY value, ending at ENDHDR.
    const h = { width: 0, height: 0, depth: 0, maxval: 0, tupltype: '' };
    let line = '';
    for (; p < b.length && p < 4096; p++) {
      if (b[p] !== 0x0a) { line += String.fromCharCode(b[p]); continue; }
      const t = line.replace(/#.*/, '').trim();
      line = '';
      if (t === 'ENDHDR') { p++; break; }
      const [key, ...rest] = t.split(/\s+/);
      const val = rest.join(' ');
      if (key === 'WIDTH') h.width = Number(val);
      else if (key === 'HEIGHT') h.height = Number(val);
      else if (key === 'DEPTH') h.depth = Number(val);
      else if (key === 'MAXVAL') h.maxval = Number(val);
      else if (key === 'TUPLTYPE') h.tupltype += (h.tupltype ? ' ' : '') + val;
    }
    if (!(h.depth >= 1 && h.depth <= 4)) bad(`This PAM picture has ${h.depth} samples a pixel, which Rutba Office does not read.`);
    if (!(h.maxval >= 1 && h.maxval <= 65535)) bad('This PAM picture\'s header is damaged.');
    return { kind, ...h, at: p };
  }
  const nums = [];
  const want = kind === 1 || kind === 4 ? 2 : 3;
  while (nums.length < want) {
    while (p < b.length && (space(b[p]) || b[p] === 0x23)) {
      if (b[p] === 0x23) while (p < b.length && b[p] !== 0x0a && b[p] !== 0x0d) p++;
      else p++;
    }
    let n = '';
    while (p < b.length && b[p] >= 0x30 && b[p] <= 0x39 && n.length < 10) n += String.fromCharCode(b[p++]);
    if (!n) bad('This Netpbm picture\'s header is damaged.');
    nums.push(Number(n));
  }
  p++; // the one whitespace before the samples
  const [width, height, maxval = 1] = nums;
  if (!(maxval >= 1 && maxval <= 65535)) bad(`This Netpbm picture says its samples go up to ${maxval}, which none does: it is damaged.`);
  const depth = kind === 3 || kind === 6 ? 3 : 1;
  return { kind, width, height, maxval, depth, at: p, tupltype: '' };
}

export function pnmInfo(b) {
  const h = header(b);
  const format = { 1: 'pbm', 4: 'pbm', 2: 'pgm', 5: 'pgm', 3: 'ppm', 6: 'ppm', 7: 'pam' }[h.kind];
  return { format, width: h.width, height: h.height };
}

export function decodePnm(b, { limits } = {}) {
  const h = header(b);
  const { kind, width, height, maxval, depth } = h;
  checkSize(width, height, limits, 'Netpbm');
  const out = rgbaOf(width, height);
  const n = width * height;
  const samples = new Uint16Array(n * depth);
  let got = 0;
  let p = h.at;
  if (kind === 4) {
    // Bits, rows padded to a byte, 1 black.
    const row = Math.ceil(width / 8);
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const at = p + y * row + (x >> 3);
        if (at >= b.length) break;
        samples[y * width + x] = (b[at] >> (7 - (x & 7))) & 1 ? 0 : 1;
        got = y * width + x + 1;
      }
    }
  } else if (kind >= 5) {
    const two = maxval > 255;
    for (let i = 0; i < n * depth; i++) {
      if (two) {
        if (p + 1 >= b.length) break;
        samples[i] = (b[p] << 8) | b[p + 1];
        p += 2;
      } else {
        if (p >= b.length) break;
        samples[i] = b[p++];
      }
      got = i + 1;
    }
  } else {
    // Text: numbers, or for P1 single digits that need not be apart.
    for (let i = 0; i < n * depth; i++) {
      while (p < b.length && (space(b[p]) || b[p] === 0x23)) {
        if (b[p] === 0x23) while (p < b.length && b[p] !== 0x0a) p++;
        else p++;
      }
      if (p >= b.length) break;
      let v = 0;
      if (kind === 1) {
        if (b[p] !== 0x30 && b[p] !== 0x31) bad('This PBM picture holds something other than 0 and 1: it is damaged.');
        v = b[p++] === 0x31 ? 0 : 1;
      } else {
        let digits = 0;
        while (p < b.length && b[p] >= 0x30 && b[p] <= 0x39) { v = v * 10 + b[p++] - 0x30; digits++; }
        if (!digits) bad('This Netpbm picture holds something other than numbers: it is damaged.');
      }
      samples[i] = v;
      got = i + 1;
    }
  }
  if (!got && n) bad('This Netpbm picture has a header and no pixels.');
  const top = kind === 1 || kind === 4 ? 1 : maxval;
  const level = (v) => (top === 255 ? v : Math.round((Math.min(v, top) * 255) / top));
  const grey = depth <= 2;
  const alpha = depth === 2 || depth === 4;
  for (let i = 0; i < n; i++) {
    const s = i * depth;
    const o = i * 4;
    if (grey) {
      out[o] = out[o + 1] = out[o + 2] = level(samples[s]);
    } else {
      out[o] = level(samples[s]);
      out[o + 1] = level(samples[s + 1]);
      out[o + 2] = level(samples[s + 2]);
    }
    if (alpha) out[o + 3] = level(samples[s + depth - 1]);
  }
  return { width, height, rgba: out, pages: 1 };
}
