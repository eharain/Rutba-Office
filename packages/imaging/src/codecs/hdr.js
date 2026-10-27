// Radiance HDR (RGBE): light as it was measured, far brighter and darker than
// a screen shows. A text header to a blank line, a line giving the size and
// which way the rows run, then each pixel as three mantissas and a shared
// exponent, rows run-length packed a channel at a time (or flat, or the old
// packing). Shown tone-mapped: the scene's average brought to a mid grey and
// its highlights rolled off (Reinhard's operator), then sRGB.

import { bad, checkSize, rgbaOf } from './bytes.js';

export function isHdr(b) {
  const head = String.fromCharCode(...b.subarray(0, 10));
  return head.startsWith('#?RADIANCE') || head.startsWith('#?RGBE') || head.startsWith('#?AUTOPANO');
}

function header(b) {
  if (!isHdr(b)) bad('This is not a Radiance HDR picture.');
  let p = 0;
  let line = '';
  let format = '';
  let blank = false;
  const lines = [];
  for (; p < b.length && p < 65536; p++) {
    if (b[p] !== 0x0a) { line += String.fromCharCode(b[p]); continue; }
    if (blank) { p++; break; } // the size line
    if (!line.trim()) blank = true;
    else {
      lines.push(line);
      const m = /^FORMAT=(.*)$/.exec(line.trim());
      if (m) format = m[1].trim();
    }
    line = '';
  }
  if (!blank) bad('This Radiance HDR picture\'s header is damaged.');
  if (format && format !== '32-bit_rle_rgbe' && format !== '32-bit_rle_xyze') bad(`This Radiance HDR picture is in the ${format} layout, which Rutba Office does not read.`);
  const m = /^\s*([-+])([XY])\s+(\d+)\s+([-+])([XY])\s+(\d+)\s*$/.exec(line);
  if (!m || m[2] === m[5]) bad('This Radiance HDR picture\'s size line is damaged.');
  // The first axis named is the one the rows run along slowest.
  const slow = { sign: m[1], axis: m[2], n: Number(m[3]) };
  const fast = { sign: m[4], axis: m[5], n: Number(m[6]) };
  const width = slow.axis === 'Y' ? fast.n : slow.n;
  const height = slow.axis === 'Y' ? slow.n : fast.n;
  return { width, height, slow, fast, xyz: format === '32-bit_rle_xyze', at: p };
}

export function hdrInfo(b) {
  const h = header(b);
  return { format: 'hdr', width: h.width, height: h.height };
}

/** One scanline of `len` RGBE pixels into line, from p; answers where it ended. */
function scanline(b, p, line, len) {
  if (len >= 8 && len < 32768 && b[p] === 2 && b[p + 1] === 2 && !(b[p + 2] & 0x80)) {
    if (((b[p + 2] << 8) | b[p + 3]) !== len) bad('A row of this Radiance HDR picture is damaged.');
    p += 4;
    for (let c = 0; c < 4; c++) {
      let x = 0;
      while (x < len) {
        if (p >= b.length) return p;
        let n = b[p++];
        if (n > 128) {
          n -= 128;
          if (n > len - x) bad('A row of this Radiance HDR picture is damaged.');
          const v = b[p++];
          for (let k = 0; k < n; k++) line[(x++) * 4 + c] = v;
        } else {
          if (n === 0 || n > len - x) bad('A row of this Radiance HDR picture is damaged.');
          for (let k = 0; k < n; k++) line[(x++) * 4 + c] = b[p++];
        }
      }
    }
    return p;
  }
  // Flat, with the old packing: a pixel of 1, 1, 1, n repeats the one before it.
  let x = 0;
  let shift = 0;
  while (x < len && p + 3 < b.length) {
    if (b[p] === 1 && b[p + 1] === 1 && b[p + 2] === 1) {
      const n = b[p + 3] << shift;
      if (!x || n > len - x) bad('A row of this Radiance HDR picture is damaged.');
      for (let k = 0; k < n; k++, x++) line.copyWithin(x * 4, (x - 1) * 4, x * 4);
      shift += 8;
      p += 4;
      continue;
    }
    line[x * 4] = b[p]; line[x * 4 + 1] = b[p + 1]; line[x * 4 + 2] = b[p + 2]; line[x * 4 + 3] = b[p + 3];
    x++;
    shift = 0;
    p += 4;
  }
  return p;
}

const srgb = (v) => {
  const c = v <= 0.0031308 ? 12.92 * v : 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
  return Math.max(0, Math.min(255, Math.round(c * 255)));
};

/** Radiance values (linear light) as an sRGB picture, by Reinhard's global operator. */
export function toneMap(light, width, height) {
  const n = width * height;
  let logSum = 0;
  let maxL = 0;
  for (let i = 0; i < n; i++) {
    const L = 0.2126 * light[i * 3] + 0.7152 * light[i * 3 + 1] + 0.0722 * light[i * 3 + 2];
    logSum += Math.log(1e-6 + L);
    if (L > maxL) maxL = L;
  }
  const avg = Math.exp(logSum / Math.max(1, n));
  const scale = 0.18 / Math.max(avg, 1e-9);
  const white = Math.max(1e-6, maxL * scale);
  const out = rgbaOf(width, height);
  for (let i = 0; i < n; i++) {
    const r = light[i * 3];
    const g = light[i * 3 + 1];
    const bl = light[i * 3 + 2];
    const L = 0.2126 * r + 0.7152 * g + 0.0722 * bl;
    if (!(L > 0)) continue;
    const Ls = L * scale;
    const Ld = (Ls * (1 + Ls / (white * white))) / (1 + Ls);
    const k = Ld / L;
    out[i * 4] = srgb(r * k);
    out[i * 4 + 1] = srgb(g * k);
    out[i * 4 + 2] = srgb(bl * k);
  }
  return out;
}

/** Where the i-th of n along an axis goes: -Y is top to bottom, +X left to right. */
const along = (axis, i, n, sign) => (axis === 'Y' ? (sign === '-' ? i : n - 1 - i) : sign === '+' ? i : n - 1 - i);

export function decodeHdr(b, { limits } = {}) {
  const h = header(b);
  const { width, height, slow, fast } = h;
  checkSize(width, height, limits, 'Radiance HDR');
  const light = new Float32Array(width * height * 3);
  const line = new Uint8Array(fast.n * 4);
  let p = h.at;
  for (let s = 0; s < slow.n && p < b.length; s++) {
    line.fill(0);
    p = scanline(b, p, line, fast.n);
    for (let f = 0; f < fast.n; f++) {
      const a = along(slow.axis, s, slow.n, slow.sign);
      const c = along(fast.axis, f, fast.n, fast.sign);
      const x = slow.axis === 'Y' ? c : a;
      const y = slow.axis === 'Y' ? a : c;
      const e = line[f * 4 + 3];
      if (!e) continue;
      const k = Math.pow(2, e - 136);
      let r = (line[f * 4] + 0.5) * k;
      let g = (line[f * 4 + 1] + 0.5) * k;
      let bl = (line[f * 4 + 2] + 0.5) * k;
      if (h.xyz) {
        // CIE XYZ to linear sRGB.
        const X = r;
        const Y = g;
        const Z = bl;
        r = Math.max(0, 3.2406 * X - 1.5372 * Y - 0.4986 * Z);
        g = Math.max(0, -0.9689 * X + 1.8758 * Y + 0.0415 * Z);
        bl = Math.max(0, 0.0557 * X - 0.204 * Y + 1.057 * Z);
      }
      const o = (y * width + x) * 3;
      light[o] = r;
      light[o + 1] = g;
      light[o + 2] = bl;
    }
  }
  return { width, height, rgba: toneMap(light, width, height), pages: 1, light };
}

/** Linear light (RGB floats) as a Radiance HDR, rows packed: for the tests and the fuzzer's good files. */
export function encodeHdr(light, width, height) {
  const head = `#?RADIANCE\nFORMAT=32-bit_rle_rgbe\n\n-Y ${height} +X ${width}\n`;
  const out = [...head].map((c) => c.charCodeAt(0));
  const rgbe = new Uint8Array(width * 4);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 3;
      const m = Math.max(light[i], light[i + 1], light[i + 2]);
      if (m < 1e-32) { rgbe.fill(0, x * 4, x * 4 + 4); continue; }
      const e = Math.floor(Math.log2(m)) + 1; // m / 2^e in [0.5, 1)
      const f = 256 / Math.pow(2, e);
      rgbe[x * 4] = Math.min(255, Math.floor(light[i] * f));
      rgbe[x * 4 + 1] = Math.min(255, Math.floor(light[i + 1] * f));
      rgbe[x * 4 + 2] = Math.min(255, Math.floor(light[i + 2] * f));
      rgbe[x * 4 + 3] = e + 128;
    }
    if (width < 8 || width >= 32768) { out.push(...rgbe); continue; }
    out.push(2, 2, width >> 8, width & 255);
    for (let c = 0; c < 4; c++) {
      for (let x = 0; x < width;) {
        let n = 1;
        while (x + n < width && n < 127 && rgbe[(x + n) * 4 + c] === rgbe[x * 4 + c]) n++;
        if (n >= 3) { out.push(128 + n, rgbe[x * 4 + c]); x += n; continue; }
        // A literal run up to the next run of three.
        let m = 0;
        while (x + m < width && m < 128) {
          const at = x + m;
          if (at + 2 < width && rgbe[at * 4 + c] === rgbe[(at + 1) * 4 + c] && rgbe[at * 4 + c] === rgbe[(at + 2) * 4 + c]) break;
          m++;
        }
        m = Math.max(1, m);
        out.push(m);
        for (let k = 0; k < m; k++) out.push(rgbe[(x + k) * 4 + c]);
        x += m;
      }
    }
  }
  return Uint8Array.from(out);
}
