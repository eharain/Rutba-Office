// SPDX-License-Identifier: AGPL-3.0-or-later
//
// A Windows drawing, played into an SVG.
//
// An EMF or a WMF is a list of calls to Windows' drawing library: choose a
// pen, choose a brush, draw a rectangle, write these words at that point.
// This is the device those calls are played onto. It keeps what Windows
// keeps between calls (the pen, brush and font chosen, the colours, how the
// drawing's units map onto the page, the clip, the path being built) and
// writes each drawing call as SVG in the picture's own device units, so a
// window draws it sharp at any size. emf.js and wmf.js read their records
// and call it.
//
// What it writes is a picture and nothing more: paths, words, PNG or JPEG
// data, clip paths, patterns and gradients, in one string. No script, no
// event attribute, no reference to anything outside it.
//
// Everything is bounded: the elements written, the points in them, the
// pixels of every bitmap and their total, the depth of saved states and of
// clips. A drawing that reaches a bound stops there and keeps what it drew.

import zlib from 'node:zlib';
import { decodeIn } from './codepage.js';

export const LIMITS = {
  bytes: 64 * 1024 * 1024,
  records: 1_000_000,
  elements: 250_000,
  points: 4_000_000,
  saves: 256,
  clipDepth: 48,
  objects: 65_536,
  imagePixels: 16 * 1024 * 1024,
  imageOut: 4 * 1024 * 1024,
  imageBytes: 96 * 1024 * 1024,
  output: 128 * 1024 * 1024,
};

/** Thrown to stop playing: a bound reached. What was drawn is kept. */
export class Enough extends Error {}

/* ── small things ─────────────────────────────────────────────────────── */

/** A number as SVG writes it: two places, no exponent, no minus zero. */
export const num = (v) => {
  if (!Number.isFinite(v)) return '0';
  const c = Math.max(-1e9, Math.min(1e9, v));
  const n = Math.round(c * 100) / 100;
  return Object.is(n, -0) ? '0' : String(n);
};

/** A COLORREF (0x00BBGGRR; a palette index read as its low bytes) as #rrggbb. */
export const hex = (c) => '#' + [c & 255, (c >>> 8) & 255, (c >>> 16) & 255].map((v) => v.toString(16).padStart(2, '0')).join('');

/** Text XML can carry: a character XML 1.0 has no place for becomes a space, so positions still line up. */
export function xmlText(s) {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length) {
      const d = s.charCodeAt(i + 1);
      if (d >= 0xdc00 && d <= 0xdfff) { out += s[i] + s[i + 1]; i++; continue; }
    }
    if ((c < 0x20) || (c >= 0xd800 && c <= 0xdfff) || c === 0xfffe || c === 0xffff) out += ' ';
    else if (c === 0x26) out += '&amp;';
    else if (c === 0x3c) out += '&lt;';
    else if (c === 0x3e) out += '&gt;';
    else if (c === 0x22) out += '&quot;';
    else out += s[i];
  }
  return out;
}

/** A font's name, kept to what a font's name is made of. */
const faceName = (s) => String(s || '').replace(/[^\p{L}\p{N} ._-]/gu, '').trim().slice(0, 64);

/** The code page a LOGFONT's character set writes its 8-bit text in. */
export const CHARSET_CODEPAGE = {
  0: 'windows-1252', 1: 'windows-1252', 77: 'macintosh', 128: 'shift_jis', 129: 'euc-kr', 130: 'euc-kr', 134: 'gbk', 136: 'big5',
  161: 'windows-1253', 162: 'windows-1254', 163: 'windows-1258', 177: 'windows-1255', 178: 'windows-1256', 186: 'windows-1257',
  204: 'windows-1251', 222: 'windows-874', 238: 'windows-1250', 255: 'cp437',
};

/** 8-bit text in a font's character set: a symbol font's bytes are its private-use letters, as Windows maps them. */
export function decodeAnsi(bytes, charset = 0) {
  if (charset === 2) { let s = ''; for (const b of bytes) s += String.fromCharCode(b < 0x20 ? b : 0xf000 + b); return s; }
  try { return decodeIn(CHARSET_CODEPAGE[charset] || 'windows-1252', bytes); } catch { return decodeIn('windows-1252', bytes); }
}

/**
 * How far above and below the baseline a font's cell reaches, as a share of
 * its em: Windows places words by the cell's top unless told otherwise, and
 * a cell's height is the em and its leading. Measured from each font's own
 * win metrics; a font not listed is taken as Arial.
 */
const FACE_METRICS = {
  arial: [0.905, 0.212], 'arial narrow': [0.936, 0.212], 'arial black': [1.1, 0.31], helvetica: [0.905, 0.212],
  'times new roman': [0.891, 0.216], times: [0.891, 0.216], calibri: [0.952, 0.269], 'calibri light': [0.952, 0.269],
  cambria: [0.95, 0.222], 'segoe ui': [1.079, 0.251], 'segoe ui semibold': [1.079, 0.251], verdana: [1.005, 0.21],
  tahoma: [1.0, 0.207], 'courier new': [0.833, 0.3], courier: [0.833, 0.3], symbol: [1.005, 0.219], wingdings: [0.899, 0.211],
  georgia: [0.917, 0.219], 'trebuchet ms': [0.939, 0.222], consolas: [0.743, 0.257], 'century gothic': [1.006, 0.22],
  garamond: [0.861, 0.263], 'comic sans ms': [1.102, 0.287], 'lucida console': [0.789, 0.211], 'microsoft sans serif': [0.922, 0.21],
  'ms sans serif': [0.922, 0.21], 'ms shell dlg': [0.922, 0.21], 'ms shell dlg 2': [1.0, 0.207], 'franklin gothic medium': [0.894, 0.2],
  'book antiqua': [0.86, 0.264], 'palatino linotype': [1.026, 0.282], aptos: [0.939, 0.272], 'bookman old style': [0.908, 0.235],
  'century schoolbook': [0.932, 0.223], 'lucida sans unicode': [1.1, 0.32], impact: [1.009, 0.212], candara: [0.952, 0.269],
  constantia: [0.952, 0.269], corbel: [0.952, 0.269], 'gill sans mt': [0.918, 0.23], 'yu gothic': [1.055, 0.299],
  'ms gothic': [0.859, 0.141], 'ms mincho': [0.859, 0.141], simsun: [0.859, 0.141], 'malgun gothic': [1.0, 0.25], mangal: [1.1, 0.4],
  'traditional arabic': [1.07, 0.5], 'simplified arabic': [1.12, 0.5], marlett: [1.0, 0.0], webdings: [0.8, 0.2],
};
const metricsOf = (face) => FACE_METRICS[String(face || '').toLowerCase()] || FACE_METRICS[String(face || '').toLowerCase().replace(/ (bold|italic|light|semibold)$/, '')] || [0.905, 0.212];

/** What a font is drawn with when its own is missing, from its pitch and family. */
function fallbackFamily(font) {
  const family = (font.pitchFamily ?? 0) & 0xf0;
  if (family === 0x10) return 'serif';
  if (family === 0x30 || ((font.pitchFamily ?? 0) & 3) === 1) return 'monospace';
  if (family === 0x40) return 'cursive';
  const face = String(font.face || '').toLowerCase();
  if (/times|roman|georgia|garamond|cambria|book|serif|mincho|song|palatino|century/.test(face)) return 'serif';
  if (/courier|consolas|mono|console|fixed/.test(face)) return 'monospace';
  return 'sans-serif';
}

/* ── affine matrices: [a b c d e f], x' = a x + c y + e, y' = b x + d y + f ── */

export const IDENTITY = Object.freeze([1, 0, 0, 1, 0, 0]);
/** A after B: B applied first. */
export const mul = (A, B) => [
  A[0] * B[0] + A[2] * B[1], A[1] * B[0] + A[3] * B[1],
  A[0] * B[2] + A[2] * B[3], A[1] * B[2] + A[3] * B[3],
  A[0] * B[4] + A[2] * B[5] + A[4], A[1] * B[4] + A[3] * B[5] + A[5],
];

/* ── pictures: DIBs read, PNGs written ────────────────────────────────── */

/** A PNG of RGBA pixels. */
export function encodePng(w, h, rgba) {
  const table = encodePng.table || (encodePng.table = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  }));
  const crc = (buf) => {
    let c = 0xffffffff;
    for (const byte of buf) c = table[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const head = Buffer.alloc(8);
    head.writeUInt32BE(data.length, 0);
    head.write(type, 4, 'latin1');
    const tail = Buffer.alloc(4);
    tail.writeUInt32BE(crc(Buffer.concat([head.subarray(4), data])), 0);
    return Buffer.concat([head, data, tail]);
  };
  const raw = Buffer.alloc((w * 4 + 1) * h);
  for (let y = 0; y < h; y++) {
    raw[y * (w * 4 + 1)] = 0;
    rgba.copy(raw, y * (w * 4 + 1) + 1, y * w * 4, (y + 1) * w * 4);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

/** Run-length coded indices (RLE8, or RLE4 with `four`), as one byte a pixel, rows from the bottom. */
function unRle(b, from, to, w, h, four) {
  const out = new Uint8Array(w * h);
  let x = 0;
  let y = 0;
  let p = from;
  while (p + 1 < to && y < h) {
    const n = b[p];
    const v = b[p + 1];
    p += 2;
    if (n > 0) {
      for (let i = 0; i < n && x < w; i++, x++) out[y * w + x] = four ? (i & 1 ? v & 15 : v >> 4) : v;
    } else if (v === 0) { x = 0; y++; } else if (v === 1) break;
    else if (v === 2) { if (p + 1 >= to) break; x += b[p]; y += b[p + 1]; p += 2; } else {
      const bytes = four ? (v + 1) >> 1 : v;
      for (let i = 0; i < v && x < w; i++, x++) {
        const at = p + (four ? i >> 1 : i);
        if (at >= to) break;
        out[y * w + x] = four ? (i & 1 ? b[at] & 15 : b[at] >> 4) : b[at];
      }
      p += bytes + (bytes & 1);
    }
  }
  return out;
}

/**
 * A device-independent bitmap: `{ w, h, topDown, pixel(x, yFromTop) ->
 * 0xRRGGBB, alpha(x, yFromTop) }`, or `{ raw, type }` for the JPEG or PNG
 * a DIB can carry, or null for a form this does not read (colours that are
 * indices into the device's palette).
 */
export function readDib(b, offBmi, offBits, usage, cbBits = Infinity) {
  if (!(offBmi >= 0 && offBmi + 12 <= b.length && offBits >= 0 && offBits <= b.length)) return null;
  if (usage !== 0) return null;
  const size = b.readUInt32LE(offBmi);
  let w; let hRaw; let bpp; let compression = 0; let used = 0; let entry = 4;
  if (size === 12) {
    w = b.readUInt16LE(offBmi + 4);
    hRaw = b.readInt16LE(offBmi + 6);
    bpp = b.readUInt16LE(offBmi + 10);
    entry = 3;
  } else if (size >= 40 && offBmi + 40 <= b.length) {
    w = b.readInt32LE(offBmi + 4);
    hRaw = b.readInt32LE(offBmi + 8);
    bpp = b.readUInt16LE(offBmi + 14);
    compression = b.readUInt32LE(offBmi + 16);
    used = b.readUInt32LE(offBmi + 32);
  } else return null;
  const h = Math.abs(hRaw);
  const topDown = hRaw < 0;
  if (!(w > 0 && h > 0) || w * h > LIMITS.imagePixels) return null;
  const end = Math.min(b.length, offBits + (Number.isFinite(cbBits) && cbBits > 0 ? cbBits : b.length));
  if (compression === 4 || compression === 5) {
    const raw = b.subarray(offBits, end);
    const type = compression === 4 ? 'image/jpeg' : 'image/png';
    const ok = compression === 4 ? raw[0] === 0xff && raw[1] === 0xd8 : raw[0] === 0x89 && raw[1] === 0x50;
    return ok ? { raw, type, w, h, topDown } : null;
  }
  if (![1, 4, 8, 16, 24, 32].includes(bpp)) return null;
  const rle = (compression === 1 && bpp === 8) || (compression === 2 && bpp === 4);
  if (compression !== 0 && !rle && !(compression === 3 && (bpp === 16 || bpp === 32))) return null;
  const stride = (((w * bpp) + 31) >> 5) * 4;
  if (!rle && offBits + stride * h > b.length) return null;
  const palette = [];
  if (bpp <= 8) {
    const n = Math.min(used || 1 << bpp, 256);
    const from = offBmi + size;
    for (let i = 0; i < n; i++) {
      const at = from + i * entry;
      palette.push(at + 2 < b.length ? (b[at + 2] << 16) | (b[at + 1] << 8) | b[at] : 0);
    }
  }
  let masks = bpp === 16 ? [0x7c00, 0x03e0, 0x001f] : [0xff0000, 0x00ff00, 0x0000ff];
  if (compression === 3) {
    if (offBmi + 52 > b.length) return null;
    masks = [0, 4, 8].map((k) => b.readUInt32LE(offBmi + 40 + k));
  }
  const channel = (v, mask) => {
    if (!mask) return 0;
    const low = 31 - Math.clz32(mask & -mask);
    const bits = 32 - Math.clz32(mask >>> low);
    const raw = (v & mask) >>> low;
    return bits >= 8 ? raw >>> (bits - 8) : Math.round((raw * 255) / ((1 << bits) - 1));
  };
  const indices = rle ? unRle(b, offBits, end, w, h, bpp === 4) : null;
  const pixel = (x, y) => {
    if (indices) return palette[indices[(h - 1 - y) * w + x]] ?? 0;
    const row = offBits + (topDown ? y : h - 1 - y) * stride;
    switch (bpp) {
      case 1: return palette[(b[row + (x >> 3)] >> (7 - (x & 7))) & 1] ?? 0;
      case 4: return palette[(b[row + (x >> 1)] >> ((x & 1) ? 0 : 4)) & 15] ?? 0;
      case 8: return palette[b[row + x]] ?? 0;
      case 16: {
        const v = b.readUInt16LE(row + x * 2);
        return (channel(v, masks[0]) << 16) | (channel(v, masks[1]) << 8) | channel(v, masks[2]);
      }
      case 24: { const at = row + x * 3; return (b[at + 2] << 16) | (b[at + 1] << 8) | b[at]; }
      default: {
        const v = b.readUInt32LE(row + x * 4);
        return compression === 3 ? (channel(v, masks[0]) << 16) | (channel(v, masks[1]) << 8) | channel(v, masks[2]) : v & 0xffffff;
      }
    }
  };
  // A 32-bit bitmap's alpha is the byte its colour masks leave over (the top one, usually).
  const alphaMask = bpp === 32 ? ~(masks[0] | masks[1] | masks[2]) >>> 0 : 0;
  const alpha = alphaMask ? (x, y) => channel(b.readUInt32LE(offBits + (topDown ? y : h - 1 - y) * stride + x * 4), alphaMask) : null;
  return { w, h, topDown, pixel, alpha, bpp, mono: bpp === 1 };
}

/**
 * A raster operation on 24-bit colours, from its ternary code: each bit of
 * the code is the answer for one combination of pattern, source and
 * destination bits. The pattern (the brush) is not followed, so an
 * operation that reads it is not played.
 */
export function ropOf(rop) {
  const code = (rop >>> 16) & 0xff;
  if (((code >> 4) & 0x0f) !== (code & 0x0f)) return null;
  const usesSource = ((code >> 2) & 3) !== (code & 3);
  return {
    usesSource,
    apply: (s, d) => {
      let out = 0;
      for (let i = 0; i < 4; i++) {
        if (!((code >> i) & 1)) continue;
        out |= ((i & 2) ? s : ~s) & ((i & 1) ? d : ~d);
      }
      return out & 0xffffff;
    },
  };
}

export const ROP = { SRCCOPY: 0xcc0020, NOTSRCCOPY: 0x330008, PATCOPY: 0xf00021, BLACKNESS: 0x000042, WHITENESS: 0xff0062, DSTCOPY: 0xaa0029 };

/* ── the objects a drawing selects ────────────────────────────────────── */

const pen = (color, style = 0, width = 0) => ({ kind: 'pen', style, width, color, cosmetic: width === 0 });
const brush = (color, style = 0, hatch = 0) => ({ kind: 'brush', style, color, hatch });
const font = (face, height, weight = 400, pitchFamily = 0) => ({ kind: 'font', face, height, width: 0, escapement: 0, orientation: 0, weight, italic: false, underline: false, strike: false, charset: 0, pitchFamily });

/** Windows' stock objects, by number. */
export const STOCK = {
  0: brush(0xffffff), 1: brush(0xc0c0c0), 2: brush(0x808080), 3: brush(0x404040), 4: brush(0), 5: brush(0, 1),
  6: pen(0xffffff), 7: pen(0), 8: pen(0, 5),
  10: font('Courier New', 12, 400, 0x31), 11: font('Courier New', 13, 400, 0x31), 12: font('Arial', 13), 13: font('Segoe UI', 16, 700),
  14: font('Segoe UI', 16, 700), 16: font('Courier New', 16, 400, 0x31), 17: font('Segoe UI', -11),
  18: brush(0xffffff), 19: pen(0),
};

/** A pen from a LOGPEN or a LOGPEN_EX. */
export function makePen({ style, width, color, ext = false, entries = null, brushStyle = 0 }) {
  const geometric = ext ? Boolean(style & 0x10000) : width > 0;
  return { kind: 'pen', style, width: geometric ? Math.abs(width) : 0, color: color & 0xffffff, cosmetic: !geometric, ext, entries, nullBrush: ext && brushStyle === 1 };
}

/* ── the device ───────────────────────────────────────────────────────── */

const FIXED_UNITS_MM = { 2: 0.1, 3: 0.01, 4: 0.254, 5: 0.0254, 6: 25.4 / 1440 };
const K = 0.5522847498307936;

/** A path's commands, in device units. */
class Figure {
  constructor(player) { this.p = player; this.parts = []; this.open = false; this.start = null; this.last = null; this.widened = false; }
  move(q) { this.p.point(); this.parts.push('M' + num(q[0]) + ' ' + num(q[1])); this.open = true; this.start = q; this.last = q; }
  line(q) { if (!this.open) this.move(this.last || q); this.p.point(); this.parts.push('L' + num(q[0]) + ' ' + num(q[1])); this.last = q; }
  curve(a, b, c) {
    if (!this.open) this.move(this.last || a);
    this.p.point(3);
    this.parts.push('C' + num(a[0]) + ' ' + num(a[1]) + ' ' + num(b[0]) + ' ' + num(b[1]) + ' ' + num(c[0]) + ' ' + num(c[1]));
    this.last = c;
  }
  close() { if (this.open) { this.parts.push('Z'); this.open = false; this.last = this.start; } }
  closeAll() {
    const out = [];
    for (const part of this.parts) {
      if (part[0] === 'M' && out.length && out[out.length - 1] !== 'Z') out.push('Z');
      out.push(part);
    }
    if (out.length && out[out.length - 1] !== 'Z') out.push('Z');
    this.parts = out;
    this.open = false;
  }
  get d() { return this.parts.join(''); }
  get empty() { return this.parts.length === 0; }
}

/**
 * The device: a Windows device context, drawn as SVG.
 *
 * `ppmm` is the reference device's pixels per millimetre, which the fixed
 * map modes (millimetres, inches, twips) are measured against.
 */
export class GdiSvg {
  constructor({ ppmm = [96 / 25.4, 96 / 25.4] } = {}) {
    this.ppmm = ppmm;
    // Device units in a screen pixel: a hatch is eight pixels a side, whatever the device's units.
    this.px = (ppmm[0] * 25.4) / 96 || 1;
    this.dc = {
      mapMode: 1, winOrg: [0, 0], winExt: [1, 1], vpOrg: [0, 0], vpExt: [1, 1], world: IDENTITY,
      pen: STOCK[7], brush: STOCK[0], font: STOCK[13], textColor: 0, bkColor: 0xffffff, bkMode: 2, polyFill: 1,
      textAlign: 0, rop2: 13, arcDir: 1, miter: 10, clip: null, cur: [0, 0], charExtra: 0,
    };
    this.saved = [];
    this.objects = new Map();
    this.body = [];
    this.defs = [];
    this.size = 0;
    this.elements = 0;
    this.points = 0;
    this.imageBytes = 0;
    this.openClip = null;
    this.ids = 0;
    this.path = null;
    this.inPath = false;
    this.line = null;
    this.raster = null;
    this.hatches = new Map();
    this.drew = 0;
  }

  /* ── state ── */

  point(n = 1) { this.points += n; if (this.points > LIMITS.points) throw new Enough('points'); }

  /** Logical units to the page: the map mode's scale, the window and the viewport. */
  pageScale() {
    const d = this.dc;
    const fixed = FIXED_UNITS_MM[d.mapMode];
    if (fixed) return [fixed * this.ppmm[0], -fixed * this.ppmm[1]];
    if (d.mapMode === 7 || d.mapMode === 8) {
      let sx = d.vpExt[0] / (d.winExt[0] || 1);
      let sy = d.vpExt[1] / (d.winExt[1] || 1);
      if (d.mapMode === 7) {
        const s = Math.min(Math.abs(sx), Math.abs(sy));
        sx = (sx < 0 ? -1 : 1) * s;
        sy = (sy < 0 ? -1 : 1) * s;
      }
      return [sx || 1, sy || 1];
    }
    return [1, 1];
  }

  /** Logical units to device units: the world transform, then the page. */
  m() {
    if (this._m && this._mFor === this.dc) return this._m;
    const d = this.dc;
    const [sx, sy] = this.pageScale();
    const page = [sx, 0, 0, sy, d.vpOrg[0] - d.winOrg[0] * sx, d.vpOrg[1] - d.winOrg[1] * sy];
    this._m = d.world === IDENTITY ? page : mul(page, d.world);
    this._mFor = this.dc;
    return this._m;
  }

  changed() { this._mFor = null; }

  pt(x, y) { const m = this.m(); const k = this.nudge || 0; return [m[0] * x + m[2] * y + m[4] + k, m[1] * x + m[3] * y + m[5] + k]; }

  /**
   * A one-pixel pen draws the pixels to the right of and below its points,
   * so a line between whole pixels lights one row of them, not half of two:
   * its points move half a pixel when nothing is filled with them.
   */
  hairline() {
    const p = this.dc.pen;
    return !this.inPath && (p.style & 0xf) !== 5 && (p.cosmetic || (p.width * this.scale()) / this.px <= 1.0001);
  }

  /** `fn` with a hairline's points moved onto the pixels it lights. */
  nudged(fn) {
    if (!this.hairline()) return fn();
    this.nudge = 0.5 * this.px;
    try { return fn(); } finally { this.nudge = 0; }
  }

  /** How much longer a logical unit is on the device, on average. */
  scale() { const m = this.m(); return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1; }

  setMapMode(mode) {
    const d = this.dc;
    if (!(mode >= 1 && mode <= 8)) return;
    // Into the scaled modes from a fixed one, the scale it had is kept as the extents.
    if ((mode === 7 || mode === 8) && !(d.mapMode === 7 || d.mapMode === 8)) {
      const [sx, sy] = this.pageScale();
      d.winExt = [1, 1];
      d.vpExt = [sx, sy];
    }
    d.mapMode = mode;
    this.changed();
  }

  /** Window and viewport extents mean something only in the two scaled modes, as in Windows. */
  setWindowExt(x, y) { if ((this.dc.mapMode === 7 || this.dc.mapMode === 8) && x && y) { this.dc.winExt = [x, y]; this.changed(); } }
  setViewportExt(x, y) { if ((this.dc.mapMode === 7 || this.dc.mapMode === 8) && x && y) { this.dc.vpExt = [x, y]; this.changed(); } }
  setWindowOrg(x, y) { this.dc.winOrg = [x, y]; this.changed(); }
  setViewportOrg(x, y) { this.dc.vpOrg = [x, y]; this.changed(); }
  offsetWindowOrg(x, y) { this.dc.winOrg = [this.dc.winOrg[0] + x, this.dc.winOrg[1] + y]; this.changed(); }
  offsetViewportOrg(x, y) { this.dc.vpOrg = [this.dc.vpOrg[0] + x, this.dc.vpOrg[1] + y]; this.changed(); }
  scaleWindowExt(xn, xd, yn, yd) { if (xd && yd) this.setWindowExt(this.dc.winExt[0] * xn / xd, this.dc.winExt[1] * yn / yd); }
  scaleViewportExt(xn, xd, yn, yd) { if (xd && yd) this.setViewportExt(this.dc.vpExt[0] * xn / xd, this.dc.vpExt[1] * yn / yd); }

  setWorld(x) { if (x.every(Number.isFinite)) { this.dc.world = x; this.changed(); } }
  /** MWT_IDENTITY 1, LEFTMULTIPLY 2 (the new one first), RIGHTMULTIPLY 3, SET 4. */
  modifyWorld(x, mode) {
    if (mode === 1) this.setWorld(IDENTITY);
    else if (!x.every(Number.isFinite)) return;
    else if (mode === 2) this.setWorld(mul(this.dc.world, x));
    else if (mode === 3) this.setWorld(mul(x, this.dc.world));
    else if (mode === 4) this.setWorld(x);
  }

  save() {
    if (this.saved.length >= LIMITS.saves) return;
    this.saved.push({ ...this.dc });
  }

  /** A negative count goes back that many saves; a positive one names a save. */
  restore(n) {
    const back = n < 0 ? -n : this.saved.length - n + 1;
    if (!(back > 0 && back <= this.saved.length)) return;
    this.flushLine();
    this.dc = this.saved.splice(this.saved.length - back)[0];
    this.changed();
  }

  setObject(i, o) {
    if (!(i >= 0 && i < LIMITS.objects)) return;
    this.objects.set(i, o);
  }

  select(i) {
    const o = i >= 0x80000000 ? STOCK[i & 0x7fffffff] : this.objects.get(i);
    if (!o) return;
    if (o.kind === 'pen') { this.flushLine(); this.dc.pen = o; } else if (o.kind === 'brush') this.dc.brush = o;
    else if (o.kind === 'font') this.dc.font = o;
  }

  remove(i) { this.objects.delete(i); }

  set(key, value) { if (key === 'textAlign' || key === 'bkMode' || key === 'rop2') this.flushLine(); this.dc[key] = value; }

  /* ── writing ── */

  push(s) {
    this.size += s.length;
    if (this.size > LIMITS.output) throw new Enough('output');
    this.body.push(s);
  }

  /** One element, inside the clip it was drawn under. */
  element(svg, clip = this.dc.clip, fromRaster = false) {
    if (!fromRaster && this.raster) this.flushRaster();
    if (this.line && !this.flushingLine) this.flushLine();
    this.elements += 1;
    if (this.elements > LIMITS.elements) throw new Enough('elements');
    const id = clip ? this.clipId(clip) : null;
    if (id !== this.openClip) {
      if (this.openClip) this.push('</g>');
      if (id) this.push(`<g clip-path="url(#${id})">`);
      this.openClip = id;
    }
    this.push(svg);
    this.drew += 1;
  }

  id(prefix) { this.ids += 1; return prefix + this.ids; }

  def(s) { this.size += s.length; if (this.size > LIMITS.output) throw new Enough('output'); this.defs.push(s); }

  clipId(node) {
    if (node.id) return node.id;
    const parent = node.parent ? this.clipId(node.parent) : null;
    node.id = this.id('c');
    this.def(`<clipPath id="${node.id}" clipPathUnits="userSpaceOnUse"${parent ? ` clip-path="url(#${parent})"` : ''}><path d="${node.d}"${node.rule === 'evenodd' ? ' clip-rule="evenodd"' : ''}/></clipPath>`);
    return node.id;
  }

  /** A pen as stroke attributes, or null when it draws nothing. */
  strokeAttrs(p = this.dc.pen) {
    const st = p.style & 0xf;
    if (st === 5 || this.dc.rop2 === 11 || p.nullBrush) return null;
    const color = this.dc.rop2 === 1 ? 0 : this.dc.rop2 === 16 ? 0xffffff : p.color;
    const w = p.width * this.scale();
    // A cosmetic pen is one pixel wide however the picture is scaled, as is
    // any pen narrower than a pixel: Windows never draws a thinner line.
    const hair = p.cosmetic || w / this.px < 0.999;
    let out = ` stroke="${hex(color)}"`;
    out += hair ? ' stroke-width="1" vector-effect="non-scaling-stroke"' : ` stroke-width="${num(w)}"`;
    let cap = 'round';
    let join = 'round';
    // A pen a pixel wide draws square to its ends, as Windows draws it.
    if (hair || w / this.px <= 1.0001) { cap = 'butt'; join = 'miter'; } else if (p.ext) {
      cap = { 0: 'round', 0x100: 'square', 0x200: 'butt' }[p.style & 0xf00] || 'round';
      join = { 0: 'round', 0x1000: 'bevel', 0x2000: 'miter' }[p.style & 0xf000] || 'round';
    }
    if (cap !== 'butt') out += ` stroke-linecap="${cap}"`;
    if (join !== 'miter') out += ` stroke-linejoin="${join}"`;
    else if (this.dc.miter && this.dc.miter !== 4) out += ` stroke-miterlimit="${num(Math.max(1, this.dc.miter))}"`;
    // Dashes: a cosmetic pen's in pixels, a geometric pen's in its own widths.
    let dash = null;
    const unit = hair ? 1 : Math.max(w, 1);
    const styled = hair || p.ext;
    if (styled && st >= 1 && st <= 4) dash = (hair ? [[18, 6], [3, 3], [9, 6, 3, 6], [9, 3, 3, 3, 3, 3]] : [[3, 1], [1, 1], [3, 1, 1, 1], [3, 1, 1, 1, 1, 1]])[st - 1].map((v) => v * unit);
    else if (st === 8) dash = [1, 1];
    else if (st === 7 && p.entries?.length) dash = p.entries.map((v) => Math.max(0.1, Math.abs(v) * (hair ? 1 : this.scale())));
    if (dash) out += ` stroke-dasharray="${dash.map(num).join(' ')}"`;
    return out;
  }

  /** A brush as a fill, or 'none'. */
  fillOf(b = this.dc.brush) {
    if (!b || b.style === 1 || this.dc.rop2 === 11) return 'none';
    if (this.dc.rop2 === 1) return '#000000';
    if (this.dc.rop2 === 16) return '#ffffff';
    if (b.style === 2 && b.hatch >= 0 && b.hatch <= 5) return `url(#${this.hatchId(b)})`;
    if (b.pattern) return `url(#${this.patternId(b)})`;
    return hex(b.color ?? 0);
  }

  hatchId(b) {
    const opaque = this.dc.bkMode === 2;
    const key = `${b.hatch}:${b.color}:${opaque ? this.dc.bkColor : 'none'}`;
    if (this.hatches.has(key)) return this.hatches.get(key);
    const id = this.id('h');
    const lines = [['M0 4.5H8'], ['M4.5 0V8'], ['M-1 -1L9 9'], ['M-1 9L9 -1'], ['M0 4.5H8', 'M4.5 0V8'], ['M-1 -1L9 9', 'M-1 9L9 -1']][b.hatch];
    const k = num(8 * this.px);
    this.def(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="${k}" height="${k}"><g transform="scale(${num(this.px)})">${opaque ? `<rect width="8" height="8" fill="${hex(this.dc.bkColor)}"/>` : ''}<path d="${lines.join('')}" stroke="${hex(b.color)}" stroke-width="1" fill="none"/></g></pattern>`);
    this.hatches.set(key, id);
    return id;
  }

  patternId(b) {
    if (b.patternId) return b.patternId;
    const id = this.id('p');
    const { href, w, h } = b.pattern;
    const W = num(w * this.px);
    const H = num(h * this.px);
    this.def(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="${W}" height="${H}"><image width="${W}" height="${H}" href="${href}"/></pattern>`);
    b.patternId = id;
    return id;
  }

  /** A brush made from a bitmap: its tile, or its average colour when the tile is large. */
  patternBrush(dib) {
    if (!dib || dib.raw) return brush(0x808080);
    if (dib.w * dib.h > 64 * 64) {
      let r = 0; let g = 0; let bl = 0; let n = 0;
      for (let y = 0; y < dib.h; y += Math.ceil(dib.h / 16)) for (let x = 0; x < dib.w; x += Math.ceil(dib.w / 16)) { const c = dib.pixel(x, y); r += c >> 16 & 255; g += c >> 8 & 255; bl += c & 255; n++; }
      return brush(Math.round(bl / n) << 16 | Math.round(g / n) << 8 | Math.round(r / n));
    }
    const png = this.dibPng(dib, 0, 0, dib.w, dib.h);
    if (!png) return brush(0x808080);
    return { ...brush(0, 3), pattern: { href: png, w: dib.w, h: dib.h } };
  }

  /* ── shapes ── */

  /** Where a shape goes: into the path being built, or a figure of its own drawn now. */
  figure() { return this.inPath ? this.path : new Figure(this); }

  finishShape(f, { fill, stroke }) {
    if (this.inPath || f.empty) return;
    const fill2 = fill ? this.fillOf() : 'none';
    const st = stroke ? this.strokeAttrs() : null;
    if (fill2 === 'none' && !st) return;
    const rule = fill2 !== 'none' && this.dc.polyFill === 1 ? ' fill-rule="evenodd"' : '';
    this.element(`<path d="${f.d}" fill="${fill2}"${rule}${st || ''}/>`);
  }

  moveTo(x, y) {
    this.dc.cur = [x, y];
    if (this.inPath) this.path.move(this.pt(x, y));
  }

  /** A line from the current position. Consecutive lines with one pen are one path. */
  lineTo(x, y) {
    if (this.raster) this.flushRaster();
    const [from, to] = this.nudged(() => [this.pt(...this.dc.cur), this.pt(x, y)]);
    this.dc.cur = [x, y];
    if (this.inPath) { if (!this.path.open) this.path.move(from); this.path.line(to); return; }
    const st = this.strokeAttrs();
    if (!st) return;
    if (this.line && this.line.st === st && this.line.clip === this.dc.clip && this.line.f.last[0] === from[0] && this.line.f.last[1] === from[1]) {
      this.line.f.line(to);
      return;
    }
    this.flushLine();
    const f = new Figure(this);
    f.move(from);
    f.line(to);
    this.line = { f, st, clip: this.dc.clip };
  }

  flushLine() {
    const l = this.line;
    if (!l) return;
    this.line = null;
    this.flushingLine = true;
    try { this.element(`<path d="${l.f.d}" fill="none"${l.st}/>`, l.clip); } finally { this.flushingLine = false; }
  }

  polyline(pts, opts) { this.nudged(() => this.polyline1(pts, opts)); }

  polyline1(pts, { to = false } = {}) {
    if (!pts.length) return;
    const f = this.figure();
    if (to) {
      if (!this.inPath || !f.open) f.move(this.pt(...this.dc.cur));
    } else f.move(this.pt(...pts[0]));
    for (let i = to ? 0 : 1; i < pts.length; i++) f.line(this.pt(...pts[i]));
    if (to) this.dc.cur = pts[pts.length - 1];
    this.finishShape(f, { fill: false, stroke: true });
  }

  polyBezier(pts, opts) { this.nudged(() => this.polyBezier1(pts, opts)); }

  polyBezier1(pts, { to = false } = {}) {
    const f = this.figure();
    let i = 0;
    if (to) { if (!this.inPath || !f.open) f.move(this.pt(...this.dc.cur)); } else {
      if (!pts.length) return;
      f.move(this.pt(...pts[0]));
      i = 1;
    }
    for (; i + 2 < pts.length; i += 3) f.curve(this.pt(...pts[i]), this.pt(...pts[i + 1]), this.pt(...pts[i + 2]));
    if (to && pts.length >= 3) this.dc.cur = pts[Math.floor(pts.length / 3) * 3 - 1];
    this.finishShape(f, { fill: false, stroke: true });
  }

  /** Closed figures, filled under the fill mode, outlined with the pen. */
  polygons(polys) {
    const f = this.figure();
    for (const pts of polys) {
      if (pts.length < 2) continue;
      f.move(this.pt(...pts[0]));
      for (let i = 1; i < pts.length; i++) f.line(this.pt(...pts[i]));
      f.close();
    }
    this.finishShape(f, { fill: true, stroke: true });
  }

  polylines(polys) { this.nudged(() => this.polylines1(polys)); }

  polylines1(polys) {
    const f = this.figure();
    for (const pts of polys) {
      if (!pts.length) continue;
      f.move(this.pt(...pts[0]));
      for (let i = 1; i < pts.length; i++) f.line(this.pt(...pts[i]));
    }
    this.finishShape(f, { fill: false, stroke: true });
  }

  /** PolyDraw: each point a move, a line or a third of a curve, perhaps closing its figure. */
  polyDraw(pts, types) {
    const f = this.figure();
    if (!f.open) f.move(this.pt(...this.dc.cur));
    for (let i = 0; i < pts.length; i++) {
      const t = types[i] & ~1;
      if (t === 6) f.move(this.pt(...pts[i]));
      else if (t === 2) f.line(this.pt(...pts[i]));
      else if (t === 4 && i + 2 < pts.length) { f.curve(this.pt(...pts[i]), this.pt(...pts[i + 1]), this.pt(...pts[i + 2])); i += 2; }
      if (types[i] & 1) f.close();
      this.dc.cur = pts[i];
    }
    this.finishShape(f, { fill: false, stroke: true });
  }

  /** The inset an inside-frame pen draws its shape at, so its line stays inside the box. */
  inset() {
    const p = this.dc.pen;
    if (this.hairline()) return (0.5 * this.px) / this.scale();
    return (p.style & 0xf) === 6 && !p.cosmetic && p.width > 1 ? p.width / 2 : 0;
  }

  rectangle(l, t, r, b) {
    const k = this.inset();
    [l, r] = [Math.min(l, r) + k, Math.max(l, r) - k];
    [t, b] = [Math.min(t, b) + k, Math.max(t, b) - k];
    this.polygons([[[l, t], [r, t], [r, b], [l, b]]]);
  }

  /** A quarter ellipse at a time, as four Béziers, through the transform. */
  ellipseInto(f, cx, cy, rx, ry) {
    const P = (x, y) => this.pt(cx + rx * x, cy + ry * y);
    f.move(P(1, 0));
    f.curve(P(1, K), P(K, 1), P(0, 1));
    f.curve(P(-K, 1), P(-1, K), P(-1, 0));
    f.curve(P(-1, -K), P(-K, -1), P(0, -1));
    f.curve(P(K, -1), P(1, -K), P(1, 0));
    f.close();
  }

  ellipse(l, t, r, b) {
    const k = this.inset();
    const f = this.figure();
    this.ellipseInto(f, (l + r) / 2, (t + b) / 2, Math.abs(r - l) / 2 - k, Math.abs(b - t) / 2 - k);
    this.finishShape(f, { fill: true, stroke: true });
  }

  roundRect(l, t, r, b, w, h) {
    const k = this.inset();
    [l, r] = [Math.min(l, r) + k, Math.max(l, r) - k];
    [t, b] = [Math.min(t, b) + k, Math.max(t, b) - k];
    const rx = Math.min(Math.abs(w) / 2, (r - l) / 2);
    const ry = Math.min(Math.abs(h) / 2, (b - t) / 2);
    const f = this.figure();
    const P = (x, y) => this.pt(x, y);
    f.move(P(l + rx, t));
    f.line(P(r - rx, t));
    f.curve(P(r - rx + rx * K, t), P(r, t + ry - ry * K), P(r, t + ry));
    f.line(P(r, b - ry));
    f.curve(P(r, b - ry + ry * K), P(r - rx + rx * K, b), P(r - rx, b));
    f.line(P(l + rx, b));
    f.curve(P(l + rx - rx * K, b), P(l, b - ry + ry * K), P(l, b - ry));
    f.line(P(l, t + ry));
    f.curve(P(l, t + ry - ry * K), P(l + rx - rx * K, t), P(l + rx, t));
    f.close();
    this.finishShape(f, { fill: true, stroke: true });
  }

  /**
   * Part of an ellipse from angle a0 through `sweep` (radians, in the
   * ellipse's own terms: x = cx + rx cos a, y = cy + ry sin a), as Béziers
   * of a quarter turn or less. `join` draws a line to its start rather than
   * starting a figure there.
   */
  arcInto(f, cx, cy, rx, ry, a0, sweep, join) {
    const P = (a) => this.pt(cx + rx * Math.cos(a), cy + ry * Math.sin(a));
    const start = P(a0);
    if (join) f.line(start); else f.move(start);
    const n = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
    const step = sweep / n;
    const k = (4 / 3) * Math.tan(step / 4);
    for (let i = 0; i < n; i++) {
      const a = a0 + step * i;
      const b = a + step;
      const c1 = this.pt(cx + rx * (Math.cos(a) - k * Math.sin(a)), cy + ry * (Math.sin(a) + k * Math.cos(a)));
      const c2 = this.pt(cx + rx * (Math.cos(b) + k * Math.sin(b)), cy + ry * (Math.sin(b) - k * Math.cos(b)));
      f.curve(c1, c2, P(b));
    }
  }

  /**
   * Arc, chord, pie, arc-to: the ellipse in the box, from the ray through the
   * start point to the ray through the end, the way the arc direction says
   * (anticlockwise on the page unless set otherwise).
   */
  arc(kind, ...box) {
    if (kind === 'arc' || kind === 'arcto') this.nudged(() => this.arc1(kind, ...box));
    else this.arc1(kind, ...box);
  }

  arc1(kind, l, t, r, b, xs, ys, xe, ye) {
    const cx = (l + r) / 2;
    const cy = (t + b) / 2;
    const rx = Math.abs(r - l) / 2;
    const ry = Math.abs(b - t) / 2;
    if (!(rx > 0 && ry > 0)) return;
    const a0 = Math.atan2((ys - cy) / ry, (xs - cx) / rx);
    const a1 = Math.atan2((ye - cy) / ry, (xe - cx) / rx);
    const m = this.m();
    const flipped = m[0] * m[3] - m[1] * m[2] < 0;
    // Anticlockwise on the page is the angle falling while y runs down the page.
    let dir = flipped ? 1 : -1;
    if (this.dc.arcDir === 2) dir = -dir;
    let sweep = (a1 - a0) * dir;
    while (sweep <= 1e-9) sweep += 2 * Math.PI;
    while (sweep > 2 * Math.PI + 1e-9) sweep -= 2 * Math.PI;
    sweep *= dir;
    const f = this.figure();
    if (kind === 'pie') { f.move(this.pt(cx, cy)); this.arcInto(f, cx, cy, rx, ry, a0, sweep, true); f.close(); } else if (kind === 'chord') { this.arcInto(f, cx, cy, rx, ry, a0, sweep, false); f.close(); } else if (kind === 'arcto') {
      if (!f.open) f.move(this.pt(...this.dc.cur));
      this.arcInto(f, cx, cy, rx, ry, a0, sweep, true);
      this.dc.cur = [cx + rx * Math.cos(a0 + sweep), cy + ry * Math.sin(a0 + sweep)];
    } else this.arcInto(f, cx, cy, rx, ry, a0, sweep, this.inPath && f.open);
    this.finishShape(f, { fill: kind === 'pie' || kind === 'chord', stroke: true });
  }

  /** AngleArc: a line from the current position, then the circle's arc, angles anticlockwise in degrees. */
  angleArc(x, y, radius, startDeg, sweepDeg) {
    const m = this.m();
    const s = m[0] * m[3] - m[1] * m[2] < 0 ? 1 : -1;
    const a0 = s * startDeg * Math.PI / 180;
    // More than a turn draws the same circle again: one is enough.
    const sweep = s * Math.max(-360, Math.min(360, Number.isFinite(sweepDeg) ? sweepDeg : 0)) * Math.PI / 180;
    if (!Number.isFinite(a0) || !(radius >= 0)) return;
    const f = this.figure();
    if (!f.open) f.move(this.pt(...this.dc.cur));
    this.arcInto(f, x, y, radius, radius, a0, sweep, true);
    this.dc.cur = [x + radius * Math.cos(a0 + sweep), y + radius * Math.sin(a0 + sweep)];
    this.finishShape(f, { fill: false, stroke: true });
  }

  setPixel(x, y, color) {
    const [px, py] = this.pt(x, y);
    this.element(`<rect x="${num(px)}" y="${num(py)}" width="1" height="1" fill="${hex(color)}"/>`);
  }

  /* ── paths ── */

  beginPath() { this.flushLine(); this.path = new Figure(this); this.inPath = true; }
  endPath() { this.inPath = false; }
  abortPath() { this.path = null; this.inPath = false; }
  closeFigure() { if (this.inPath && this.path) this.path.close(); }
  widenPath() { if (this.path) this.path.widened = true; }

  /** FillPath, StrokePath, StrokeAndFillPath: the path drawn, then let go. */
  drawPath(fill, stroke) {
    const p = this.path;
    this.path = null;
    this.inPath = false;
    if (!p || p.empty) return;
    if (p.widened && fill && !stroke) {
      // A widened path is the pen's outline: filling it draws the line in the brush's colour.
      const st = this.strokeAttrs();
      const colour = this.fillOf();
      if (st && colour !== 'none') this.element(`<path d="${p.d}" fill="none"${st.replace(/ stroke="[^"]*"/, () => ` stroke="${colour}"`)}/>`);
      return;
    }
    // Filling closes every figure, as Windows does before it fills.
    if (fill) p.closeAll();
    this.finishShape(p, { fill, stroke });
  }

  /* ── clipping ── */

  /** A clip in device units, combined with the current one by a region mode: AND 1, OR 2, XOR 3, DIFF 4, COPY 5. */
  clipWith(d, rule, mode, rect = null) {
    this.flushLine();
    const cur = this.dc.clip;
    const depth = (n) => { let k = 0; for (let x = n; x; x = x.parent) k++; return k; };
    if (mode === 5 || (!cur && mode !== 4)) {
      // Everything joined with a region is still everything.
      this.dc.clip = mode === 2 && !cur ? null : { d, rule, parent: null, rect };
      return;
    }
    if (mode === 1) {
      if (cur?.rect && rect && !cur.parent) {
        const r = [Math.max(cur.rect[0], rect[0]), Math.max(cur.rect[1], rect[1]), Math.min(cur.rect[2], rect[2]), Math.min(cur.rect[3], rect[3])];
        if (r[2] < r[0]) r[2] = r[0];
        if (r[3] < r[1]) r[3] = r[1];
        this.dc.clip = { d: rectD(r), rule: 'nonzero', parent: null, rect: r };
        return;
      }
      if (depth(cur) >= LIMITS.clipDepth) return;
      this.dc.clip = { d, rule, parent: cur, rect: null };
    } else if (mode === 4) {
      if (depth(cur) >= LIMITS.clipDepth) return;
      const big = 'M-1e9 -1e9H1e9V1e9H-1e9Z';
      this.dc.clip = { d: big + d, rule: 'evenodd', parent: cur, rect: null };
    } else if (mode === 2) {
      // A union is exact for a clip of one shape; past that, the new region is added to the last.
      this.dc.clip = { d: cur.d + d, rule: 'nonzero', parent: cur.parent, rect: null };
    } else this.dc.clip = { d, rule, parent: null, rect };
  }

  /** IntersectClipRect / ExcludeClipRect, in logical units. */
  clipRect(l, t, r, b, exclude = false) {
    const m = this.m();
    const axis = Math.abs(m[1]) < 1e-12 && Math.abs(m[2]) < 1e-12;
    const pts = [this.pt(l, t), this.pt(r, t), this.pt(r, b), this.pt(l, b)];
    const rect = axis ? [Math.min(pts[0][0], pts[2][0]), Math.min(pts[0][1], pts[2][1]), Math.max(pts[0][0], pts[2][0]), Math.max(pts[0][1], pts[2][1])] : null;
    const d = rect ? rectD(rect) : polyD(pts);
    this.clipWith(d, 'nonzero', exclude ? 4 : 1, rect);
  }

  /** Rectangles in device units (a region), as a clip. */
  clipRegion(rects, mode) {
    if (!rects.length) { if (mode === 5) { this.flushLine(); this.dc.clip = null; } return; }
    const rect = rects.length === 1 ? rects[0] : null;
    this.clipWith(rects.map(rectD).join(''), 'nonzero', mode, rect);
  }

  clipPath(mode) {
    const p = this.path;
    this.path = null;
    this.inPath = false;
    if (!p || p.empty) return;
    p.close();
    this.clipWith(p.d, this.dc.polyFill === 1 ? 'evenodd' : 'nonzero', mode);
  }

  /** A region (logical rectangles) filled with a brush, or framed with one. */
  fillRegion(rects, b = this.dc.brush) {
    if (!rects.length) return;
    const fill = this.fillOf(b);
    if (fill === 'none') return;
    this.element(`<path d="${rects.map((r) => polyD([this.pt(r[0], r[1]), this.pt(r[2], r[1]), this.pt(r[2], r[3]), this.pt(r[0], r[3])])).join('')}" fill="${fill}"/>`);
  }

  frameRegion(rects, b, w) {
    if (!rects.length) return;
    const fill = this.fillOf(b);
    if (fill === 'none') return;
    const width = Math.max(1, Math.abs(w) * this.scale());
    this.element(`<path d="${rects.map((r) => polyD([this.pt(r[0], r[1]), this.pt(r[2], r[1]), this.pt(r[2], r[3]), this.pt(r[0], r[3])])).join('')}" fill="none" stroke="${fill}" stroke-width="${num(width)}"/>`);
  }

  /** GradientFill: rectangles shaded across or down, triangles in their colours' average. */
  gradient(vertices, shapes, mode) {
    for (const s of shapes) {
      if (mode === 2) {
        const v = s.map((i) => vertices[i]).filter(Boolean);
        if (v.length !== 3) continue;
        const avg = [0, 1, 2].map((k) => Math.round(v.reduce((a, x) => a + x.rgb[k], 0) / 3));
        this.element(`<path d="${polyD(v.map((x) => this.pt(x.x, x.y)))}" fill="${hex(avg[0] | avg[1] << 8 | avg[2] << 16)}"/>`);
        continue;
      }
      const a = vertices[s[0]];
      const b = vertices[s[1]];
      if (!a || !b) continue;
      const id = this.id('g');
      const from = this.pt(a.x, a.y);
      const to = mode === 0 ? this.pt(b.x, a.y) : this.pt(a.x, b.y);
      const c = (x) => hex(x.rgb[0] | x.rgb[1] << 8 | x.rgb[2] << 16);
      this.def(`<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${num(from[0])}" y1="${num(from[1])}" x2="${num(to[0])}" y2="${num(to[1])}"><stop offset="0" stop-color="${c(a)}"/><stop offset="1" stop-color="${c(b)}"/></linearGradient>`);
      this.element(`<path d="${polyD([this.pt(a.x, a.y), this.pt(b.x, a.y), this.pt(b.x, b.y), this.pt(a.x, b.y)])}" fill="url(#${id})"/>`);
    }
  }

  /* ── words ── */

  /**
   * Words at a point, as ExtTextOut writes them: `dx` the advance after each
   * character in logical units (or null, and the font's own advances are
   * used), `options` the ETO_ flags, `rect` the opaque or clip rectangle.
   */
  text({ x, y, str, dx = null, options = 0, rect = null }) {
    if (this.inPath || !str) return;
    const dc = this.dc;
    if (options & 0x10) return; // glyph indices: letters only the font file can name
    if (dc.textAlign & 1) [x, y] = dc.cur;
    const m = this.m();
    const f = dc.font || STOCK[13];
    const [asc, desc] = metricsOf(f.face);
    const along = Math.hypot(m[0], m[1]) || 1;
    const det = Math.abs(m[0] * m[3] - m[1] * m[2]);
    const across = det / along || along;
    let em = f.height < 0 ? -f.height * across : f.height > 0 ? (f.height * across) / (asc + desc) : 12;
    if (!(em > 0.05 && em < 1e5)) return;
    // The baseline's direction on the page: the transform's, turned by the font's escapement.
    let angle = Math.atan2(m[1], m[0]) * 180 / Math.PI;
    if (m[0] < 0 && Math.abs(m[1]) < 1e-12) angle = 0;
    angle -= (f.escapement || 0) / 10;
    const chars = Array.from(str);
    let xs = null;
    let total = 0;
    if (dx) {
      xs = [];
      let unit = 0;
      for (const ch of chars) {
        xs.push(total);
        const n = ch.length;
        let adv = 0;
        for (let k = 0; k < n; k++) adv += dx[unit + k] ?? 0;
        unit += n;
        total += adv * along + (dc.charExtra || 0) * along;
      }
    } else total = chars.length * em * 0.5;
    const h = dc.textAlign & 6;
    const shift = h === 6 ? -total / 2 : h === 2 ? -total : 0;
    const v = dc.textAlign & 24;
    const base = v === 24 ? 0 : v === 8 ? -desc * em : asc * em;
    const [ox, oy] = this.pt(x, y);
    const place = `translate(${num(ox)} ${num(oy)})${Math.abs(angle) > 0.01 ? ` rotate(${num(angle)})` : ''}`;
    let clip = dc.clip;
    if (rect && (options & 4)) {
      const pts = [this.pt(rect[0], rect[1]), this.pt(rect[2], rect[1]), this.pt(rect[2], rect[3]), this.pt(rect[0], rect[3])];
      clip = { d: polyD(pts), rule: 'nonzero', parent: dc.clip, rect: null };
    }
    if (rect && (options & 2)) {
      const pts = [this.pt(rect[0], rect[1]), this.pt(rect[2], rect[1]), this.pt(rect[2], rect[3]), this.pt(rect[0], rect[3])];
      if (rect[2] !== rect[0] && rect[3] !== rect[1]) this.element(`<path d="${polyD(pts)}" fill="${hex(dc.bkColor)}"/>`, clip);
    } else if (dc.bkMode === 2 && dx) {
      this.element(`<rect transform="${place}" x="${num(shift)}" y="${num(base - asc * em)}" width="${num(total)}" height="${num((asc + desc) * em)}" fill="${hex(dc.bkColor)}"/>`, clip);
    }
    const face = faceName(f.face);
    const family = `${face ? `'${xmlText(face)}', ` : ''}${fallbackFamily(f)}`;
    const weight = f.weight > 0 ? Math.max(100, Math.min(900, Math.round(f.weight / 100) * 100)) : 400;
    const deco = [f.underline && 'underline', f.strike && 'line-through'].filter(Boolean).join(' ');
    const anchor = !xs && h ? ` text-anchor="${h === 6 ? 'middle' : 'end'}"` : '';
    const xAttr = xs ? xs.map((v2) => num(v2 + shift)).join(' ') : '0';
    this.element(`<text transform="${place}" x="${xAttr}" y="${num(base)}" font-family="${family}" font-size="${num(em)}"${weight !== 400 ? ` font-weight="${weight}"` : ''}${f.italic ? ' font-style="italic"' : ''}${deco ? ` text-decoration="${deco}"` : ''} fill="${hex(dc.textColor)}"${anchor} xml:space="preserve" style="white-space:pre">${xmlText(chars.join(''))}</text>`, clip);
    if (dc.textAlign & 1) {
      let adv = 0;
      if (dx) for (let i = 0; i < str.length; i++) adv += dx[i] ?? 0;
      dc.cur = [x + adv, y];
    }
  }

  /* ── bitmaps ── */

  /**
   * A DIB's rectangle (rows from the top) as a PNG data URI; `alpha` 'premultiplied'
   * reads a 32-bit bitmap's own alpha, `key` makes one colour clear.
   */
  dibPng(dib, sx, sy, scx, scy, { alpha = null, key = null, invert = false } = {}) {
    if (dib.raw) {
      this.imageBytes += dib.raw.length;
      if (this.imageBytes > LIMITS.imageBytes) throw new Enough('images');
      return `data:${dib.type};base64,${Buffer.from(dib.raw).toString('base64')}`;
    }
    let x0 = Math.max(0, Math.min(sx, sx + scx));
    let y0 = Math.max(0, Math.min(sy, sy + scy));
    const x1 = Math.min(dib.w, Math.max(sx, sx + scx));
    const y1 = Math.min(dib.h, Math.max(sy, sy + scy));
    x0 = Math.floor(x0); y0 = Math.floor(y0);
    const W0 = Math.ceil(x1) - x0;
    const H0 = Math.ceil(y1) - y0;
    if (!(W0 > 0 && H0 > 0)) return null;
    const step = Math.max(1, Math.ceil(Math.sqrt((W0 * H0) / LIMITS.imageOut)));
    const W = Math.ceil(W0 / step);
    const H = Math.ceil(H0 / step);
    const rgba = Buffer.alloc(W * H * 4);
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        const px = x0 + x * step;
        const py = y0 + y * step;
        let c = dib.pixel(px, py);
        if (invert) c = ~c & 0xffffff;
        const at = (y * W + x) * 4;
        let a = 255;
        let r = (c >> 16) & 255; let g = (c >> 8) & 255; let b = c & 255;
        if (alpha && dib.alpha) {
          a = dib.alpha(px, py);
          if (a > 0 && a < 255) { r = Math.min(255, Math.round(r * 255 / a)); g = Math.min(255, Math.round(g * 255 / a)); b = Math.min(255, Math.round(b * 255 / a)); }
        }
        if (key != null && c === key) a = 0;
        rgba[at] = r; rgba[at + 1] = g; rgba[at + 2] = b; rgba[at + 3] = a;
      }
    }
    const png = encodePng(W, H, rgba);
    this.imageBytes += png.length;
    if (this.imageBytes > LIMITS.imageBytes) throw new Enough('images');
    return `data:image/png;base64,${png.toString('base64')}`;
  }

  /** An image filling the logical rectangle (x, y, cx, cy), however the transform turns it. */
  imageAt(href, x, y, cx, cy, opacity = 1, clip = this.dc.clip, fromRaster = false) {
    const p0 = this.pt(x, y);
    const p1 = this.pt(x + cx, y);
    const p2 = this.pt(x, y + cy);
    const mtx = [p1[0] - p0[0], p1[1] - p0[1], p2[0] - p0[0], p2[1] - p0[1], p0[0], p0[1]];
    if (!(Math.abs(mtx[0] * mtx[3] - mtx[1] * mtx[2]) > 1e-9)) return;
    const axis = Math.abs(mtx[1]) < 1e-9 && Math.abs(mtx[2]) < 1e-9 && mtx[0] > 0 && mtx[3] > 0;
    const place = axis
      ? `x="${num(p0[0])}" y="${num(p0[1])}" width="${num(mtx[0])}" height="${num(mtx[3])}"`
      : `width="1" height="1" transform="matrix(${mtx.map((v) => Number.isFinite(v) ? String(Math.round(v * 10000) / 10000) : '0').join(' ')})"`;
    this.element(`<image ${place} preserveAspectRatio="none"${opacity < 1 ? ` opacity="${num(opacity)}"` : ''} href="${href}"/>`, clip, fromRaster);
  }

  /** A rectangle filled by a raster operation that reads no bitmap: the brush, black or white. */
  patBlt(x, y, cx, cy, rop) {
    let fill = null;
    if (rop === ROP.PATCOPY) fill = this.fillOf();
    else if (rop === ROP.BLACKNESS) fill = '#000000';
    else if (rop === ROP.WHITENESS) fill = '#ffffff';
    if (!fill || fill === 'none') return;
    this.element(`<path d="${polyD([this.pt(x, y), this.pt(x + cx, y), this.pt(x + cx, y + cy), this.pt(x, y + cy)])}" fill="${fill}"/>`);
  }

  /**
   * A bitmap copied onto the page. A plain copy is an image; one combined
   * with what is under it (a mask ANDed, a picture ORed: the old way to make
   * a background see-through) is played with its neighbours onto a white
   * page and a black one, and what differs between them is clear.
   */
  blit({ x, y, cx, cy, rop, source = null, alpha = null, key = null }) {
    if (!source) { this.patBlt(x, y, cx, cy, rop); return; }
    const { dib } = source;
    this.flushLine();
    const plain = rop === ROP.SRCCOPY || rop === ROP.NOTSRCCOPY;
    // A copy over nothing combined, a blend, a keyed colour or a JPEG is an image of its own.
    if (alpha || key != null || dib.raw || (plain && !(this.raster && this.overlapsRaster(x, y, cx, cy)))) {
      const href = this.dibPng(dib, source.x, source.y, source.cx, source.cy, { alpha: alpha && (alpha.format & 1) ? 'premultiplied' : null, key, invert: rop === ROP.NOTSRCCOPY });
      if (href) this.imageAt(href, x, y, cx, cy, alpha ? alpha.constant / 255 : 1);
      return;
    }
    const op = ropOf(rop);
    if (!op || !op.usesSource) { if (rop === ROP.DSTCOPY) return; this.patBlt(x, y, cx, cy, rop); return; }
    const m = this.m();
    if (Math.abs(m[1]) > 1e-12 || Math.abs(m[2]) > 1e-12) {
      const href = this.dibPng(dib, source.x, source.y, source.cx, source.cy);
      if (href) this.imageAt(href, x, y, cx, cy);
      return;
    }
    const p0 = this.pt(x, y);
    const p1 = this.pt(x + cx, y + cy);
    if (this.raster && this.raster.clip !== this.dc.clip) this.flushRaster();
    if (!this.raster) this.raster = { items: [], clip: this.dc.clip, box: [Infinity, Infinity, -Infinity, -Infinity] };
    const g = this.raster;
    g.items.push({ p0, p1, op, rop, source });
    g.box = [Math.min(g.box[0], p0[0], p1[0]), Math.min(g.box[1], p0[1], p1[1]), Math.max(g.box[2], p0[0], p1[0]), Math.max(g.box[3], p0[1], p1[1])];
    if (g.items.length >= 16) this.flushRaster();
  }

  overlapsRaster(x, y, cx, cy) {
    const a = this.pt(x, y);
    const b = this.pt(x + cx, y + cy);
    const box = this.raster.box;
    return Math.max(a[0], b[0]) > box[0] && Math.min(a[0], b[0]) < box[2] && Math.max(a[1], b[1]) > box[1] && Math.min(a[1], b[1]) < box[3];
  }

  /** The combined bitmaps played onto a white page and a black one; where they differ is see-through. */
  flushRaster() {
    const g = this.raster;
    this.raster = null;
    if (!g || !g.items.length) return;
    const [bx0, by0, bx1, by1] = g.box;
    if (!(bx1 > bx0 && by1 > by0)) return;
    let k = 1;
    for (const it of g.items) k = Math.max(k, Math.abs(it.source.cx) / Math.max(1e-6, Math.abs(it.p1[0] - it.p0[0])));
    k = Math.min(k, 16);
    let W = Math.ceil((bx1 - bx0) * k);
    let H = Math.ceil((by1 - by0) * k);
    if (W * H > LIMITS.imageOut) { const s = Math.sqrt(LIMITS.imageOut / (W * H)); k *= s; W = Math.ceil((bx1 - bx0) * k); H = Math.ceil((by1 - by0) * k); }
    W = Math.max(1, W); H = Math.max(1, H);
    const onWhite = new Uint32Array(W * H).fill(0xffffff);
    const onBlack = new Uint32Array(W * H);
    for (const it of g.items) {
      const x0 = (it.p0[0] - bx0) * k; const y0 = (it.p0[1] - by0) * k;
      const x1 = (it.p1[0] - bx0) * k; const y1 = (it.p1[1] - by0) * k;
      const fromX = Math.max(0, Math.floor(Math.min(x0, x1)));
      const toX = Math.min(W, Math.ceil(Math.max(x0, x1)));
      const fromY = Math.max(0, Math.floor(Math.min(y0, y1)));
      const toY = Math.min(H, Math.ceil(Math.max(y0, y1)));
      const s = it.source;
      for (let py = fromY; py < toY; py++) {
        const v = (py + 0.5 - y0) / (y1 - y0);
        const sy = s.y + Math.min(Math.abs(s.cy) - 1, Math.max(0, Math.floor(v * Math.abs(s.cy))));
        for (let px = fromX; px < toX; px++) {
          const u = (px + 0.5 - x0) / (x1 - x0);
          const sx = s.x + Math.min(Math.abs(s.cx) - 1, Math.max(0, Math.floor(u * Math.abs(s.cx))));
          const c = sx >= 0 && sx < s.dib.w && sy >= 0 && sy < s.dib.h ? s.dib.pixel(sx, sy) : 0;
          const i = py * W + px;
          onWhite[i] = it.op.apply(c, onWhite[i]);
          onBlack[i] = it.op.apply(c, onBlack[i]);
        }
      }
    }
    const rgba = Buffer.alloc(W * H * 4);
    for (let i = 0; i < W * H; i++) {
      const w = onWhite[i];
      const d = onBlack[i];
      let a = 0;
      for (const shift of [16, 8, 0]) a = Math.max(a, 255 - (((w >> shift) & 255) - ((d >> shift) & 255)));
      a = Math.max(0, Math.min(255, a));
      const at = i * 4;
      if (a) {
        rgba[at] = Math.min(255, Math.round((((d >> 16) & 255) * 255) / a));
        rgba[at + 1] = Math.min(255, Math.round((((d >> 8) & 255) * 255) / a));
        rgba[at + 2] = Math.min(255, Math.round(((d & 255) * 255) / a));
      }
      rgba[at + 3] = a;
    }
    const png = encodePng(W, H, rgba);
    this.imageBytes += png.length;
    if (this.imageBytes > LIMITS.imageBytes) throw new Enough('images');
    this.element(`<image x="${num(bx0)}" y="${num(by0)}" width="${num(bx1 - bx0)}" height="${num(by1 - by0)}" preserveAspectRatio="none" href="data:image/png;base64,${png.toString('base64')}"/>`, g.clip, true);
  }

  /* ── the end ── */

  /**
   * The SVG: `view` the device rectangle [x, y, width, height] the picture
   * shows, `width` and `height` its size in CSS pixels. Null when nothing
   * was drawn.
   */
  finish({ view, width, height }) {
    try { this.flushLine(); this.flushRaster(); } catch (err) { if (!(err instanceof Enough)) throw err; }
    if (!this.drew) return null;
    if (this.openClip) this.body.push('</g>');
    this.openClip = null;
    const head = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${view.map(num).join(' ')}" width="${num(width)}" height="${num(height)}" preserveAspectRatio="none">`;
    return head + (this.defs.length ? `<defs>${this.defs.join('')}</defs>` : '') + this.body.join('') + '</svg>';
  }
}

export const rectD = (r) => `M${num(r[0])} ${num(r[1])}H${num(r[2])}V${num(r[3])}H${num(r[0])}Z`;
export const polyD = (pts) => 'M' + pts.map((p) => num(p[0]) + ' ' + num(p[1])).join('L') + 'Z';
