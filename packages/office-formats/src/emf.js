// SPDX-License-Identifier: AGPL-3.0-or-later
//
// An Enhanced Metafile's pictures, as a PNG a window can draw.
//
// A picture pasted into PowerPoint or Word from another program is often
// kept as an EMF: a list of drawing calls rather than pixels. A browser
// draws no EMF, so a deck that holds one showed a broken picture where
// PowerPoint showed the picture. Most such files are a bitmap wrapped in a
// metafile: the bitmap, and often a black-and-white mask laid down before it
// to make its background see-through. Those calls are played here.
//
// The calls combine pixels with raster operations (OR the mask, AND the
// picture), which say nothing about transparency. So they are played twice,
// onto a white page and onto a black one: a pixel the same on both is solid,
// one white on the first and black on the second is not there, and one in
// between is that far see-through. That is exactly what PowerPoint shows
// over a slide of any colour.
//
// A metafile that draws lines, shapes or words is not played: null, and the
// caller shows the picture's frame rather than half a drawing.

import zlib from 'node:zlib';

const EMR = {
  HEADER: 1, EOF: 14, SETMAPMODE: 17, SAVEDC: 33, RESTOREDC: 34,
  SETWINDOWEXTEX: 9, SETWINDOWORGEX: 10, SETVIEWPORTEXTEX: 11, SETVIEWPORTORGEX: 12,
  INTERSECTCLIPRECT: 30, SETWORLDTRANSFORM: 35, MODIFYWORLDTRANSFORM: 36, EXTSELECTCLIPRGN: 75,
  BITBLT: 76, STRETCHBLT: 77, STRETCHDIBITS: 81,
};

/** Records that draw something other than a bitmap: lines, shapes, paths, regions, words, blends. */
const DRAWS = new Set([
  2, 3, 4, 5, 6, 7, 8, 41, 42, 43, 44, 45, 46, 47, 54, 55, 56, 62, 63, 64, 71, 72, 73, 74,
  78, 79, 80, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 95, 96, 108, 114, 116, 118,
]);

/** The longest side the PNG is drawn at: a slide never shows more, and a 16 MB scan stays a quick open. */
const MOST = 2048;

export const isEmf = (bytes) => Boolean(bytes && bytes.length >= 88 && bytes.readUInt32LE(0) === EMR.HEADER && bytes.readUInt32LE(40) === 0x464d4520);

/**
 * The bitmap a metafile's source rectangle reads, as `pixel(x, yFromTop)` ->
 * 0xRRGGBB, or null for a form this does not read (JPEG or PNG inside a
 * DIB, run-length coding, palette indices into the device's palette).
 */
function readDib(b, offBmi, offBits, usage) {
  if (usage !== 0) return null;
  const size = b.readUInt32LE(offBmi);
  const w = b.readInt32LE(offBmi + 4);
  const hRaw = b.readInt32LE(offBmi + 8);
  const bpp = b.readUInt16LE(offBmi + 14);
  const compression = b.readUInt32LE(offBmi + 16);
  const used = b.readUInt32LE(offBmi + 32);
  const h = Math.abs(hRaw);
  const topDown = hRaw < 0;
  if (!(w > 0 && h > 0) || ![1, 4, 8, 16, 24, 32].includes(bpp)) return null;
  if (compression !== 0 && !(compression === 3 && (bpp === 16 || bpp === 32))) return null;
  const stride = (((w * bpp) + 31) >> 5) * 4;
  if (offBits + stride * h > b.length) return null;
  const palette = [];
  if (bpp <= 8) {
    const n = used || 1 << bpp;
    for (let i = 0; i < n; i++) {
      const at = offBmi + size + i * 4;
      palette.push((b[at + 2] << 16) | (b[at + 1] << 8) | b[at]);
    }
  }
  // 16 and 32 bits: 5-5-5 and 8-8-8 unless the header gives its own masks.
  let masks = bpp === 16 ? [0x7c00, 0x03e0, 0x001f] : [0xff0000, 0x00ff00, 0x0000ff];
  if (compression === 3) masks = [0, 4, 8].map((k) => b.readUInt32LE(offBmi + size + k));
  // One colour from its mask, stretched to eight bits.
  const channel = (v, mask) => {
    if (!mask) return 0;
    const low = 31 - Math.clz32(mask & -mask);
    const bits = 32 - Math.clz32(mask >>> low);
    const raw = (v & mask) >>> low;
    return bits >= 8 ? raw >>> (bits - 8) : Math.round((raw * 255) / ((1 << bits) - 1));
  };
  const pixel = (x, y) => {
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
  return { w, h, topDown, pixel };
}

/**
 * A raster operation on 24-bit colours, from its ternary code: each bit of
 * the code is the answer for one combination of pattern, source and
 * destination bits. The pattern (the brush) is not followed, so an
 * operation that reads it is not played.
 */
function ropOf(rop) {
  const code = (rop >>> 16) & 0xff;
  if (((code >> 4) & 0x0f) !== (code & 0x0f)) return null;
  // Bit (2·source + destination) of the low four is the answer.
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

/** A PNG of RGBA pixels. */
function encodePng(w, h, rgba) {
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

/**
 * The metafile's bitmaps, played onto its own bounds, as
 * `{ png, width, height }`, or null when it draws anything else or holds
 * nothing to draw.
 */
export function emfToPng(bytes) {
  const b = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer || bytes, bytes.byteOffset || 0, bytes.byteLength ?? bytes.length);
  if (!isEmf(b)) return null;
  const left = b.readInt32LE(8);
  const top = b.readInt32LE(12);
  const fullW = b.readInt32LE(16) - left + 1;
  const fullH = b.readInt32LE(20) - top + 1;
  if (!(fullW > 0 && fullH > 0)) return null;
  const k = Math.min(1, MOST / Math.max(fullW, fullH));
  const W = Math.max(1, Math.round(fullW * k));
  const H = Math.max(1, Math.round(fullH * k));
  const onWhite = new Uint32Array(W * H).fill(0xffffff);
  const onBlack = new Uint32Array(W * H);

  // Logical units to device units: the window and viewport, and the map mode that relates them.
  let dc = { mode: 1, winOrg: [0, 0], winExt: [1, 1], vpOrg: [0, 0], vpExt: [1, 1], clip: null };
  const saved = [];
  const toDevice = (x, y) => {
    if (dc.mode === 1) return [x - dc.winOrg[0] + dc.vpOrg[0], y - dc.winOrg[1] + dc.vpOrg[1]];
    const sx = dc.vpExt[0] / (dc.winExt[0] || 1);
    let sy = dc.vpExt[1] / (dc.winExt[1] || 1);
    // Isotropic: one unit is as long across as down.
    if (dc.mode === 7) sy = Math.sign(sy) * Math.abs(sx);
    return [(x - dc.winOrg[0]) * sx + dc.vpOrg[0], (y - dc.winOrg[1]) * sy + dc.vpOrg[1]];
  };
  // A device rectangle on the canvas: [x0, y0, x1, y1), drawn order kept for a mirrored one.
  const onCanvas = (x, y, cx, cy) => {
    const [ax, ay] = toDevice(x, y);
    const [bx, by] = toDevice(x + cx, y + cy);
    return [(ax - left) * k, (ay - top) * k, (bx - left) * k, (by - top) * k];
  };

  let drew = false;
  const blit = ({ x, y, cx, cy, rop, source }) => {
    const op = ropOf(rop);
    if (!op || (op.usesSource && !source)) return false;
    const [x0, y0, x1, y1] = onCanvas(x, y, cx, cy);
    let lo = [Math.min(x0, x1), Math.min(y0, y1)];
    let hi = [Math.max(x0, x1), Math.max(y0, y1)];
    if (dc.clip) {
      lo = [Math.max(lo[0], dc.clip[0]), Math.max(lo[1], dc.clip[1])];
      hi = [Math.min(hi[0], dc.clip[2]), Math.min(hi[1], dc.clip[3])];
    }
    const fromX = Math.max(0, Math.floor(lo[0]));
    const fromY = Math.max(0, Math.floor(lo[1]));
    const toX = Math.min(W, Math.ceil(hi[0]));
    const toY = Math.min(H, Math.ceil(hi[1]));
    for (let py = fromY; py < toY; py++) {
      const v = (py + 0.5 - y0) / (y1 - y0);
      const sy = source ? source.y + Math.min(source.cy - 1, Math.max(0, Math.floor(v * source.cy))) : 0;
      for (let px = fromX; px < toX; px++) {
        const i = py * W + px;
        let s = 0;
        if (source) {
          const u = (px + 0.5 - x0) / (x1 - x0);
          const sx = source.x + Math.min(source.cx - 1, Math.max(0, Math.floor(u * source.cx)));
          s = sx >= 0 && sx < source.dib.w && sy >= 0 && sy < source.dib.h ? source.dib.pixel(sx, sy) : 0;
        }
        onWhite[i] = op.apply(s, onWhite[i]);
        onBlack[i] = op.apply(s, onBlack[i]);
      }
    }
    drew = true;
    return true;
  };
  // The source rectangle, its rows counted from the top of the bitmap: a
  // bottom-up bitmap measures its source rectangle from the bottom row.
  const sourceOf = (o, offBmi, offBits, usage, xSrc, ySrc, cxSrc, cySrc) => {
    const dib = readDib(b, o + offBmi, o + offBits, usage);
    if (!dib) return null;
    return { dib, x: xSrc, y: dib.topDown ? ySrc : dib.h - ySrc - cySrc, cx: cxSrc, cy: cySrc };
  };

  let o = 0;
  while (o + 8 <= b.length) {
    const type = b.readUInt32LE(o);
    const n = b.readUInt32LE(o + 4);
    if (n < 8 || o + n > b.length) return null;
    const r = o + 8;
    const int = (at) => b.readInt32LE(r + at);
    const uint = (at) => b.readUInt32LE(r + at);
    if (DRAWS.has(type)) return null;
    switch (type) {
      case EMR.SETMAPMODE: dc.mode = uint(0); if (![1, 7, 8].includes(dc.mode)) return null; break;
      case EMR.SETWINDOWEXTEX: dc.winExt = [int(0), int(4)]; break;
      case EMR.SETWINDOWORGEX: dc.winOrg = [int(0), int(4)]; break;
      case EMR.SETVIEWPORTEXTEX: dc.vpExt = [int(0), int(4)]; break;
      case EMR.SETVIEWPORTORGEX: dc.vpOrg = [int(0), int(4)]; break;
      case EMR.SAVEDC: saved.push({ ...dc }); break;
      case EMR.RESTOREDC: {
        // A negative index counts back from the latest save.
        const back = int(0) < 0 ? -int(0) : saved.length - int(0) + 1;
        if (back > 0 && back <= saved.length) dc = saved.splice(saved.length - back)[0];
        break;
      }
      case EMR.INTERSECTCLIPRECT: {
        const [x0, y0, x1, y1] = onCanvas(int(0), int(4), int(8) - int(0), int(12) - int(4));
        const box = [Math.min(x0, x1), Math.min(y0, y1), Math.max(x0, x1), Math.max(y0, y1)];
        dc.clip = dc.clip ? [Math.max(dc.clip[0], box[0]), Math.max(dc.clip[1], box[1]), Math.min(dc.clip[2], box[2]), Math.min(dc.clip[3], box[3])] : box;
        break;
      }
      // Copying no region is how a metafile lets its clip go again.
      case EMR.EXTSELECTCLIPRGN: if (uint(0) === 0 && uint(4) === 5) dc.clip = null; break;
      case EMR.SETWORLDTRANSFORM:
      case EMR.MODIFYWORLDTRANSFORM: {
        const m = [0, 4, 8, 12, 16, 20].map((at) => b.readFloatLE(r + at));
        const identity = m[0] === 1 && m[1] === 0 && m[2] === 0 && m[3] === 1 && m[4] === 0 && m[5] === 0;
        // Setting the identity, or MWT_IDENTITY, is no transform; any other is one this does not follow.
        if (!(identity || (type === EMR.MODIFYWORLDTRANSFORM && uint(24) === 1))) return null;
        break;
      }
      case EMR.BITBLT:
      case EMR.STRETCHBLT: {
        const cbBmi = uint(80);
        const stretch = type === EMR.STRETCHBLT;
        const source = cbBmi ? sourceOf(o, uint(76), uint(84), uint(72), int(36), int(40), stretch ? int(92) : int(24), stretch ? int(96) : int(28)) : null;
        if (cbBmi && !source) return null;
        if (!blit({ x: int(16), y: int(20), cx: int(24), cy: int(28), rop: uint(32), source })) return null;
        break;
      }
      case EMR.STRETCHDIBITS: {
        const source = sourceOf(o, uint(40), uint(48), uint(56), int(24), int(28), int(32), int(36));
        if (!source) return null;
        if (!blit({ x: int(16), y: int(20), cx: int(64), cy: int(68), rop: uint(60), source })) return null;
        break;
      }
      default: break;
    }
    o += n;
    if (type === EMR.EOF) break;
  }
  if (!drew) return null;

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
  return { png: encodePng(W, H, rgba), width: W, height: H };
}

export default { emfToPng, isEmf };
