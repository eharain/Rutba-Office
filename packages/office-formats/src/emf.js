// SPDX-License-Identifier: AGPL-3.0-or-later
//
// An Enhanced Metafile, as an SVG a window can draw, or its bitmaps as a PNG.
//
// A picture pasted into Word, Excel or PowerPoint from another program is
// often kept as an EMF: a list of drawing calls rather than pixels. A
// chart, a diagram, a document's icon with its name under it, a signature
// scanned into a bitmap. A browser draws no EMF, so each showed as a broken
// picture or an empty frame where Office showed the picture.
//
// emfToSvg plays every call onto the SVG device in gdi-svg.js: the lines,
// shapes, words and bitmaps come out as SVG in the metafile's own units, so
// the drawing stays sharp at any size and its words stay words.
//
// emfToPng is the older, narrower way, kept for what the SVG cannot be made
// of: a metafile that is only bitmaps, played into one PNG. The calls
// combine pixels with raster operations (OR the mask, AND the picture),
// which say nothing about transparency, so they are played twice, onto a
// white page and onto a black one: a pixel the same on both is solid, one
// white on the first and black on the second is not there, and one in
// between is that far see-through. A metafile that draws lines, shapes or
// words gives no PNG.

import { readDib, ropOf, encodePng, GdiSvg, Enough, LIMITS, ROP, STOCK, makePen, decodeAnsi } from './gdi-svg.js';

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
    if (!dib || dib.raw) return null;
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

/* ── the metafile as SVG ──────────────────────────────────────────────── */

/** A region's rectangles (an RGNDATA at `at` in the record), as [l, t, r, b]. */
function regionRects(rd, at) {
  const count = rd.u32(at + 8);
  const most = Math.max(0, Math.floor((rd.n - at - 32) / 16));
  const rects = [];
  for (let i = 0; i < Math.min(count, most); i++) {
    const q = at + 32 + i * 16;
    rects.push([rd.i32(q), rd.i32(q + 4), rd.i32(q + 8), rd.i32(q + 12)]);
  }
  return rects;
}

/** A record's fields, read from its start; reading past its end is a damaged record. */
export class RecordReader {
  constructor(b) { this.b = b; this.o = 0; this.n = 0; }
  at(o, n) { this.o = o; this.n = n; return this; }
  check(k, size) { if (k < 0 || k + size > this.n) throw new RangeError('a record shorter than its fields'); return this.o + k; }
  i32(k) { return this.b.readInt32LE(this.check(k, 4)); }
  u32(k) { return this.b.readUInt32LE(this.check(k, 4)); }
  i16(k) { return this.b.readInt16LE(this.check(k, 2)); }
  u16(k) { return this.b.readUInt16LE(this.check(k, 2)); }
  u8(k) { return this.b[this.check(k, 1)]; }
  f32(k) { return this.b.readFloatLE(this.check(k, 4)); }
  /** `count` points from `k`, 32-bit or 16-bit, as many as the record holds. */
  points(k, count, small) {
    const size = small ? 4 : 8;
    const n = Math.min(count >>> 0, Math.max(0, Math.floor((this.n - k) / size)));
    const out = new Array(n);
    for (let i = 0; i < n; i++) {
      const q = k + i * size;
      out[i] = small ? [this.i16(q), this.i16(q + 2)] : [this.i32(q), this.i32(q + 4)];
    }
    return out;
  }
}

const BLIT = { BITBLT: 76, STRETCHBLT: 77, SETDIBITSTODEVICE: 80, STRETCHDIBITS: 81, ALPHABLEND: 114, TRANSPARENTBLT: 116 };

/**
 * The metafile drawn as an SVG: `{ svg, width, height }` (its size in CSS
 * pixels, from the frame its header gives), or null when it is not an EMF
 * or draws nothing this reads.
 *
 * Every drawing record of an ordinary EMF is played: pens, brushes and
 * fonts chosen and let go, lines, curves, polygons, rectangles, rounded
 * rectangles, ellipses, arcs, chords and pies, paths filled and outlined,
 * words where ExtTextOut put them with the advances it gave, the world
 * transform, saved and restored states, clips, the map modes with their
 * window and viewport, gradients, and bitmaps as embedded pictures. EMF+
 * records (GDI+ drawing kept in comments) are not read: a file that carries
 * both draws from its ordinary records, as every program that reads only
 * those does.
 */
export function emfToSvg(bytes) {
  if (!bytes) return null;
  const b = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer || bytes, bytes.byteOffset || 0, bytes.byteLength ?? bytes.length);
  if (!isEmf(b) || b.length > LIMITS.bytes) return null;
  const hsize = b.readUInt32LE(4);
  const bounds = [8, 12, 16, 20].map((k) => b.readInt32LE(k));
  const frame = [24, 28, 32, 36].map((k) => b.readInt32LE(k));
  const dev = [b.readInt32LE(72), b.readInt32LE(76)];
  const mm = [b.readInt32LE(80), b.readInt32LE(84)];
  const known = dev[0] > 0 && dev[1] > 0 && mm[0] > 0 && mm[1] > 0;
  let ppmm = known ? [dev[0] / mm[0], dev[1] / mm[1]] : [96 / 25.4, 96 / 25.4];
  if (hsize >= 108 && b.length >= 108) {
    const um = [b.readUInt32LE(100), b.readUInt32LE(104)];
    if (um[0] > 0 && um[1] > 0 && dev[0] > 0 && dev[1] > 0) ppmm = [dev[0] / (um[0] / 1000), dev[1] / (um[1] / 1000)];
  }
  // The frame (hundredths of a millimetre) is the picture's extent, as Office
  // places it; the bounds stand in for a frame that is empty, or one with no
  // reference device to measure it against.
  let view;
  let css;
  const fw = frame[2] - frame[0];
  const fh = frame[3] - frame[1];
  if (known && fw > 0 && fh > 0) {
    view = [frame[0] / 100 * ppmm[0], frame[1] / 100 * ppmm[1], fw / 100 * ppmm[0], fh / 100 * ppmm[1]];
    css = [fw / 100 / 25.4 * 96, fh / 100 / 25.4 * 96];
  } else {
    view = [bounds[0], bounds[1], bounds[2] - bounds[0] + 1, bounds[3] - bounds[1] + 1];
    css = [view[2] / ppmm[0] / 25.4 * 96, view[3] / ppmm[1] / 25.4 * 96];
  }
  if (!(view[2] > 0 && view[3] > 0 && view.every(Number.isFinite) && css.every((v) => v > 0 && Number.isFinite(v)))) return null;

  const g = new GdiSvg({ ppmm });
  const rd = new RecordReader(b);
  const sourceOf = (offBmi, offBits, usage, cbBits, xSrc, ySrc, cxSrc, cySrc) => {
    if (!(offBmi >= 8 && offBmi + 12 <= rd.n && offBits >= 8 && offBits <= rd.n)) return null;
    const dib = readDib(b, rd.o + offBmi, rd.o + offBits, usage, Math.min(cbBits, rd.n - offBits));
    if (!dib) return null;
    return { dib, x: xSrc, y: dib.raw || dib.topDown ? ySrc : dib.h - ySrc - cySrc, cx: cxSrc, cy: cySrc };
  };
  const textAt = (e, wide) => {
    const x = rd.i32(e);
    const y = rd.i32(e + 4);
    const count = rd.u32(e + 8);
    const offString = rd.u32(e + 12);
    const options = rd.u32(e + 16);
    const hasRect = !(options & 0x100);
    const rect = hasRect ? [rd.i32(e + 20), rd.i32(e + 24), rd.i32(e + 28), rd.i32(e + 32)] : null;
    const offDx = rd.u32(e + (hasRect ? 36 : 20));
    const unit = wide ? 2 : 1;
    if (offString < 8 || offString > rd.n) return;
    const chars = Math.min(count, Math.floor((rd.n - offString) / unit));
    if (!chars) return;
    const raw = b.subarray(rd.o + offString, rd.o + offString + chars * unit);
    const str = wide ? raw.toString('utf16le') : decodeAnsi(raw, g.dc.font?.charset ?? 0);
    let dx = null;
    const pdy = options & 0x2000 ? 2 : 1;
    if (offDx >= 8 && offDx + chars * 4 * pdy <= rd.n) {
      dx = new Array(chars);
      for (let i = 0; i < chars; i++) dx[i] = rd.i32(offDx + i * 4 * pdy);
    }
    g.text({ x, y, str, dx, options, rect });
  };
  const brushAt = (k) => {
    const ih = rd.u32(k);
    const o2 = ih >= 0x80000000 ? STOCK[ih & 0x7fffffff] : g.objects.get(ih);
    return o2?.kind === 'brush' ? o2 : g.dc.brush;
  };

  let o = 0;
  let records = 0;
  try {
    while (o + 8 <= b.length) {
      const type = b.readUInt32LE(o);
      const n = b.readUInt32LE(o + 4);
      if (n < 8 || o + n > b.length) break;
      if (++records > LIMITS.records) break;
      rd.at(o, n);
      switch (type) {
        case 2: case 85: g.polyBezier(rd.points(28, rd.u32(24), type === 85)); break;
        case 3: case 86: g.polygons([rd.points(28, rd.u32(24), type === 86)]); break;
        case 4: case 87: g.polyline(rd.points(28, rd.u32(24), type === 87)); break;
        case 5: case 88: g.polyBezier(rd.points(28, rd.u32(24), type === 88), { to: true }); break;
        case 6: case 89: g.polyline(rd.points(28, rd.u32(24), type === 89), { to: true }); break;
        case 7: case 8: case 90: case 91: {
          const small = type >= 90;
          const polys = Math.min(rd.u32(24), Math.max(0, Math.floor((n - 32) / 4)));
          const counts = [];
          for (let i = 0; i < polys; i++) counts.push(rd.u32(32 + i * 4));
          const all = rd.points(32 + polys * 4, rd.u32(28), small);
          const groups = [];
          let at = 0;
          for (const c of counts) {
            if (at >= all.length) break;
            groups.push(all.slice(at, at + c));
            at += c;
          }
          if (type === 8 || type === 91) g.polygons(groups); else g.polylines(groups);
          break;
        }
        case 56: case 92: {
          const small = type === 92;
          const pts = rd.points(28, rd.u32(24), small);
          const typesAt = 28 + pts.length * (small ? 4 : 8);
          const types = [];
          for (let i = 0; i < pts.length && typesAt + i < n; i++) types.push(rd.u8(typesAt + i));
          g.polyDraw(pts.slice(0, types.length), types);
          break;
        }
        case 9: g.setWindowExt(rd.i32(8), rd.i32(12)); break;
        case 10: g.setWindowOrg(rd.i32(8), rd.i32(12)); break;
        case 11: g.setViewportExt(rd.i32(8), rd.i32(12)); break;
        case 12: g.setViewportOrg(rd.i32(8), rd.i32(12)); break;
        case 15: g.setPixel(rd.i32(8), rd.i32(12), rd.u32(16)); break;
        case 17: g.setMapMode(rd.u32(8)); break;
        case 18: g.set('bkMode', rd.u32(8)); break;
        case 19: g.set('polyFill', rd.u32(8)); break;
        case 20: g.set('rop2', rd.u32(8)); break;
        case 22: g.set('textAlign', rd.u32(8)); break;
        case 24: g.set('textColor', rd.u32(8) & 0xffffff); break;
        case 25: g.set('bkColor', rd.u32(8) & 0xffffff); break;
        case 27: g.moveTo(rd.i32(8), rd.i32(12)); break;
        case 29: g.clipRect(rd.i32(8), rd.i32(12), rd.i32(16), rd.i32(20), true); break;
        case 30: g.clipRect(rd.i32(8), rd.i32(12), rd.i32(16), rd.i32(20)); break;
        case 31: g.scaleViewportExt(rd.i32(8), rd.i32(12), rd.i32(16), rd.i32(20)); break;
        case 32: g.scaleWindowExt(rd.i32(8), rd.i32(12), rd.i32(16), rd.i32(20)); break;
        case 33: g.save(); break;
        case 34: g.restore(rd.i32(8)); break;
        case 35: case 36: {
          const x = [8, 12, 16, 20, 24, 28].map((k) => rd.f32(k));
          if (type === 35) g.setWorld(x); else g.modifyWorld(x, rd.u32(32));
          break;
        }
        case 37: g.select(rd.u32(8)); break;
        case 38: g.setObject(rd.u32(8), makePen({ style: rd.u32(12), width: rd.i32(16), color: rd.u32(24) })); break;
        case 39: g.setObject(rd.u32(8), { kind: 'brush', style: rd.u32(12), color: rd.u32(16) & 0xffffff, hatch: rd.u32(20) }); break;
        case 40: g.remove(rd.u32(8)); break;
        case 41: g.angleArc(rd.i32(8), rd.i32(12), rd.u32(16), rd.f32(20), rd.f32(24)); break;
        case 42: g.ellipse(rd.i32(8), rd.i32(12), rd.i32(16), rd.i32(20)); break;
        case 43: g.rectangle(rd.i32(8), rd.i32(12), rd.i32(16), rd.i32(20)); break;
        case 44: g.roundRect(rd.i32(8), rd.i32(12), rd.i32(16), rd.i32(20), rd.i32(24), rd.i32(28)); break;
        case 45: case 46: case 47: case 55:
          g.arc({ 45: 'arc', 46: 'chord', 47: 'pie', 55: 'arcto' }[type], rd.i32(8), rd.i32(12), rd.i32(16), rd.i32(20), rd.i32(24), rd.i32(28), rd.i32(32), rd.i32(36));
          break;
        case 54: g.lineTo(rd.i32(8), rd.i32(12)); break;
        case 57: g.set('arcDir', rd.u32(8)); break;
        case 58: { const u = rd.u32(8); g.set('miter', u < 1000 ? u : rd.f32(8)); break; }
        case 59: g.beginPath(); break;
        case 60: g.endPath(); break;
        case 61: g.closeFigure(); break;
        case 62: g.drawPath(true, false); break;
        case 63: g.drawPath(true, true); break;
        case 64: g.drawPath(false, true); break;
        case 66: g.widenPath(); break;
        case 67: g.clipPath(rd.u32(8)); break;
        case 68: g.abortPath(); break;
        case 71: g.fillRegion(regionRects(rd, 32), brushAt(28)); break;
        case 72: g.frameRegion(regionRects(rd, 40), brushAt(28), rd.i32(32)); break;
        case 74: g.fillRegion(regionRects(rd, 28)); break;
        case 75: {
          // Region rectangles are in device units, which here are the picture's own.
          const cb = rd.u32(8);
          g.clipRegion(cb >= 32 ? regionRects(rd, 16) : [], rd.u32(12));
          break;
        }
        case 82: {
          let face = '';
          for (let k = 40; k + 2 <= Math.min(n, 104); k += 2) {
            const c = rd.u16(k);
            if (!c) break;
            face += String.fromCharCode(c);
          }
          g.setObject(rd.u32(8), {
            kind: 'font', height: rd.i32(12), width: rd.i32(16), escapement: rd.i32(20), orientation: rd.i32(24), weight: rd.i32(28),
            italic: Boolean(rd.u8(32)), underline: Boolean(rd.u8(33)), strike: Boolean(rd.u8(34)), charset: rd.u8(35), pitchFamily: rd.u8(39), face,
          });
          break;
        }
        case 83: case 84: textAt(36, type === 84); break;
        case 96: case 97: {
          const strings = Math.min(rd.u32(36), Math.max(0, Math.floor((n - 40) / 40)));
          for (let i = 0; i < strings; i++) textAt(40 + i * 40, type === 97);
          break;
        }
        case 108: {
          const options = rd.u32(20);
          const hasRect = !(options & 0x100);
          const at = hasRect ? 52 : 36;
          const small = Boolean(options & 0x200);
          const chars = Math.min(rd.u32(16), Math.max(0, Math.floor((n - at) / (small ? 1 : 2))));
          const raw = b.subarray(o + at, o + at + chars * (small ? 1 : 2));
          const str = small ? decodeAnsi(raw, g.dc.font?.charset ?? 0) : raw.toString('utf16le');
          g.text({ x: rd.i32(8), y: rd.i32(12), str, options, rect: hasRect ? [rd.i32(36), rd.i32(40), rd.i32(44), rd.i32(48)] : null });
          break;
        }
        case 93: case 94: {
          const offBmi = rd.u32(16);
          const offBits = rd.u32(24);
          const ok = offBmi >= 8 && offBmi + 12 <= n && offBits >= 8 && offBits <= n;
          const dib = ok ? readDib(b, o + offBmi, o + offBits, type === 93 ? 0 : rd.u32(12), Math.min(rd.u32(28), n - offBits)) : null;
          g.setObject(rd.u32(8), g.patternBrush(dib));
          break;
        }
        case 95: {
          const count = Math.min(rd.u32(48), Math.max(0, Math.floor((n - 52) / 4)), 64);
          const entries = [];
          for (let i = 0; i < count; i++) entries.push(rd.u32(52 + i * 4));
          g.setObject(rd.u32(8), makePen({ style: rd.u32(28), width: rd.u32(32), color: rd.u32(40), ext: true, entries, brushStyle: rd.u32(36) }));
          break;
        }
        case BLIT.BITBLT: case BLIT.STRETCHBLT: case BLIT.ALPHABLEND: case BLIT.TRANSPARENTBLT: {
          const stretch = type !== BLIT.BITBLT;
          const cbBmi = rd.u32(88);
          const cx = rd.i32(32);
          const cy = rd.i32(36);
          const source = cbBmi ? sourceOf(rd.u32(84), rd.u32(92), rd.u32(80), rd.u32(96), rd.i32(44), rd.i32(48), stretch ? rd.i32(100) : cx, stretch ? rd.i32(104) : cy) : null;
          if (cbBmi && !source) break;
          const word = rd.u32(40);
          const x = rd.i32(24);
          const y = rd.i32(28);
          if (type === BLIT.ALPHABLEND) {
            if (source) g.blit({ x, y, cx, cy, rop: ROP.SRCCOPY, source, alpha: { constant: (word >>> 16) & 255, format: (word >>> 24) & 255 } });
          } else if (type === BLIT.TRANSPARENTBLT) {
            // The clear colour is a COLORREF; the bitmap's pixels read as RGB.
            if (source) g.blit({ x, y, cx, cy, rop: ROP.SRCCOPY, source, key: ((word & 255) << 16) | (word & 0xff00) | ((word >>> 16) & 255) });
          } else g.blit({ x, y, cx, cy, rop: word, source });
          break;
        }
        case BLIT.STRETCHDIBITS: {
          const source = sourceOf(rd.u32(48), rd.u32(56), rd.u32(64), rd.u32(60), rd.i32(32), rd.i32(36), rd.i32(40), rd.i32(44));
          if (source) g.blit({ x: rd.i32(24), y: rd.i32(28), cx: rd.i32(72), cy: rd.i32(76), rop: rd.u32(68), source });
          break;
        }
        case BLIT.SETDIBITSTODEVICE: {
          const source = sourceOf(rd.u32(48), rd.u32(56), rd.u32(64), rd.u32(60), rd.i32(32), rd.i32(36), rd.i32(40), rd.i32(44));
          if (source) g.blit({ x: rd.i32(24), y: rd.i32(28), cx: rd.i32(40), cy: rd.i32(44), rop: ROP.SRCCOPY, source });
          break;
        }
        case 118: {
          const nv = Math.min(rd.u32(24), Math.max(0, Math.floor((n - 36) / 16)));
          const ns = rd.u32(28);
          const mode = rd.u32(32);
          const vertices = [];
          for (let i = 0; i < nv; i++) {
            const q = 36 + i * 16;
            vertices.push({ x: rd.i32(q), y: rd.i32(q + 4), rgb: [rd.u16(q + 8) >> 8, rd.u16(q + 10) >> 8, rd.u16(q + 12) >> 8] });
          }
          const per = mode === 2 ? 3 : 2;
          const from = 36 + nv * 16;
          const shapes = [];
          for (let i = 0; i < Math.min(ns, Math.max(0, Math.floor((n - from) / (per * 4)))); i++) {
            const s2 = [];
            for (let k = 0; k < per; k++) s2.push(rd.u32(from + (i * per + k) * 4));
            shapes.push(s2);
          }
          g.gradient(vertices, shapes, mode);
          break;
        }
        default: break;
      }
      o += n;
      if (type === 14) break;
    }
  } catch (err) {
    // A record shorter than its fields, or a bound reached: what was drawn before it stands.
    if (!(err instanceof Enough) && !(err instanceof RangeError)) throw err;
  }
  const svg = g.finish({ view, width: css[0], height: css[1] });
  return svg ? { svg, width: Math.round(css[0]), height: Math.round(css[1]) } : null;
}

export default { emfToPng, emfToSvg, isEmf };
