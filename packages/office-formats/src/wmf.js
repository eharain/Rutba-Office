// SPDX-License-Identifier: AGPL-3.0-or-later
//
// A Windows Metafile, the 16-bit sibling of the EMF, as an SVG a window can
// draw.
//
// Word 6, Word 97 and PowerPoint 97 kept pasted drawings, clip art and
// equations as WMFs, and a .doc or .ppt carries them still. A WMF is the
// same kind of list of drawing calls an EMF is, written smaller: 16-bit
// coordinates, fields in reverse order, objects numbered by the first free
// slot rather than by the file. Its records are read here and played on the
// same device emf.js uses (gdi-svg.js), so the two draw alike.
//
// The picture's size comes from the placeable header Word and PowerPoint
// write in front (its box, in units an inch); without one, from the window
// the records set. The viewport is the picture's box, as every program that
// places a WMF makes it: the records' own viewport calls are not followed.

import { GdiSvg, Enough, LIMITS, ROP, makePen, readDib, decodeAnsi } from './gdi-svg.js';

const PLACEABLE = 0x9ac6cdd7;

/** Whether these bytes are a WMF: the placeable key, or a memory or disk metafile's header. */
export function isWmf(bytes) {
  if (!bytes || bytes.length < 18) return false;
  const b = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer || bytes, bytes.byteOffset || 0, bytes.byteLength ?? bytes.length);
  if (b.readUInt32LE(0) === PLACEABLE) return b.length >= 40;
  const type = b.readUInt16LE(0);
  return (type === 1 || type === 2) && b.readUInt16LE(2) === 9 && [0x100, 0x300].includes(b.readUInt16LE(4));
}

/** A record's parameters, word by word from its start; reading past its end is a damaged record. */
class Params {
  constructor(b) { this.b = b; this.at = 0; this.words = 0; }
  set(at, words) { this.at = at; this.words = words; return this; }
  check(i, n = 1) { if (i < 0 || i + n > this.words) throw new RangeError('a record shorter than its fields'); return this.at + i * 2; }
  i16(i) { return this.b.readInt16LE(this.check(i)); }
  u16(i) { return this.b.readUInt16LE(this.check(i)); }
  u32(i) { return this.b.readUInt32LE(this.check(i, 2)); }
  byte(i) { return this.b[this.check(i >> 1) + (i & 1)]; }
  bytes(i, n) { const from = this.check(i, Math.ceil(n / 2)); return this.b.subarray(from, from + n); }
}

/** A packed DIB (BITMAPINFO, then its bits) at `at`, read up to `end`. */
function packedDib(b, at, end, usage) {
  if (at + 12 > end) return null;
  const size = b.readUInt32LE(at);
  if (!(size === 12 || (size >= 40 && at + 40 <= end))) return null;
  const bpp = size === 12 ? b.readUInt16LE(at + 10) : b.readUInt16LE(at + 14);
  const compression = size === 12 ? 0 : b.readUInt32LE(at + 16);
  const used = size === 12 ? 0 : b.readUInt32LE(at + 32);
  const entries = bpp <= 8 ? Math.min(used || 1 << bpp, 256) : 0;
  const entry = size === 12 ? 3 : usage === 1 ? 2 : 4;
  const masks = compression === 3 && size === 40 ? 12 : 0;
  const offBits = at + size + masks + entries * entry;
  if (offBits > end) return null;
  return readDib(b, at, offBits, usage, end - offBits);
}

/**
 * The metafile drawn as an SVG: `{ svg, width, height }` (its size in CSS
 * pixels), or null when it is not a WMF, has no size, or draws nothing.
 */
export function wmfToSvg(bytes) {
  if (!isWmf(bytes)) return null;
  const b = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer || bytes, bytes.byteOffset || 0, bytes.byteLength ?? bytes.length);
  if (b.length > LIMITS.bytes) return null;
  const placeable = b.readUInt32LE(0) === PLACEABLE;
  const start = placeable ? 22 : 0;
  if (start + 18 > b.length) return null;
  const headerWords = b.readUInt16LE(start + 2);
  const first = start + Math.max(9, headerWords) * 2;
  const slots = [];
  const most = Math.min(Math.max(1, b.readUInt16LE(start + 10)), LIMITS.objects);

  // The picture's box: the placeable header's, or the first window the records set.
  let box = null;
  let inch = 1440;
  if (placeable) {
    const [l, t, r, bt] = [6, 8, 10, 12].map((k) => b.readInt16LE(k));
    inch = b.readUInt16LE(14) || 1440;
    if (r !== l && bt !== t) box = { org: [Math.min(l, r), Math.min(t, bt)], ext: [Math.abs(r - l), Math.abs(bt - t)] };
  }
  let org = null;
  let ext = null;
  for (let p = first, n = 0; p + 6 <= b.length && n < 4096; n++) {
    const size = b.readUInt32LE(p) * 2;
    const fn = b.readUInt16LE(p + 4);
    if (size < 6 || p + size > b.length || fn === 0) break;
    if (fn === 0x020b && size >= 10 && !org) org = [b.readInt16LE(p + 8), b.readInt16LE(p + 6)];
    if (fn === 0x020c && size >= 10 && !ext) ext = [b.readInt16LE(p + 8), b.readInt16LE(p + 6)];
    p += size;
  }
  if (!box) {
    if (!ext || !ext[0] || !ext[1]) return null;
    box = { org: org || [0, 0], ext: [Math.abs(ext[0]), Math.abs(ext[1])] };
  }
  const view = [0, 0, box.ext[0], box.ext[1]];
  const css = [box.ext[0] / inch * 96, box.ext[1] / inch * 96];
  if (!(css[0] > 0 && css[1] > 0)) return null;

  // The device is the picture's box; the window starts as the box and the records move it.
  const g = new GdiSvg({ ppmm: [inch / 25.4, inch / 25.4] });
  g.dc.mapMode = 8;
  g.dc.winOrg = [...box.org];
  g.dc.winExt = [...box.ext];
  g.dc.vpOrg = [0, 0];
  g.dc.vpExt = [...box.ext];
  g.changed();

  const create = (o) => {
    let i = slots.indexOf(null);
    if (i < 0) {
      if (slots.length >= most && slots.length >= LIMITS.objects) return;
      i = slots.length;
      slots.push(null);
    }
    slots[i] = o;
    g.setObject(i, o);
  };
  const objectAt = (i) => (i >= 0 && i < slots.length ? slots[i] : null);
  const points = (q, from, count) => {
    const n = Math.min(Math.max(0, count), Math.floor((q.words - from) / 2));
    const out = new Array(n);
    for (let i = 0; i < n; i++) out[i] = [q.i16(from + i * 2), q.i16(from + i * 2 + 1)];
    return out;
  };
  const region = (o) => (o?.kind === 'region' ? o.rects : []);
  const sourceDib = (q, from, usage) => packedDib(b, q.at + from * 2, q.at + q.words * 2, usage);
  const blit = (dib, xs, ys, sw, sh, x, y, w, h, rop) => {
    if (!dib) { g.blit({ x, y, cx: w, cy: h, rop }); return; }
    const source = { dib, x: xs, y: dib.raw || dib.topDown ? ys : dib.h - ys - sh, cx: sw, cy: sh };
    g.blit({ x, y, cx: w, cy: h, rop, source });
  };

  const q = new Params(b);
  let p = first;
  let records = 0;
  try {
    while (p + 6 <= b.length) {
      const size = b.readUInt32LE(p) * 2;
      const fn = b.readUInt16LE(p + 4);
      if (fn === 0) break;
      if (size < 6 || p + size > b.length) break;
      if (++records > LIMITS.records) break;
      q.set(p + 6, (size - 6) / 2);
      switch (fn) {
        case 0x0201: g.set('bkColor', q.u32(0) & 0xffffff); break;
        case 0x0102: g.set('bkMode', q.u16(0)); break;
        case 0x0103: if (q.u16(0) === 7) g.setMapMode(7); break;
        case 0x0104: g.set('rop2', q.u16(0)); break;
        case 0x0106: g.set('polyFill', q.u16(0)); break;
        case 0x0108: g.set('charExtra', q.i16(0)); break;
        case 0x0209: g.set('textColor', q.u32(0) & 0xffffff); break;
        case 0x020b: g.setWindowOrg(q.i16(1), q.i16(0)); break;
        case 0x020c: g.setWindowExt(q.i16(1), q.i16(0)); break;
        case 0x020f: g.offsetWindowOrg(q.i16(1), q.i16(0)); break;
        case 0x0410: g.scaleWindowExt(q.i16(3), q.i16(2), q.i16(1), q.i16(0)); break;
        case 0x0213: g.lineTo(q.i16(1), q.i16(0)); break;
        case 0x0214: g.moveTo(q.i16(1), q.i16(0)); break;
        case 0x0415: g.clipRect(q.i16(3), q.i16(2), q.i16(1), q.i16(0), true); break;
        case 0x0416: g.clipRect(q.i16(3), q.i16(2), q.i16(1), q.i16(0)); break;
        case 0x0817: case 0x081a: case 0x0830:
          g.arc({ 0x0817: 'arc', 0x081a: 'pie', 0x0830: 'chord' }[fn], q.i16(7), q.i16(6), q.i16(5), q.i16(4), q.i16(3), q.i16(2), q.i16(1), q.i16(0));
          break;
        case 0x0418: g.ellipse(q.i16(3), q.i16(2), q.i16(1), q.i16(0)); break;
        case 0x041b: g.rectangle(q.i16(3), q.i16(2), q.i16(1), q.i16(0)); break;
        case 0x061c: g.roundRect(q.i16(5), q.i16(4), q.i16(3), q.i16(2), q.i16(1), q.i16(0)); break;
        case 0x061d: g.blit({ x: q.i16(5), y: q.i16(4), cx: q.i16(3), cy: q.i16(2), rop: q.u32(0) }); break;
        case 0x001e: g.save(); break;
        case 0x0127: g.restore(q.i16(0)); break;
        case 0x041f: g.setPixel(q.i16(3), q.i16(2), q.u32(0)); break;
        case 0x0521: {
          const count = q.i16(0);
          if (count <= 0) break;
          const words = Math.ceil(count / 2);
          const raw = q.bytes(1, count);
          g.text({ x: q.i16(2 + words), y: q.i16(1 + words), str: decodeAnsi(raw, g.dc.font?.charset ?? 0) });
          break;
        }
        case 0x0a32: {
          const y = q.i16(0);
          const x = q.i16(1);
          const count = q.i16(2);
          const options = q.u16(3);
          let at = 4;
          let rect = null;
          if (options & 6) { rect = [q.i16(4), q.i16(5), q.i16(6), q.i16(7)]; at = 8; }
          if (count <= 0) { if (rect && (options & 2)) g.text({ x, y, str: ' ', options, rect }); break; }
          const raw = q.bytes(at, Math.min(count, (q.words - at) * 2));
          const str = decodeAnsi(raw, g.dc.font?.charset ?? 0);
          const dxAt = at + Math.ceil(count / 2);
          let dx = null;
          if (dxAt + count <= q.words && str.length === raw.length) {
            dx = new Array(count);
            for (let i = 0; i < count; i++) dx[i] = q.i16(dxAt + i);
          }
          g.text({ x, y, str, dx, options, rect });
          break;
        }
        case 0x0324: g.polygons([points(q, 1, q.i16(0))]); break;
        case 0x0325: g.polyline(points(q, 1, q.i16(0))); break;
        case 0x0538: {
          const n = Math.min(q.u16(0), q.words - 1);
          const counts = [];
          for (let i = 0; i < n; i++) counts.push(q.u16(1 + i));
          let from = 1 + n;
          const polys = [];
          for (const c of counts) {
            const pts = points(q, from, c);
            if (!pts.length) break;
            polys.push(pts);
            from += c * 2;
          }
          g.polygons(polys);
          break;
        }
        case 0x012d: g.select(q.u16(0)); break;
        case 0x01f0: {
          const i = q.u16(0);
          if (i < slots.length) { slots[i] = null; g.remove(i); }
          break;
        }
        case 0x012e: g.set('textAlign', q.u16(0)); break;
        case 0x02fa: create(makePen({ style: q.u16(0), width: q.i16(1), color: q.u32(3) })); break;
        case 0x02fc: create({ kind: 'brush', style: q.u16(0), color: q.u32(1) & 0xffffff, hatch: q.u16(3) }); break;
        case 0x02fb: {
          let face = '';
          for (let i = 18; i < Math.min(50, q.words * 2); i++) {
            const c = q.byte(i);
            if (!c) break;
            face += String.fromCharCode(c);
          }
          const charset = q.byte(13);
          create({
            kind: 'font', height: q.i16(0), width: q.i16(1), escapement: q.i16(2), orientation: q.i16(3), weight: q.i16(4),
            italic: Boolean(q.byte(10)), underline: Boolean(q.byte(11)), strike: Boolean(q.byte(12)), charset, pitchFamily: q.byte(17),
            face: decodeAnsi(Buffer.from(face, 'latin1'), charset === 2 ? 0 : charset),
          });
          break;
        }
        case 0x0142: {
          const style = q.u16(0);
          create(style === 3 ? { kind: 'brush', style: 0, color: 0x808080, hatch: 0 } : g.patternBrush(sourceDib(q, 2, q.u16(1))));
          break;
        }
        case 0x01f9: create({ kind: 'brush', style: 0, color: 0x808080, hatch: 0 }); break;
        case 0x00f7: create({ kind: 'palette' }); break;
        case 0x06ff: {
          // A region: scans of rectangles, each a top and bottom with left-right pairs.
          const rects = [];
          const scans = Math.min(q.i16(5), 4096);
          let i = 11;
          for (let s = 0; s < scans && i + 3 <= q.words; s++) {
            const count = q.u16(i);
            const top = q.i16(i + 1);
            const bottom = q.i16(i + 2);
            for (let k = 0; k + 1 < count && i + 3 + k + 1 < q.words; k += 2) rects.push([q.i16(i + 3 + k), top, q.i16(i + 4 + k), bottom]);
            i += 3 + count + 1;
          }
          create({ kind: 'region', rects });
          break;
        }
        case 0x012c: {
          const r = region(objectAt(q.u16(0)));
          const dev = r.map(([l, t, rr, bb]) => {
            const a = g.pt(l, t);
            const c = g.pt(rr, bb);
            return [Math.min(a[0], c[0]), Math.min(a[1], c[1]), Math.max(a[0], c[0]), Math.max(a[1], c[1])];
          });
          g.clipRegion(dev, 5);
          break;
        }
        case 0x0228: { const br = objectAt(q.u16(1)); g.fillRegion(region(objectAt(q.u16(0))), br?.kind === 'brush' ? br : g.dc.brush); break; }
        case 0x0429: { const br = objectAt(q.u16(1)); g.frameRegion(region(objectAt(q.u16(0))), br?.kind === 'brush' ? br : g.dc.brush, q.i16(3)); break; }
        case 0x012b: g.fillRegion(region(objectAt(q.u16(0)))); break;
        case 0x0940: {
          const rop = q.u32(0);
          if (size / 2 === (fn >> 8) + 3) { g.blit({ x: q.i16(8), y: q.i16(7), cx: q.i16(6), cy: q.i16(5), rop }); break; }
          const w = q.i16(5);
          const h = q.i16(4);
          blit(sourceDib(q, 8, 0), q.i16(3), q.i16(2), w, h, q.i16(7), q.i16(6), w, h, rop);
          break;
        }
        case 0x0b41: {
          const rop = q.u32(0);
          if (size / 2 === (fn >> 8) + 3) { g.blit({ x: q.i16(9), y: q.i16(8), cx: q.i16(7), cy: q.i16(6), rop }); break; }
          blit(sourceDib(q, 10, 0), q.i16(5), q.i16(4), q.i16(3), q.i16(2), q.i16(9), q.i16(8), q.i16(7), q.i16(6), rop);
          break;
        }
        case 0x0f43:
          blit(sourceDib(q, 11, q.u16(2)), q.i16(6), q.i16(5), q.i16(4), q.i16(3), q.i16(10), q.i16(9), q.i16(8), q.i16(7), q.u32(0));
          break;
        case 0x0d33: {
          const w = q.i16(6);
          const h = q.i16(5);
          blit(sourceDib(q, 9, q.u16(0)), q.i16(4), q.i16(3), w, h, q.i16(8), q.i16(7), w, h, ROP.SRCCOPY);
          break;
        }
        default: break;
      }
      p += size;
    }
  } catch (err) {
    // A record shorter than its fields, or a bound reached: what was drawn before it stands.
    if (!(err instanceof Enough) && !(err instanceof RangeError)) throw err;
  }
  const svg = g.finish({ view, width: css[0], height: css[1] });
  return svg ? { svg, width: Math.round(css[0]), height: Math.round(css[1]) } : null;
}

export default { wmfToSvg, isWmf };
