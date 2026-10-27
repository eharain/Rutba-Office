// Windows metafiles made at test time: an EMF and a WMF writer, and two
// drawings made with them that use most of what a metafile can say.
//
// The writers put down records exactly as Windows' metafile device does:
// the EMF header with its bounds, frame and reference device, each record
// with its type and size; the WMF with the placeable header Word and
// PowerPoint read, its own header, and records whose fields run backwards.
// The tests, the fuzzer and the window checks all draw from here, so a
// picture checked in a window is the picture the engine tests read.

const u32 = (v) => { const b = Buffer.alloc(4); b.writeUInt32LE(v >>> 0, 0); return b; };
const i32 = (v) => { const b = Buffer.alloc(4); b.writeInt32LE(v | 0, 0); return b; };
const f32 = (v) => { const b = Buffer.alloc(4); b.writeFloatLE(v, 0); return b; };
const u16 = (v) => { const b = Buffer.alloc(2); b.writeUInt16LE(v & 0xffff, 0); return b; };
const i16 = (v) => { const b = Buffer.alloc(2); b.writeInt16LE(v | 0, 0); return b; };
const pad4 = (b) => (b.length % 4 ? Buffer.concat([b, Buffer.alloc(4 - (b.length % 4))]) : b);

/** A COLORREF from #rrggbb. */
export const rgb = (hex) => { const v = parseInt(hex.replace('#', ''), 16); return ((v & 255) << 16) | (v & 0xff00) | ((v >> 16) & 255); };

/** A bottom-up 24-bit DIB (BITMAPINFOHEADER and bits) of `w` x `h`, `colour(x, yFromTop)` -> 0xRRGGBB. */
export function dib24(w, h, colour) {
  const info = Buffer.alloc(40);
  info.writeUInt32LE(40, 0);
  info.writeInt32LE(w, 4);
  info.writeInt32LE(h, 8);
  info.writeUInt16LE(1, 12);
  info.writeUInt16LE(24, 14);
  const stride = ((w * 24 + 31) >> 5) * 4;
  const bits = Buffer.alloc(stride * h);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const c = colour(x, y);
      const at = (h - 1 - y) * stride + x * 3;
      bits[at] = c & 255;
      bits[at + 1] = (c >> 8) & 255;
      bits[at + 2] = (c >> 16) & 255;
    }
  }
  return { info, bits, w, h };
}

/**
 * An EMF writer on a 96 dpi reference device. Coordinates are logical
 * units; `build()` answers the file.
 */
export class EmfWriter {
  constructor(width, height) {
    this.width = width;
    this.height = height;
    this.records = [];
    this.handles = 1;
  }

  add(type, ...fields) {
    const body = pad4(Buffer.concat(fields));
    this.records.push(Buffer.concat([u32(type), u32(8 + body.length), body]));
    return this;
  }

  bounds() { return [i32(0), i32(0), i32(this.width - 1), i32(this.height - 1)]; }

  mapMode(m) { return this.add(17, u32(m)); }
  windowExt(x, y) { return this.add(9, i32(x), i32(y)); }
  windowOrg(x, y) { return this.add(10, i32(x), i32(y)); }
  viewportExt(x, y) { return this.add(11, i32(x), i32(y)); }
  viewportOrg(x, y) { return this.add(12, i32(x), i32(y)); }
  bkMode(m) { return this.add(18, u32(m)); }
  polyFill(m) { return this.add(19, u32(m)); }
  textAlign(a) { return this.add(22, u32(a)); }
  textColor(c) { return this.add(24, u32(c)); }
  bkColor(c) { return this.add(25, u32(c)); }
  save() { return this.add(33); }
  restore(n = -1) { return this.add(34, i32(n)); }
  world(m) { return this.add(35, ...m.map(f32)); }
  modifyWorld(m, mode) { return this.add(36, ...m.map(f32), u32(mode)); }
  select(ih) { return this.add(37, u32(ih)); }
  remove(ih) { return this.add(40, u32(ih)); }
  /** CreatePen: style, width in logical units (0 for a hairline), colour. */
  pen(ih, style, width, color) { this.handles = Math.max(this.handles, ih + 1); return this.add(38, u32(ih), u32(style), i32(width), i32(0), u32(color)); }
  /** ExtCreatePen, geometric, with a solid brush. */
  extPen(ih, style, width, color, entries = []) { this.handles = Math.max(this.handles, ih + 1); return this.add(95, u32(ih), u32(0), u32(0), u32(0), u32(0), u32(style | 0x10000), u32(width), u32(0), u32(color), u32(0), u32(entries.length), ...entries.map(u32)); }
  brush(ih, style, color, hatch = 0) { this.handles = Math.max(this.handles, ih + 1); return this.add(39, u32(ih), u32(style), u32(color), u32(hatch)); }
  font(ih, { height, weight = 400, italic = false, underline = false, escapement = 0, face = 'Arial', charset = 0 }) {
    this.handles = Math.max(this.handles, ih + 1);
    const name = Buffer.alloc(64);
    name.write(face.slice(0, 31), 0, 'utf16le');
    return this.add(82, u32(ih), i32(height), i32(0), i32(escapement), i32(escapement), i32(weight),
      Buffer.from([italic ? 1 : 0, underline ? 1 : 0, 0, charset, 0, 0, 4, 0x22]), name);
  }
  moveTo(x, y) { return this.add(27, i32(x), i32(y)); }
  lineTo(x, y) { return this.add(54, i32(x), i32(y)); }
  rectangle(l, t, r, b) { return this.add(43, i32(l), i32(t), i32(r), i32(b)); }
  ellipse(l, t, r, b) { return this.add(42, i32(l), i32(t), i32(r), i32(b)); }
  roundRect(l, t, r, b, w, h) { return this.add(44, i32(l), i32(t), i32(r), i32(b), i32(w), i32(h)); }
  arc(kind, l, t, r, b, xs, ys, xe, ye) { return this.add({ arc: 45, chord: 46, pie: 47 }[kind], i32(l), i32(t), i32(r), i32(b), i32(xs), i32(ys), i32(xe), i32(ye)); }
  polygon(pts) { return this.add(3, ...this.bounds(), u32(pts.length), ...pts.flatMap(([x, y]) => [i32(x), i32(y)])); }
  polyline(pts) { return this.add(4, ...this.bounds(), u32(pts.length), ...pts.flatMap(([x, y]) => [i32(x), i32(y)])); }
  polygon16(pts) { return this.add(86, ...this.bounds(), u32(pts.length), ...pts.flatMap(([x, y]) => [i16(x), i16(y)])); }
  polyline16(pts) { return this.add(87, ...this.bounds(), u32(pts.length), ...pts.flatMap(([x, y]) => [i16(x), i16(y)])); }
  polyBezier16(pts) { return this.add(85, ...this.bounds(), u32(pts.length), ...pts.flatMap(([x, y]) => [i16(x), i16(y)])); }
  intersectClip(l, t, r, b) { return this.add(30, i32(l), i32(t), i32(r), i32(b)); }
  beginPath() { return this.add(59); }
  endPath() { return this.add(60); }
  closeFigure() { return this.add(61); }
  fillPath() { return this.add(62, ...this.bounds()); }
  strokeAndFillPath() { return this.add(63, ...this.bounds()); }
  /** ExtTextOutW at (x, y), with each character's advance when `dx` is given. */
  text(x, y, str, { dx = null, options = 0 } = {}) {
    const chars = Buffer.from(str, 'utf16le');
    const offString = 76;
    const strBytes = pad4(chars);
    const offDx = dx ? offString + strBytes.length : 0;
    return this.add(84, ...this.bounds(), u32(1), f32(1), f32(1),
      i32(x), i32(y), u32(str.length), u32(offString), u32(options), i32(0), i32(0), i32(-1), i32(-1), u32(offDx),
      strBytes, ...(dx ? dx.map(i32) : []));
  }
  /** StretchDIBits of a whole bitmap into (x, y, cx, cy) with SRCCOPY. */
  stretchDib(x, y, cx, cy, { info, bits, w, h }, rop = 0xcc0020) {
    const offBmi = 80;
    const offBits = offBmi + info.length;
    return this.add(81, ...this.bounds(), i32(x), i32(y), i32(0), i32(0), i32(w), i32(h),
      u32(offBmi), u32(info.length), u32(offBits), u32(bits.length), u32(0), u32(rop), i32(cx), i32(cy), info, bits);
  }
  /** GradientFill: one rectangle shaded from `from` to `to`, across (0) or down (1). */
  gradientRect(l, t, r, b, from, to, mode = 0) {
    const vertex = (x, y, c) => [i32(x), i32(y), u16((c & 255) << 8), u16(((c >> 8) & 255) << 8), u16(((c >> 16) & 255) << 8), u16(0)];
    // A rectangle's two indices are followed by four bytes of padding, as Windows writes and checks them.
    return this.add(118, ...this.bounds(), u32(2), u32(1), u32(mode), ...vertex(l, t, from), ...vertex(r, b, to), u32(0), u32(1), u32(0));
  }
  comment(bytes) { return this.add(70, u32(bytes.length), bytes); }

  build() {
    const body = Buffer.concat([...this.records, Buffer.concat([u32(14), u32(20), u32(0), u32(16), u32(20)])]);
    const header = Buffer.alloc(108);
    header.writeUInt32LE(1, 0);
    header.writeUInt32LE(108, 4);
    [0, 0, this.width - 1, this.height - 1].forEach((v, i) => header.writeInt32LE(v, 8 + i * 4));
    // The frame, in hundredths of a millimetre, at 96 pixels an inch.
    [0, 0, Math.round(this.width / 96 * 2540), Math.round(this.height / 96 * 2540)].forEach((v, i) => header.writeInt32LE(v, 24 + i * 4));
    header.writeUInt32LE(0x464d4520, 40);
    header.writeUInt32LE(0x10000, 44);
    header.writeUInt32LE(108 + body.length, 48);
    header.writeUInt32LE(this.records.length + 2, 52);
    header.writeUInt16LE(this.handles, 56);
    header.writeInt32LE(1920, 72);
    header.writeInt32LE(1080, 76);
    header.writeInt32LE(508, 80);
    header.writeInt32LE(286, 84);
    header.writeUInt32LE(508000, 100);
    header.writeUInt32LE(285750, 104);
    return Buffer.concat([header, body]);
  }
}

/**
 * A WMF writer: the placeable header (a picture `width` x `height` units at
 * `inch` units an inch), then records. Its own fields are written in the
 * order the format keeps them, which is backwards.
 */
export class WmfWriter {
  constructor(width, height, inch = 1440) {
    this.width = width;
    this.height = height;
    this.inch = inch;
    this.records = [];
    this.slots = 0;
    this.biggest = 0;
  }

  add(fn, ...words) {
    const body = Buffer.concat(words.map((w) => (Buffer.isBuffer(w) ? w : i16(w))));
    const padded = body.length % 2 ? Buffer.concat([body, Buffer.alloc(1)]) : body;
    const size = 3 + padded.length / 2;
    this.biggest = Math.max(this.biggest, size);
    this.records.push(Buffer.concat([u32(size), u16(fn), padded]));
    return this;
  }

  windowOrg(x, y) { return this.add(0x020b, y, x); }
  windowExt(x, y) { return this.add(0x020c, y, x); }
  mapMode(m) { return this.add(0x0103, m); }
  bkMode(m) { return this.add(0x0102, m); }
  textAlign(a) { return this.add(0x012e, a); }
  textColor(c) { return this.add(0x0209, u32(c)); }
  polyFill(m) { return this.add(0x0106, m); }
  pen(style, width, color) { this.slots++; return this.add(0x02fa, style, width, 0, u32(color)); }
  brush(style, color, hatch = 0) { this.slots++; return this.add(0x02fc, style, u32(color), hatch); }
  font({ height, weight = 400, italic = false, escapement = 0, face = 'Arial', charset = 0 }) {
    this.slots++;
    const name = Buffer.alloc(32);
    name.write(face.slice(0, 31), 0, 'latin1');
    return this.add(0x02fb, height, 0, escapement, escapement, weight, Buffer.from([italic ? 1 : 0, 0, 0, charset, 0, 0, 4, 0x22]), name);
  }
  select(i) { return this.add(0x012d, i); }
  remove(i) { return this.add(0x01f0, i); }
  moveTo(x, y) { return this.add(0x0214, y, x); }
  lineTo(x, y) { return this.add(0x0213, y, x); }
  rectangle(l, t, r, b) { return this.add(0x041b, b, r, t, l); }
  ellipse(l, t, r, b) { return this.add(0x0418, b, r, t, l); }
  roundRect(l, t, r, b, w, h) { return this.add(0x061c, h, w, b, r, t, l); }
  arc(kind, l, t, r, b, xs, ys, xe, ye) { return this.add({ arc: 0x0817, pie: 0x081a, chord: 0x0830 }[kind], ye, xe, ys, xs, b, r, t, l); }
  polygon(pts) { return this.add(0x0324, pts.length, ...pts.flat()); }
  polyline(pts) { return this.add(0x0325, pts.length, ...pts.flat()); }
  polyPolygon(polys) { return this.add(0x0538, polys.length, ...polys.map((p) => p.length), ...polys.flat(2)); }
  save() { return this.add(0x001e); }
  restore(n = -1) { return this.add(0x0127, n); }
  intersectClip(l, t, r, b) { return this.add(0x0416, b, r, t, l); }
  /** ExtTextOut at (x, y), its 8-bit text, and each character's advance when `dx` is given. */
  text(x, y, str, { dx = null } = {}) {
    const bytes = Buffer.from(str, 'latin1');
    const padded = bytes.length % 2 ? Buffer.concat([bytes, Buffer.alloc(1)]) : bytes;
    return this.add(0x0a32, y, x, str.length, 0, padded, ...(dx ? dx : []));
  }
  /** TextOut, the oldest call: the length, the text, then the point. */
  textOut(x, y, str) {
    const bytes = Buffer.from(str, 'latin1');
    const padded = bytes.length % 2 ? Buffer.concat([bytes, Buffer.alloc(1)]) : bytes;
    return this.add(0x0521, str.length, padded, y, x);
  }
  /** StretchDIB of a whole bitmap into (x, y, cx, cy). */
  stretchDib(x, y, cx, cy, { info, bits, w, h }) {
    return this.add(0x0f43, u32(0xcc0020), 0, h, w, 0, 0, cy, cx, y, x, info, bits);
  }

  build({ placeable = true } = {}) {
    const body = Buffer.concat([...this.records, Buffer.concat([u32(3), u16(0)])]);
    const header = Buffer.concat([u16(1), u16(9), u16(0x300), u32((18 + body.length) / 2), u16(Math.max(this.slots, 1)), u32(this.biggest), u16(0)]);
    if (!placeable) return Buffer.concat([header, body]);
    const ph = Buffer.alloc(22);
    ph.writeUInt32LE(0x9ac6cdd7, 0);
    ph.writeInt16LE(0, 6);
    ph.writeInt16LE(0, 8);
    ph.writeInt16LE(this.width, 10);
    ph.writeInt16LE(this.height, 12);
    ph.writeUInt16LE(this.inch, 14);
    let sum = 0;
    for (let i = 0; i < 20; i += 2) sum ^= ph.readUInt16LE(i);
    ph.writeUInt16LE(sum, 20);
    return Buffer.concat([ph, header, body]);
  }
}

// Arial's advances for printable ASCII, in thousandths of an em.
const ARIAL = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556,
  278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722,
  667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556,
  333, 500, 278, 556, 500, 722, 500, 500, 500];

/** Each character's advance in Arial (a little wider when bold), as a program that measured them gives it. */
export function advances(str, em, { bold = false } = {}) {
  return Array.from(str).map((c) => {
    const code = c.charCodeAt(0);
    const w = code >= 32 && code < 32 + ARIAL.length ? ARIAL[code - 32] : 556;
    return Math.round((em * w * (bold ? 1.06 : 1)) / 1000);
  });
}

/**
 * A chart as a program pastes one into Office: a title, gridlines, bars with
 * a gradient and a hatched one, a line series with round markers, a pie, a
 * legend in a rounded box, a turned axis title and a small bitmap logo.
 * 480 x 300 device pixels.
 */
export function showcaseEmf() {
  const e = new EmfWriter(480, 300);
  const black = rgb('#000000');
  e.bkMode(1);
  // White page with a thin grey frame.
  e.pen(1, 0, 0, rgb('#a0a0a0')).brush(2, 0, rgb('#ffffff')).select(1).select(2).rectangle(0, 0, 479, 299);
  // The title, bold, centred on the top.
  e.font(3, { height: -20, weight: 700, face: 'Arial' }).select(3).textColor(rgb('#1f3864')).textAlign(6 | 0);
  e.text(240, 10, 'Quarterly sales', { dx: advances('Quarterly sales', 20, { bold: true }) });
  // Gridlines, dotted hairlines, and the axes.
  e.pen(4, 2, 0, rgb('#bfbfbf')).select(4);
  for (const y of [80, 130, 180, 230]) e.moveTo(60, y).lineTo(300, y);
  e.pen(5, 0, 2, black).select(5).moveTo(60, 50).lineTo(60, 260).lineTo(300, 260);
  // Labels down the axis, right-aligned to it, on their baselines.
  e.font(6, { height: -12, face: 'Arial' }).select(6).textColor(black).textAlign(2 | 24);
  [['300', 84], ['200', 134], ['100', 184], ['0', 264]].forEach(([s, y]) => e.text(54, y, s, { dx: advances(s, 12) }));
  // Bars: a solid one, a gradient one, a hatched one.
  e.pen(7, 0, 1, rgb('#2f5597')).brush(8, 0, rgb('#4472c4')).select(7).select(8).rectangle(80, 140, 120, 260);
  e.gradientRect(150, 100, 190, 260, rgb('#ed7d31'), rgb('#ffc000'), 1);
  e.brush(9, 2, rgb('#70ad47'), 5).select(9).rectangle(220, 170, 260, 260);
  // A line series, 3 units wide with round ends, and round markers.
  e.extPen(10, 0, 3, rgb('#c00000')).select(10).polyline([[100, 200], [170, 120], [240, 150], [290, 70]]);
  e.pen(11, 0, 1, rgb('#c00000')).brush(12, 0, rgb('#ffffff')).select(11).select(12);
  for (const [x, y] of [[100, 200], [170, 120], [240, 150], [290, 70]]) e.ellipse(x - 5, y - 5, x + 5, y + 5);
  // A pie of three slices.
  e.pen(13, 0, 1, rgb('#ffffff')).select(13);
  e.brush(14, 0, rgb('#4472c4')).select(14).arc('pie', 330, 60, 450, 180, 450, 120, 390, 60);
  e.brush(15, 0, rgb('#ed7d31')).select(15).arc('pie', 330, 60, 450, 180, 390, 60, 340, 150);
  e.brush(16, 0, rgb('#a5a5a5')).select(16).arc('pie', 330, 60, 450, 180, 340, 150, 450, 120);
  // A legend in a rounded box.
  e.pen(17, 0, 1, rgb('#7f7f7f')).brush(18, 0, rgb('#f2f2f2')).select(17).select(18).roundRect(330, 200, 460, 280, 16, 16);
  e.textAlign(0).font(19, { height: -12, face: 'Arial', italic: true }).select(19);
  [['North', '#4472c4', 210], ['South', '#ed7d31', 232], ['West', '#a5a5a5', 254]].forEach(([s, c, y], i) => {
    e.brush(20 + i, 0, rgb(c)).select(20 + i).pen(30 + i, 5, 0, 0).select(30 + i).rectangle(342, y + 2, 354, y + 14);
    e.text(362, y, s, { dx: advances(s, 12) });
  });
  // The axis title, turned a quarter (escapement 900 is anticlockwise).
  e.font(24, { height: -13, weight: 700, face: 'Arial', escapement: 900 }).select(24).textAlign(6 | 24);
  e.text(22, 155, 'Sales (k)', { dx: advances('Sales (k)', 13, { bold: true }) });
  // A 16 x 16 logo, a blue square with a white cross, stretched to 24 x 24.
  e.stretchDib(446, 6, 24, 24, dib24(16, 16, (x, y) => (x === 7 || x === 8 || y === 7 || y === 8 ? 0xffffff : 0x2f5597)));
  return e.build();
}

/**
 * A drawing as a Windows 3.1 program wrote one, in twips on a 3 x 2 inch
 * picture: shapes in pens and brushes, an arc, a polygon of two parts, and
 * words placed with TextOut and ExtTextOut.
 */
export function showcaseWmf() {
  const w = new WmfWriter(4320, 2880, 1440);
  w.windowOrg(0, 0).windowExt(4320, 2880).bkMode(1);
  w.pen(0, 15, rgb('#1f3864')).select(0); // 0
  w.brush(0, rgb('#dae3f3')).select(1); // 1
  w.roundRect(120, 120, 4200, 2760, 360, 360);
  w.brush(0, rgb('#ffc000')).select(2); // 2
  w.ellipse(360, 480, 1560, 1680);
  w.pen(0, 45, rgb('#c00000')).select(3); // 3
  w.arc('arc', 1800, 480, 3000, 1680, 3000, 1080, 1800, 1080);
  w.brush(2, rgb('#548235'), 3).select(4); // 4: hatched
  w.pen(0, 15, rgb('#385723')).select(5); // 5
  w.polyFill(1).polyPolygon([[[3120, 600], [3960, 600], [3960, 1560], [3120, 1560]], [[3360, 840], [3720, 840], [3720, 1320], [3360, 1320]]]);
  w.font({ height: -300, weight: 700, face: 'Times New Roman' }).select(6); // 6
  w.textColor(rgb('#1f3864')).textAlign(6 | 24);
  w.text(2160, 2160, 'Rutba drawing');
  w.font({ height: -220, italic: true, face: 'Arial' }).select(7); // 7
  w.textColor(rgb('#404040')).textAlign(0);
  w.textOut(360, 2340, 'a WMF made at test time');
  return w.build();
}
