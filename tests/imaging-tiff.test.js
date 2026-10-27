// TIFF everywhere, and icons: the suite's own TIFF reader and writer, and
// its icon and cursor reader. The Windows-made files in tests/fixtures/pictures
// were written by Windows' own encoder (tools/make-picture-fixtures.ps1) and
// are read here pixel for pixel against the PNG Windows wrote of the same
// picture; every other file is made by the test.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import { decodePicture, pictureInfo, encodeTiff, encodePng } from '@rutba/imaging/decode';
import { encodeIco, encodePicture, iconSizes } from '@rutba/imaging/encode';
import { lzwEncode, lzwDecode, packbitsDecode } from '../packages/imaging/src/codecs/lzw.js';
import { runTable, CODES } from '../packages/imaging/src/codecs/ccitt.js';
import { judge } from '../tools/fuzz-open.js';

/** A test picture: a gradient with a see-through column. */
function gradient(w, h, { alpha = true } = {}) {
  const rgba = new Uint8Array(w * h * 4);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 4;
      rgba[i] = (x * 255) / Math.max(1, w - 1);
      rgba[i + 1] = (y * 255) / Math.max(1, h - 1);
      rgba[i + 2] = (x * y) & 255;
      rgba[i + 3] = alpha && x === 1 ? 90 : 255;
    }
  }
  return rgba;
}

const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'pictures');
const fixture = (name) => fs.readFileSync(path.join(FIXTURES, name));
const same = (a, b) => Buffer.compare(Buffer.from(a), Buffer.from(b)) === 0;
const opaque = (rgba) => {
  const c = Uint8Array.from(rgba);
  for (let i = 3; i < c.length; i += 4) c[i] = 255;
  return c;
};

/* ── a TIFF writer for the tests: any tags, any chunks, either byte order, BigTIFF too ── */

const SIZE = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 11: 4, 16: 8 };
function valueBytes(type, vals, little) {
  const out = new Uint8Array(vals.length * SIZE[type]);
  const v = new DataView(out.buffer);
  vals.forEach((x, i) => {
    const at = i * SIZE[type];
    if (type === 1 || type === 2) out[at] = x;
    else if (type === 3) v.setUint16(at, x, little);
    else if (type === 4) v.setUint32(at, x, little);
    else if (type === 11) v.setFloat32(at, x, little);
    else if (type === 16) v.setBigUint64(at, BigInt(x), little);
    else if (type === 5) { v.setUint32(at, x, little); v.setUint32(at + 4, 1, little); }
  });
  return out;
}

/**
 * pages: [{ tags: { [tag]: [type, values] }, chunks: [Uint8Array] }]; the
 * strip or tile offsets and counts are filled in from the chunks.
 */
function buildTiff(pages, { little = true, big = false } = {}) {
  const head = big ? 16 : 8;
  const blobs = [];
  let size = head;
  const put = (bytes) => {
    const at = size;
    blobs.push([at, bytes]);
    size += bytes.length + (bytes.length & 1);
    return at;
  };
  const ifds = pages.map((pg) => {
    const offs = pg.chunks.map(put);
    const tags = { ...pg.tags };
    const tiled = 322 in tags;
    tags[tiled ? 324 : 273] = [big ? 16 : 4, offs];
    tags[tiled ? 325 : 279] = [4, pg.chunks.map((c) => c.length)];
    return Object.keys(tags).map(Number).sort((a, b) => a - b).map((tag) => ({ tag, type: tags[tag][0], bytes: valueBytes(tags[tag][0], tags[tag][1], little), count: tags[tag][1].length }));
  });
  const entry = big ? 20 : 12;
  const inline = big ? 8 : 4;
  const at = [];
  for (const list of ifds) {
    at.push(size);
    const ifdLen = (big ? 8 : 2) + list.length * entry + (big ? 8 : 4);
    let ext = size + ifdLen;
    for (const e of list) {
      if (e.bytes.length > inline) {
        e.at = ext;
        ext += e.bytes.length + (e.bytes.length & 1);
      }
    }
    size = ext;
  }
  const out = new Uint8Array(size);
  const v = new DataView(out.buffer);
  out[0] = out[1] = little ? 0x49 : 0x4d;
  v.setUint16(2, big ? 43 : 42, little);
  if (big) {
    v.setUint16(4, 8, little);
    v.setBigUint64(8, BigInt(at[0]), little);
  } else v.setUint32(4, at[0], little);
  for (const [o, bytes] of blobs) out.set(bytes, o);
  ifds.forEach((list, n) => {
    let p = at[n];
    if (big) { v.setBigUint64(p, BigInt(list.length), little); p += 8; } else { v.setUint16(p, list.length, little); p += 2; }
    for (const e of list) {
      v.setUint16(p, e.tag, little);
      v.setUint16(p + 2, e.type, little);
      if (big) v.setBigUint64(p + 4, BigInt(e.count), little); else v.setUint32(p + 4, e.count, little);
      const slot = p + (big ? 12 : 8);
      if (e.at !== undefined) {
        if (big) v.setBigUint64(slot, BigInt(e.at), little); else v.setUint32(slot, e.at, little);
        out.set(e.bytes, e.at);
      } else out.set(e.bytes, slot);
      p += entry;
    }
    const next = n + 1 < at.length ? at[n + 1] : 0;
    if (big) v.setBigUint64(p, BigInt(next), little); else v.setUint32(p, next, little);
  });
  return out;
}

const S = (...v) => [3, v];
const L = (...v) => [4, v];
const basic = (w, h, extra = {}) => ({ 256: L(w), 257: L(h), 278: L(h), ...extra });

/* ── Windows' own files ── */

test('TIFFs Windows wrote read pixel for pixel as Windows draws them: none, PackBits, LZW, Deflate, and fax in CCITT Group 3 and 4', () => {
  const colour = decodePicture(fixture('wic-colour.png'), 'wic-colour.png');
  for (const name of ['wic-none.tif', 'wic-rle.tif', 'wic-lzw.tif', 'wic-zip.tif']) {
    const pic = decodePicture(fixture(name), name);
    assert.equal(pic.format, 'tiff', name);
    assert.deepEqual([pic.width, pic.height, pic.pages], [colour.width, colour.height, 1], name);
    assert.ok(same(opaque(pic.rgba), opaque(colour.rgba)), `${name} is the picture`);
  }
  const bw = decodePicture(fixture('wic-bw.png'), 'wic-bw.png');
  for (const name of ['wic-bw-none.tif', 'wic-bw-rle.tif', 'wic-bw-lzw.tif', 'wic-bw-ccitt3.tif', 'wic-bw-ccitt4.tif']) {
    const pic = decodePicture(fixture(name), name);
    assert.deepEqual([pic.width, pic.height], [bw.width, bw.height], name);
    assert.ok(same(opaque(pic.rgba), opaque(bw.rgba)), `${name} is the picture`);
  }
});

test('a TIFF of three pages Windows wrote: each page read as itself, and a page past the last refused', () => {
  const bytes = fixture('wic-pages.tif');
  assert.deepEqual(pictureInfo(bytes, 'wic-pages.tif'), { pages: 3, format: 'tiff', width: 160, height: 120 });
  for (let page = 0; page < 3; page++) {
    const want = decodePicture(fixture(`wic-page-${page + 1}.png`), 'p.png');
    const pic = decodePicture(bytes, 'wic-pages.tif', { page });
    assert.equal(pic.page, page);
    assert.deepEqual([pic.width, pic.height], [want.width, want.height]);
    assert.ok(same(opaque(pic.rgba), opaque(want.rgba)), `page ${page + 1}`);
  }
  assert.throws(() => decodePicture(bytes, 'wic-pages.tif', { page: 3 }), /has 3 pages; there is no page 4/);
});

/* ── the suite's own writer ── */

test('TIFFs the suite writes read back exactly: LZW, Deflate, PackBits and none, with and without see-through, several pages', () => {
  const a = gradient(37, 23);
  const b = gradient(19, 41, { alpha: false });
  for (const compression of ['lzw', 'deflate', 'packbits', 'none']) {
    const bytes = encodeTiff([{ rgba: a, width: 37, height: 23 }, { rgba: b, width: 19, height: 41 }], { compression });
    assert.deepEqual(pictureInfo(bytes, 'x.tif'), { pages: 2, format: 'tiff', width: 37, height: 23 }, compression);
    const one = decodePicture(bytes, 'x.tif');
    assert.ok(same(one.rgba, a), `${compression}: page 1 with its alpha`);
    const two = decodePicture(bytes, 'x.tif', { page: 1 });
    assert.deepEqual([two.width, two.height], [19, 41]);
    assert.ok(same(two.rgba, b), `${compression}: page 2`);
  }
  // The writer's LZW is a stream the reader (and every other reader) undoes; a long run grows the codes past 12 bits' reset.
  const long = new Uint8Array(200000).map((_, i) => (i * 31) ^ (i >> 7));
  assert.ok(same(lzwDecode(lzwEncode(long), long.length), long));
  assert.ok(same(packbitsDecode(Uint8Array.from([2, 1, 2, 3, 0xfe, 9, 0x80]), 6), [1, 2, 3, 9, 9, 9]));
  assert.ok(same(encodePicture(a, 37, 23, 'tiff').subarray(0, 4), [0x49, 0x49, 42, 0]));
  assert.throws(() => encodeTiff([{ rgba: a, width: 37, height: 23 }], { compression: 'jbig' }), /no TIFF compression called jbig/);
});

/* ── TIFFs of every other kind, made here ── */

test('grey, bilevel, 4-bit and palette TIFFs, in either byte order', () => {
  const w = 9;
  const h = 5;
  const grey = Uint8Array.from({ length: w * h }, (_, i) => (i * 23) & 255);
  for (const little of [true, false]) {
    for (const photometric of [0, 1]) {
      const bytes = buildTiff([{ tags: basic(w, h, { 258: S(8), 259: S(1), 262: S(photometric), 277: S(1) }), chunks: [grey] }], { little });
      const pic = decodePicture(bytes, 'g.tif');
      for (let i = 0; i < w * h; i++) assert.equal(pic.rgba[i * 4], photometric ? grey[i] : 255 - grey[i]);
    }
  }
  // 1 bit, rows padded to a byte: 1010 1010 1 then 0 1 0 ...
  const bits = new Uint8Array(2 * h);
  for (let y = 0; y < h; y++) { bits[y * 2] = y & 1 ? 0x55 : 0xaa; bits[y * 2 + 1] = y & 1 ? 0 : 0x80; }
  const bi = decodePicture(buildTiff([{ tags: basic(w, h, { 258: S(1), 259: S(1), 262: S(1) }), chunks: [bits] }]), 'b.tif');
  assert.deepEqual([bi.rgba[0], bi.rgba[4], bi.rgba[8 * 4], bi.rgba[(w + 8) * 4]], [255, 0, 255, 0]);
  // 4 bits: 0..15 to 0..255
  const nib = Uint8Array.from([0x0f, 0x8a, 0x30]);
  const four = decodePicture(buildTiff([{ tags: basic(5, 1, { 258: S(4), 259: S(1), 262: S(1) }), chunks: [nib] }]), 'n.tif');
  assert.deepEqual([0, 1, 2, 3, 4].map((x) => four.rgba[x * 4]), [0, 255, 136, 170, 51]);
  // A palette of 16 (16-bit levels), 4 bits a pixel.
  const map = [];
  for (const c of [0, 1, 2]) for (let k = 0; k < 16; k++) map.push(((k * 16 + c * 5) & 255) * 257);
  const pal = decodePicture(buildTiff([{ tags: basic(5, 1, { 258: S(4), 259: S(1), 262: S(3), 320: [3, map] }), chunks: [nib] }], { little: false }), 'p.tif');
  assert.deepEqual([...pal.rgba.subarray(4, 8)], [240, 245, 250, 255]);
  assert.deepEqual([...pal.rgba.subarray(8, 12)], [128, 133, 138, 255]);
});

test('16-bit and 32-bit floating-point TIFFs, with their predictors, read at their levels', () => {
  const w = 6;
  const h = 4;
  for (const little of [true, false]) {
    // 16-bit RGB, horizontal differencing.
    const vals = Array.from({ length: w * h * 3 }, (_, i) => (i * 2777) & 0xffff);
    const raw = new Uint8Array(vals.length * 2);
    const dv = new DataView(raw.buffer);
    for (let y = 0; y < h; y++) {
      for (let i = w * 3 - 1; i >= 0; i--) {
        const at = y * w * 3 + i;
        dv.setUint16(at * 2, i >= 3 ? (vals[at] - vals[at - 3]) & 0xffff : vals[at], little);
      }
    }
    const deep = decodePicture(buildTiff([{ tags: basic(w, h, { 258: S(16, 16, 16), 259: S(1), 262: S(2), 277: S(3), 317: S(2) }), chunks: [raw] }], { little }), 'd.tif');
    for (let p = 0; p < w * h; p++) for (let c = 0; c < 3; c++) assert.equal(deep.rgba[p * 4 + c], vals[p * 3 + c] >> 8);
    // 32-bit float grey, 0..1, with the floating-point predictor.
    const fl = Array.from({ length: w * h }, (_, i) => i / (w * h - 1));
    const rowBytes = w * 4;
    const data = new Uint8Array(rowBytes * h);
    for (let y = 0; y < h; y++) {
      const sb = new Uint8Array(rowBytes);
      const sv = new DataView(sb.buffer);
      for (let x = 0; x < w; x++) sv.setFloat32(x * 4, fl[y * w + x], little);
      const planes = new Uint8Array(rowBytes);
      for (let i = 0; i < w; i++) for (let byte = 0; byte < 4; byte++) planes[(little ? 3 - byte : byte) * w + i] = sb[i * 4 + byte];
      for (let i = rowBytes - 1; i >= 1; i--) planes[i] = (planes[i] - planes[i - 1]) & 0xff;
      data.set(planes, y * rowBytes);
    }
    const float = decodePicture(buildTiff([{ tags: basic(w, h, { 258: S(32), 259: S(1), 262: S(1), 317: S(3), 339: S(3) }), chunks: [data] }], { little }), 'f.tif');
    for (let i = 0; i < w * h; i++) assert.equal(float.rgba[i * 4], Math.round(fl[i] * 255));
  }
});

test('tiled, planar and multi-strip TIFFs, and Deflate tiles with edges past the picture', () => {
  const w = 40;
  const h = 30;
  const src = gradient(w, h, { alpha: false });
  const rgb = (x, y) => [...src.subarray((y * w + x) * 4, (y * w + x) * 4 + 3)];
  // Tiles of 16 × 16, Deflate.
  const tiles = [];
  for (let ty = 0; ty < 2; ty++) {
    for (let tx = 0; tx < 3; tx++) {
      const t = new Uint8Array(16 * 16 * 3);
      for (let y = 0; y < 16; y++) for (let x = 0; x < 16; x++) if (tx * 16 + x < w && ty * 16 + y < h) t.set(rgb(tx * 16 + x, ty * 16 + y), (y * 16 + x) * 3);
      tiles.push(zlib.deflateSync(t));
    }
  }
  const tiled = decodePicture(buildTiff([{ tags: { 256: L(w), 257: L(h), 258: S(8, 8, 8), 259: S(8), 262: S(2), 277: S(3), 322: L(16), 323: L(16) }, chunks: tiles }]), 't.tif');
  assert.ok(same(tiled.rgba, src), 'tiled');
  // Planar: three planes, each in strips of 7 rows, LZW.
  const strips = [];
  for (let c = 0; c < 3; c++) {
    for (let y0 = 0; y0 < h; y0 += 7) {
      const rows = Math.min(7, h - y0);
      const s = new Uint8Array(rows * w);
      for (let y = 0; y < rows; y++) for (let x = 0; x < w; x++) s[y * w + x] = rgb(x, y0 + y)[c];
      strips.push(lzwEncode(s));
    }
  }
  const planar = decodePicture(buildTiff([{ tags: { 256: L(w), 257: L(h), 258: S(8, 8, 8), 259: S(5), 262: S(2), 277: S(3), 278: L(7), 284: S(2) }, chunks: strips }], { little: false }), 'pl.tif');
  assert.ok(same(planar.rgba, src), 'planar');
});

test('YCbCr, CMYK, see-through (associated and not) and a turned TIFF', () => {
  // YCbCr 2 × 2: one block of four lumas and one Cb, Cr.
  const Y = [40, 80, 120, 160];
  const ycc = decodePicture(buildTiff([{ tags: basic(2, 2, { 258: S(8, 8, 8), 259: S(1), 262: S(6), 277: S(3), 530: S(2, 2) }), chunks: [Uint8Array.from([...Y, 100, 170])] }]), 'y.tif');
  const cb = 100 - 128;
  const cr = 170 - 128;
  const clamp = (v) => Math.max(0, Math.min(255, Math.round(v)));
  Y.forEach((y, i) => assert.deepEqual([...ycc.rgba.subarray(i * 4, i * 4 + 3)], [clamp(y + 1.402 * cr), clamp(y - 0.344136 * cb - 0.714136 * cr), clamp(y + 1.772 * cb)]));
  // CMYK: ink taken away from white.
  const cmyk = decodePicture(buildTiff([{ tags: basic(2, 1, { 258: S(8, 8, 8, 8), 259: S(1), 262: S(5), 277: S(4) }), chunks: [Uint8Array.from([0, 0, 0, 0, 255, 0, 128, 51])] }]), 'c.tif');
  assert.deepEqual([...cmyk.rgba.subarray(0, 4)], [255, 255, 255, 255]);
  assert.deepEqual([...cmyk.rgba.subarray(4, 8)], [0, 204, Math.round((127 * 204) / 255), 255]);
  // RGBA, unassociated alpha (2) kept; associated (1) unpremultiplied.
  const straight = decodePicture(buildTiff([{ tags: basic(1, 1, { 258: S(8, 8, 8, 8), 259: S(1), 262: S(2), 277: S(4), 338: S(2) }), chunks: [Uint8Array.from([200, 100, 50, 128])] }]), 'a.tif');
  assert.deepEqual([...straight.rgba], [200, 100, 50, 128]);
  const assoc = decodePicture(buildTiff([{ tags: basic(1, 1, { 258: S(8, 8, 8, 8), 259: S(1), 262: S(2), 277: S(4), 338: S(1) }), chunks: [Uint8Array.from([100, 50, 25, 128])] }]), 'a.tif');
  assert.ok(Math.abs(assoc.rgba[0] - 199) <= 1 && Math.abs(assoc.rgba[1] - 100) <= 1 && assoc.rgba[3] === 128, String([...assoc.rgba]));
  // Orientation 6: stored 3 × 1, shown 1 × 3 (unless asked not to turn it).
  const turned = buildTiff([{ tags: basic(3, 1, { 258: S(8), 259: S(1), 262: S(1), 274: S(6) }), chunks: [Uint8Array.from([10, 20, 30])] }]);
  const up = decodePicture(turned, 't.tif');
  assert.deepEqual([up.width, up.height, up.rgba[0], up.rgba[8]], [1, 3, 10, 30]);
  assert.deepEqual([decodePicture(turned, 't.tif', { orient: false }).width], [3]);
});

test('BigTIFF in either byte order, and the reduced-size copies a TIFF carries are not counted as pages', () => {
  const px = Uint8Array.from([1, 2, 3, 4, 5, 6]);
  for (const little of [true, false]) {
    const pic = decodePicture(buildTiff([{ tags: basic(2, 1, { 258: S(8, 8, 8), 259: S(1), 262: S(2), 277: S(3) }), chunks: [px] }], { little, big: true }), 'big.tif');
    assert.deepEqual([...pic.rgba], [1, 2, 3, 255, 4, 5, 6, 255]);
  }
  const page = (n, sub) => ({ tags: basic(1, 1, { 254: L(sub), 258: S(8), 259: S(1), 262: S(1) }), chunks: [Uint8Array.from([n])] });
  const bytes = buildTiff([page(10, 0), page(20, 1), page(30, 2)]);
  assert.equal(pictureInfo(bytes, 'r.tif').pages, 2);
  assert.equal(decodePicture(bytes, 'r.tif', { page: 1 }).rgba[0], 30);
});

/* ── fax codes ── */

test('the fax code tables are prefix codes, and a one-dimensional (Modified Huffman) fax reads', () => {
  for (const white of [true, false]) {
    const codes = [...runTable(white).keys()].map((k) => (k & 0xffff).toString(2).padStart(k >> 16, '0'));
    assert.equal(new Set(codes).size, codes.length);
    for (const a of codes) for (const b of codes) if (a !== b) assert.ok(!b.startsWith(a), `${white ? 'white' : 'black'} ${a} begins ${b}`);
  }
  assert.equal(CODES.WHITE_TERM.split(' ').length, 64);
  assert.equal(CODES.BLACK_MAKEUP.split(' ').length, 27);
  // Two rows of 8: all white (10011), then 4 white (1011) and 4 black (011); each row starts on a byte.
  const fax = Uint8Array.from([0b10011000, 0b10110110]);
  const pic = decodePicture(buildTiff([{ tags: basic(8, 2, { 258: S(1), 259: S(2), 262: S(0) }), chunks: [fax] }]), 'mh.tif');
  const row = (y) => [0, 1, 2, 3, 4, 5, 6, 7].map((x) => pic.rgba[(y * 8 + x) * 4]);
  assert.deepEqual(row(0), [255, 255, 255, 255, 255, 255, 255, 255]);
  assert.deepEqual(row(1), [255, 255, 255, 255, 0, 0, 0, 0]);
});

/* ── icons and cursors ── */

test('icons and cursors: every size read, the largest first, PNG or bitmap inside, and the icon Image writes', () => {
  const sizes = [16, 48, 32];
  const ico = encodeIco(sizes.map((s) => ({ width: s, height: s, png: encodePng(gradient(s, s), s, s) })));
  assert.deepEqual(pictureInfo(ico, 'a.ico'), { pages: 3, format: 'ico', width: 48, height: 48 });
  assert.deepEqual([0, 1, 2].map((page) => decodePicture(ico, 'a.ico', { page }).width), [48, 32, 16]);
  assert.ok(same(decodePicture(ico, 'a.ico', { page: 2 }).rgba, gradient(16, 16)));
  assert.throws(() => decodePicture(ico, 'a.ico', { page: 3 }), /has 3 sizes; there is no size 4/);
  // A cursor: the same directory, type 2, a hotspot where an icon has planes and depth.
  const cur = encodeIco([{ width: 32, height: 32, png: encodePng(gradient(32, 32), 32, 32) }], { cursor: true, hotspots: [{ x: 5, y: 7 }] });
  assert.deepEqual([cur[2], cur[10], cur[12]], [2, 5, 7]);
  assert.equal(decodePicture(cur, 'a.cur').format, 'cur');
  // A classic bitmap icon: 2 × 2, 32 bits, its height doubled for the AND mask.
  const dib = new Uint8Array(40 + 16 + 8);
  const dv = new DataView(dib.buffer);
  dv.setUint32(0, 40, true); dv.setInt32(4, 2, true); dv.setInt32(8, 4, true); dv.setUint16(12, 1, true); dv.setUint16(14, 32, true);
  dib.set([0, 0, 255, 255, 0, 255, 0, 255, 255, 0, 0, 255, 255, 255, 255, 0], 40); // bottom row: red, green; top row: blue, see-through white
  const one = new Uint8Array(6 + 16 + dib.length);
  const ov = new DataView(one.buffer);
  ov.setUint16(2, 1, true); ov.setUint16(4, 1, true);
  one[6] = 2; one[7] = 2; ov.setUint16(10, 1, true); ov.setUint16(12, 32, true); ov.setUint32(14, dib.length, true); ov.setUint32(18, 22, true);
  one.set(dib, 22);
  const bmp = decodePicture(one, 'b.ico');
  assert.deepEqual([...bmp.rgba], [0, 0, 255, 255, 255, 255, 255, 0, 255, 0, 0, 255, 0, 255, 0, 255]);
  // Image's icon: every standard size up to the picture's, never enlarged.
  assert.deepEqual(iconSizes(300, 100), [256, 128, 64, 48, 32, 24, 16]);
  assert.deepEqual(iconSizes(40, 30), [32, 24, 16]);
  assert.deepEqual(iconSizes(10, 9), [10]);
  const own = encodePicture(gradient(20, 20), 20, 20, 'ico');
  assert.ok(same(decodePicture(own, 'own.ico').rgba, gradient(20, 20)));
});

/* ── damage ── */

test('a damaged TIFF or icon is a refusal in a sentence, never a fault, and an absurd size is refused before memory is taken', () => {
  const huge = buildTiff([{ tags: basic(60000, 60000, { 258: S(8, 8, 8), 259: S(1), 262: S(2), 277: S(3) }), chunks: [new Uint8Array(4)] }]);
  assert.throws(() => decodePicture(huge, 'huge.tif'), (err) => err.picture && /too large|larger/i.test(err.message));
  // An IFD that points at itself is one page, not a hang.
  const loop = buildTiff([{ tags: basic(1, 1, { 258: S(8), 259: S(1), 262: S(1) }), chunks: [Uint8Array.from([7])] }]);
  const ifd = new DataView(loop.buffer).getUint32(4, true);
  const n = new DataView(loop.buffer).getUint16(ifd, true);
  new DataView(loop.buffer).setUint32(ifd + 2 + n * 12, ifd, true);
  assert.equal(pictureInfo(loop, 'loop.tif').pages, 1);
  assert.throws(() => decodePicture(buildTiff([{ tags: basic(4, 4, { 258: S(8), 259: S(34712), 262: S(1) }), chunks: [new Uint8Array(4)] }]), 'j2k.tif'), /JPEG 2000/);
  assert.throws(() => decodePicture(Uint8Array.from([0x49, 0x49, 42, 0, 200, 0, 0, 0]), 'short.tif'), (err) => judge(err).outcome === 'refused');

  const goods = [
    ['x.tif', encodeTiff([{ rgba: gradient(24, 16), width: 24, height: 16 }, { rgba: gradient(8, 8), width: 8, height: 8 }])],
    ['x.tif', Buffer.from(fixture('wic-bw-ccitt4.tif'))],
    ['x.tif', Buffer.from(fixture('wic-bw-ccitt3.tif'))],
    ['x.ico', encodeIco([{ width: 16, height: 16, png: encodePng(gradient(16, 16), 16, 16) }])],
  ];
  let refused = 0;
  for (const [name, good] of goods) {
    for (let i = 0; i < 300; i++) {
      const b = Buffer.from(good);
      const at = (i * 7919) % b.length;
      b[at] ^= 1 << (i % 8);
      if (i % 3 === 0) b[(at * 31) % b.length] = 0xff;
      try {
        decodePicture(i % 5 === 0 ? b.subarray(0, at) : b, name, { page: i % 2 });
      } catch (err) {
        assert.equal(judge(err).outcome, 'refused', `${name}: ${err.name}: ${err.message} ${err.cause?.stack || ''}`);
        refused++;
      }
    }
  }
  assert.ok(refused > 0);
});
