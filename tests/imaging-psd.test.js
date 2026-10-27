// Photoshop documents, flattened: the picture Photoshop keeps beside the
// layers, in every colour mode and depth it saves, packed every way it packs
// it, PSD and the large PSB. Every document is laid out here, byte by byte.
import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { decodePicture, pictureInfo, needsDecoding } from '@rutba/imaging/decode';
import { packbitsEncode } from '../packages/imaging/src/codecs/lzw.js';
import { psdOf, sample } from '../tools/picture-sources.js';
import { judge } from '../tools/fuzz-open.js';

const px = (pic, x, y) => [...pic.rgba.subarray((y * pic.width + x) * 4, (y * pic.width + x) * 4 + 4)];

/**
 * A document: mode (1 grey, 2 indexed, 3 RGB, 4 CMYK, 9 Lab, 0 bitmap),
 * depth, planes (one Uint8Array of a channel's rows each), and how they are
 * packed (0 raw, 1 PackBits, 2 zip, 3 zip with prediction).
 */
function psd({ w, h, mode, depth = 8, planes, compression = 0, big = false, modeData = [], resources = [], layers = [] }) {
  const out = [];
  const u16 = (v) => out.push((v >> 8) & 255, v & 255);
  const u32 = (v) => out.push((v >>> 24) & 255, (v >> 16) & 255, (v >> 8) & 255, v & 255);
  out.push(0x38, 0x42, 0x50, 0x53);
  u16(big ? 2 : 1);
  out.push(0, 0, 0, 0, 0, 0);
  u16(planes.length);
  u32(h);
  u32(w);
  u16(depth);
  u16(mode);
  u32(modeData.length);
  out.push(...modeData);
  u32(resources.length);
  out.push(...resources);
  if (big) u32(0);
  u32(layers.length);
  out.push(...layers);
  u16(compression);
  const rowBytes = depth === 1 ? Math.ceil(w / 8) : (w * depth) / 8;
  if (compression === 0) for (const p of planes) out.push(...p);
  else if (compression === 1) {
    const rows = [];
    for (const p of planes) for (let y = 0; y < h; y++) rows.push(packbitsEncode(p.subarray(y * rowBytes, (y + 1) * rowBytes)));
    for (const r of rows) (big ? u32 : u16)(r.length);
    for (const r of rows) out.push(...r);
  } else {
    const all = new Uint8Array(planes.reduce((n, p) => n + p.length, 0));
    let at = 0;
    for (const p of planes) {
      const q = Uint8Array.from(p);
      if (compression === 3 && depth === 8) for (let y = 0; y < h; y++) for (let x = w - 1; x >= 1; x--) q[y * w + x] = (q[y * w + x] - q[y * w + x - 1]) & 255;
      all.set(q, at);
      at += q.length;
    }
    out.push(...zlib.deflateSync(all));
  }
  return Uint8Array.from(out);
}

const plane = (w, h, f) => Uint8Array.from({ length: w * h }, (_, i) => f(i % w, Math.floor(i / w)));

test('RGB documents, raw, PackBits, zipped and zipped with prediction, PSD and PSB, read as the picture', () => {
  const w = 21;
  const h = 7;
  const planes = [plane(w, h, (x) => x * 12), plane(w, h, (_, y) => y * 30), plane(w, h, (x, y) => (x * y) & 255)];
  for (const compression of [0, 1, 2, 3]) {
    for (const big of [false, true]) {
      const pic = decodePicture(psd({ w, h, mode: 3, planes, compression, big }), big ? 'a.psb' : 'a.psd');
      assert.equal(pic.format, 'psd');
      for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) assert.deepEqual(px(pic, x, y), [x * 12, y * 30, (x * y) & 255, 255], `${compression} ${big} ${x},${y}`);
    }
  }
  assert.deepEqual(pictureInfo(psd({ w, h, mode: 3, planes }), 'a.psd'), { pages: 1, format: 'psd', width: w, height: h });
  assert.equal(needsDecoding('a.psd'), true);
  assert.equal(needsDecoding('a.psb'), true);
});

test('grey, bitmap, indexed, CMYK and Lab documents, and 16-bit and 32-bit ones', () => {
  const grey = decodePicture(psd({ w: 2, h: 1, mode: 1, planes: [Uint8Array.from([10, 240])] }), 'g.psd');
  assert.deepEqual([px(grey, 0, 0), px(grey, 1, 0)], [[10, 10, 10, 255], [240, 240, 240, 255]]);
  // Bitmap: 1 is black.
  const bits = decodePicture(psd({ w: 3, h: 1, mode: 0, depth: 1, planes: [Uint8Array.from([0b01000000])] }), 'b.psd');
  assert.deepEqual([0, 1, 2].map((x) => px(bits, x, 0)[0]), [255, 0, 255]);
  // Indexed: 768 bytes of palette, reds then greens then blues.
  const pal = new Array(768).fill(0);
  pal[5] = 200; pal[256 + 5] = 100; pal[512 + 5] = 50;
  const ix = decodePicture(psd({ w: 1, h: 1, mode: 2, planes: [Uint8Array.from([5])], modeData: pal }), 'i.psd');
  assert.deepEqual(px(ix, 0, 0), [200, 100, 50, 255]);
  // CMYK: Photoshop keeps the ink inverted, 255 for none.
  const cmyk = decodePicture(psd({ w: 2, h: 1, mode: 4, planes: [[255, 0], [255, 255], [255, 255], [255, 204]].map((v) => Uint8Array.from(v)) }), 'c.psd');
  assert.deepEqual([px(cmyk, 0, 0), px(cmyk, 1, 0)], [[255, 255, 255, 255], [0, 204, 204, 255]]);
  // Lab: L 0..255 for 0..100, a and b about 128.
  const lab = decodePicture(psd({ w: 2, h: 1, mode: 9, planes: [[255, 0], [128, 128], [128, 128]].map((v) => Uint8Array.from(v)) }), 'l.psd');
  assert.ok(px(lab, 0, 0).slice(0, 3).every((v) => v >= 250) && px(lab, 1, 0).slice(0, 3).every((v) => v <= 5), String(px(lab, 0, 0)));
  // 16-bit: two bytes, high first; zipped with prediction too.
  const deep = (v) => Uint8Array.from([v >> 8, v & 255, (v + 0x1000) >> 8, (v + 0x1000) & 255]);
  for (const compression of [0, 1, 2]) {
    const d16 = decodePicture(psd({ w: 2, h: 1, mode: 3, depth: 16, planes: [deep(0x8000), deep(0x2000), deep(0x0000)], compression }), 'd.psd');
    assert.deepEqual(px(d16, 0, 0), [0x80, 0x20, 0x00, 255]);
    assert.deepEqual(px(d16, 1, 0), [0x90, 0x30, 0x10, 255]);
  }
  // 32-bit: linear light, shown through sRGB's curve.
  const f32 = (v) => { const b = new Uint8Array(4); new DataView(b.buffer).setFloat32(0, v); return b; };
  const lin = decodePicture(psd({ w: 1, h: 1, mode: 1, depth: 32, planes: [f32(0.214)] }), 'f.psd');
  assert.ok(Math.abs(px(lin, 0, 0)[0] - 128) <= 1, String(px(lin, 0, 0)));
});

test('transparency: kept where the layer count says the flattened picture has it, the white matte taken out; a saved selection is not shown', () => {
  const w = 3;
  const colour = [plane(w, 1, () => 255), plane(w, 1, () => 128), plane(w, 1, () => 128)];
  const alpha = Uint8Array.from([255, 128, 0]);
  // A negative layer count: the fourth channel is the picture's transparency.
  const negative = [0, 0, 0, 2, 0xff, 0xff, 0, 0]; // layer info of 2 bytes: a count of -1
  const see = decodePicture(psd({ w, h: 1, mode: 3, planes: [...colour, alpha], layers: negative }), 't.psd');
  assert.deepEqual(px(see, 0, 0), [255, 128, 128, 255]);
  assert.equal(px(see, 1, 0)[3], 128);
  assert.equal(px(see, 2, 0)[3], 0);
  // Matted on white: a half-clear (255, 128, 128) over white is stored as (255, 191, 191); read back as it was.
  const matted = [Uint8Array.from([255]), Uint8Array.from([191]), Uint8Array.from([191])];
  const back = decodePicture(psd({ w: 1, h: 1, mode: 3, planes: [...matted, Uint8Array.from([128])], layers: negative }), 'm.psd');
  assert.ok(Math.abs(px(back, 0, 0)[1] - 128) <= 1, String(px(back, 0, 0)));
  // A positive layer count: the fourth channel is a saved selection; the picture is opaque.
  const positive = [0, 0, 0, 2, 0, 1, 0, 0];
  const sel = decodePicture(psd({ w, h: 1, mode: 3, planes: [...colour, alpha], layers: positive }), 's.psd');
  assert.deepEqual([0, 1, 2].map((x) => px(sel, x, 0)[3]), [255, 255, 255]);
  // The writer the fuzzer and the window checks use: every pixel back within the matte's rounding.
  const src = sample(17, 9);
  const pic = decodePicture(psdOf(src, 17, 9, { alpha: true }), 'w.psd');
  for (let i = 0; i < src.length; i++) assert.ok(Math.abs(pic.rgba[i] - src[i]) <= 1, `${i}`);
  const big = decodePicture(psdOf(src, 17, 9, { alpha: true, big: true }), 'w.psb');
  assert.equal(big.rgba[2 * 4 + 3], src[2 * 4 + 3]);
});

test('a document saved without its flattened picture, or damaged, is a refusal in a sentence, never a fault', () => {
  // Version info (resource 1057) saying there is no real flattened picture.
  const res = [0x38, 0x42, 0x49, 0x4d, 0x04, 0x21, 0, 0, 0, 0, 0, 5, 0, 0, 0, 1, 0, 0];
  assert.throws(() => decodePicture(psd({ w: 1, h: 1, mode: 3, planes: [[1], [2], [3]].map((v) => Uint8Array.from(v)), resources: res }), 'x.psd'), /Maximize Compatibility/);
  assert.throws(() => decodePicture(psd({ w: 1, h: 1, mode: 5, planes: [Uint8Array.from([1])] }), 'x.psd'), /colour mode Rutba Office does not read/);
  const huge = psd({ w: 1, h: 1, mode: 3, planes: [[1], [2], [3]].map((v) => Uint8Array.from(v)) });
  new DataView(huge.buffer).setUint32(14, 300000);
  new DataView(huge.buffer).setUint32(18, 300000);
  assert.throws(() => decodePicture(huge, 'x.psd'), /larger than/);
  let refused = 0;
  const goods = [psdOf(sample(24, 16), 24, 16, { alpha: true }), psdOf(sample(24, 16), 24, 16, { packed: false })];
  for (const good of goods) {
    for (let i = 0; i < 400; i++) {
      const b = Buffer.from(good);
      const at = (i * 7919) % b.length;
      b[at] ^= 1 << (i % 8);
      if (i % 3 === 0) b[(at * 31) % b.length] = 0xff;
      try {
        decodePicture(i % 5 === 0 ? b.subarray(0, at) : b, 'x.psd');
      } catch (err) {
        assert.equal(judge(err).outcome, 'refused', `${err.name}: ${err.message} ${err.cause?.stack || ''}`);
        refused++;
      }
    }
  }
  assert.ok(refused > 0);
});
