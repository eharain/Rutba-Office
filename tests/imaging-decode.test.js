// One way to read a picture (decodePicture): pixels from any format it
// reads, in any window or the main process, held to its limits, and a
// sentence where it cannot. Every picture here is made by the test.
import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { decodePicture, pictureInfo, pictureFormat, needsDecoding, orient, shrink, encodePng, LIMITS } from '@rutba/imaging/decode';
import { inflateRawJs, crc32 } from '../packages/imaging/src/codecs/zlib.js';
import { encodeBmp } from '../packages/imaging/src/codecs/bmp.js';
import { judge } from '../tools/fuzz-open.js';

/** A test picture: a gradient with a see-through column. */
export function gradient(w, h, { alpha = true } = {}) {
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

const same = (a, b) => Buffer.compare(Buffer.from(a), Buffer.from(b)) === 0;

/** A PNG of any colour type, depth and interlace, written by the test with Node's zlib. */
function png({ w, h, colour, depth, interlace = 0, sample, palette = null, trns = null }) {
  const ch = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colour];
  const rows = [];
  const passes = interlace
    ? [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]]
    : [[0, 0, 1, 1]];
  for (const [x0, y0, dx, dy] of passes) {
    const pw = Math.ceil((w - x0) / dx);
    const ph = Math.ceil((h - y0) / dy);
    if (pw <= 0 || ph <= 0) continue;
    for (let y = y0; y < h; y += dy) {
      const bits = [];
      for (let x = x0; x < w; x += dx) for (let c = 0; c < ch; c++) bits.push(sample(x, y, c));
      const rowBytes = Math.ceil((pw * ch * depth) / 8);
      const row = Buffer.alloc(rowBytes + 1);
      bits.forEach((v, i) => {
        if (depth === 16) row.writeUInt16BE(v, 1 + i * 2);
        else if (depth === 8) row[1 + i] = v;
        else row[1 + ((i * depth) >> 3)] |= v << (8 - depth - ((i * depth) & 7));
      });
      rows.push(row);
    }
  }
  const chunk = (type, body) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(body.length);
    const tb = Buffer.concat([Buffer.from(type, 'latin1'), body]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(tb));
    return Buffer.concat([len, tb, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = depth;
  ihdr[9] = colour;
  ihdr[12] = interlace;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10]),
    chunk('IHDR', ihdr),
    ...(palette ? [chunk('PLTE', Buffer.from(palette))] : []),
    ...(trns ? [chunk('tRNS', Buffer.from(trns))] : []),
    chunk('IDAT', zlib.deflateSync(Buffer.concat(rows))),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

test('PNG: what it writes it reads back exactly, opaque as RGB and see-through as RGBA', () => {
  for (const alpha of [true, false]) {
    const rgba = gradient(37, 23, { alpha });
    const bytes = encodePng(rgba, 37, 23);
    assert.equal(bytes[25], alpha ? 6 : 2, 'colour type');
    const pic = decodePicture(bytes, 'x.png');
    assert.equal(pic.format, 'png');
    assert.deepEqual([pic.width, pic.height, pic.pages, pic.page], [37, 23, 1, 0]);
    assert.ok(same(pic.rgba, rgba));
  }
});

test('PNG: grey, palette, grey with alpha, 1 to 16 bits, interlaced, and a transparent colour', () => {
  const w = 19;
  const h = 13;
  for (const interlace of [0, 1]) {
    for (const depth of [1, 2, 4, 8, 16]) {
      const max = 2 ** depth - 1;
      const v = (x, y) => (x + y * 3) % (max + 1);
      const pic = decodePicture(png({ w, h, colour: 0, depth, interlace, sample: (x, y) => v(x, y) }));
      for (const [x, y] of [[0, 0], [5, 7], [18, 12]]) {
        const want = depth === 16 ? v(x, y) >> 8 : Math.round((v(x, y) * 255) / max);
        assert.equal(pic.rgba[(y * w + x) * 4], want, `grey ${depth}-bit ${interlace ? 'interlaced' : ''} at ${x},${y}`);
      }
    }
    const palette = [255, 0, 0, 0, 255, 0, 0, 0, 255, 9, 9, 9];
    const pal = decodePicture(png({ w, h, colour: 3, depth: 2, interlace, sample: (x) => x % 4, palette, trns: [255, 128] }));
    assert.deepEqual([...pal.rgba.subarray(0, 8)], [255, 0, 0, 255, 0, 255, 0, 128]);
    assert.deepEqual([...pal.rgba.subarray(12, 16)], [9, 9, 9, 255]);
    const ga = decodePicture(png({ w, h, colour: 4, depth: 8, interlace, sample: (x, y, c) => (c ? 200 : x * 10) }));
    assert.deepEqual([...ga.rgba.subarray(4 * 3, 4 * 4)], [30, 30, 30, 200]);
    const key = decodePicture(png({ w, h, colour: 2, depth: 16, interlace, sample: (x, y, c) => (x === 2 ? 0x1234 : c * 0x4000), trns: [0x12, 0x34, 0x12, 0x34, 0x12, 0x34] }));
    assert.equal(key.rgba[2 * 4 + 3], 0, 'the transparent colour is clear');
    assert.equal(key.rgba[3 * 4 + 3], 255);
  }
});

test('BMP: 24 and 32 bits written and read back; a palette, RLE8, 16-bit fields, top-down and a core header read', () => {
  const rgba = gradient(13, 9);
  const back = decodePicture(encodeBmp(rgba, 13, 9), 'a.bmp');
  assert.ok(same(back.rgba, rgba), '32-bit with alpha');
  const opaque = gradient(13, 9, { alpha: false });
  const bmp24 = encodeBmp(opaque, 13, 9);
  assert.equal(bmp24[28], 24);
  assert.ok(same(decodePicture(bmp24, 'a.bmp').rgba, opaque));

  const file = (dib, pixelsAt) => {
    const head = Buffer.alloc(14);
    head.write('BM', 0, 'latin1');
    head.writeUInt32LE(14 + dib.length, 2);
    head.writeUInt32LE(pixelsAt, 10);
    return Buffer.concat([head, dib]);
  };
  const info = (w, h, bpp, compression, colours = 0) => {
    const b = Buffer.alloc(40);
    b.writeUInt32LE(40, 0);
    b.writeInt32LE(w, 4);
    b.writeInt32LE(h, 8);
    b.writeUInt16LE(1, 12);
    b.writeUInt16LE(bpp, 14);
    b.writeUInt32LE(compression, 16);
    b.writeUInt32LE(colours, 32);
    return b;
  };
  // 8-bit palette, top-down (negative height), rows padded to four bytes
  const pal = Buffer.from([0, 0, 255, 0, 0, 255, 0, 0]); // red, green (BGRA order)
  const rows8 = Buffer.from([0, 1, 0, 0, 1, 0, 1, 0]);
  const p8 = decodePicture(file(Buffer.concat([info(3, -2, 8, 0, 2), pal, rows8]), 14 + 40 + 8), 'p.bmp');
  assert.deepEqual([...p8.rgba.subarray(0, 12)], [255, 0, 0, 255, 0, 255, 0, 255, 255, 0, 0, 255]);
  assert.deepEqual([...p8.rgba.subarray(12, 16)], [0, 255, 0, 255]);
  // RLE8, bottom-up: the bottom row a run of 3 green, the top row literal red-green-red
  const rle = Buffer.from([3, 1, 0, 0, 0, 3, 0, 1, 0, 0, 0, 0, 0, 1]);
  const r8 = decodePicture(file(Buffer.concat([info(3, 2, 8, 1, 2), pal, rle]), 14 + 40 + 8), 'r.bmp');
  assert.deepEqual([...r8.rgba.subarray(0, 12)], [255, 0, 0, 255, 0, 255, 0, 255, 255, 0, 0, 255]);
  assert.deepEqual([...r8.rgba.subarray(12, 24)], [0, 255, 0, 255, 0, 255, 0, 255, 0, 255, 0, 255]);
  // 16-bit 5-6-5 bit fields
  const masks = Buffer.alloc(12);
  masks.writeUInt32LE(0xf800, 0);
  masks.writeUInt32LE(0x07e0, 4);
  masks.writeUInt32LE(0x001f, 8);
  const px = Buffer.alloc(4);
  px.writeUInt16LE(0xf800, 0);
  px.writeUInt16LE(0x07e0, 2);
  const b16 = decodePicture(file(Buffer.concat([info(2, 1, 16, 3), masks, px]), 14 + 40 + 12), 'f.bmp');
  assert.deepEqual([...b16.rgba], [255, 0, 0, 255, 0, 255, 0, 255]);
  // OS/2 core header, 24-bit
  const core = Buffer.alloc(12);
  core.writeUInt32LE(12, 0);
  core.writeUInt16LE(1, 4);
  core.writeUInt16LE(1, 6);
  core.writeUInt16LE(1, 8);
  core.writeUInt16LE(24, 10);
  const c24 = decodePicture(file(Buffer.concat([core, Buffer.from([10, 20, 30, 0])]), 26), 'c.bmp');
  assert.deepEqual([...c24.rgba], [30, 20, 10, 255]);
});

test('inflate in JavaScript, where a window has no zlib, unpacks what zlib packs at every level', () => {
  const data = Buffer.alloc(200000);
  for (let i = 0; i < data.length; i++) data[i] = i % 7 ? (i * 31) & 255 : i & 15;
  for (const level of [0, 1, 6, 9]) {
    const out = inflateRawJs(zlib.deflateRawSync(data, { level }), { size: data.length });
    assert.ok(same(out, data), `level ${level}`);
  }
  assert.throws(() => inflateRawJs(zlib.deflateRawSync(data), { size: 1000, max: 1000 }), /more than it says/);
  assert.throws(() => inflateRawJs(Buffer.from([0xff, 0xff, 0xff])), /damaged/);
});

test('limits: a header claiming a huge picture is refused before anything is allocated for it', () => {
  const huge = encodePng(gradient(4, 4), 4, 4);
  const b = Buffer.from(huge);
  b.writeUInt32BE(100000, 16);
  b.writeUInt32BE(100000, 20);
  assert.throws(() => decodePicture(b, 'big.png'), /larger than the 120 megapixels/);
  assert.throws(() => decodePicture(encodePng(gradient(40, 40), 40, 40), 'a.png', { limits: { pixels: 1000 } }), /megapixels/);
  assert.equal(LIMITS.pixels, 120e6);
  assert.throws(() => decodePicture(new Uint8Array(4), 'x.png'), /too short/);
});

test('formats it does not read are a sentence, or a decoder lent for them', () => {
  const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 16, 74, 70, 73, 70, 0, 1, 1, 0, 0, 1]);
  assert.equal(pictureFormat(jpeg), 'jpeg');
  assert.throws(() => decodePicture(jpeg, 'a.jpg'), /cannot draw JPEG pictures here/);
  const lent = decodePicture(jpeg, 'a.jpg', { system: () => encodePng(gradient(3, 2), 3, 2) });
  assert.deepEqual([lent.width, lent.height, lent.format], [3, 2, 'jpeg']);
  assert.throws(() => decodePicture(Buffer.from('just some words in a file'), 'notes.xyz'), /cannot draw \.xyz files here/);
  try {
    decodePicture(Buffer.from('just some words in a file'), 'notes.xyz');
  } catch (err) {
    assert.equal(judge(err).outcome, 'refused');
  }
  assert.equal(needsDecoding('a.png'), false);
  assert.equal(needsDecoding('a.bmp'), false);
  assert.deepEqual(pictureInfo(encodePng(gradient(5, 3), 5, 3)), { pages: 1, format: 'png', width: 5, height: 3 });
});

test('a damaged PNG is a refusal in a sentence, never a fault', () => {
  const good = Buffer.from(encodePng(gradient(30, 20), 30, 20));
  let refused = 0;
  for (let i = 0; i < 400; i++) {
    const b = Buffer.from(good);
    const at = 8 + ((i * 7919) % (b.length - 8));
    b[at] ^= 1 << (i % 8);
    const cut = i % 3 === 0 ? b.subarray(0, 10 + ((i * 31) % b.length)) : b;
    try {
      decodePicture(cut, 'd.png');
    } catch (err) {
      assert.equal(judge(err).outcome, 'refused', `${err.name}: ${err.message} ${err.cause?.stack || ''}`);
      refused++;
    }
  }
  assert.ok(refused > 0);
});

test('turned upright by an orientation, and shrunk by averaging', () => {
  // 2 × 1: red, blue
  const px = Uint8Array.from([255, 0, 0, 255, 0, 0, 255, 255]);
  const six = orient(px, 2, 1, 6); // a quarter clockwise
  assert.deepEqual([six.width, six.height], [1, 2]);
  assert.deepEqual([...six.rgba], [255, 0, 0, 255, 0, 0, 255, 255]);
  const three = orient(px, 2, 1, 3);
  assert.deepEqual([...three.rgba], [0, 0, 255, 255, 255, 0, 0, 255]);
  const s = shrink(gradient(100, 50, { alpha: false }), 100, 50, 10);
  assert.deepEqual([s.width, s.height], [10, 5]);
  assert.equal(decodePicture(encodePng(gradient(64, 32), 64, 32), 'a.png', { fit: 16 }).width, 16);
  const withPng = decodePicture(encodePng(gradient(8, 8), 8, 8), 'a.png', { png: true });
  assert.ok(same(decodePicture(withPng.png).rgba, withPng.rgba));
});
