// A metafile's bitmaps, drawn into a PNG a window can show.
//
// A picture pasted into PowerPoint as an EMF showed as a broken picture in
// Presentations: no window draws a metafile. A bitmap in a metafile, alone
// or with the mask that makes its background see-through, now comes out as
// a PNG, see-through where PowerPoint shows the slide behind it. One that
// draws lines or words gives nothing, and the slide shows its frame.
import test from 'node:test';
import assert from 'node:assert/strict';
import { emfToPng, isEmf } from '../packages/office-formats/src/emf.js';
import { decodePng } from '../packages/ocr/src/png.js';

/** A record: its type, its size, its body. */
const record = (type, body) => {
  const head = Buffer.alloc(8);
  head.writeUInt32LE(type, 0);
  head.writeUInt32LE(8 + body.length, 4);
  return Buffer.concat([head, body]);
};
const ints = (...v) => {
  const b = Buffer.alloc(v.length * 4);
  v.forEach((n, i) => b.writeInt32LE(n, i * 4));
  return b;
};

/** A metafile of `w` x `h` device pixels holding these records. */
function emf(w, h, records) {
  const header = Buffer.alloc(100);
  header.writeUInt32LE(1, 0);
  header.writeUInt32LE(100, 4);
  ints(0, 0, w - 1, h - 1).copy(header, 8);
  ints(0, 0, w * 10, h * 10).copy(header, 24);
  header.writeUInt32LE(0x464d4520, 40);
  const body = Buffer.concat([...records, record(14, ints(0, 0, 20))]);
  header.writeUInt32LE(100 + body.length, 48);
  return Buffer.concat([header, body]);
}

/** A bottom-up DIB of `w` x `h`, 24 bits a pixel, `colour(x, yFromTop)` -> 0xRRGGBB. */
function dib24(w, h, colour) {
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
  return { info, bits };
}

/** A one-bit DIB, white where `on(x, yFromTop)`. */
function dib1(w, h, on) {
  const info = Buffer.alloc(48);
  info.writeUInt32LE(40, 0);
  info.writeInt32LE(w, 4);
  info.writeInt32LE(h, 8);
  info.writeUInt16LE(1, 12);
  info.writeUInt16LE(1, 14);
  info.writeUInt32LE(2, 32);
  info.writeUInt32LE(0x000000, 40);
  info.writeUInt32LE(0xffffff, 44);
  const stride = ((w + 31) >> 5) * 4;
  const bits = Buffer.alloc(stride * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) if (on(x, y)) bits[(h - 1 - y) * stride + (x >> 3)] |= 0x80 >> (x & 7);
  return { info, bits };
}

/** EMR_STRETCHDIBITS of the whole bitmap onto (x, y, cx, cy). */
function stretchDibits({ info, bits }, rop, [x, y, cx, cy]) {
  const w = info.readInt32LE(4);
  const h = Math.abs(info.readInt32LE(8));
  const fixed = 72;
  const body = Buffer.alloc(fixed);
  ints(x, y, x + cx - 1, y + cy - 1, x, y, 0, 0, w, h, 8 + fixed, info.length, 8 + fixed + info.length, bits.length, 0).copy(body, 0);
  body.writeUInt32LE(rop, 60);
  body.writeInt32LE(cx, 64);
  body.writeInt32LE(cy, 68);
  return record(81, Buffer.concat([body, info, bits]));
}

const pixelAt = (img, x, y) => {
  const at = (y * img.width + x) * 4;
  return [...img.data.subarray(at, at + 4)];
};

test('a bitmap in a metafile comes out as its pixels, stretched to the bounds', () => {
  const red = dib24(2, 2, (x, y) => (x === 0 && y === 0 ? 0xff0000 : 0x0000ff));
  const out = emfToPng(emf(4, 4, [stretchDibits(red, 0xcc0020, [0, 0, 4, 4])]));
  assert.ok(out, 'drawn');
  assert.deepEqual([out.width, out.height], [4, 4]);
  const img = decodePng(out.png);
  assert.deepEqual(pixelAt(img, 0, 0), [255, 0, 0, 255], 'the top-left source pixel fills the top-left quarter');
  assert.deepEqual(pixelAt(img, 1, 1), [255, 0, 0, 255]);
  assert.deepEqual(pixelAt(img, 3, 3), [0, 0, 255, 255], 'a bottom-up bitmap is drawn the right way up');
});

test('a mask then its picture, as PowerPoint keeps a picture with a see-through background', () => {
  // The mask is white where the picture is; the picture is white around it.
  const inside = (x, y) => x >= 1 && x <= 2 && y >= 1 && y <= 2;
  const mask = dib1(4, 4, inside);
  const picture = dib24(4, 4, (x, y) => (inside(x, y) ? 0x208040 : 0xffffff));
  const out = emfToPng(emf(4, 4, [stretchDibits(mask, 0xee0086, [0, 0, 4, 4]), stretchDibits(picture, 0x8800c6, [0, 0, 4, 4])]));
  const img = decodePng(out.png);
  assert.deepEqual(pixelAt(img, 1, 1), [0x20, 0x80, 0x40, 255], 'the picture is solid');
  assert.equal(pixelAt(img, 0, 0)[3], 0, 'around it the slide shows through');
  assert.equal(pixelAt(img, 3, 3)[3], 0);
});

test('a metafile that draws lines or words gives nothing, and so does one that is not a metafile', () => {
  const line = record(54, ints(10, 10));
  assert.equal(emfToPng(emf(4, 4, [line])), null);
  assert.equal(emfToPng(emf(4, 4, [])), null, 'nothing drawn');
  assert.equal(isEmf(Buffer.from('not a metafile at all, just words and more words to be long enough to read a header from it, really')), false);
  assert.equal(emfToPng(Buffer.alloc(10)), null);
});

test('a large metafile is drawn at most 2048 pixels on its longer side, its shape kept', () => {
  const grey = dib24(2, 2, () => 0x808080);
  const out = emfToPng(emf(4000, 3000, [stretchDibits(grey, 0xcc0020, [0, 0, 4000, 3000])]));
  assert.deepEqual([out.width, out.height], [2048, 1536]);
});
