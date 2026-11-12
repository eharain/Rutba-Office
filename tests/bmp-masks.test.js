// A bitmap's colour masks (packages/imaging/src/codecs/bmp.js): each channel
// read from the bits its mask names, scaled to 0 to 255. A mask with a gap in
// it (0xFF0000FF) read the bits above the gap as a value far past 255, which
// wrapped round to a wrong colour; its bits are now gathered and the value
// held to 255. The usual 5-6-5, 5-5-5, 8-8-8-8 and 10-10-10-2 masks read as
// they always have.
import test from 'node:test';
import assert from 'node:assert/strict';
import { decodeBmp, maskOf, channelOf } from '../packages/imaging/src/codecs/bmp.js';

/** A one-row bit-fields BMP: a 40-byte header, its masks after it, then the pixels, each `bpp` bits. */
function bitFields(bpp, masks, pixels) {
  const bytes = bpp >> 3;
  const stride = Math.floor((pixels.length * bpp + 31) / 32) * 4;
  const maskBytes = masks.length * 4;
  const pixelsAt = 14 + 40 + maskBytes;
  const b = Buffer.alloc(pixelsAt + stride);
  b.write('BM', 0, 'latin1');
  b.writeUInt32LE(b.length, 2);
  b.writeUInt32LE(pixelsAt, 10);
  b.writeUInt32LE(40, 14);
  b.writeInt32LE(pixels.length, 18);
  b.writeInt32LE(1, 22);
  b.writeUInt16LE(1, 26);
  b.writeUInt16LE(bpp, 28);
  b.writeUInt32LE(masks.length === 4 ? 6 : 3, 30); // BI_ALPHABITFIELDS with four masks, BI_BITFIELDS with three
  masks.forEach((m, i) => b.writeUInt32LE(m >>> 0, 54 + i * 4));
  pixels.forEach((v, i) => {
    if (bytes === 2) b.writeUInt16LE(v, pixelsAt + i * 2);
    else b.writeUInt32LE(v >>> 0, pixelsAt + i * 4);
  });
  return new Uint8Array(b);
}

const pixel = (rgba, i) => [...rgba.subarray(i * 4, i * 4 + 4)];

test('BMP: 5-6-5 bit fields read each channel from its own bits', () => {
  const b = decodeBmp(bitFields(16, [0xf800, 0x07e0, 0x001f], [0xf800, 0x07e0, 0x001f, 0xffff, 0x0000, 0x8410]));
  assert.equal(b.width, 6);
  assert.deepEqual(pixel(b.rgba, 0), [255, 0, 0, 255]);
  assert.deepEqual(pixel(b.rgba, 1), [0, 255, 0, 255]);
  assert.deepEqual(pixel(b.rgba, 2), [0, 0, 255, 255]);
  assert.deepEqual(pixel(b.rgba, 3), [255, 255, 255, 255]);
  assert.deepEqual(pixel(b.rgba, 4), [0, 0, 0, 255]);
  // The top bit of each: 16 of 31, 32 of 63, 16 of 31.
  assert.deepEqual(pixel(b.rgba, 5), [132, 130, 132, 255]);
});

test('BMP: 5-5-5 bit fields read five bits to each channel, the top bit unused', () => {
  const b = decodeBmp(bitFields(16, [0x7c00, 0x03e0, 0x001f], [0x7e01, 0xffff, 0x4210]));
  assert.deepEqual(pixel(b.rgba, 0), [255, 132, 8, 255], 'red 31, green 16, blue 1 of 31');
  assert.deepEqual(pixel(b.rgba, 1), [255, 255, 255, 255], 'the unused top bit does not count');
  assert.deepEqual(pixel(b.rgba, 2), [132, 132, 132, 255]);
});

test('BMP: 10-10-10-2 bit fields scale ten bits and two to 0 to 255', () => {
  // Red 512, green 256, blue 768 of 1023; alpha 2 of 3.
  const v = (2 * 2 ** 30) + (512 << 20) + (256 << 10) + 768;
  const b = decodeBmp(bitFields(32, [0x3ff00000, 0x000ffc00, 0x000003ff, 0xc0000000], [v]));
  assert.deepEqual(pixel(b.rgba, 0), [128, 64, 191, 170]);
});

test('BMP: 8-8-8-8 bit fields read the colour and the alpha', () => {
  const b = decodeBmp(bitFields(32, [0x00ff0000, 0x0000ff00, 0x000000ff, 0xff000000], [0x80ff4020, 0xff000000, 0x00102030]));
  assert.deepEqual(pixel(b.rgba, 0), [255, 64, 32, 128]);
  assert.deepEqual(pixel(b.rgba, 1), [0, 0, 0, 255]);
  assert.deepEqual(pixel(b.rgba, 2), [16, 32, 48, 0], 'see-through where its alpha is nought');
});

test('BMP: a mask with a gap gives every channel a value from 0 to 255, its bits gathered in order', () => {
  // Red's mask is the low byte and the top byte; green and blue the middle two.
  const gapped = 0xff0000ff;
  const masks = [gapped, 0x0000ff00, 0x00ff0000];
  const b = decodeBmp(bitFields(32, masks, [0xff0000ff, 0x000000ff, 0xff000000, 0x00000000, 0x80000000]));
  for (let i = 0; i < 5; i++) for (const c of pixel(b.rgba, i)) assert.ok(c >= 0 && c <= 255);
  assert.equal(pixel(b.rgba, 0)[0], 255, 'all its bits set: full red, not a value wrapped round');
  assert.equal(pixel(b.rgba, 1)[0], 1, 'the low byte alone: the low eight of sixteen bits');
  assert.equal(pixel(b.rgba, 2)[0], 254, 'the top byte alone: the high eight of sixteen bits');
  assert.equal(pixel(b.rgba, 3)[0], 0);
  assert.equal(pixel(b.rgba, 4)[0], 128, 'its highest bit alone: half, as the top bit of any channel is');
  assert.deepEqual(pixel(b.rgba, 0).slice(1), [0, 0, 255], 'green and blue untouched by it');
  const m = maskOf(gapped);
  assert.deepEqual({ shift: m.shift, bits: m.bits, max: m.max, gapped: m.gapped }, { shift: 0, bits: 16, max: 65535, gapped: true });
});

test('BMP masks: a channel value never leaves 0 to 255, whatever the mask', () => {
  for (const mask of [0x0000001f, 0x000007e0, 0xff000000, 0xffffffff, 0x80000001, 0x0f0f0f0f, 0x3ff00000, 0]) {
    const m = maskOf(mask);
    for (const v of [0, 1, 0x7fffffff, 0x80000000, 0xffffffff, 0x12345678]) {
      const c = channelOf(v >>> 0, m);
      assert.ok(Number.isInteger(c) && c >= 0 && c <= 255, `mask ${mask.toString(16)}, pixel ${v.toString(16)}: ${c}`);
    }
  }
  assert.equal(channelOf(0xffffffff, maskOf(0xffffffff)), 255, 'all ones, as a damaged header has');
  assert.equal(channelOf(0x3ff00000, maskOf(0x3ff00000)), 255, 'a ten-bit channel at full');
  assert.equal(maskOf(0x07e0).gapped, false);
});
