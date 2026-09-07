'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const zlib = require('zlib');
const { decodePng, isPng, PdfDocument, renderPaper } = require('../src/index');

/**
 * PNGs built by hand, chunk by chunk, so the decoder is tested against the
 * format rather than against another library's opinion of it. The decoder
 * ignores CRCs on purpose (a signature upload with a bad CRC but good pixels
 * should still sign), so these write zeros there.
 */
function chunk(type, body) {
  const out = Buffer.alloc(12 + body.length);
  out.writeUInt32BE(body.length, 0);
  out.write(type, 4, 'latin1');
  body.copy(out, 8);
  return out;
}

function png({ width, height, colorType, raw, palette = null, trns = null }) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = colorType;
  const parts = [
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
  ];
  if (palette) parts.push(chunk('PLTE', palette));
  if (trns) parts.push(chunk('tRNS', trns));
  parts.push(chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0)));
  return Buffer.concat(parts);
}

test('an RGBA png decodes to RGB samples plus a separate alpha channel', () => {
  // 2x1: an opaque red pixel and a half-transparent blue one, filter 0.
  const raw = Buffer.from([0, 255, 0, 0, 255, 0, 0, 255, 128]);
  const image = decodePng(png({ width: 2, height: 1, colorType: 6, raw }));
  assert.equal(image.colorSpace, 'DeviceRGB');
  assert.deepEqual([...image.data], [255, 0, 0, 0, 0, 255]);
  assert.deepEqual([...image.alpha], [255, 128]);
});

test('the sub filter reconstructs left-relative scanlines', () => {
  // 2x1 RGB, filter 1: second pixel stored as deltas from the first.
  const raw = Buffer.from([1, 10, 20, 30, 5, 5, 5]);
  const image = decodePng(png({ width: 2, height: 1, colorType: 2, raw }));
  assert.deepEqual([...image.data], [10, 20, 30, 15, 25, 35]);
});

test('the up and paeth filters reconstruct against the prior line', () => {
  // 1x3 grey: line 1 plain, line 2 up-filtered, line 3 paeth-filtered.
  const raw = Buffer.from([0, 100, 2, 28, 4, 100]);
  const image = decodePng(png({ width: 1, height: 3, colorType: 0, raw }));
  assert.equal(image.colorSpace, 'DeviceGray');
  assert.deepEqual([...image.data], [100, 128, 228]);
});

test('a palette png expands through PLTE, with tRNS becoming alpha', () => {
  const palette = Buffer.from([255, 0, 0, 0, 0, 255]);
  const raw = Buffer.from([0, 0, 1]); // one row: red, blue
  const image = decodePng(png({ width: 2, height: 1, colorType: 3, raw, palette, trns: Buffer.from([200]) }));
  assert.deepEqual([...image.data], [255, 0, 0, 0, 0, 255]);
  assert.deepEqual([...image.alpha], [200, 255]);
});

test('what it cannot decode it refuses by name', () => {
  assert.throws(() => decodePng(Buffer.from('not a png')), /signature/);
  const ihdr16 = png({ width: 1, height: 1, colorType: 0, raw: Buffer.from([0, 0]) });
  ihdr16[24] = 16; // bit depth byte inside IHDR
  assert.throws(() => decodePng(ihdr16), /8-bit/);
  assert.equal(isPng(Buffer.from('plain text')), false);
});

test('an embedded image lands in the file as a flate XObject with an SMask', () => {
  const raw = Buffer.from([0, 255, 0, 0, 255, 0, 0, 255, 128]);
  const image = decodePng(png({ width: 2, height: 1, colorType: 6, raw }));
  const doc = new PdfDocument({ created: '2026-09-02T00:00:00Z' });
  const page = doc.addPage();
  const ref = doc.addImage(image);
  page.image(ref, 50, 50, 100, 50);
  page.text('beside the mark', 50, 120);
  const bytes = doc.toBuffer().toString('latin1');
  assert.ok(bytes.includes('/Subtype /Image'));
  assert.ok(bytes.includes('/SMask'));
  assert.ok(bytes.includes('/XObject << /Im1'));
  assert.ok(bytes.includes('/Im1 Do'));
});

test('renderPaper draws a signatures block from images and typed adoptions', () => {
  const raw = Buffer.from([0, 255, 0, 0, 255, 0, 0, 255, 128]);
  const image = decodePng(png({ width: 2, height: 1, colorType: 6, raw }));
  const bytes = renderPaper({
    letterhead: { name: 'Test' },
    title: 'Certificate of Completion',
    created: '2026-09-02T00:00:00Z',
    signatures: [
      { name: 'Alma Signer', detail: 'signer - signed', image },
      { name: 'Kai Typed', detail: 'signer - signed', typedName: 'Kai Typed' },
    ],
  }).toString('latin1');
  assert.ok(bytes.includes('/Subtype /Image'), 'the drawn mark is embedded');
  assert.ok(bytes.includes('Signatures as adopted'));
  assert.ok(bytes.includes('Kai Typed'));
});
