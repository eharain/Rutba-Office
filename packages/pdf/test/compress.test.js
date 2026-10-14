'use strict';

// Reduce File Size (compress.js): a picture shown at more detail than the
// page needs is planned for at the size the target resolution asks, drawn
// smaller (here, resampled and deflated; in the window, encoded as JPEG and
// handed back), and the file written whole with it in place of the old,
// which is gone; a picture already small enough, and the words, are left as
// they were.

const test = require('node:test');
const assert = require('node:assert');
const { PdfDocument } = require('../src/document');
const { openPdf, Stream } = require('../src/reader');
const { compressPlan, compressPdf, resample } = require('../src/compress');

/** A page with a large photograph shown small, a small one, and words. */
function photos() {
  const doc = new PdfDocument({ title: 'Photos', created: '2026-10-14T09:00:00Z' });
  const big = 1200;
  const rgb = Buffer.alloc(big * big * 3);
  // Noise, so deflate cannot make it small by itself.
  let s = 7;
  for (let i = 0; i < rgb.length; i++) { s ^= s << 13; s ^= s >>> 17; s ^= s << 5; rgb[i] = s & 255; }
  const large = doc.addImage({ width: big, height: big, colorSpace: 'DeviceRGB', data: rgb });
  const small = doc.addImage({ width: 40, height: 40, colorSpace: 'DeviceGray', data: Buffer.alloc(1600, 128) });
  const page = doc.addPage();
  page.text('Two photographs.', 72, 80, { font: 'Helvetica', size: 14 });
  page.image(large, 72, 100, 144, 144); // two inches: 600 dots an inch
  page.image(small, 300, 100, 144, 144);
  return doc.toBuffer();
}

test('a picture shown at more detail than the target is planned for at the size the target asks', () => {
  const plan = compressPlan(photos(), { dpi: 150 });
  assert.equal(plan.length, 1, 'the small picture is left alone');
  assert.equal(plan[0].width, 1200);
  assert.equal(plan[0].toWidth, 300);
  assert.equal(plan[0].toHeight, 300);
  assert.equal(plan[0].comps, 3);
});

test('resampled here, the file is written whole, smaller, with the words and the small picture as they were', () => {
  const original = photos();
  const { bytes, replaced, before, after } = compressPdf(original, { dpi: 150 });
  assert.equal(replaced, 1);
  assert.ok(after < before / 4, `${before} to ${after}`);
  const file = openPdf(bytes);
  assert.match(file.pageText(0), /Two photographs/);
  const xo = file.resolve(file.resolve(file.pages()[0].dict.get('Resources')).get('XObject'));
  const sizes = [...xo.keys()].map((k) => file.resolve(xo.get(k))).filter((x) => x instanceof Stream).map((x) => x.dict.get('Width')).sort((a, b) => a - b);
  assert.deepEqual(sizes, [40, 300]);
});

test('a picture handed back encoded is written as it was handed, and only when it is smaller', () => {
  const original = photos();
  const plan = compressPlan(original, { dpi: 96 });
  const jpeg = Buffer.concat([Buffer.from([0xff, 0xd8, 0xff]), Buffer.alloc(5000, 1)]);
  const { bytes, replaced } = compressPdf(original, { dpi: 96, images: { [plan[0].ref]: { data: jpeg, width: 192, height: 192, filter: 'DCTDecode', comps: 3 } } });
  assert.equal(replaced, 1);
  const file = openPdf(bytes);
  const xo = file.resolve(file.resolve(file.pages()[0].dict.get('Resources')).get('XObject'));
  const pic = [...xo.keys()].map((k) => file.resolve(xo.get(k))).find((x) => x.dict.get('Filter')?.name === 'DCTDecode');
  assert.equal(pic.dict.get('Width'), 192);
  assert.ok(pic.data.equals(jpeg));
});

test('resampling averages each box of pixels', () => {
  const data = Buffer.from([0, 100, 200, 100, 0, 0, 0, 0]);
  assert.deepEqual([...resample({ width: 4, height: 2, comps: 1, data }, 2, 1)], [25, 75]);
});
