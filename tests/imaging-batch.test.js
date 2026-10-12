// Batch from Pictures, and saving a picture: the size each is written at,
// the name a pattern gives it, a free name rather than an overwrite, the
// JPEG metadata kept or stripped, and the quality that brings a file under
// a size.
import test from 'node:test';
import assert from 'node:assert/strict';
import { sizeFor, nameFromPattern, planBatch, keepBoth, freeName, formatOf, stripJpeg, withExifFrom, exifSegment, underSize } from '@rutba/imaging/batch';
import { readExif, withOrientation } from '@rutba/imaging/exif';

test('sizes: as it is, a percent, inside a box never larger, the long side, exact', () => {
  const s = { width: 4000, height: 3000 };
  assert.deepEqual(sizeFor(s, { mode: 'none' }), s);
  assert.deepEqual(sizeFor(s, { mode: 'percent', percent: 25 }), { width: 1000, height: 750 });
  assert.deepEqual(sizeFor(s, { mode: 'fit', width: 1920, height: 1080 }), { width: 1440, height: 1080 });
  assert.deepEqual(sizeFor({ width: 800, height: 600 }, { mode: 'fit', width: 1920, height: 1080 }), { width: 800, height: 600 });
  assert.deepEqual(sizeFor(s, { mode: 'long', size: 1000 }), { width: 1000, height: 750 });
  assert.deepEqual(sizeFor(s, { mode: 'exact', width: 300, height: 300 }), { width: 300, height: 300 });
});

test('a name from a pattern: the name, a number padded, the date, the size; no character a file system refuses', () => {
  const at = { name: 'IMG_0042', n: 7, date: '2026-10-12T09:30:00', width: 1024, height: 768 };
  assert.equal(nameFromPattern('{name}', at), 'IMG_0042');
  assert.equal(nameFromPattern('Holiday {n:3}', at), 'Holiday 007');
  assert.equal(nameFromPattern('{date} {name} {w}x{h}', at), '2026-10-12 IMG_0042 1024x768');
  assert.equal(nameFromPattern('a/b:c*?', at), 'abc');
  assert.equal(nameFromPattern('   ', at), 'IMG_0042');
  // A $ in a name is kept as it is, never read as a replacement pattern.
  assert.equal(nameFromPattern('{name} copy', { name: "Sales $'000 $&" }), "Sales $'000 $& copy");
  assert.equal(formatOf('C:\\x\\a.JPEG'), 'jpeg');
  assert.equal(formatOf('a.heic'), null);
});

test('a free name beside a file already there, as Explorer names a copy', () => {
  const taken = new Set(['C:\\p\\a.jpg', 'C:\\p\\a (2).jpg']);
  assert.equal(freeName('C:\\p\\a.jpg', (p) => taken.has(p)), 'C:\\p\\a (3).jpg');
  assert.equal(freeName('C:\\p\\b.jpg', (p) => taken.has(p)), 'C:\\p\\b.jpg');
  assert.equal(freeName('/home/x/a.png', (p) => p === '/home/x/a.png'), '/home/x/a (2).png');
});

test('a batch plan: formats, names and folders; clashes found before anything is written; Keep both moves them aside', () => {
  const files = [
    { path: 'C:\\p\\one.png', size: { width: 400, height: 200 } },
    { path: 'C:\\p\\two.jpg', size: { width: 200, height: 400 } },
    { path: 'C:\\p\\three.jpg', size: { width: 100, height: 100 } },
  ];
  const existing = new Set(['C:\\p\\two.jpg', 'C:\\p\\three.jpg', 'C:\\p\\one.png', 'C:\\p\\one-small.jpg']);
  const exists = (p) => existing.has(p);
  const plan = planBatch(files, { format: 'jpeg', pattern: '{name}-small', resize: { mode: 'fit', width: 100, height: 100 }, exists });
  assert.deepEqual(plan.items.map((i) => i.target), ['C:\\p\\one-small.jpg', 'C:\\p\\two-small.jpg', 'C:\\p\\three-small.jpg']);
  assert.deepEqual(plan.items.map((i) => i.size), [{ width: 100, height: 50 }, { width: 50, height: 100 }, { width: 100, height: 100 }]);
  assert.deepEqual(plan.clashes.map((i) => i.target), ['C:\\p\\one-small.jpg']);
  const kept = keepBoth(plan, exists);
  assert.equal(kept.items[0].target, 'C:\\p\\one-small (2).jpg');
  assert.equal(kept.clashes.length, 0);

  // The same name as the original, in the same format, beside it: that is the original, and a clash.
  const same = planBatch(files, { format: 'same', pattern: '{name}', exists });
  assert.ok(same.items.every((i) => i.clash && i.original));
  assert.deepEqual(same.items.map((i) => i.format), ['png', 'jpeg', 'jpeg']);

  // Two pictures that would get one name are told apart within the batch.
  const one = planBatch(files, { format: 'png', pattern: 'picture', folder: 'D:\\out' });
  assert.deepEqual(one.items.map((i) => i.target), ['D:\\out\\picture.png', 'D:\\out\\picture (2).png', 'D:\\out\\picture (3).png']);
});

// A small JPEG: SOI, JFIF, an EXIF with an orientation, a comment, then a scan.
const JFIF = [0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00];
const COMMENT = [0xff, 0xfe, 0x00, 0x07, 0x68, 0x65, 0x6c, 0x6c, 0x6f];
const SCAN = [0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3f, 0x00, 0x12, 0x34, 0x56, 0xff, 0xd9];
const plain = new Uint8Array([0xff, 0xd8, ...JFIF, ...COMMENT, ...SCAN]);
const camera = withOrientation(plain, 6);

test('strip: a JPEG keeps its JFIF header and its picture, and loses its EXIF and comments', () => {
  assert.equal(readExif(camera).orientation, 6);
  const stripped = stripJpeg(camera);
  assert.equal(exifSegment(stripped), null);
  assert.equal(readExif(stripped).orientation, undefined);
  assert.deepEqual([...stripped.subarray(2, 2 + JFIF.length)], JFIF);
  assert.deepEqual([...stripped.subarray(stripped.length - SCAN.length)], SCAN);
  assert.equal(stripped.length, plain.length - COMMENT.length);
});

test('keep: a freshly encoded JPEG is given the original EXIF, its orientation set upright', () => {
  const encoded = new Uint8Array([0xff, 0xd8, ...JFIF, ...SCAN]);
  const kept = withExifFrom(encoded, camera);
  assert.ok(exifSegment(kept));
  assert.equal(readExif(kept).orientation, 1);
  assert.deepEqual([...kept.subarray(kept.length - SCAN.length)], SCAN);
  assert.deepEqual([...withExifFrom(encoded, plain)], [...encoded], 'nothing to keep, nothing added');
});

test('a size cap: the best quality under the size, or the lowest and a plain "does not fit"', async () => {
  // A stand-in encoder: a file of 1,000 bytes a tenth of quality.
  const encode = async (q) => new Uint8Array(Math.round(q * 10000));
  const fits = await underSize(encode, 5000);
  assert.ok(fits.fits && fits.bytes.length <= 5000 && fits.quality > 0.45, JSON.stringify({ q: fits.quality, n: fits.bytes.length }));
  const none = await underSize(encode, 100);
  assert.equal(none.fits, false);
  assert.equal(none.quality, 0.05);
  const easy = await underSize(encode, 1e6);
  assert.equal(easy.quality, 0.92);
});
