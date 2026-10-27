// EXIF read from every file that carries it (2.21 D1): a JPEG's APP1, a
// TIFF itself, a PNG's eXIf chunk (before or after the pixels), a WebP's
// EXIF chunk, and the Exif item of a HEIC (in mdat) and an AVIF (in the meta
// box's idat). Each walk is bounded, and a damaged file is a sentence, never
// a throw or a hang. And the reader that fetches only the parts it needs.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readExif, readMetadata, readMetadataFrom, containerOf, describeExif } from '@rutba/imaging/exif';
import { cameraBlock, jpegWith, pngWith, webpWith, heifWith, plainTiff, concat, tiffBlock, XMP } from './exif-fixtures.js';

const camera = (exif) => ({ make: exif.make, model: exif.model, orientation: exif.orientation, taken: exif.dateTimeOriginal, iso: exif.iso, lens: exif.lens, lat: exif.latitude != null ? Math.round(exif.latitude * 1000) / 1000 : null, lon: exif.longitude != null ? Math.round(exif.longitude * 1000) / 1000 : null });
const EXPECTED = { make: 'Rutba', model: 'Test Phone 7', orientation: 6, taken: '2026:10:27 09:15:00', iso: 100, lens: 'Wide 26mm', lat: 51.507, lon: -0.128 };

const block = cameraBlock();
const FILES = {
  jpeg: jpegWith({ block, xmp: XMP }),
  png: pngWith({ block, xmp: XMP }),
  'png, after the pixels': pngWith({ block, xmp: XMP, after: true }),
  webp: webpWith({ block, xmp: XMP }),
  'webp, with the JPEG lead': webpWith({ block, lead: true }),
  heic: heifWith({ block, xmp: XMP }),
  avif: heifWith({ block, xmp: XMP, avif: true }),
  tiff: block,
  'big-endian tiff': cameraBlock({ little: false }),
};

for (const [name, file] of Object.entries(FILES)) {
  test(`the camera, the date and the place are read from a ${name}`, () => {
    const meta = readMetadata(file);
    assert.equal(meta.problem, null, meta.problem);
    assert.deepEqual(camera(meta.exif), EXPECTED);
    assert.equal(readExif(file).model, 'Test Phone 7', 'readExif answers the same fields');
  });
}

test('each container is told by its first bytes, and a video in the same boxes is not taken for a picture', () => {
  assert.equal(containerOf(FILES.jpeg), 'jpeg');
  assert.equal(containerOf(FILES.png), 'png');
  assert.equal(containerOf(FILES.webp), 'webp');
  assert.equal(containerOf(FILES.heic), 'heif');
  assert.equal(containerOf(FILES.avif), 'heif');
  assert.equal(containerOf(FILES.tiff), 'tiff');
  const mp4 = new Uint8Array([0, 0, 0, 24, ...[...'ftypisom'].map((c) => c.charCodeAt(0)), 0, 0, 2, 0, ...[...'isommp41'].map((c) => c.charCodeAt(0))]);
  assert.equal(containerOf(mp4), null);
});

test('the Details rows read the same from a HEIC as from a JPEG', () => {
  const rows = (f) => describeExif(readExif(f)).map((r) => `${r.label}: ${r.value}`);
  assert.deepEqual(rows(FILES.heic), rows(FILES.jpeg));
  assert.ok(rows(FILES.heic).includes('Camera: Rutba Test Phone 7'));
});

test('a file with no EXIF answers nothing, and no problem', () => {
  for (const f of [jpegWith(), pngWith(), webpWith(), heifWith(), plainTiff()]) {
    const meta = readMetadata(f);
    assert.equal(meta.problem, null);
    assert.equal(meta.exif.make, undefined);
  }
});

test('a damaged EXIF block is a sentence, and what came before the damage still reads', () => {
  // The first directory pointed past the block's end.
  const broken = new Uint8Array(block);
  new DataView(broken.buffer).setUint32(4, 0xfffff0, true);
  const meta = readMetadata(jpegWith({ block: broken }));
  assert.match(meta.problem, /^The EXIF block of this file is damaged: .+\.$/);
  assert.deepEqual(meta.exif, {});
  // Not a TIFF header at all.
  const notTiff = readMetadata(pngWith({ block: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]) }));
  assert.match(notTiff.problem, /does not begin with a TIFF header/);
  // An inner directory broken: IFD0's camera still stands.
  const inner = tiffBlock({ ifd0: [[0x010f, 2, 'Rutba'], [0x8769, 4, 0x7ffffff0]] });
  const partly = readMetadata(jpegWith({ block: inner }));
  assert.equal(partly.exif.make, 'Rutba');
  assert.match(partly.problem, /damaged/);
  assert.ok(describeExif(readExif(jpegWith({ block: broken }))).some((r) => r.label === 'Metadata' && /damaged/.test(r.value)), 'the panel says so');
});

test('a directory that claims more entries than any camera writes is refused, not walked', () => {
  const huge = tiffBlock({ ifd0: [[0x010f, 2, 'Rutba']] });
  new DataView(huge.buffer).setUint16(8, 60000, true);
  assert.match(readMetadata(huge).problem, /claims 60000 entries/);
});

test('a directory that points back at itself is read once', () => {
  const loop = tiffBlock({ ifd0: [[0x010f, 2, 'Rutba'], [0x8769, 4, 8]] });
  const started = Date.now();
  assert.equal(readMetadata(loop).exif.make, 'Rutba');
  assert.ok(Date.now() - started < 200);
});

test('truncated and mutated files of every container never throw and never hang', () => {
  let seed = 7;
  const rand = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
  const started = Date.now();
  let refused = 0;
  for (const file of Object.values(FILES)) {
    for (let cut = 1; cut < file.length; cut += Math.max(1, Math.floor(file.length / 40))) {
      const meta = readMetadata(file.subarray(0, cut));
      assert.ok(meta && typeof meta === 'object');
      if (meta.problem) refused++;
    }
    for (let i = 0; i < 150; i++) {
      const bad = new Uint8Array(file);
      for (let k = 0; k < 1 + Math.floor(rand() * 6); k++) bad[Math.floor(rand() * bad.length)] = Math.floor(rand() * 256);
      const meta = readMetadata(bad);
      assert.ok(meta && typeof meta === 'object');
      if (meta.problem) assert.match(meta.problem, /^The (EXIF|XMP|file).*\.$/);
    }
  }
  assert.ok(refused > 0, 'some cuts are refused in a sentence');
  assert.ok(Date.now() - started < 5000, `${Date.now() - started} ms`);
});

test('a WebP keeps its EXIF after the picture: the reader fetches its start and the chunk, never the picture', async () => {
  const file = webpWith({ block, xmp: XMP, filler: 3 * 1024 * 1024 });
  const reads = [];
  const read = async (offset, length) => { reads.push([offset, length]); return file.subarray(offset, offset + length); };
  const meta = await readMetadataFrom(read, file.length, { first: 64 * 1024, chunk: 16 * 1024 });
  assert.deepEqual(camera(meta.exif), EXPECTED);
  assert.equal(meta.marks.rating, 3);
  const total = reads.reduce((n, [, l]) => n + l, 0);
  assert.ok(total < 200 * 1024, `${total} bytes read of ${file.length}`);
  assert.ok(reads.length <= 4, `${reads.length} reads`);
  // From the start alone, the same file says nothing wrong: what is past the start is not read.
  const head = readMetadata(file.subarray(0, 64 * 1024), { size: file.length });
  assert.equal(head.problem, null);
  assert.equal(head.partial, true);
});

test('a TIFF whose directory is at its end, and a HEIC whose Exif item follows the picture, are read the same way', async () => {
  const pixels = new Uint8Array(1024 * 1024).fill(9);
  const own = cameraBlock();
  const tiff = tiffBlock({ ifd0: [[0x010f, 2, 'Rutba'], [0x0110, 2, 'Far Directory']] });
  const far = concat([tiff.subarray(0, 8), new Uint8Array(pixels.length), tiff.subarray(8)]);
  const dv = new DataView(far.buffer);
  dv.setUint32(4, 8 + pixels.length, true);
  // The directory's string offsets moved too.
  const at = 8 + pixels.length;
  for (let i = 0; i < 2; i++) {
    const p = at + 2 + i * 12;
    dv.setUint32(p + 8, dv.getUint32(p + 8, true) + pixels.length, true);
  }
  const reads = [];
  const meta = await readMetadataFrom(async (o, l) => { reads.push(l); return far.subarray(o, o + l); }, far.length, { first: 4096, chunk: 4096 });
  assert.equal(meta.exif.model, 'Far Directory');
  assert.ok(reads.reduce((a, b) => a + b, 0) < 64 * 1024);
  const heic = heifWith({ block: own });
  const fromParts = await readMetadataFrom(async (o, l) => heic.subarray(o, o + l), heic.length, { first: 200, chunk: 64 });
  assert.deepEqual(camera(fromParts.exif), EXPECTED);
});

test('a file that ends before its EXIF block does is told so', async () => {
  const file = webpWith({ block });
  const meta = await readMetadataFrom(async (o, l) => file.subarray(o, Math.min(o + l, file.length - 40)), file.length, { first: 32, chunk: 4096 });
  assert.match(meta.problem || "", /ended before|cut short|damaged/);
});
