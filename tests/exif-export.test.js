// The camera's details carried into an exported picture (2.21 D3), and
// Remove location (D2): JPEG, PNG (eXIf), WebP (EXIF chunk) and TIFF exports
// keep the original's EXIF from any file that has one, made true of the new
// picture (upright, its own size, the original's thumbnail left behind);
// Remove location takes the GPS directory and its pointer out and keeps the
// rest. Every output is read back, and a PNG's checksums and a WebP's sizes
// are what a strict reader wants.
import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { readMetadata, readExif, exifForExport, exifBlockOf, keepExif, withoutLocation, jpegWithExif, pngWithExif, webpWithExif, tiffWithExif } from '@rutba/imaging/exif';
import { cameraBlock, jpegWith, pngWith, webpWith, heifWith, simpleWebp, plainTiff, tiffBlock, XMP, THUMB } from './exif-fixtures.js';

const block = cameraBlock();
const ORIGINALS = {
  jpeg: jpegWith({ block, xmp: XMP }),
  png: pngWith({ block }),
  webp: webpWith({ block }),
  heic: heifWith({ block }),
  avif: heifWith({ block, avif: true }),
  tiff: block,
};

const has = (bytes, needle) => {
  outer: for (let i = 0; i + needle.length <= bytes.length; i++) {
    for (let k = 0; k < needle.length; k++) if (bytes[i + k] !== needle[k]) continue outer;
    return true;
  }
  return false;
};
// The latitude's seconds as the block stores them (2650/100, little-endian): gone means gone.
const LAT_SECONDS = [0x5a, 0x0a, 0, 0, 100, 0, 0, 0];

test('the block an export carries: upright, its new size, no thumbnail of the picture before the edit', () => {
  const out = exifForExport(ORIGINALS.jpeg, { width: 800, height: 600 });
  const exif = readMetadata(out).exif;
  assert.equal(exif.make, 'Rutba');
  assert.equal(exif.orientation, 1);
  assert.equal(exif.pixelWidth, 800);
  assert.equal(exif.pixelHeight, 600);
  assert.equal(exif.dateTimeOriginal, '2026:10:27 09:15:00');
  assert.ok(exif.latitude > 51, 'the place kept unless asked');
  assert.ok(has(block, THUMB) && !has(out, THUMB), 'the old thumbnail zeroed');
  assert.equal(out.length, block.length, 'changed in place: every offset still true');
});

test('Remove location: no GPS directory, no pointer to one, no coordinates left in the bytes; the camera, date and rating kept', () => {
  const out = exifForExport(ORIGINALS.jpeg, { location: false });
  const meta = readMetadata(out);
  assert.equal(meta.problem, null);
  assert.equal(meta.exif.latitude, undefined);
  assert.equal(meta.exif.lat, undefined);
  assert.equal(meta.exif.latRef, undefined);
  assert.equal(meta.exif.make, 'Rutba');
  assert.equal(meta.exif.model, 'Test Phone 7');
  assert.equal(meta.exif.lens, 'Wide 26mm');
  assert.equal(meta.exif.rating, 4);
  assert.deepEqual(meta.marks.tags, ['beach', 'family']);
  assert.ok(has(block, LAT_SECONDS) && !has(out, LAT_SECONDS), 'the coordinates are zeroed, not just unlinked');
  // IFD0 lists no GPS pointer.
  const dv = new DataView(out.buffer, out.byteOffset);
  const ifd = dv.getUint32(4, true);
  const tags = Array.from({ length: dv.getUint16(ifd, true) }, (_, i) => dv.getUint16(ifd + 2 + i * 12, true));
  assert.ok(!tags.includes(0x8825), tags.map((t) => t.toString(16)).join(','));
  assert.ok(tags.includes(0x8769), 'the Exif directory kept');
});

test('withoutLocation on a block alone keeps everything else exactly, the orientation and thumbnail too', () => {
  const out = withoutLocation(block);
  const exif = readMetadata(out).exif;
  assert.equal(exif.orientation, 6);
  assert.equal(exif.latitude, undefined);
  assert.ok(has(out, THUMB));
  assert.equal(withoutLocation(new Uint8Array([1, 2, 3])), null);
});

for (const [name, original] of Object.entries(ORIGINALS)) {
  test(`a ${name}'s camera goes into JPEG, PNG, WebP and TIFF exports, upright, and without its place when asked`, () => {
    assert.ok(exifBlockOf(original), 'the original has a block');
    const targets = {
      jpeg: jpegWith(),
      png: pngWith({ width: 3, height: 2 }),
      webp: simpleWebp(3, 2),
      tiff: plainTiff(3, 2),
    };
    for (const [format, encoded] of Object.entries(targets)) {
      for (const location of [true, false]) {
        const out = keepExif(encoded, format, original, { width: 3, height: 2, location });
        assert.notEqual(out.length, encoded.length, `${format}: something added`);
        const meta = readMetadata(out);
        assert.equal(meta.problem, null, `${format}: ${meta.problem}`);
        assert.equal(meta.container, format === 'jpeg' ? 'jpeg' : format);
        assert.equal(meta.exif.model, 'Test Phone 7', `${format} from ${name}`);
        assert.equal(meta.exif.orientation, 1, `${format}: upright`);
        assert.equal(meta.exif.dateTimeOriginal, '2026:10:27 09:15:00');
        assert.equal(meta.exif.latitude != null, location, `${format}: the place ${location ? 'kept' : 'gone'}`);
        if (!location) assert.ok(!has(out, LAT_SECONDS), `${format}: no coordinates in the bytes`);
      }
    }
  });
}

test('a PNG export is valid: eXIf straight after IHDR, every checksum right, the pixels still inflate', () => {
  const out = pngWithExif(pngWith({ width: 3, height: 2 }), exifForExport(ORIGINALS.jpeg));
  const dv = new DataView(out.buffer, out.byteOffset);
  const types = [];
  for (let at = 8; at < out.length;) {
    const length = dv.getUint32(at);
    const type = String.fromCharCode(...out.subarray(at + 4, at + 8));
    types.push(type);
    assert.equal(dv.getUint32(at + 8 + length), zlib.crc32(out.subarray(at + 4, at + 8 + length)), `${type}: checksum`);
    if (type === 'IDAT') assert.equal(zlib.inflateSync(out.subarray(at + 8, at + 8 + length)).length, 4 * 2);
    at += 12 + length;
  }
  assert.deepEqual(types, ['IHDR', 'eXIf', 'IDAT', 'IEND']);
  // Given again, the old eXIf goes: one block, never two.
  const twice = pngWithExif(out, exifForExport(ORIGINALS.png));
  assert.equal(twice.length, out.length);
});

test('a simple WebP becomes an extended one: VP8X with the canvas size and the EXIF flag, the RIFF size true', () => {
  for (const alpha of [false, true]) {
    const out = webpWithExif(simpleWebp(300, 200, alpha), exifForExport(ORIGINALS.jpeg));
    const dv = new DataView(out.buffer, out.byteOffset);
    assert.equal(dv.getUint32(4, true), out.length - 8, 'the RIFF size');
    assert.equal(String.fromCharCode(...out.subarray(12, 16)), 'VP8X');
    assert.equal(out[20] & 0x08, 0x08, 'EXIF flag');
    assert.equal(Boolean(out[20] & 0x10), alpha, 'alpha flag as the picture has');
    assert.equal(out[24] | (out[25] << 8) | (out[26] << 16), 299);
    assert.equal(out[27] | (out[28] << 8) | (out[29] << 16), 199);
    assert.equal(String.fromCharCode(...out.subarray(30, 34)), 'VP8L', 'the picture next, as it was');
    assert.equal(readExif(out).make, 'Rutba');
  }
  // An extended WebP keeps its other chunks and gets its EXIF replaced, the XMP last.
  const ext = webpWithExif(webpWith({ block: tiffBlock({ ifd0: [[0x010f, 2, 'Old']] }), xmp: XMP }), exifForExport(ORIGINALS.jpeg));
  assert.equal(readExif(ext).make, 'Rutba');
  assert.equal(readMetadata(ext).xmp.rating, 3);
});

test('a TIFF export keeps the writer\'s own fields and pixels, and gains the camera', () => {
  const tiff = plainTiff(4, 3, { value: 77 });
  const out = tiffWithExif(tiff, exifForExport(ORIGINALS.jpeg, { location: false }));
  const dv = new DataView(out.buffer, out.byteOffset);
  const ifd = dv.getUint32(4, true);
  const entries = new Map(Array.from({ length: dv.getUint16(ifd, true) }, (_, i) => {
    const p = ifd + 2 + i * 12;
    return [dv.getUint16(p, true), { type: dv.getUint16(p + 2, true), value: dv.getUint32(p + 8, true) }];
  }));
  assert.equal(entries.get(0x0100).value & 0xffff, 4, 'width as written');
  const strip = entries.get(0x0111).value;
  assert.ok(out.subarray(strip, strip + 12).every((v) => v === 77), 'the pixels where the strip says');
  const exif = readExif(out);
  assert.equal(exif.software, 'Rutba Image', 'the writer\'s own Software wins');
  assert.equal(exif.make, 'Rutba');
  assert.equal(exif.dateTimeOriginal, '2026:10:27 09:15:00');
  assert.equal(exif.latitude, undefined);
  const keys = [...entries.keys()];
  assert.deepEqual(keys, [...keys].sort((a, b) => a - b), 'entries in tag order, as TIFF requires');
  // Big-endian writer, little-endian camera: each value turned.
  const be = tiffWithExif(plainTiff(2, 2, { little: false }), exifForExport(ORIGINALS.jpeg));
  const beExif = readExif(be);
  assert.equal(beExif.model, 'Test Phone 7');
  assert.equal(beExif.iso, 100);
  assert.ok(Math.abs(beExif.fNumber - 1.8) < 1e-9);
  assert.ok(beExif.latitude > 51);
});

test('nothing to keep, nothing added; a damaged original adds nothing rather than a broken block', () => {
  const encoded = jpegWith();
  assert.deepEqual([...keepExif(encoded, 'jpeg', pngWith())], [...encoded]);
  const broken = new Uint8Array(block);
  new DataView(broken.buffer).setUint32(4, 0xfffff0, true);
  assert.deepEqual([...keepExif(encoded, 'jpeg', jpegWith({ block: broken }))], [...encoded]);
  assert.deepEqual([...keepExif(encoded, 'bmp', ORIGINALS.jpeg)], [...encoded], 'a format with no place for it');
});

test('a block too large for a JPEG segment is carried as its descriptive fields alone', () => {
  const huge = tiffBlock({ ifd0: [[0x010f, 2, 'Rutba'], [0x0110, 2, 'Big Notes']], exif: [[0x927c, 7, new Array(70000).fill(7)], [0x8827, 3, 200]] });
  const original = pngWith({ block: huge });
  const out = keepExif(jpegWith(), 'jpeg', original);
  const exif = readExif(out);
  assert.equal(exif.model, 'Big Notes');
  assert.equal(exif.iso, 200);
  assert.ok(out.length < 1000, 'the maker\'s notes left out');
});

test('a rebuilt block from a TIFF original carries no strips of the original', () => {
  const tiffOriginal = tiffBlock({ ifd0: [[0x0100, 3, 4000], [0x0111, 4, 8], [0x010f, 2, 'Scanner Co'], [0x0110, 2, 'S1']] });
  const out = exifForExport(tiffOriginal);
  const dv = new DataView(out.buffer, out.byteOffset);
  const ifd = dv.getUint32(4, true);
  const tags = Array.from({ length: dv.getUint16(ifd, true) }, (_, i) => dv.getUint16(ifd + 2 + i * 12, true));
  assert.deepEqual(tags, [0x010f, 0x0110]);
});
