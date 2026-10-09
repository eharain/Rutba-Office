// Rotate the picture, for a JPEG: its EXIF orientation set, its pixels not
// touched — the tag rewritten where the file has one, a small EXIF block
// added after the JFIF header where it has none.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readExif, withOrientation, TURN_CLOCKWISE } from '@rutba/imaging/exif';

const JFIF = [0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00];
const SCAN = [0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3f, 0x00, 0x12, 0x34, 0x56, 0xff, 0xd9];
const jpeg = (...segments) => new Uint8Array([0xff, 0xd8, ...segments.flat(), ...SCAN]);

/** An EXIF APP1 whose IFD0 holds the entries given, little-endian. */
function exifSegment(entries) {
  const tiff = [0x49, 0x49, 42, 0, 8, 0, 0, 0, entries.length, 0];
  for (const [tag, type, value] of entries) tiff.push(tag & 0xff, tag >> 8, type, 0, 1, 0, 0, 0, value & 0xff, value >> 8, 0, 0);
  tiff.push(0, 0, 0, 0);
  const payload = [0x45, 0x78, 0x69, 0x66, 0, 0, ...tiff];
  return [0xff, 0xe1, (payload.length + 2) >> 8, (payload.length + 2) & 0xff, ...payload];
}

test('a JPEG with no EXIF gets one holding only its orientation, after the JFIF header, the rest as it was', () => {
  const plain = jpeg(JFIF);
  assert.equal(readExif(plain).orientation, undefined);
  const turned = withOrientation(plain, 6);
  assert.equal(readExif(turned).orientation, 6);
  assert.deepEqual([...turned.subarray(0, 2 + JFIF.length)], [0xff, 0xd8, ...JFIF], 'the JFIF header first, as it was');
  assert.equal(turned[2 + JFIF.length + 1], 0xe1, 'the EXIF block straight after it');
  assert.deepEqual([...turned.subarray(turned.length - SCAN.length)], SCAN, 'the picture itself untouched');
});

test('a JPEG whose EXIF has an orientation has that one value rewritten, nothing else', () => {
  const original = jpeg(JFIF, exifSegment([[0x010f, 3, 7], [0x0112, 3, 6]]));
  const turned = withOrientation(original, 3);
  assert.equal(turned.length, original.length);
  assert.equal(readExif(turned).orientation, 3);
  assert.equal([...turned].filter((b, i) => b !== original[i]).length, 1, 'one byte changed');
});

test('EXIF without an orientation tag, or a file that is not a JPEG, is not changed here', () => {
  assert.equal(withOrientation(jpeg(JFIF, exifSegment([[0x010f, 3, 7]])), 6), null);
  assert.equal(withOrientation(new Uint8Array([0x89, 0x50, 0x4e, 0x47]), 6), null);
  assert.throws(() => withOrientation(jpeg(JFIF), 9), /1 to 8/);
});

test('four quarter turns come back round, and a mirrored picture stays mirrored', () => {
  for (const start of [1, 2, 3, 4, 5, 6, 7, 8]) {
    let o = start;
    for (let i = 0; i < 4; i++) o = TURN_CLOCKWISE[o];
    assert.equal(o, start);
  }
  assert.deepEqual([1, 6, 3, 8].map((o) => TURN_CLOCKWISE[o]), [6, 3, 8, 1]);
  assert.ok([2, 4, 5, 7].every((o) => [2, 4, 5, 7].includes(TURN_CLOCKWISE[o])));
});
