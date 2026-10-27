// Pictures through Windows' own codecs (2.21 C5): the camera's JPEG found
// inside a RAW's TIFF directories (a DNG's preview strip, a NEF's SubIFD, a
// Fujifilm RAF's header), a bitmap turned upright by each orientation, and
// on Windows the Imaging Component's decode at the screen's size, with null
// (never a throw) for a file it has no codec for.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import zlib from 'node:zlib';
import { rawPreviews, turnBitmap, decodeThroughWindows, CODEC_EXTS } from '../packages/office-shell/src/electron/windows-codec.js';

/** A little-endian TIFF directory list: each IFD [[tag, type, count, value]]; `value` a number, or 'IFD:n' for the n-th directory's offset. */
function tiff(ifds, { magic = 42, tailBytes = 0 } = {}) {
  const sizes = ifds.map((e) => 2 + e.length * 12 + 4);
  const starts = [];
  let at = 8;
  for (const s of sizes) { starts.push(at); at += s; }
  const buf = Buffer.alloc(at + tailBytes);
  buf.write('II', 0, 'latin1');
  buf.writeUInt16LE(magic, 2);
  buf.writeUInt32LE(starts[0], 4);
  ifds.forEach((entries, k) => {
    const o = starts[k];
    buf.writeUInt16LE(entries.length, o);
    entries.forEach(([tag, type, count, value], i) => {
      const p = o + 2 + i * 12;
      buf.writeUInt16LE(tag, p);
      buf.writeUInt16LE(type, p + 2);
      buf.writeUInt32LE(count, p + 4);
      const v = typeof value === 'string' ? starts[Number(value.slice(4))] : value;
      if (type === 3) buf.writeUInt16LE(v, p + 8);
      else buf.writeUInt32LE(v, p + 8);
    });
    buf.writeUInt32LE(0, o + 2 + entries.length * 12);
  });
  return buf;
}

test('a DNG: its preview strip (JPEG compression) and orientation found in IFD0, the raw data in a SubIFD listed too', () => {
  const file = tiff([
    [[0x0103, 3, 1, 7], [0x0111, 4, 1, 4000], [0x0112, 3, 1, 6], [0x0117, 4, 1, 900], [0x014a, 4, 1, 'IFD:1']],
    [[0x0103, 3, 1, 7], [0x0111, 4, 1, 5000], [0x0117, 4, 1, 3000]],
  ], { tailBytes: 9000 });
  const found = rawPreviews(file, file.length);
  assert.equal(found.orientation, 6);
  assert.deepEqual(found.previews, [{ offset: 4000, length: 900 }, { offset: 5000, length: 3000 }]);
});

test('a NEF-like file: the JPEG by its interchange format in a SubIFD list, an Olympus magic number accepted', () => {
  const file = tiff([
    [[0x0112, 3, 1, 8], [0x014a, 4, 1, 'IFD:1']],
    [[0x0201, 4, 1, 3000], [0x0202, 4, 1, 1200]],
  ], { magic: 0x4f52, tailBytes: 5000 });
  const found = rawPreviews(file, file.length);
  assert.equal(found.orientation, 8);
  assert.deepEqual(found.previews, [{ offset: 3000, length: 1200 }]);
});

test('a Fujifilm RAF: the JPEG\'s place in its header', () => {
  const file = Buffer.alloc(200);
  file.write('FUJIFILMCCD-RAW 0201', 0, 'latin1');
  file.writeUInt32BE(120, 84);
  file.writeUInt32BE(60, 88);
  assert.deepEqual(rawPreviews(file, 200).previews, [{ offset: 120, length: 60 }]);
});

test('previews that point past the file, directories in a loop, and garbage are left out quietly', () => {
  const loop = tiff([[[0x0201, 4, 1, 999999], [0x0202, 4, 1, 10], [0x014a, 4, 1, 'IFD:0']]]);
  assert.deepEqual(rawPreviews(loop, loop.length), { previews: [], orientation: 1 });
  assert.deepEqual(rawPreviews(Buffer.from('not a raw file at all, just words'), 33), { previews: [], orientation: 1 });
  let seed = 3;
  const rand = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
  for (let i = 0; i < 300; i++) {
    const b = Buffer.from(loop);
    for (let k = 0; k < 4; k++) b[Math.floor(rand() * b.length)] = Math.floor(rand() * 256);
    const r = rawPreviews(b, b.length);
    assert.ok(Array.isArray(r.previews));
  }
});

test('a bitmap turned upright by each of the eight orientations', () => {
  // 3 × 2, each pixel's first byte its index: 0 1 2 / 3 4 5.
  const data = Buffer.alloc(3 * 2 * 4);
  for (let i = 0; i < 6; i++) data[i * 4] = i;
  const rows = (t) => Array.from({ length: t.height }, (_, y) => Array.from({ length: t.width }, (_, x) => t.data[(y * t.width + x) * 4]).join('')).join('/');
  assert.equal(rows(turnBitmap(data, 3, 2, 1)), '012/345');
  assert.equal(rows(turnBitmap(data, 3, 2, 2)), '210/543');
  assert.equal(rows(turnBitmap(data, 3, 2, 3)), '543/210');
  assert.equal(rows(turnBitmap(data, 3, 2, 4)), '345/012');
  assert.equal(rows(turnBitmap(data, 3, 2, 5)), '03/14/25');
  assert.equal(rows(turnBitmap(data, 3, 2, 6)), '30/41/52');
  assert.equal(rows(turnBitmap(data, 3, 2, 7)), '52/41/30');
  assert.equal(rows(turnBitmap(data, 3, 2, 8)), '25/14/03');
});

test('the kinds of file it is for: HEIC, HEIF, JPEG XR and the cameras\' RAW', () => {
  for (const ext of ['.heic', '.heif', '.jxr', '.dng', '.cr2', '.nef', '.arw', '.raf', '.orf', '.rw2']) assert.ok(CODEC_EXTS.includes(ext), ext);
  assert.ok(!CODEC_EXTS.includes('.jpg'), 'what Chromium draws itself is not sent to Windows');
});

test('on Windows: a picture decoded by the Imaging Component at 2048 on its long side, and null for a file with no codec', { skip: process.platform !== 'win32' }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-codec-'));
  try {
    // A 3000 × 1000 PNG made here: grey with a red band.
    const W = 3000;
    const H = 1000;
    const raw = Buffer.alloc((W * 3 + 1) * H, 128);
    for (let y = 0; y < H; y++) {
      raw[y * (W * 3 + 1)] = 0;
      if (y > 400 && y < 600) for (let x = 0; x < W; x++) { const i = y * (W * 3 + 1) + 1 + x * 3; raw[i] = 220; raw[i + 1] = 30; raw[i + 2] = 30; }
    }
    const chunk = (type, d) => { const len = Buffer.alloc(4); len.writeUInt32BE(d.length); const td = Buffer.concat([Buffer.from(type), d]); const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(td)); return Buffer.concat([len, td, crc]); };
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(W, 0);
    ihdr.writeUInt32BE(H, 4);
    ihdr[8] = 8;
    ihdr[9] = 2;
    const png = path.join(dir, 'wide.png');
    fs.writeFileSync(png, Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]));
    const out = await decodeThroughWindows(png, 2048);
    assert.ok(out, 'decoded');
    assert.equal(out.readUInt32BE(16), 2048);
    assert.equal(out.readUInt32BE(20), 683);
    const fake = path.join(dir, 'not-a-picture.heic');
    fs.writeFileSync(fake, 'words, not a picture');
    assert.equal(await decodeThroughWindows(fake, 2048), null);
    assert.equal(await decodeThroughWindows(path.join(dir, 'missing.jxr'), 2048), null);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
