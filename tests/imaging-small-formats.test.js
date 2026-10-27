// The small formats the suite reads with no library: QOI, Netpbm (PBM, PGM,
// PPM, PAM), Targa, PCX and Radiance HDR (tone-mapped). Each is read from
// files laid out byte by byte here, as its specification has them, and from
// the writers the fuzzer's good files come from; a damaged one is a refusal
// in a sentence.
import test from 'node:test';
import assert from 'node:assert/strict';
import { decodePicture, pictureInfo, pictureFormat, needsDecoding } from '@rutba/imaging/decode';
import { encodeQoi } from '../packages/imaging/src/codecs/qoi.js';
import { encodePcx } from '../packages/imaging/src/codecs/pcx.js';
import { encodeHdr, decodeHdr, toneMap } from '../packages/imaging/src/codecs/hdr.js';
import { pnmOf, tgaOf, lightOf, sample, pictureSources } from '../tools/picture-sources.js';
import { judge } from '../tools/fuzz-open.js';

const bytes = (...parts) => Uint8Array.from(parts.flatMap((p) => (typeof p === 'string' ? [...p].map((c) => c.charCodeAt(0)) : Array.isArray(p) ? p : [p])));
const px = (pic, x, y) => [...pic.rgba.subarray((y * pic.width + x) * 4, (y * pic.width + x) * 4 + 4)];
const same = (a, b) => Buffer.compare(Buffer.from(a), Buffer.from(b)) === 0;

test('QOI: every operation read as the specification has it, and what the writer packs read back exactly', () => {
  // 2 × 2: RGB, a run of one, a small difference, then the first colour by its index (hash 21).
  const qoi = bytes('qoif', [0, 0, 0, 2, 0, 0, 0, 2, 3, 0], [0xfe, 16, 32, 48], 0xc0, 0x40 | (3 << 4) | (2 << 2) | 1, 0x15, [0, 0, 0, 0, 0, 0, 0, 1]);
  const pic = decodePicture(qoi, 'a.qoi');
  assert.equal(pic.format, 'qoi');
  assert.deepEqual([px(pic, 0, 0), px(pic, 1, 0), px(pic, 0, 1), px(pic, 1, 1)], [[16, 32, 48, 255], [16, 32, 48, 255], [17, 32, 47, 255], [16, 32, 48, 255]]);
  // The luma operation and RGBA.
  const luma = bytes('qoif', [0, 0, 0, 3, 0, 0, 0, 1, 4, 1], [0xff, 100, 100, 100, 128], 0x80 | (32 + 10), (8 + 3) << 4 | (8 - 2), [0xfe, 1, 2, 3], [0, 0, 0, 0, 0, 0, 0, 1]);
  const l = decodePicture(luma, 'l.qoi');
  assert.deepEqual([px(l, 0, 0), px(l, 1, 0), px(l, 2, 0)], [[100, 100, 100, 128], [113, 110, 108, 128], [1, 2, 3, 128]]);
  for (const [w, h] of [[1, 1], [37, 23], [64, 3]]) {
    const src = sample(w, h);
    for (let i = 0; i < src.length; i += 7) src[i] = (i * 13) & 255; // noise, so every operation is used
    assert.ok(same(decodePicture(encodeQoi(src, w, h), 'x.qoi').rgba, src), `${w} × ${h}`);
  }
  assert.throws(() => decodePicture(bytes('qoif', [0, 0, 0, 2, 0, 0, 0, 2, 5, 0], [0, 0, 0, 0, 0, 0, 0, 1]), 'bad.qoi'), /has 5 channels/);
});

test('Netpbm: PBM, PGM and PPM as text and as bytes, sixteen-bit samples, and PAM with alpha', () => {
  // P1: digits need not be apart; 1 is black. A comment between the numbers.
  const p1 = decodePicture(bytes('P1\n# a comment\n3 2\n010\n1 0 1\n'), 'a.pbm');
  assert.deepEqual([0, 1, 2].map((x) => px(p1, x, 0)[0]), [255, 0, 255]);
  assert.deepEqual([0, 1, 2].map((x) => px(p1, x, 1)[0]), [0, 255, 0]);
  // P2 with a maxval of 15.
  const p2 = decodePicture(bytes('P2 2 1 15 0 15'), 'a.pgm');
  assert.deepEqual([px(p2, 0, 0)[0], px(p2, 1, 0)[0]], [0, 255]);
  // P3, colour as text.
  const p3 = decodePicture(bytes('P3\n2 1\n255\n255 0 0  0 0 255\n'), 'a.ppm');
  assert.deepEqual([px(p3, 0, 0), px(p3, 1, 0)], [[255, 0, 0, 255], [0, 0, 255, 255]]);
  // P4: bits, rows padded to a byte, 1 black.
  const p4 = decodePicture(bytes('P4\n10 2\n', [0b10100000, 0b01000000, 0xff, 0xc0]), 'a.pbm');
  assert.deepEqual([0, 1, 2, 8, 9].map((x) => px(p4, x, 0)[0]), [0, 255, 0, 255, 0]);
  assert.deepEqual([0, 9].map((x) => px(p4, x, 1)[0]), [0, 0]);
  // P5 sixteen-bit, high byte first.
  const p5 = decodePicture(bytes('P5 2 1 65535\n', [0xff, 0xff, 0x80, 0x00]), 'a.pgm');
  assert.deepEqual([px(p5, 0, 0)[0], px(p5, 1, 0)[0]], [255, 128]);
  // P6 and PAM, from the writer.
  const src = sample(9, 4);
  const p6 = decodePicture(pnmOf(src, 9, 4), 'a.ppm');
  for (let i = 0; i < 36; i++) assert.deepEqual(px(p6, i % 9, (i / 9) | 0).slice(0, 3), [...src.subarray(i * 4, i * 4 + 3)]);
  assert.ok(same(decodePicture(pnmOf(src, 9, 4, { pam: true }), 'a.pam').rgba, src));
  const ga = decodePicture(bytes('P7\nWIDTH 1\nHEIGHT 1\nDEPTH 2\nMAXVAL 255\nTUPLTYPE GRAYSCALE_ALPHA\nENDHDR\n', [200, 100]), 'g.pam');
  assert.deepEqual(px(ga, 0, 0), [200, 200, 200, 100]);
  assert.deepEqual(pictureInfo(pnmOf(src, 9, 4), 'a.ppm'), { pages: 1, format: 'ppm', width: 9, height: 4 });
  assert.throws(() => decodePicture(bytes('P1 2 1 0 7'), 'a.pbm'), /other than 0 and 1/);
});

test('Targa: colour, grey and palette pictures, plain and packed, any row order, and alpha that is really there', () => {
  const head = (type, w, h, bits, desc, map = [0, 0, 0, 0, 0]) => [0, map[0] ? 1 : 0, type, map[1], map[2], map[3], map[4], map[0], 0, 0, 0, 0, w, 0, h, 0, bits, desc];
  // 24-bit, bottom row first: the first pixel in the file is the bottom left.
  const t24 = decodePicture(bytes(head(2, 2, 2, 24, 0), [255, 0, 0], [0, 255, 0], [0, 0, 255], [9, 9, 9]), 'a.tga');
  assert.deepEqual([px(t24, 0, 1), px(t24, 1, 1), px(t24, 0, 0), px(t24, 1, 0)], [[0, 0, 255, 255], [0, 255, 0, 255], [255, 0, 0, 255], [9, 9, 9, 255]]);
  // The same, right to left and top down.
  const flipped = decodePicture(bytes(head(2, 2, 1, 24, 0x30), [255, 0, 0], [0, 255, 0]), 'a.tga');
  assert.deepEqual([px(flipped, 0, 0), px(flipped, 1, 0)], [[0, 255, 0, 255], [0, 0, 255, 255]]);
  // Grey, and a palette of three 24-bit colours starting at entry 0.
  const grey = decodePicture(bytes(head(3, 2, 1, 8, 0x20), 10, 250), 'g.tga');
  assert.deepEqual([px(grey, 0, 0), px(grey, 1, 0)], [[10, 10, 10, 255], [250, 250, 250, 255]]);
  const pal = decodePicture(bytes(head(1, 3, 1, 8, 0x20, [24, 0, 0, 3, 0]), [0, 0, 255], [0, 255, 0], [255, 0, 0], 2, 0, 1), 'p.tga');
  assert.deepEqual([px(pal, 0, 0), px(pal, 1, 0), px(pal, 2, 0)], [[0, 0, 255, 255], [255, 0, 0, 255], [0, 255, 0, 255]]);
  // 16-bit 1-5-5-5 with its alpha bit said to be used.
  const s16 = decodePicture(bytes(head(2, 2, 1, 16, 0x21), [0x1f, 0x80], [0xe0, 0x03]), 's.tga');
  assert.deepEqual([px(s16, 0, 0), px(s16, 1, 0)], [[0, 0, 255, 255], [0, 255, 0, 0]]);
  // 32-bit with every alpha zero: the byte was unused, the picture opaque.
  const unused = decodePicture(bytes(head(2, 1, 1, 32, 0x28), [1, 2, 3, 0]), 'u.tga');
  assert.deepEqual(px(unused, 0, 0), [3, 2, 1, 255]);
  // Packed, from the writer, with real alpha.
  const src = sample(23, 7);
  assert.ok(same(decodePicture(tgaOf(src, 23, 7), 'r.tga').rgba, src));
  // Known by its name only, and a header no Targa has is a sentence.
  assert.equal(pictureFormat(tgaOf(src, 23, 7), 'r.tga'), 'tga');
  assert.equal(pictureFormat(tgaOf(src, 23, 7), 'r.bin'), null);
  assert.throws(() => decodePicture(bytes(head(7, 1, 1, 24, 0), [1, 2, 3]), 'x.tga'), /not a TGA picture Rutba Office knows/);
});

test('PCX: 256 colours with the palette at the end, black and white, sixteen colours in four planes, and 24-bit colour with alpha', () => {
  const header = (bits, planes, w, h, line, palette16 = []) => {
    const hd = new Uint8Array(128);
    hd.set([0x0a, 5, 1, bits, 0, 0, 0, 0, w - 1, 0, h - 1, 0]);
    hd.set(palette16, 16);
    hd[65] = planes;
    hd[66] = line;
    return [...hd];
  };
  // 8-bit: a run of three of index 2 (0xC3 0x02) then index 1, which needs no count; palette at the end.
  const pal = new Array(768).fill(0);
  pal.splice(3, 3, 10, 20, 30);
  pal.splice(6, 3, 200, 100, 50);
  const vga = decodePicture(bytes(header(8, 1, 4, 1, 4), [0xc3, 2, 1], 0x0c, pal), 'v.pcx');
  assert.deepEqual([0, 1, 2, 3].map((x) => px(vga, x, 0)), [[200, 100, 50, 255], [200, 100, 50, 255], [200, 100, 50, 255], [10, 20, 30, 255]]);
  // 1 bit: white where the bit is set.
  const mono = decodePicture(bytes(header(1, 1, 3, 1, 2), [0b10100000, 0]), 'm.pcx');
  assert.deepEqual([0, 1, 2].map((x) => px(mono, x, 0)[0]), [255, 0, 255]);
  // 1 bit in four planes: an index from the planes' bits, into the header's sixteen colours.
  const ega = [];
  for (let k = 0; k < 16; k++) ega.push(k * 16, 255 - k * 16, k);
  const four = decodePicture(bytes(header(1, 4, 2, 1, 2, ega), [0b10000000, 0, 0b01000000, 0, 0xc1, 0b11000000, 0, 0, 0]), 'e.pcx');
  assert.deepEqual([px(four, 0, 0), px(four, 1, 0)], [[80, 175, 5, 255], [96, 159, 6, 255]]);
  // 24-bit with alpha, from the writer (rows padded to an even length).
  const src = sample(13, 5);
  assert.ok(same(decodePicture(encodePcx(src, 13, 5, { alpha: true }), 'c.pcx').rgba, src));
  const opaque = decodePicture(encodePcx(src, 13, 5), 'c.pcx');
  assert.deepEqual(px(opaque, 2, 0), [...src.subarray(8, 11), 255]);
});

test('Radiance HDR: light read back within its precision, every row order, CIE XYZ, and shown tone-mapped without clipping the highlights', () => {
  const w = 40;
  const h = 6;
  const light = lightOf(sample(w, h, { alpha: false }), w, h);
  const file = encodeHdr(light, w, h);
  const pic = decodePicture(file, 'a.hdr');
  assert.deepEqual([pic.format, pic.width, pic.height], ['hdr', w, h]);
  // The light itself, through the reader's own path (RGBE keeps about 1 part in 256 of the brightest channel).
  const back = decodeHdr(file).light;
  for (let i = 0; i < light.length; i++) {
    const top = Math.max(light[i - (i % 3)], light[i - (i % 3) + 1], light[i - (i % 3) + 2]);
    assert.ok(Math.abs(back[i] - light[i]) <= top / 128 + 1e-6, `${i}: ${back[i]} for ${light[i]}`);
  }
  // Tone-mapped: brighter light is a brighter pixel, and the brightest is not burnt to white.
  const shown = toneMap(Float32Array.from([0.01, 0.01, 0.01, 1, 1, 1, 100, 100, 100]), 3, 1);
  assert.ok(shown[0] < shown[4] && shown[4] < shown[8] && shown[8] <= 255, String(shown));
  // A flat file of one pixel, rows from the bottom (+Y): the first row in the file is the bottom one.
  const flat = bytes('#?RADIANCE\nFORMAT=32-bit_rle_rgbe\n\n+Y 2 +X 1\n', [128, 0, 0, 129], [0, 0, 128, 129]);
  const up = decodePicture(flat, 'f.hdr');
  assert.ok(px(up, 0, 0)[2] > px(up, 0, 0)[0] && px(up, 0, 1)[0] > px(up, 0, 1)[2], 'blue on top, red below');
  // XYZ: daylight white (D65) stays grey.
  const xyz = decodePicture(bytes('#?RADIANCE\nFORMAT=32-bit_rle_xyze\n\n-Y 1 +X 1\n', [121, 128, 139, 129]), 'x.hdr');
  const [r, g, b] = px(xyz, 0, 0);
  assert.ok(Math.abs(r - g) < 12 && Math.abs(g - b) < 12, String([r, g, b]));
  assert.throws(() => decodePicture(bytes('#?RADIANCE\nFORMAT=48-bit\n\n-Y 1 +X 1\n', [1, 2, 3, 4]), 'x.hdr'), /48-bit layout/);
  assert.throws(() => decodePicture(bytes('#?RADIANCE\n\n-Y 1 -Y 1\n', [1, 2, 3, 4]), 'x.hdr'), /size line is damaged/);
});

test('the small formats are served decoded to the windows, and a damaged one is a refusal in a sentence, never a fault', () => {
  for (const ext of ['qoi', 'pbm', 'pgm', 'ppm', 'pnm', 'pam', 'tga', 'pcx', 'hdr', 'cur', 'tif']) assert.equal(needsDecoding(`a.${ext}`), true, ext);
  let refused = 0;
  for (const [name, good] of Object.entries(pictureSources())) {
    if (!/\.(qoi|ppm|pam|tga|pcx|hdr)$/.test(name)) continue;
    for (let i = 0; i < 300; i++) {
      const b = Buffer.from(good);
      const at = (i * 7919) % b.length;
      b[at] ^= 1 << (i % 8);
      if (i % 3 === 0) b[(at * 31) % b.length] = 0xff;
      try {
        decodePicture(i % 5 === 0 ? b.subarray(0, at) : b, name);
      } catch (err) {
        assert.equal(judge(err).outcome, 'refused', `${name}: ${err.name}: ${err.message} ${err.cause?.stack || ''}`);
        refused++;
      }
    }
  }
  assert.ok(refused > 0);
  // A huge size in a header is refused before memory is taken.
  assert.throws(() => decodePicture(bytes('P6 60000 60000 255\n', [1, 2, 3]), 'a.ppm'), /larger than/);
  assert.throws(() => decodePicture(bytes('qoif', [0, 1, 0, 0, 0, 1, 0, 0, 3, 0], [0, 0, 0, 0, 0, 0, 0, 1]), 'a.qoi'), /larger than/);
});
