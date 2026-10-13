// QR Codes, encoded by the suite: the standard's own worked example
// (ISO/IEC 18004 Annex I, "01234567" at 1-M) to the codeword, the format and
// version information to the table, symbols module for module against an
// independent encoder's, the masks chosen by the penalty rules, every
// version and level read back by the reader, and pictures of symbols read
// back off their pixels.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import { encodeQr, formatBits, versionBits, reedSolomon, dataCodewords, alignmentPositions, modeFor } from '@rutba/barcode';
import { interleave, penalty, functionPatterns } from '@rutba/barcode/qr';
import { decodeQr, modulesFromPicture } from '@rutba/barcode/decode';
import { readPng, isDark } from '@rutba/barcode/png-read';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const REFERENCE = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/qr-reference.json'), 'utf8')).symbols;
const bin = (s) => parseInt(s, 2);

test('the standard\'s example: 01234567 at 1-M makes its data and error correction codewords', () => {
  const data = ['00010000', '00100000', '00001100', '01010110', '01100001', '10000000', '11101100', '00010001', '11101100', '00010001', '11101100', '00010001', '11101100', '00010001', '11101100', '00010001'].map(bin);
  const ecc = ['10100101', '00100100', '11010100', '11000001', '11101101', '00110110', '11000111', '10000111', '00101100', '01010101'].map(bin);
  assert.deepEqual([...reedSolomon(data, 10)], ecc);
  assert.deepEqual(interleave(data, 1, 'M'), [...data, ...ecc]);
  const q = encodeQr('01234567', { level: 'M' });
  assert.equal(q.mode, 'numeric');
  assert.equal(q.version, 1);
  assert.equal(decodeQr(q.modules).text, '01234567');
});

test('HELLO WORLD at 1-Q: alphanumeric, the codewords every worked example shows, and mask 6', () => {
  const q = encodeQr('HELLO WORLD', { level: 'Q' });
  assert.equal(q.mode, 'alphanumeric');
  assert.equal(q.mask, 6);
  assert.deepEqual([...reedSolomon([32, 91, 11, 120, 209, 114, 220, 77, 67, 64, 236, 17, 236], 13)], [168, 72, 22, 82, 217, 54, 156, 0, 46, 15, 180, 122, 16]);
});

test('the format information is the standard\'s table, and the version information its BCH code', () => {
  const table = {
    L: ['111011111000100', '111001011110011', '111110110101010', '111100010011101', '110011000101111', '110001100011000', '110110001000001', '110100101110110'],
    M: ['101010000010010', '101000100100101', '101111001111100', '101101101001011', '100010111111001', '100000011001110', '100111110010111', '100101010100000'],
    Q: ['011010101011111', '011000001101000', '011111100110001', '011101000000110', '010010010110100', '010000110000011', '010111011011010', '010101111101101'],
    H: ['001011010001001', '001001110111110', '001110011100111', '001100111010000', '000011101100010', '000001001010101', '000110100001100', '000100000111011'],
  };
  for (const [level, masks] of Object.entries(table)) masks.forEach((bits, mask) => assert.equal(formatBits(level, mask), bin(bits), `${level}${mask}`));
  assert.equal(versionBits(7), 0x07c94);
  assert.equal(versionBits(8), 0x085bc);
  assert.equal(versionBits(40), 0x28c69);
});

test('capacities and alignment patterns where the standard puts them', () => {
  assert.deepEqual([dataCodewords(1, 'L'), dataCodewords(1, 'H'), dataCodewords(10, 'M'), dataCodewords(40, 'L'), dataCodewords(40, 'H')], [19, 9, 216, 2956, 1276]);
  assert.deepEqual(alignmentPositions(1), []);
  assert.deepEqual(alignmentPositions(2), [6, 18]);
  assert.deepEqual(alignmentPositions(7), [6, 22, 38]);
  assert.deepEqual(alignmentPositions(32), [6, 34, 60, 86, 112, 138]);
  assert.deepEqual(alignmentPositions(40), [6, 30, 58, 86, 114, 142, 170]);
});

test('symbols are module for module those an independent encoder writes, the same mask chosen', () => {
  for (const ref of REFERENCE) {
    const q = encodeQr(ref.text, { level: ref.level });
    assert.equal(q.version, ref.version, `${ref.level} "${ref.text.slice(0, 16)}": version`);
    assert.equal(q.mask, ref.mask, `${ref.level} "${ref.text.slice(0, 16)}": mask`);
    const rows = q.modules.map((row) => row.map((b) => (b ? '1' : '0')).join('').padEnd(Math.ceil(q.size / 4) * 4, '0').match(/.{4}/g).map((b) => parseInt(b, 2).toString(16)).join(''));
    assert.deepEqual(rows, ref.rows, `${ref.level} "${ref.text.slice(0, 16)}": modules`);
  }
});

test('the mask kept is the one of the eight the penalty rules score lowest', () => {
  for (const text of ['https://rutba.io', 'Rutba Office', '1234567890']) {
    const best = encodeQr(text, { level: 'M' });
    for (let mask = 0; mask < 8; mask++) {
      const other = encodeQr(text, { level: 'M', mask });
      const flat = (q) => Uint8Array.from(q.modules.flat(), (b) => (b ? 1 : 0));
      assert.ok(penalty(flat(best), best.size) <= penalty(flat(other), other.size), `${text}: mask ${mask}`);
      assert.equal(decodeQr(other.modules).text, text, `${text} reads back under mask ${mask}`);
    }
  }
});

test('every version and level reads back to its text, in each mode', () => {
  const alnum = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:';
  for (let version = 1; version <= 40; version++) {
    for (const level of ['L', 'M', 'Q', 'H']) {
      const room = dataCodewords(version, level);
      const kind = (version + level.charCodeAt(0)) % 3;
      // A text that fills most of the version's room, so the version chosen is this one.
      const text = kind === 0
        ? Array.from({ length: Math.floor((room * 8 - 30) / 3.34) }, (_, i) => String(i % 10)).join('')
        : kind === 1
          ? Array.from({ length: Math.floor((room * 8 - 30) / 5.5) }, (_, i) => alnum[i % alnum.length]).join('')
          : Array.from({ length: room - 3 }, (_, i) => String.fromCharCode(97 + (i % 26))).join('');
      const q = encodeQr(text, { level, version });
      assert.equal(q.size, version * 4 + 17);
      const back = decodeQr(q.modules);
      assert.equal(back.text, text, `version ${version}-${level}`);
      assert.equal(back.level, level);
      assert.equal(back.mode, modeFor(text));
    }
  }
});

test('words in any script go as UTF-8 bytes, and a text too long for the level is refused', () => {
  for (const text of ['اردو کا ایک جملہ', 'हिन्दी', 'emoji 🙂 too', 'mixed Case 123']) assert.equal(decodeQr(encodeQr(text, { level: 'H' }).modules).text, text);
  assert.throws(() => encodeQr('a'.repeat(3000), { level: 'H' }), /too long/);
  assert.throws(() => encodeQr('x', { level: 'Z' }), /error correction level/);
});

test('the function patterns: finders, timing, and the dark module', () => {
  const { size, dark, fixed } = functionPatterns(7);
  assert.equal(size, 45);
  assert.equal(dark[8 * size + 6], 1, 'timing is dark on even places');
  assert.equal(dark[(size - 8) * size + 8], 1, 'the dark module');
  assert.equal(fixed[0], 1);
});

/** A PNG of a symbol at `scale` pixels a module, with a quiet zone, in a colour on white. */
function symbolPng(modules, scale, [r, g, b]) {
  const n = modules.length + 8;
  const w = n * scale;
  const raw = Buffer.alloc((w * 4 + 1) * w, 255);
  for (let y = 0; y < w; y++) {
    raw[y * (w * 4 + 1)] = 0;
    for (let x = 0; x < w; x++) {
      const my = Math.floor(y / scale) - 4;
      const mx = Math.floor(x / scale) - 4;
      if (my >= 0 && mx >= 0 && my < modules.length && mx < modules.length && modules[my][mx]) {
        const o = y * (w * 4 + 1) + 1 + x * 4;
        raw[o] = r; raw[o + 1] = g; raw[o + 2] = b;
      }
    }
  }
  const crc = (buf) => { let c = -1; for (const byte of buf) { c ^= byte; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; } return (c ^ -1) >>> 0; };
  const chunk = (type, data) => { const h = Buffer.alloc(8); h.writeUInt32BE(data.length); h.write(type, 4, 'latin1'); const t = Buffer.alloc(4); t.writeUInt32BE(crc(Buffer.concat([h.subarray(4), data]))); return Buffer.concat([h, data, t]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(w, 4); ihdr[8] = 8; ihdr[9] = 6;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

test('a picture of a symbol, in a dark colour, is read back off its pixels', () => {
  const text = 'https://rutba.io/office?from=qr';
  const q = encodeQr(text, { level: 'Q' });
  const png = readPng(symbolPng(q.modules, 6, [0x1f, 0x38, 0x64]));
  const modules = modulesFromPicture({ width: png.width, height: png.height, dark: (x, y) => isDark(png.rgba(x, y)) });
  assert.equal(modules.length, q.size);
  assert.equal(decodeQr(modules).text, text);
});
