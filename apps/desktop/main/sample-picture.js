// A picture to photograph, drawn here.
//
// The Pictures window opened with no file shows the operating system's own
// pictures folder, which is the right thing for a person and the wrong thing
// for a capture: `npm run smoke` photographed whatever was in the developer's
// Pictures folder and wrote it into build/smoke. The website captures had
// already been given generated pictures for this reason
// (build/make-screens.js); the smoke run had not, and it is the one that runs
// before every release.
//
// So: a real, valid PNG, written where the run can point a window at it. No
// dependency, no file checked in, nothing from the machine it happens to be
// on.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import zlib from 'node:zlib';

const crcTable = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

function chunk(type, data) {
  const head = Buffer.alloc(8);
  head.writeUInt32BE(data.length, 0);
  head.write(type, 4, 'ascii');
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, 'ascii'), data])), 0);
  return Buffer.concat([head, data, crc]);
}

/** A diagonal two-colour gradient as a real, valid PNG. */
export function gradientPng(width, height, from, to) {
  const raw = Buffer.alloc((width * 3 + 1) * height);
  let o = 0;
  for (let y = 0; y < height; y++) {
    raw[o++] = 0; // filter: none
    for (let x = 0; x < width; x++) {
      const t = (x / width + y / height) / 2;
      for (let c = 0; c < 3; c++) raw[o++] = Math.round(from[c] + (to[c] - from[c]) * t);
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // colour type: truecolour
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

const SAMPLES = [
  ['Harbour at dawn.png', [24, 40, 88], [232, 148, 96]],
  ['Long grass.png', [30, 84, 52], [206, 224, 138]],
  ['Blue hour.png', [12, 24, 60], [96, 130, 200]],
];

/**
 * A folder of sample pictures, made once per run and left in the temporary
 * directory. Answers the path of the first, which is what a window is opened
 * on: the app lists the folder that file is in, so it shows these and nothing
 * else.
 */
export function samplePictures() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-pictures-'));
  for (const [name, from, to] of SAMPLES) fs.writeFileSync(path.join(dir, name), gradientPng(1200, 800, from, to));
  return { dir, first: path.join(dir, SAMPLES[0][0]) };
}
