// Where to cut a large file into pieces: FastCDC, the published algorithm
// (Xia et al., USENIX ATC 2016), with its gear hash and normalised chunking.
// Pieces are 1 to 4 MB, 2 MB on average. The cut points depend only on the
// file's bytes and the folder's gear table, so a small edit moves the pieces
// around it and leaves the rest named the same.
//
// The gear table is the only part made from a secret: 1,024 bytes from HKDF of
// the folder's key, so a list of piece sizes does not by itself identify a
// file someone else also has. Nothing else about the algorithm changes.

import crypto from 'node:crypto';

export const MIN = 1 << 20;
export const AVG = 1 << 21;
export const MAX = 1 << 22;

// The masks of normalised chunking, top-aligned because the low bits of a
// 32-bit gear hash depend only on the last few bytes.
const MASK_HARD = (((1 << 23) - 1) << 9) >>> 0;
const MASK_EASY = (((1 << 19) - 1) << 13) >>> 0;

/** The gear table for a folder: 256 u32s from HKDF of its key. */
export function gearTable(folderKey) {
  const bytes = Buffer.from(crypto.hkdfSync('sha256', folderKey, Buffer.alloc(32), Buffer.from('rutba-locker:v1:gear', 'utf8'), 1024));
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) table[i] = bytes.readUInt32LE(i * 4);
  return table;
}

/**
 * The length of the next piece of `buf` (at most `max`), the point where the
 * hash says to cut. Bytes before `min` are never a cut point; a run with no
 * cut point ends at `max`.
 */
export function cutPoint(buf, length, table, min = MIN, avg = AVG, max = MAX) {
  if (length <= min) return length;
  const n = Math.min(length, max);
  let fp = 0;
  let i = min;
  const hardEnd = Math.min(n, avg);
  for (; i < hardEnd; i += 1) {
    fp = ((fp << 1) + table[buf[i]]) >>> 0;
    if ((fp & MASK_HARD) === 0) return i + 1;
  }
  for (; i < n; i += 1) {
    fp = ((fp << 1) + table[buf[i]]) >>> 0;
    if ((fp & MASK_EASY) === 0) return i + 1;
  }
  return n;
}
