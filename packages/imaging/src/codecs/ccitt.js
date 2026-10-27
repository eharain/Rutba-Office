// CCITT fax pictures: Group 3 (one- and two-dimensional, ITU-T T.4) and
// Group 4 (T.6), as TIFF carries them (compression 2, 3 and 4). A fax, and
// most black-and-white scans an office has, are one of these.
//
// The answer is one bit a pixel, rows padded to a byte, a 1 for each pixel
// of a black run: what libtiff answers, so the TIFF's photometric tag says
// which way round to show it. A row the file does not reach stays white; a
// damaged code ends the picture there, keeping the rows above it.

import { bad } from './bytes.js';

const WHITE_TERM = '00110101 000111 0111 1000 1011 1100 1110 1111 10011 10100 00111 01000 001000 000011 110100 110101 101010 101011 0100111 0001100 0001000 0010111 0000011 0000100 0101000 0101011 0010011 0100100 0011000 00000010 00000011 00011010 00011011 00010010 00010011 00010100 00010101 00010110 00010111 00101000 00101001 00101010 00101011 00101100 00101101 00000100 00000101 00001010 00001011 01010010 01010011 01010100 01010101 00100100 00100101 01011000 01011001 01011010 01011011 01001010 01001011 00110010 00110011 00110100';
const WHITE_MAKEUP = '11011 10010 010111 0110111 00110110 00110111 01100100 01100101 01101000 01100111 011001100 011001101 011010010 011010011 011010100 011010101 011010110 011010111 011011000 011011001 011011010 011011011 010011000 010011001 010011010 011000 010011011';
const BLACK_TERM = '0000110111 010 11 10 011 0011 0010 00011 000101 000100 0000100 0000101 0000111 00000100 00000111 000011000 0000010111 0000011000 0000001000 00001100111 00001101000 00001101100 00000110111 00000101000 00000010111 00000011000 000011001010 000011001011 000011001100 000011001101 000001101000 000001101001 000001101010 000001101011 000011010010 000011010011 000011010100 000011010101 000011010110 000011010111 000001101100 000001101101 000011011010 000011011011 000001010100 000001010101 000001010110 000001010111 000001100100 000001100101 000001010010 000001010011 000000100100 000000110111 000000111000 000000100111 000000101000 000001011000 000001011001 000000101011 000000101100 000001011010 000001100110 000001100111';
const BLACK_MAKEUP = '0000001111 000011001000 000011001001 000001011011 000000110011 000000110100 000000110101 0000001101100 0000001101101 0000001001010 0000001001011 0000001001100 0000001001101 0000001110010 0000001110011 0000001110100 0000001110101 0000001110110 0000001110111 0000001010010 0000001010011 0000001010100 0000001010101 0000001011010 0000001011011 0000001100100 0000001100101';
const EXTENDED = '00000001000 00000001100 00000001101 000000010010 000000010011 000000010100 000000010101 000000010110 000000010111 000000011100 000000011101 000000011110 000000011111';

/** A code table: (length << 16 | bits) -> run length. Exported for the tests, which hold it to being a prefix code. */
export function runTable(white) {
  const m = new Map();
  const add = (list, run0, step) => list.split(' ').forEach((code, i) => m.set((code.length << 16) | parseInt(code, 2), run0 + i * step));
  add(white ? WHITE_TERM : BLACK_TERM, 0, 1);
  add(white ? WHITE_MAKEUP : BLACK_MAKEUP, 64, 64);
  add(EXTENDED, 1792, 64);
  return m;
}

const TABLES = [runTable(true), runTable(false)];

/** The codes as strings, for an encoder in the tests. */
export const CODES = { WHITE_TERM, WHITE_MAKEUP, BLACK_TERM, BLACK_MAKEUP, EXTENDED };

class Bits {
  constructor(src, reverse) {
    this.src = src;
    this.reverse = reverse;
    this.pos = 0; // in bits
    this.end = src.length * 8;
  }
  bit(at) {
    const byte = this.src[at >> 3];
    const k = at & 7;
    return this.reverse ? (byte >> k) & 1 : (byte >> (7 - k)) & 1;
  }
  peek(n) {
    let v = 0;
    for (let i = 0; i < n; i++) v = (v << 1) | (this.pos + i < this.end ? this.bit(this.pos + i) : 0);
    return v;
  }
  read() {
    if (this.pos >= this.end) throw OUT;
    return this.bit(this.pos++);
  }
  align() {
    this.pos = (this.pos + 7) & ~7;
  }
  left() {
    return this.end - this.pos;
  }
}

const OUT = { out: true };

/** One run: makeup codes then a terminating one; -1 on an EOL found where a run should be. */
function readRun(bits, colour) {
  const table = TABLES[colour];
  let total = 0;
  for (;;) {
    let code = 0;
    let len = 0;
    let run = -1;
    while (len < 13) {
      code = (code << 1) | bits.read();
      len++;
      const hit = table.get((len << 16) | code);
      if (hit !== undefined) {
        run = hit;
        break;
      }
    }
    if (run < 0) throw BROKEN;
    total += run;
    if (run < 64) return total;
  }
}

const BROKEN = { broken: true };

/** Past an EOL (eleven or more zeros, then a one), if one is next. Answers whether there was. */
function skipEol(bits) {
  let zeros = 0;
  while (zeros < bits.left() && bits.bit(bits.pos + zeros) === 0) zeros++;
  if (zeros >= 11 && zeros < bits.left()) {
    bits.pos += zeros + 1;
    return true;
  }
  return false;
}

/** The changing elements of a row (where its colour changes) set into packed bits. */
function paint(out, row, rowBytes, width, changes) {
  let colour = 0;
  let at = 0;
  for (let i = 0; i <= changes.length; i++) {
    const to = i < changes.length ? Math.min(width, changes[i]) : width;
    if (colour) {
      for (let x = at; x < to; x++) out[row * rowBytes + (x >> 3)] |= 0x80 >> (x & 7);
    }
    at = Math.max(at, to);
    colour ^= 1;
  }
}

function oneDimensional(bits, width) {
  const changes = [];
  let a0 = 0;
  let colour = 0;
  while (a0 < width) {
    const run = readRun(bits, colour);
    a0 += run;
    if (a0 > width) throw BROKEN;
    changes.push(a0);
    colour ^= 1;
  }
  // A row's last run reaches its edge; that edge is not a change.
  if (changes[changes.length - 1] === width) changes.pop();
  return changes;
}

function twoDimensional(bits, width, ref) {
  const changes = [];
  let a0 = -1;
  let colour = 0;
  // ref: the changing elements of the row above, then the edge twice.
  const r = [...ref, width, width];
  let guard = 0;
  while (a0 < width) {
    if (++guard > 4 * width + 16) throw BROKEN;
    // b1: the first change on the row above, right of a0, to the colour opposite a0's.
    let i = 0;
    while (i < r.length - 2 && (r[i] <= a0 || (i & 1) !== colour)) i++;
    const b1 = r[i];
    const b2 = r[Math.min(i + 1, r.length - 1)];
    // The mode code.
    let mode;
    if (bits.read()) mode = 'V0';
    else if (bits.read()) mode = bits.read() ? 'VR1' : 'VL1';
    else if (bits.read()) mode = 'H';
    else if (bits.read()) mode = 'P';
    else if (bits.read()) mode = bits.read() ? 'VR2' : 'VL2';
    else if (bits.read()) mode = bits.read() ? 'VR3' : 'VL3';
    else throw BROKEN; // an extension (uncompressed mode) or an EOL mid-row
    if (mode === 'P') {
      a0 = b2;
    } else if (mode === 'H') {
      const start = a0 < 0 ? 0 : a0;
      const a1 = start + readRun(bits, colour);
      const a2 = a1 + readRun(bits, colour ^ 1);
      if (a2 > width) throw BROKEN;
      changes.push(a1, a2);
      a0 = a2;
    } else {
      const k = { V0: 0, VR1: 1, VR2: 2, VR3: 3, VL1: -1, VL2: -2, VL3: -3 }[mode];
      const a1 = b1 + k;
      if (a1 < (a0 < 0 ? 0 : a0) || a1 > width) throw BROKEN;
      changes.push(a1);
      a0 = a1;
      colour ^= 1;
    }
  }
  while (changes.length && changes[changes.length - 1] >= width) changes.pop();
  return changes;
}

/**
 * @param {Uint8Array} src the strip's bytes
 * @param {{ width, rows, kind: 2 | 3 | 4, options?: number, reverse?: boolean }} o
 *   kind 2 = Modified Huffman (each row on a byte), 3 = T.4 (options: bit 0 two-dimensional), 4 = T.6
 */
export function ccittDecode(src, { width, rows, kind, options = 0, reverse = false }) {
  const rowBytes = Math.ceil(width / 8);
  const out = new Uint8Array(rowBytes * rows);
  if (options & 2) bad('This fax picture uses the uncompressed mode, which Rutba Office does not read yet.');
  const bits = new Bits(src, reverse);
  let ref = [];
  for (let y = 0; y < rows; y++) {
    try {
      let changes;
      if (kind === 2) {
        if (y) bits.align();
        changes = oneDimensional(bits, width);
      } else if (kind === 3) {
        // EOLs come before rows; the end of the page is six of them.
        let eols = 0;
        while (skipEol(bits)) eols++;
        if (eols >= 2 && y) break;
        const twoD = options & 1 ? bits.read() === 0 : false;
        changes = twoD ? twoDimensional(bits, width, ref) : oneDimensional(bits, width);
      } else {
        if (bits.peek(24) === 0x001001) break; // EOFB
        changes = twoDimensional(bits, width, ref);
      }
      paint(out, y, rowBytes, width, changes);
      ref = changes;
    } catch (err) {
      if (err === OUT || err === BROKEN) break;
      throw err;
    }
  }
  return out;
}
