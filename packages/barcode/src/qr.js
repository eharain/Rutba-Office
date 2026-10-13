// QR Code, encoded here: ISO/IEC 18004.
//
// Text in, a square of modules out. The text goes in one segment, in the
// densest mode that holds it (numeric, alphanumeric, or bytes of UTF-8), at
// the error correction asked for (L, M, Q or H, which restore about 7, 15, 25
// and 30 per cent of the symbol), in the smallest of the forty versions it
// fits. The data and its Reed-Solomon codewords are interleaved and laid in
// the symbol's zigzag, and of the eight masks the one the four penalty rules
// score lowest is kept, as the standard has it.
//
// Plain arithmetic, no platform: the windows draw the modules on a canvas,
// the tests read them back through ./decode.js.

/** The four levels, with the two bits the format information writes for each. */
export const LEVELS = { L: 1, M: 0, Q: 3, H: 2 };
const LEVEL_INDEX = { L: 0, M: 1, Q: 2, H: 3 };

/** Error correction codewords in each block, by level then version (index 0 unused). */
const ECC_PER_BLOCK = [
  [-1, 7, 10, 15, 20, 26, 18, 20, 24, 30, 18, 20, 24, 26, 30, 22, 24, 28, 30, 28, 28, 28, 28, 30, 30, 26, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
  [-1, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28],
  [-1, 13, 22, 18, 26, 18, 24, 18, 22, 20, 24, 28, 26, 24, 20, 30, 24, 28, 28, 26, 30, 28, 30, 30, 30, 30, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
  [-1, 17, 28, 22, 16, 22, 28, 26, 26, 24, 28, 24, 28, 22, 24, 24, 30, 28, 28, 26, 28, 30, 24, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
];

/** How many blocks the codewords are split into, by level then version. */
const BLOCKS = [
  [-1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8, 8, 9, 9, 10, 12, 12, 12, 13, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22, 24, 25],
  [-1, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49],
  [-1, 1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20, 23, 23, 25, 27, 29, 34, 34, 35, 38, 40, 43, 45, 48, 51, 53, 56, 59, 62, 65, 68],
  [-1, 1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25, 25, 34, 30, 32, 35, 37, 40, 42, 45, 48, 51, 54, 57, 60, 63, 66, 70, 74, 77, 81],
];

/** The modes: their four-bit indicator and the length field's width in the three version bands. */
export const MODES = {
  numeric: { bits: 1, count: [10, 12, 14] },
  alphanumeric: { bits: 2, count: [9, 11, 13] },
  byte: { bits: 4, count: [8, 16, 16] },
};
export const ALPHANUMERIC = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:';

const countBits = (mode, version) => MODES[mode].count[version <= 9 ? 0 : version <= 26 ? 1 : 2];

/** The modules a version's symbol has for data and error correction, after its function patterns. */
export function rawDataModules(version) {
  let n = (16 * version + 128) * version + 64;
  if (version >= 2) {
    const align = Math.floor(version / 7) + 2;
    n -= (25 * align - 10) * align - 55;
    if (version >= 7) n -= 36;
  }
  return n;
}

/** The data codewords a version holds at a level. */
export function dataCodewords(version, level) {
  const l = LEVEL_INDEX[level];
  return Math.floor(rawDataModules(version) / 8) - ECC_PER_BLOCK[l][version] * BLOCKS[l][version];
}

/** The block structure, for the decoder too: { eccPerBlock, blocks, total }. */
export function blockLayout(version, level) {
  const l = LEVEL_INDEX[level];
  return { eccPerBlock: ECC_PER_BLOCK[l][version], blocks: BLOCKS[l][version], total: Math.floor(rawDataModules(version) / 8) };
}

/** Where the alignment patterns' centres are, along either axis. */
export function alignmentPositions(version) {
  if (version === 1) return [];
  const count = Math.floor(version / 7) + 2;
  const size = version * 4 + 17;
  const step = version === 32 ? 26 : Math.ceil((version * 4 + 4) / (count * 2 - 2)) * 2;
  const out = [6];
  for (let pos = size - 7; out.length < count; pos -= step) out.splice(1, 0, pos);
  return out;
}

/* ---- Reed-Solomon over GF(256), polynomial 0x11D ------------------------- */

const EXP = new Uint8Array(512);
const LOG = new Uint8Array(256);
{
  let x = 1;
  for (let i = 0; i < 255; i++) {
    EXP[i] = x;
    LOG[x] = i;
    x <<= 1;
    if (x & 0x100) x ^= 0x11d;
  }
  for (let i = 255; i < 512; i++) EXP[i] = EXP[i - 255];
}
export const gfMul = (a, b) => (a && b ? EXP[LOG[a] + LOG[b]] : 0);

const divisors = new Map();
/** The generator polynomial of `degree`, highest power first without its leading 1. */
function divisor(degree) {
  if (divisors.has(degree)) return divisors.get(degree);
  const out = new Uint8Array(degree);
  out[degree - 1] = 1;
  let root = 1;
  for (let i = 0; i < degree; i++) {
    for (let j = 0; j < degree; j++) {
      out[j] = gfMul(out[j], root);
      if (j + 1 < degree) out[j] ^= out[j + 1];
    }
    root = gfMul(root, 0x02);
  }
  divisors.set(degree, out);
  return out;
}

/** The error correction codewords of a block of data. */
export function reedSolomon(data, degree) {
  const div = divisor(degree);
  const out = new Uint8Array(degree);
  for (const b of data) {
    const factor = b ^ out[0];
    out.copyWithin(0, 1);
    out[degree - 1] = 0;
    for (let i = 0; i < degree; i++) out[i] ^= gfMul(div[i], factor);
  }
  return out;
}

/* ---- the data --------------------------------------------------------------- */

/** UTF-8 bytes of a text, without a platform encoder. */
export function utf8(text) {
  const out = [];
  for (const ch of String(text)) {
    const c = ch.codePointAt(0);
    if (c < 0x80) out.push(c);
    else if (c < 0x800) out.push(0xc0 | (c >> 6), 0x80 | (c & 63));
    else if (c < 0x10000) out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    else out.push(0xf0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  }
  return out;
}

/** The densest single mode a text can be written in. */
export function modeFor(text) {
  if (/^[0-9]*$/.test(text)) return 'numeric';
  if ([...text].every((ch) => ALPHANUMERIC.includes(ch))) return 'alphanumeric';
  return 'byte';
}

/** The segment's bits after its header: [value, width] pairs. */
function segmentBits(text, mode) {
  const out = [];
  if (mode === 'numeric') {
    for (let i = 0; i < text.length; i += 3) {
      const chunk = text.slice(i, i + 3);
      out.push([Number(chunk), chunk.length * 3 + 1]);
    }
    return { bits: out, count: text.length };
  }
  if (mode === 'alphanumeric') {
    for (let i = 0; i < text.length; i += 2) {
      const a = ALPHANUMERIC.indexOf(text[i]);
      if (i + 1 < text.length) out.push([a * 45 + ALPHANUMERIC.indexOf(text[i + 1]), 11]);
      else out.push([a, 6]);
    }
    return { bits: out, count: text.length };
  }
  const bytes = utf8(text);
  for (const b of bytes) out.push([b, 8]);
  return { bits: out, count: bytes.length };
}

/** A version, the smallest at least `min` the text fits at the level, and its data codewords. */
function plan(text, level, mode, min = 1, max = 40) {
  const seg = segmentBits(text, mode);
  const body = seg.bits.reduce((n, [, w]) => n + w, 0);
  for (let v = min; v <= max; v++) {
    const cb = countBits(mode, v);
    if (seg.count >= 1 << cb) continue;
    const need = 4 + cb + body;
    if (need <= dataCodewords(v, level) * 8) return { version: v, seg, need };
  }
  return null;
}

/** The data codewords: mode, length, the text, terminator and padding. */
function dataBytes(version, level, mode, seg) {
  const bits = [];
  const push = (value, width) => { for (let i = width - 1; i >= 0; i--) bits.push((value >>> i) & 1); };
  push(MODES[mode].bits, 4);
  push(seg.count, countBits(mode, version));
  for (const [value, width] of seg.bits) push(value, width);
  const capacity = dataCodewords(version, level) * 8;
  push(0, Math.min(4, capacity - bits.length));
  while (bits.length % 8) bits.push(0);
  const out = [];
  for (let i = 0; i < bits.length; i += 8) out.push(bits.slice(i, i + 8).reduce((n, b) => (n << 1) | b, 0));
  for (let pad = 0xec; out.length < capacity / 8; pad ^= 0xec ^ 0x11) out.push(pad);
  return out;
}

/** Data codewords split into blocks, each given its error correction, then interleaved. */
export function interleave(data, version, level) {
  const { eccPerBlock, blocks, total } = blockLayout(version, level);
  const short = blocks - (total % blocks);
  const shortLen = Math.floor(total / blocks);
  const parts = [];
  for (let i = 0, k = 0; i < blocks; i++) {
    const len = shortLen - eccPerBlock + (i < short ? 0 : 1);
    const d = data.slice(k, k + len);
    k += len;
    parts.push({ data: d, ecc: reedSolomon(d, eccPerBlock) });
  }
  const out = [];
  const longest = shortLen - eccPerBlock + 1;
  for (let i = 0; i < longest; i++) for (const p of parts) if (i < p.data.length) out.push(p.data[i]);
  for (let i = 0; i < eccPerBlock; i++) for (const p of parts) out.push(p.ecc[i]);
  return out;
}

/* ---- the symbol --------------------------------------------------------------- */

/** The fifteen format bits for a level and mask, BCH-protected and masked. */
export function formatBits(level, mask) {
  const data = (LEVELS[level] << 3) | mask;
  let rem = data;
  for (let i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
  return ((data << 10) | rem) ^ 0x5412;
}

/** The eighteen version bits (versions 7 and up), BCH-protected. */
export function versionBits(version) {
  let rem = version;
  for (let i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >>> 11) * 0x1f25);
  return (version << 12) | rem;
}

/** The eight mask conditions: true where a module is flipped. `x` is the column, `y` the row. */
export const MASKS = [
  (x, y) => (x + y) % 2 === 0,
  (x, y) => y % 2 === 0,
  (x) => x % 3 === 0,
  (x, y) => (x + y) % 3 === 0,
  (x, y) => (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0,
  (x, y) => ((x * y) % 2) + ((x * y) % 3) === 0,
  (x, y) => (((x * y) % 2) + ((x * y) % 3)) % 2 === 0,
  (x, y) => (((x + y) % 2) + ((x * y) % 3)) % 2 === 0,
];

/**
 * A version's function patterns: finders with their separators, timing,
 * alignment, the dark module, and the areas the format and version
 * information take. Answers { size, dark, fixed } as row-major arrays.
 */
export function functionPatterns(version) {
  const size = version * 4 + 17;
  const dark = new Uint8Array(size * size);
  const fixed = new Uint8Array(size * size);
  const set = (x, y, on) => { dark[y * size + x] = on ? 1 : 0; fixed[y * size + x] = 1; };
  for (let i = 0; i < size; i++) { set(6, i, i % 2 === 0); set(i, 6, i % 2 === 0); }
  const finder = (cx, cy) => {
    for (let dy = -4; dy <= 4; dy++) {
      for (let dx = -4; dx <= 4; dx++) {
        const x = cx + dx;
        const y = cy + dy;
        if (x < 0 || y < 0 || x >= size || y >= size) continue;
        const d = Math.max(Math.abs(dx), Math.abs(dy));
        set(x, y, d !== 2 && d !== 4);
      }
    }
  };
  finder(3, 3);
  finder(size - 4, 3);
  finder(3, size - 4);
  const align = alignmentPositions(version);
  for (let i = 0; i < align.length; i++) {
    for (let j = 0; j < align.length; j++) {
      if ((i === 0 && j === 0) || (i === 0 && j === align.length - 1) || (i === align.length - 1 && j === 0)) continue;
      for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) set(align[i] + dx, align[j] + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);
    }
  }
  // The format areas, filled in once the mask is chosen, and the dark module.
  for (let i = 0; i <= 8; i++) { if (!fixed[8 * size + i]) set(i, 8, false); if (!fixed[i * size + 8]) set(8, i, false); }
  for (let i = 0; i < 8; i++) { set(size - 1 - i, 8, false); set(8, size - 1 - i, false); }
  set(8, size - 8, true);
  if (version >= 7) {
    const bits = versionBits(version);
    for (let i = 0; i < 18; i++) {
      const on = ((bits >>> i) & 1) === 1;
      const a = size - 11 + (i % 3);
      const b = Math.floor(i / 3);
      set(a, b, on);
      set(b, a, on);
    }
  }
  return { size, dark, fixed };
}

/** The format information written into both of its places. */
function placeFormat(dark, size, level, mask) {
  const bits = formatBits(level, mask);
  const bit = (i) => (bits >>> i) & 1;
  const set = (x, y, on) => { dark[y * size + x] = on; };
  for (let i = 0; i <= 5; i++) set(8, i, bit(i));
  set(8, 7, bit(6));
  set(8, 8, bit(7));
  set(7, 8, bit(8));
  for (let i = 9; i < 15; i++) set(14 - i, 8, bit(i));
  for (let i = 0; i < 8; i++) set(size - 1 - i, 8, bit(i));
  for (let i = 8; i < 15; i++) set(8, size - 15 + i, bit(i));
  set(8, size - 8, 1);
}

/** The modules the codewords take, in the order they are laid: [x, y] pairs. */
export function zigzag(size, fixed) {
  const out = [];
  for (let right = size - 1; right >= 1; right -= 2) {
    if (right === 6) right = 5;
    const upward = ((right + 1) & 2) === 0;
    for (let v = 0; v < size; v++) {
      const y = upward ? size - 1 - v : v;
      for (let j = 0; j < 2; j++) {
        const x = right - j;
        if (!fixed[y * size + x]) out.push([x, y]);
      }
    }
  }
  return out;
}

/** The penalty the standard's four rules give a finished symbol: the lower, the easier to read. */
export function penalty(dark, size) {
  let score = 0;
  const at = (x, y) => dark[y * size + x];
  // Rule 1: five or more of a colour in a row or a column.
  for (let pass = 0; pass < 2; pass++) {
    for (let a = 0; a < size; a++) {
      let run = 1;
      for (let b = 1; b <= size; b++) {
        const same = b < size && (pass ? at(a, b) === at(a, b - 1) : at(b, a) === at(b - 1, a));
        if (same) run++;
        else {
          if (run >= 5) score += 3 + (run - 5);
          run = 1;
        }
      }
    }
  }
  // Rule 2: every 2x2 block of one colour.
  for (let y = 0; y < size - 1; y++) {
    for (let x = 0; x < size - 1; x++) {
      const c = at(x, y);
      if (c === at(x + 1, y) && c === at(x, y + 1) && c === at(x + 1, y + 1)) score += 3;
    }
  }
  // Rule 3: the finder's 1:1:3:1:1 with four light modules on either side.
  const A = [1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0];
  const B = [0, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1];
  for (let a = 0; a < size; a++) {
    for (let b = 0; b + 11 <= size; b++) {
      let rowA = true; let rowB = true; let colA = true; let colB = true;
      for (let k = 0; k < 11; k++) {
        const r = at(b + k, a);
        const c = at(a, b + k);
        if (r !== A[k]) rowA = false;
        if (r !== B[k]) rowB = false;
        if (c !== A[k]) colA = false;
        if (c !== B[k]) colB = false;
      }
      score += 40 * (rowA + rowB + colA + colB);
    }
  }
  // Rule 4: how far the dark share strays from half, in steps of five per cent.
  let darkCount = 0;
  for (const m of dark) darkCount += m;
  const k = Math.ceil(Math.abs(darkCount * 20 - size * size * 10) / (size * size)) - 1;
  score += Math.max(0, k) * 10;
  return score;
}

/**
 * A QR Code for `text`. Options: `level` ('L' | 'M' | 'Q' | 'H', 'M' when
 * left out), `mask` (0 to 7 to choose one; the best by penalty otherwise),
 * `minVersion` / `version`. Answers { version, level, mask, mode, size,
 * modules }, `modules` an array of rows of booleans, dark true.
 */
export function encodeQr(text, { level = 'M', mask = null, minVersion = 1, version = null, mode = null } = {}) {
  const value = String(text ?? '');
  if (!(level in LEVELS)) throw new Error(`"${level}" is not an error correction level: L, M, Q or H`);
  const m = mode || modeFor(value);
  const chosen = plan(value, level, m, version || minVersion, version || 40);
  if (!chosen) throw new Error(version ? `The text is too long for a version ${version} QR Code at level ${level}.` : `The text is too long for a QR Code at level ${level}.`);
  const v = chosen.version;
  const codewords = interleave(dataBytes(v, level, m, chosen.seg), v, level);
  const { size, dark: base, fixed } = functionPatterns(v);
  const placed = new Uint8Array(base);
  const order = zigzag(size, fixed);
  let i = 0;
  for (const [x, y] of order) {
    const byte = codewords[i >>> 3];
    placed[y * size + x] = byte === undefined ? 0 : (byte >>> (7 - (i & 7))) & 1;
    i++;
  }
  const masked = (k) => {
    const out = new Uint8Array(placed);
    for (const [x, y] of order) if (MASKS[k](x, y)) out[y * size + x] ^= 1;
    placeFormat(out, size, level, k);
    return out;
  };
  let best = null;
  for (let k = 0; k < 8; k++) {
    if (mask != null && k !== Number(mask)) continue;
    const out = masked(k);
    const score = penalty(out, size);
    if (!best || score < best.score) best = { k, out, score };
  }
  const modules = [];
  for (let y = 0; y < size; y++) modules.push(Array.from(best.out.subarray(y * size, (y + 1) * size), (b) => b === 1));
  return { version: v, level, mask: best.k, mode: m, size, modules, penalty: best.score };
}
