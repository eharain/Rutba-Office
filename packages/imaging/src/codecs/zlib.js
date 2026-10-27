// Deflate, both ways, wherever a picture is read.
//
// In the main process and Node's workers the platform's own zlib does it
// (reached without an import, so the windows' bundle never asks for
// node:zlib); in a window, which has no synchronous zlib, the inflater here
// does, written after Mark Adler's puff. Both are held to the size the
// caller expects, so a small file that claims to unpack into gigabytes is a
// refusal, not a window that runs out of memory.

import { bad } from './bytes.js';

const NODE = (() => {
  try {
    return globalThis.process?.getBuiltinModule?.('node:zlib') || null;
  } catch {
    return null;
  }
})();

/* ── checksums ─────────────────────────────────────────────────────────── */

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(bytes, crc = 0) {
  let c = (crc ^ 0xffffffff) >>> 0;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

export function adler32(bytes) {
  let a = 1;
  let b = 0;
  for (let i = 0; i < bytes.length; ) {
    const end = Math.min(bytes.length, i + 5552);
    for (; i < end; i++) {
      a += bytes[i];
      b += a;
    }
    a %= 65521;
    b %= 65521;
  }
  return ((b << 16) | a) >>> 0;
}

/* ── inflate ───────────────────────────────────────────────────────────── */

const LBASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEXT = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DBASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DEXT = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const CLORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

const damaged = () => bad('The compressed pixels in this picture are damaged.');

/** A canonical Huffman code from its code lengths: counts per length and the symbols in order. */
function huffman(lengths) {
  const counts = new Uint16Array(16);
  for (let i = 0; i < lengths.length; i++) counts[lengths[i]]++;
  counts[0] = 0;
  let left = 1;
  for (let len = 1; len < 16; len++) {
    left <<= 1;
    left -= counts[len];
    if (left < 0) damaged();
  }
  const offs = new Uint16Array(16);
  for (let len = 1; len < 15; len++) offs[len + 1] = offs[len] + counts[len];
  const symbols = new Uint16Array(lengths.length);
  for (let i = 0; i < lengths.length; i++) if (lengths[i]) symbols[offs[lengths[i]]++] = i;
  return { counts, symbols };
}

let FIXED = null;
function fixed() {
  if (FIXED) return FIXED;
  const l = new Uint8Array(288);
  for (let i = 0; i < 144; i++) l[i] = 8;
  for (let i = 144; i < 256; i++) l[i] = 9;
  for (let i = 256; i < 280; i++) l[i] = 7;
  for (let i = 280; i < 288; i++) l[i] = 8;
  const d = new Uint8Array(30).fill(5);
  FIXED = { lit: huffman(l), dist: huffman(d) };
  return FIXED;
}

/**
 * Raw deflate unpacked, in JavaScript.
 * @param {Uint8Array} src
 * @param {{ size?: number, max?: number }} o the size expected (a first guess) and the most allowed
 */
export function inflateRawJs(src, { size = 0, max = 512 * 1024 * 1024 } = {}) {
  let pos = 0;
  let bitbuf = 0;
  let bitcnt = 0;
  let out = new Uint8Array(Math.max(256, Math.min(size || src.length * 4, max)));
  let n = 0;
  const bits = (need) => {
    while (bitcnt < need) {
      if (pos >= src.length) damaged();
      bitbuf |= src[pos++] << bitcnt;
      bitcnt += 8;
    }
    const v = bitbuf & ((1 << need) - 1);
    bitbuf >>>= need;
    bitcnt -= need;
    return v;
  };
  const room = (k) => {
    if (n + k <= out.length) return;
    if (n + k > max) bad('This picture unpacks to more than it says it holds: it is damaged.');
    let cap = out.length * 2;
    while (cap < n + k) cap *= 2;
    const grown = new Uint8Array(Math.min(cap, max));
    grown.set(out.subarray(0, n));
    out = grown;
  };
  const decode = (h) => {
    let code = 0;
    let first = 0;
    let index = 0;
    for (let len = 1; len < 16; len++) {
      code |= bits(1);
      const count = h.counts[len];
      if (code - count < first) return h.symbols[index + (code - first)];
      index += count;
      first += count;
      first <<= 1;
      code <<= 1;
    }
    return damaged();
  };
  const codes = (lit, dist) => {
    for (;;) {
      let sym = decode(lit);
      if (sym < 256) {
        if (n >= out.length) room(1);
        out[n++] = sym;
      } else if (sym === 256) {
        return;
      } else {
        sym -= 257;
        if (sym >= 29) damaged();
        const len = LBASE[sym] + bits(LEXT[sym]);
        const ds = decode(dist);
        if (ds >= 30) damaged();
        const d = DBASE[ds] + bits(DEXT[ds]);
        if (d > n) damaged();
        room(len);
        for (let i = 0; i < len; i++, n++) out[n] = out[n - d];
      }
    }
  };
  let last = 0;
  do {
    last = bits(1);
    const type = bits(2);
    if (type === 0) {
      // A stored block starts on a byte: what is left of the bit buffer is whole bytes unread.
      pos -= bitcnt >> 3;
      bitbuf = 0;
      bitcnt = 0;
      if (pos + 4 > src.length) damaged();
      const len = src[pos] | (src[pos + 1] << 8);
      const nlen = src[pos + 2] | (src[pos + 3] << 8);
      if ((len ^ 0xffff) !== nlen) damaged();
      pos += 4;
      if (pos + len > src.length) damaged();
      room(len);
      out.set(src.subarray(pos, pos + len), n);
      n += len;
      pos += len;
    } else if (type === 1) {
      const f = fixed();
      codes(f.lit, f.dist);
    } else if (type === 2) {
      const nlen = bits(5) + 257;
      const ndist = bits(5) + 1;
      const ncode = bits(4) + 4;
      if (nlen > 286 || ndist > 30) damaged();
      const lengths = new Uint8Array(nlen + ndist);
      const cl = new Uint8Array(19);
      for (let i = 0; i < ncode; i++) cl[CLORDER[i]] = bits(3);
      const lencode = huffman(cl);
      for (let i = 0; i < nlen + ndist; ) {
        const sym = decode(lencode);
        if (sym < 16) {
          lengths[i++] = sym;
        } else {
          let len = 0;
          let rep;
          if (sym === 16) {
            if (i === 0) damaged();
            len = lengths[i - 1];
            rep = 3 + bits(2);
          } else if (sym === 17) rep = 3 + bits(3);
          else rep = 11 + bits(7);
          if (i + rep > nlen + ndist) damaged();
          while (rep--) lengths[i++] = len;
        }
      }
      if (!lengths[256]) damaged();
      codes(huffman(lengths.subarray(0, nlen)), huffman(lengths.subarray(nlen)));
    } else {
      damaged();
    }
  } while (!last);
  return out.subarray(0, n);
}

/**
 * Zlib-wrapped (or raw) deflate unpacked.
 * @param {Uint8Array} src
 * @param {{ size?: number, max?: number, raw?: boolean }} o
 */
export function inflate(src, { size = 0, max = 512 * 1024 * 1024, raw = false } = {}) {
  let body = src;
  if (!raw) {
    if (src.length < 2 || (src[0] & 0x0f) !== 8 || ((src[0] << 8) | src[1]) % 31 !== 0) {
      // Not a zlib header: some writers put raw deflate where zlib belongs.
      raw = true;
    } else {
      if (src[1] & 0x20) damaged();
      body = src.subarray(2);
    }
  }
  if (NODE) {
    try {
      const out = NODE.inflateRawSync(body, { maxOutputLength: Math.max(1, max), finishFlush: NODE.constants.Z_SYNC_FLUSH });
      return new Uint8Array(out.buffer, out.byteOffset, out.byteLength);
    } catch (err) {
      if (err?.code === 'ERR_BUFFER_TOO_LARGE') bad('This picture unpacks to more than it says it holds: it is damaged.');
      damaged();
    }
  }
  return inflateRawJs(body, { size, max });
}

/* ── deflate ───────────────────────────────────────────────────────────── */

/**
 * Bytes deflated in a zlib wrapper: the platform's zlib where there is one,
 * else stored blocks (correct, and larger: a window encodes through its
 * canvas instead).
 */
export function deflate(bytes, { level = 6 } = {}) {
  if (NODE) {
    const out = NODE.deflateSync(bytes, { level });
    return new Uint8Array(out.buffer, out.byteOffset, out.byteLength);
  }
  const blocks = Math.max(1, Math.ceil(bytes.length / 65535));
  const out = new Uint8Array(2 + bytes.length + blocks * 5 + 4);
  out[0] = 0x78;
  out[1] = 0x01;
  let at = 2;
  for (let i = 0; i < blocks; i++) {
    const start = i * 65535;
    const len = Math.min(65535, bytes.length - start);
    out[at] = i === blocks - 1 ? 1 : 0;
    out[at + 1] = len & 0xff;
    out[at + 2] = len >> 8;
    out[at + 3] = ~len & 0xff;
    out[at + 4] = (~len >> 8) & 0xff;
    out.set(bytes.subarray(start, start + len), at + 5);
    at += 5 + len;
  }
  const a = adler32(bytes);
  out[at] = a >>> 24;
  out[at + 1] = (a >>> 16) & 0xff;
  out[at + 2] = (a >>> 8) & 0xff;
  out[at + 3] = a & 0xff;
  return out;
}

/** Whether the platform's zlib is here (the main process, Node's workers), for the tests. */
export const hasPlatformZlib = Boolean(NODE);
