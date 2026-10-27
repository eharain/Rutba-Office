// TIFF's LZW and PackBits, both ways.
//
// LZW as TIFF writes it: codes most significant bit first, 9 to 12 bits,
// the width growing one code early; and the old LSB-first variant some 1990s
// writers left, which libtiff still reads. Each is held to the size its
// strip should unpack to, so a damaged table cannot grow without end.

import { bad } from './bytes.js';

const damaged = () => bad('The compressed pixels in this TIFF picture are damaged.');

/** LZW unpacked into at most `size` bytes; a strip cut short keeps what it had. */
export function lzwDecode(src, size) {
  const out = new Uint8Array(size);
  // The old LSB-first codes start with a clear code read backwards (libtiff's test).
  const old = src.length > 1 && src[0] === 0 && src[1] & 1;
  const prefix = new Int16Array(4096);
  const suffix = new Uint8Array(4096);
  const lengthOf = new Uint16Array(4096);
  for (let i = 0; i < 256; i++) {
    prefix[i] = -1;
    suffix[i] = i;
    lengthOf[i] = 1;
  }
  let next = 258;
  let width = 9;
  let n = 0;
  let pos = 0;
  let buf = 0;
  let cnt = 0;
  let prev = -1;
  const read = () => {
    while (cnt < width) {
      if (pos >= src.length) return -1;
      if (old) buf |= src[pos++] << cnt;
      else buf = ((buf << 8) | src[pos++]) >>> 0;
      cnt += 8;
    }
    let code;
    if (old) {
      code = buf & ((1 << width) - 1);
      buf >>>= width;
    } else {
      code = (buf >>> (cnt - width)) & ((1 << width) - 1);
      buf &= (1 << (cnt - width)) - 1;
    }
    cnt -= width;
    return code;
  };
  const emit = (code) => {
    const len = lengthOf[code];
    if (n + len > size) {
      // Past the strip's end: what fits is kept, the rest dropped (writers pad).
      let c = code;
      for (let i = len - 1; i >= 0; i--) {
        if (n + i < size) out[n + i] = suffix[c];
        c = prefix[c];
      }
      n = size;
      return;
    }
    let c = code;
    for (let i = len - 1; i >= 0; i--) {
      out[n + i] = suffix[c];
      c = prefix[c];
    }
    n += len;
  };
  const firstOf = (code) => {
    let c = code;
    while (prefix[c] >= 0) c = prefix[c];
    return suffix[c];
  };
  while (n < size) {
    const code = read();
    if (code < 0 || code === 257) break;
    if (code === 256) {
      next = 258;
      width = 9;
      prev = -1;
      continue;
    }
    if (prev < 0) {
      if (code > 255) damaged();
      emit(code);
      prev = code;
      continue;
    }
    if (code < next) {
      emit(code);
      if (next < 4096) {
        prefix[next] = prev;
        suffix[next] = firstOf(code);
        lengthOf[next] = lengthOf[prev] + 1;
        next++;
      }
    } else if (code === next && next < 4096) {
      prefix[next] = prev;
      suffix[next] = firstOf(prev);
      lengthOf[next] = lengthOf[prev] + 1;
      next++;
      emit(code);
    } else {
      damaged();
    }
    prev = code;
    const grow = old ? next : next + 1;
    if (grow >= 1 << width && width < 12) width++;
  }
  return out;
}

/** LZW packed as TIFF writes it (new style, MSB first). */
export function lzwEncode(src) {
  const out = [];
  let buf = 0;
  let cnt = 0;
  let width = 9;
  const put = (code) => {
    buf = (buf << width) | code;
    cnt += width;
    while (cnt >= 8) {
      out.push((buf >>> (cnt - 8)) & 0xff);
      cnt -= 8;
    }
    buf &= (1 << cnt) - 1;
  };
  // A dictionary of (prefix code, byte) -> code, in a flat table.
  let dict = new Map();
  let next = 258;
  put(256);
  if (!src.length) {
    put(257);
    if (cnt) out.push((buf << (8 - cnt)) & 0xff);
    return Uint8Array.from(out);
  }
  let w = src[0];
  for (let i = 1; i < src.length; i++) {
    const c = src[i];
    const key = w * 256 + c;
    const hit = dict.get(key);
    if (hit !== undefined) {
      w = hit;
      continue;
    }
    put(w);
    dict.set(key, next++);
    if (next + 1 > 1 << width) {
      if (width < 12) width++;
    }
    if (next >= 4094) {
      put(256);
      dict = new Map();
      next = 258;
      width = 9;
    }
    w = c;
  }
  put(w);
  // The end code's width follows the table as the decoder will have grown it.
  if (next + 1 >= 1 << width && width < 12) width++;
  put(257);
  if (cnt) out.push((buf << (8 - cnt)) & 0xff);
  return Uint8Array.from(out);
}

/** PackBits unpacked into `size` bytes. */
export function packbitsDecode(src, size) {
  const out = new Uint8Array(size);
  let n = 0;
  let i = 0;
  while (n < size && i < src.length) {
    const h = src[i++];
    if (h < 128) {
      const len = Math.min(h + 1, size - n, src.length - i);
      out.set(src.subarray(i, i + len), n);
      n += len;
      i += h + 1;
    } else if (h > 128) {
      if (i >= src.length) break;
      const len = Math.min(257 - h, size - n);
      out.fill(src[i++], n, n + len);
      n += len;
    }
  }
  return out;
}

/** PackBits, a row at a time (runs never cross a row, as the TIFF spec asks). */
export function packbitsEncode(raw, rowBytes = raw.length) {
  const out = [];
  for (let row = 0; row < raw.length; row += rowBytes) {
    const end = Math.min(raw.length, row + rowBytes);
    let i = row;
    while (i < end) {
      let run = 1;
      while (i + run < end && run < 128 && raw[i + run] === raw[i]) run++;
      if (run >= 2) {
        out.push(257 - run, raw[i]);
        i += run;
        continue;
      }
      let lit = 1;
      while (i + lit < end && lit < 128 && !(i + lit + 1 < end && raw[i + lit] === raw[i + lit + 1])) lit++;
      out.push(lit - 1);
      for (let k = 0; k < lit; k++) out.push(raw[i + k]);
      i += lit;
    }
  }
  return Uint8Array.from(out);
}
