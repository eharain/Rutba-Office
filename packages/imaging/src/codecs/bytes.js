// What every picture reader here shares: a refusal in a sentence, and reads
// of a file's bytes that cannot reach past its end.
//
// Every reader takes untrusted files, so none of them indexes a byte array
// it has not first checked the length of: a read past the end is a damaged
// file, said so, never a silent `undefined` turned into NaN arithmetic or a
// RangeError from a DataView.

/** A picture the reader will not draw, and why, in words a person can act on. */
export class PictureError extends Error {
  constructor(message, options) {
    super(message, options);
    this.name = 'PictureError';
    this.picture = true;
  }
}

/** Throw a refusal. */
export function bad(message) {
  throw new PictureError(message);
}

/** A file's bytes as a Uint8Array, whatever they came as. */
export function asBytes(bytes) {
  if (bytes instanceof Uint8Array) return bytes;
  if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes);
  if (ArrayBuffer.isView(bytes)) return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (typeof bytes === 'string') {
    // base64, as the document packages sometimes hand pictures over
    const bin = atob(bytes.replace(/\s+/g, ''));
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  bad('This picture arrived as something other than bytes.');
}

/** Latin-1 text of a run of bytes, clipped to what is there. */
export function ascii(b, at, len) {
  let s = '';
  const end = Math.min(b.length, at + len);
  for (let i = Math.max(0, at); i < end; i++) s += String.fromCharCode(b[i]);
  return s;
}

/**
 * A reader over a file's bytes: every read checks the end first, and a read
 * past it is a refusal naming the kind of file it was.
 */
export function reader(b, kind, little = true) {
  const short = () => bad(`This ${kind} file ends early: it may be a download that stopped, or damaged.`);
  const r = {
    b,
    little,
    length: b.length,
    need(at, n) {
      if (!(at >= 0) || at + n > b.length) short();
    },
    u8(at) {
      if (!(at >= 0) || at >= b.length) short();
      return b[at];
    },
    u16(at, le = r.little) {
      r.need(at, 2);
      return le ? b[at] | (b[at + 1] << 8) : (b[at] << 8) | b[at + 1];
    },
    u32(at, le = r.little) {
      r.need(at, 4);
      return le
        ? (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0
        : ((b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3]) >>> 0;
    },
    i16(at, le = r.little) {
      const v = r.u16(at, le);
      return v & 0x8000 ? v - 0x10000 : v;
    },
    i32(at, le = r.little) {
      return r.u32(at, le) | 0;
    },
    /** 8 bytes as a Number; offsets past 2^53 are a damaged file. */
    u64(at, le = r.little) {
      const lo = r.u32(at + (le ? 0 : 4), le);
      const hi = r.u32(at + (le ? 4 : 0), le);
      if (hi > 0x1fffff) short();
      return hi * 0x100000000 + lo;
    },
    f32(at, le = r.little) {
      r.need(at, 4);
      return new DataView(b.buffer, b.byteOffset + at, 4).getFloat32(0, le);
    },
    f64(at, le = r.little) {
      r.need(at, 8);
      return new DataView(b.buffer, b.byteOffset + at, 8).getFloat64(0, le);
    },
    slice(at, n) {
      r.need(at, n);
      return b.subarray(at, at + n);
    },
    short,
  };
  return r;
}

/** The limits a reader holds a file to before it allocates anything. */
export const LIMITS = Object.freeze({
  /** The largest file read at all. */
  bytes: 1024 * 1024 * 1024,
  /** The most pixels in one page: 120 megapixels is 480 MB of RGBA. */
  pixels: 120e6,
  /** The longest side of a page. */
  side: 65535,
  /** The most pages a file is counted to. */
  pages: 10000,
});

/** Refuse a page too large to draw before anything is allocated for it. */
export function checkSize(width, height, limits, kind) {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) {
    bad(`This ${kind} file says its picture is ${width} × ${height} pixels, which no picture is: it is damaged.`);
  }
  const side = limits?.side ?? LIMITS.side;
  const pixels = limits?.pixels ?? LIMITS.pixels;
  if (width > side || height > side || width * height > pixels) {
    bad(`This ${kind} picture is ${width} × ${height} pixels, larger than the ${Math.round(pixels / 1e6)} megapixels Rutba Office draws in one piece.`);
  }
}

/** A fresh RGBA buffer for a page, opaque black. */
export function rgbaOf(width, height) {
  const out = new Uint8Array(width * height * 4);
  for (let i = 3; i < out.length; i += 4) out[i] = 255;
  return out;
}

/** Straight colour from colour multiplied by its alpha (TIFF's associated alpha, a PSD's matte). */
export function unpremultiply(rgba) {
  for (let i = 0; i < rgba.length; i += 4) {
    const a = rgba[i + 3];
    if (a === 0 || a === 255) continue;
    rgba[i] = Math.min(255, Math.round((rgba[i] * 255) / a));
    rgba[i + 1] = Math.min(255, Math.round((rgba[i + 1] * 255) / a));
    rgba[i + 2] = Math.min(255, Math.round((rgba[i + 2] * 255) / a));
  }
  return rgba;
}
