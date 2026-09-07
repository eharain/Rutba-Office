'use strict';

/**
 * PNG in, PDF image out - with nothing installed to do it.
 *
 * The README promised image support would be "an XObject and a `Do`, not a
 * redesign", and this file is the half that earns it: a PNG decoder small
 * enough to read, producing exactly what a PDF image XObject wants - raw
 * 8-bit samples, deflated, plus a separate alpha channel for /SMask.
 *
 * == The honest subset ==================================================
 *
 * Bit depth 8, non-interlaced, colour types 0 (grey), 2 (RGB), 3 (palette),
 * 4 (grey+alpha) and 6 (RGBA). That covers what canvases, screenshots and
 * signature pads produce. Everything else - 16-bit, 1/2/4-bit, Adam7
 * interlacing - is REFUSED BY NAME rather than decoded wrongly: a signature
 * image that renders as noise on a certificate is worse than an error at
 * capture time.
 *
 * zlib is Node's own; the only work here is chunk walking and scanline
 * unfiltering (the five standard filters, Paeth included).
 */

const zlib = require('zlib');

const SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/** Samples per pixel, before any palette expansion. */
const CHANNELS = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

function paeth(a, b, c) {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  return pb <= pc ? b : c;
}

/**
 * buffer → { width, height, colorSpace, bitsPerComponent, data, alpha }
 *
 * `data` is raw unfiltered samples in `colorSpace` order; `alpha` is a raw
 * 8-bit grey channel or null. Both are UNCOMPRESSED here - the document
 * assembler deflates streams itself, so tests can assert on pixels.
 */
function decodePng(buffer) {
  const bytes = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  if (bytes.length < 8 || !bytes.subarray(0, 8).equals(SIGNATURE)) {
    throw new Error('not a PNG - the eight-byte signature is missing');
  }

  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let interlace = 0;
  let palette = null;
  let transparency = null;
  const idat = [];

  let at = 8;
  while (at + 8 <= bytes.length) {
    const length = bytes.readUInt32BE(at);
    const type = bytes.toString('latin1', at + 4, at + 8);
    const body = bytes.subarray(at + 8, at + 8 + length);
    if (type === 'IHDR') {
      width = body.readUInt32BE(0);
      height = body.readUInt32BE(4);
      bitDepth = body[8];
      colorType = body[9];
      interlace = body[12];
    } else if (type === 'PLTE') {
      palette = Buffer.from(body);
    } else if (type === 'tRNS') {
      transparency = Buffer.from(body);
    } else if (type === 'IDAT') {
      idat.push(Buffer.from(body));
    } else if (type === 'IEND') {
      break;
    }
    at += 12 + length; // length + type + body + crc
  }

  if (!width || !height) throw new Error('the PNG has no IHDR');
  if (bitDepth !== 8) throw new Error(`only 8-bit PNGs are supported (this one is ${bitDepth}-bit)`);
  if (interlace !== 0) throw new Error('interlaced PNGs are not supported - re-export without Adam7');
  if (!(colorType in CHANNELS)) throw new Error(`PNG colour type ${colorType} is not supported`);
  if (colorType === 3 && !palette) throw new Error('a palette PNG with no PLTE chunk');

  const channels = CHANNELS[colorType];
  const stride = width * channels;
  const raw = zlib.inflateSync(Buffer.concat(idat));
  if (raw.length < (stride + 1) * height) throw new Error('the PNG pixel data is truncated');

  // Unfilter, scanline by scanline. `prior` is the previous UNFILTERED line.
  const pixels = Buffer.alloc(stride * height);
  for (let row = 0; row < height; row++) {
    const filter = raw[row * (stride + 1)];
    const src = raw.subarray(row * (stride + 1) + 1, (row + 1) * (stride + 1));
    const out = pixels.subarray(row * stride, (row + 1) * stride);
    const prior = row ? pixels.subarray((row - 1) * stride, row * stride) : null;
    for (let i = 0; i < stride; i++) {
      const left = i >= channels ? out[i - channels] : 0;
      const up = prior ? prior[i] : 0;
      const upLeft = prior && i >= channels ? prior[i - channels] : 0;
      let value = src[i];
      if (filter === 1) value += left;
      else if (filter === 2) value += up;
      else if (filter === 3) value += (left + up) >> 1;
      else if (filter === 4) value += paeth(left, up, upLeft);
      else if (filter !== 0) throw new Error(`PNG filter ${filter} is not in the specification`);
      out[i] = value & 0xff;
    }
  }

  // Split into what PDF wants: opaque samples + a separate alpha channel.
  const count = width * height;
  if (colorType === 2) {
    return { width, height, colorSpace: 'DeviceRGB', bitsPerComponent: 8, data: pixels, alpha: null };
  }
  if (colorType === 0) {
    return { width, height, colorSpace: 'DeviceGray', bitsPerComponent: 8, data: pixels, alpha: null };
  }
  if (colorType === 6) {
    const data = Buffer.alloc(count * 3);
    const alpha = Buffer.alloc(count);
    for (let i = 0; i < count; i++) {
      data[i * 3] = pixels[i * 4];
      data[i * 3 + 1] = pixels[i * 4 + 1];
      data[i * 3 + 2] = pixels[i * 4 + 2];
      alpha[i] = pixels[i * 4 + 3];
    }
    return { width, height, colorSpace: 'DeviceRGB', bitsPerComponent: 8, data, alpha };
  }
  if (colorType === 4) {
    const data = Buffer.alloc(count);
    const alpha = Buffer.alloc(count);
    for (let i = 0; i < count; i++) {
      data[i] = pixels[i * 2];
      alpha[i] = pixels[i * 2 + 1];
    }
    return { width, height, colorSpace: 'DeviceGray', bitsPerComponent: 8, data, alpha };
  }
  // colorType 3: palette entries are RGB triples; tRNS, when present, is a
  // parallel table of alpha values for the first N entries.
  const data = Buffer.alloc(count * 3);
  let alpha = null;
  for (let i = 0; i < count; i++) {
    const index = pixels[i] * 3;
    if (index + 2 >= palette.length) throw new Error('a palette index points past the PLTE table');
    data[i * 3] = palette[index];
    data[i * 3 + 1] = palette[index + 1];
    data[i * 3 + 2] = palette[index + 2];
    if (transparency && pixels[i] < transparency.length) {
      if (!alpha) alpha = Buffer.alloc(count, 0xff);
      alpha[i] = transparency[pixels[i]];
    }
  }
  return { width, height, colorSpace: 'DeviceRGB', bitsPerComponent: 8, data, alpha };
}

/** Is this buffer a PNG at all? The cheap gate callers ask first. */
function isPng(buffer) {
  return Buffer.isBuffer(buffer) && buffer.length >= 8 && buffer.subarray(0, 8).equals(SIGNATURE);
}

module.exports = { decodePng, isPng };
