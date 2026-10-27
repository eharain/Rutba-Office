// TIFF, read and written: every page, in any of the compressions an office
// meets (none, PackBits, LZW, Deflate, CCITT fax Group 3 and 4, and JPEG
// through a decoder the caller lends), in strips or tiles, chunky or
// planar, 1 to 16 bits a sample and 32-bit floats, as bilevel, grey,
// palette, RGB, CMYK, YCbCr or CIE L*a*b*, with alpha.
//
// Written here rather than taken from UTIF.js: the package is not in the
// tree and nothing is downloaded, and the common cases are a few hundred
// lines held to this suite's rule for untrusted files (every count and
// offset checked before it is used, every page's pixels counted before
// anything is allocated for them).

import { bad, reader, checkSize, rgbaOf, unpremultiply } from './bytes.js';
import { inflate, deflate } from './zlib.js';
import { lzwDecode, lzwEncode, packbitsDecode, packbitsEncode } from './lzw.js';
import { ccittDecode } from './ccitt.js';

const TYPE_SIZE = [0, 1, 1, 2, 4, 8, 1, 1, 2, 4, 8, 4, 8, 4, 0, 0, 8, 8, 8];

export function isTiff(b) {
  if (b.length < 8) return false;
  const le = b[0] === 0x49 && b[1] === 0x49;
  const be = b[0] === 0x4d && b[1] === 0x4d;
  if (!le && !be) return false;
  const magic = le ? b[2] | (b[3] << 8) : (b[2] << 8) | b[3];
  return magic === 42 || magic === 43;
}

/** The directory chain: each IFD's tags, read lazily. */
function directories(b, maxPages = 10000) {
  if (!isTiff(b)) bad('This is not a TIFF picture.');
  const r = reader(b, 'TIFF', b[0] === 0x49);
  const big = r.u16(2) === 43;
  if (big && (r.u16(4) !== 8 || r.u16(6) !== 0)) bad('This BigTIFF picture\'s header is damaged.');
  let next = big ? r.u64(8) : r.u32(4);
  const seen = new Set();
  const ifds = [];
  while (next && ifds.length < maxPages) {
    if (seen.has(next) || next >= b.length) break; // a loop, or a pointer past the end: the pages so far
    seen.add(next);
    const count = big ? r.u64(next) : r.u16(next);
    const entry = big ? 20 : 12;
    const first = next + (big ? 8 : 2);
    if (count > 4096 || first + count * entry > b.length) {
      if (!ifds.length) bad('This TIFF picture\'s first page is damaged.');
      break;
    }
    const tags = new Map();
    for (let i = 0; i < count; i++) {
      const at = first + i * entry;
      const tag = r.u16(at);
      const type = r.u16(at + 2);
      const n = big ? r.u64(at + 4) : r.u32(at + 4);
      const size = TYPE_SIZE[type] || 0;
      if (!size) continue;
      const inline = n * size <= (big ? 8 : 4);
      const valueAt = inline ? at + (big ? 12 : 8) : big ? r.u64(at + 12) : r.u32(at + 8);
      tags.set(tag, { type, n, valueAt });
    }
    ifds.push({ at: next, tags });
    const after = first + count * entry;
    next = after + (big ? 8 : 4) <= b.length ? (big ? r.u64(after) : r.u32(after)) : 0;
  }
  return { r, big, ifds };
}

function values(r, t, max = 1 << 24) {
  if (!t) return null;
  const size = TYPE_SIZE[t.type];
  if (t.n > max || t.valueAt + t.n * size > r.length) bad('A list in this TIFF picture runs past the end of the file: it is damaged.');
  const out = new Array(t.n);
  for (let i = 0; i < t.n; i++) {
    const at = t.valueAt + i * size;
    switch (t.type) {
      case 1: case 2: case 7: out[i] = r.b[at]; break;
      case 6: out[i] = (r.b[at] << 24) >> 24; break;
      case 3: out[i] = r.u16(at); break;
      case 8: out[i] = r.i16(at); break;
      case 4: case 13: out[i] = r.u32(at); break;
      case 9: out[i] = r.i32(at); break;
      case 5: { const d = r.u32(at + 4); out[i] = d ? r.u32(at) / d : 0; break; }
      case 10: { const d = r.i32(at + 4); out[i] = d ? r.i32(at) / d : 0; break; }
      case 11: out[i] = r.f32(at); break;
      case 12: out[i] = r.f64(at); break;
      default: out[i] = r.u64(at);
    }
  }
  return out;
}

const one = (r, tags, tag, fallback) => {
  const v = values(r, tags.get(tag), 1 << 16);
  return v && v.length ? v[0] : fallback;
};

/** The pages a person would call pages: the chain's IFDs, less the reduced-size previews beside a full one. */
function pagesOf(d) {
  const full = d.ifds.filter((ifd) => !(one(d.r, ifd.tags, 254, 0) & 1));
  return full.length ? full : d.ifds;
}

/** Page count and first-page size, read from the directories only. */
export function tiffInfo(b, { maxPages = 10000 } = {}) {
  const d = directories(b, maxPages);
  const pages = pagesOf(d);
  const t = pages[0].tags;
  return { format: 'tiff', width: one(d.r, t, 256, 0), height: one(d.r, t, 257, 0), pages: pages.length };
}

const COMPRESSION_NAMES = { 6: 'old-style JPEG', 34712: 'JPEG 2000', 34892: 'lossy JPEG', 34887: 'LERC', 34925: 'LZMA', 50000: 'Zstandard', 50001: 'WebP', 50002: 'JPEG XL', 32809: 'ThunderScan', 32895: 'IT8', 34661: 'JBIG', 34676: 'SGI Log', 34677: 'SGI Log24' };

function reverseBits(src) {
  const out = new Uint8Array(src.length);
  for (let i = 0; i < src.length; i++) {
    let v = src[i];
    v = ((v & 0xf0) >> 4) | ((v & 0x0f) << 4);
    v = ((v & 0xcc) >> 2) | ((v & 0x33) << 2);
    v = ((v & 0xaa) >> 1) | ((v & 0x55) << 1);
    out[i] = v;
  }
  return out;
}

/** sRGB from CIE L*a*b* (D65 white), each 0..255. */
export function labToRgb(L, a, bb) {
  let y = (L + 16) / 116;
  let x = a / 500 + y;
  let z = y - bb / 200;
  const f = (t) => (t ** 3 > 0.008856 ? t ** 3 : (t - 16 / 116) / 7.787);
  x = 0.95047 * f(x);
  y = 1.0 * f(y);
  z = 1.08883 * f(z);
  const g = (c) => {
    const v = c > 0.0031308 ? 1.055 * c ** (1 / 2.4) - 0.055 : 12.92 * c;
    return Math.max(0, Math.min(255, Math.round(v * 255)));
  };
  return [g(x * 3.2406 + y * -1.5372 + z * -0.4986), g(x * -0.9689 + y * 1.8758 + z * 0.0415), g(x * 0.0557 + y * -0.204 + z * 1.057)];
}

/**
 * One page of a TIFF as straight RGBA.
 * @param {Uint8Array} b
 * @param {{ page?: number, limits?: object, jpeg?: (bytes: Uint8Array) => ({ width, height, rgba } | null) }} o
 */
export function decodeTiff(b, { page = 0, limits, jpeg = null } = {}) {
  const d = directories(b, limits?.pages);
  const { r } = d;
  const pages = pagesOf(d);
  if (!(page >= 0 && page < pages.length)) bad(`This TIFF picture has ${pages.length} page${pages.length === 1 ? '' : 's'}; there is no page ${page + 1}.`);
  const tags = pages[page].tags;
  const width = one(r, tags, 256, 0);
  const height = one(r, tags, 257, 0);
  checkSize(width, height, limits, 'TIFF');
  const compression = one(r, tags, 259, 1);
  const spp = one(r, tags, 277, 1);
  const bpsList = values(r, tags.get(258), 64) || [1];
  const bps = bpsList[0];
  const format = one(r, tags, 339, 1); // 1 uint, 2 int, 3 float
  const planar = one(r, tags, 284, 1);
  const predictor = one(r, tags, 317, 1);
  const fill = one(r, tags, 266, 1);
  const orientation = one(r, tags, 274, 1);
  const extra = values(r, tags.get(338), 64) || [];
  let photometric = one(r, tags, 262, -1);
  if (photometric < 0) photometric = spp >= 3 ? 2 : 1;
  if (!(spp >= 1 && spp <= 16)) bad('This TIFF picture says each pixel has ' + spp + ' samples, which none has: it is damaged.');
  if (bpsList.some((v) => v !== bps)) bad('This TIFF picture gives its colours different depths, which Rutba Office does not read yet.');
  if (!(bps === 1 || bps === 2 || bps === 4 || bps === 8 || bps === 16 || bps === 32 || (bps > 0 && bps < 16))) {
    bad(`This TIFF picture has ${bps}-bit samples, which Rutba Office does not read yet.`);
  }
  if (format === 3 && bps !== 32) bad(`This TIFF picture holds ${bps}-bit floating-point samples, which Rutba Office does not read yet.`);
  if (photometric === 32844 || photometric === 32845) bad('This TIFF picture is in SGI LogLuv colour, which Rutba Office does not read yet.');
  if (photometric === 4) bad('This TIFF page is a transparency mask, not a picture.');

  // Old-style JPEG with the whole stream in one place: drawn as that JPEG.
  if (compression === 6) {
    const at = one(r, tags, 513, 0);
    const len = one(r, tags, 514, 0);
    if (at && len && at + len <= b.length && jpeg) {
      const pic = jpeg(b.subarray(at, at + len));
      if (pic && pic.width && pic.height) return { width: pic.width, height: pic.height, rgba: pic.rgba, pages: pages.length, orientation };
    }
    bad('This TIFF picture is compressed as old-style JPEG, which Rutba Office does not read yet.');
  }
  if (COMPRESSION_NAMES[compression]) bad(`This TIFF picture is compressed as ${COMPRESSION_NAMES[compression]}, which Rutba Office does not read yet.`);
  if (![1, 2, 3, 4, 5, 7, 8, 32946, 32773].includes(compression)) bad(`This TIFF picture uses compression ${compression}, which Rutba Office does not know.`);
  if (compression === 7 && !jpeg) bad('This TIFF picture\'s pages are JPEG-compressed, which this window cannot unpack; open it in Pictures or Image.');
  if ((compression === 2 || compression === 3 || compression === 4) && (bps !== 1 || spp !== 1)) bad('This fax TIFF picture is damaged: a fax is one bit a pixel.');

  const tiled = tags.has(322);
  const cw = tiled ? one(r, tags, 322, 0) : width;
  const ch = tiled ? one(r, tags, 323, 0) : Math.min(height, one(r, tags, 278, height) || height);
  if (!(cw >= 1 && ch >= 1) || cw * ch > 4 * (limits?.pixels ?? 120e6)) bad('This TIFF picture\'s strips or tiles are damaged.');
  const across = Math.ceil(width / cw);
  const down = Math.ceil(height / ch);
  const planes = planar === 2 ? spp : 1;
  const chunks = across * down * planes;
  const offsets = values(r, tags.get(tiled ? 324 : 273), 1 << 22);
  const counts = values(r, tags.get(tiled ? 325 : 279), 1 << 22);
  if (!offsets || offsets.length < chunks) bad('This TIFF picture does not say where its pixels are: it is damaged.');
  const chunkSpp = planar === 2 ? 1 : spp;
  const rowBytes = Math.ceil((cw * chunkSpp * bps) / 8);
  const chunkBytes = rowBytes * ch;
  const little = r.little;

  // YCbCr outside JPEG: subsampled blocks.
  const sub = photometric === 6 && compression !== 7 ? values(r, tags.get(530), 2) || [2, 2] : [1, 1];
  const ycc = photometric === 6 && compression !== 7 && planar === 1;
  const [sh, sv] = sub;
  if (ycc && (bps !== 8 || spp < 3 || ![1, 2, 4].includes(sh) || ![1, 2, 4].includes(sv))) bad('This TIFF picture\'s YCbCr layout is one Rutba Office does not read yet.');
  const yccRowBytes = ycc ? Math.ceil(cw / sh) * (sh * sv + 2) : 0;
  const tables = compression === 7 && tags.has(347) ? Uint8Array.from(values(r, tags.get(347), 1 << 20)) : null;

  const isFloat = format === 3;
  const samples = isFloat ? new Float32Array(width * height * spp) : null;
  const rgba = rgbaOf(width, height);
  const maxVal = bps >= 32 ? 0xffffffff : (1 << bps) - 1;
  const cmap = photometric === 3 ? values(r, tags.get(320), 3 * 65536) : null;
  if (photometric === 3 && (!cmap || cmap.length < 3 * (1 << Math.min(bps, 16)))) bad('This TIFF picture has no palette for its colours: it is damaged.');
  const paletteSize = photometric === 3 ? 1 << bps : 0;
  // Alpha is the first extra sample that says it is (1 associated, 2 not); extra samples come last.
  const alphaExtra = extra.findIndex((e) => e === 1 || e === 2);
  const alphaIndex = alphaExtra >= 0 && spp - extra.length + alphaExtra >= 1 ? spp - extra.length + alphaExtra : -1;
  const associated = alphaIndex >= 0 && extra[alphaExtra] === 1;

  // Every sample of the page, gathered into one place per pixel before colour is made.
  const plane = new Uint8Array(width * height * spp * (bps > 8 ? 2 : 1)); // 8-bit, or 16-bit as two bytes (high first)
  const wide = bps > 8 && !isFloat;
  const putSample = (x, y, s, v) => {
    const i = (y * width + x) * spp + s;
    if (wide) {
      // Depths of 9 to 15 bits widened to 16, so the high byte is the level (a palette keeps its index).
      if (bps < 16 && photometric !== 3) v = (v << (16 - bps)) & 0xffff;
      plane[i * 2] = v >>> 8;
      plane[i * 2 + 1] = v & 0xff;
    } else plane[i] = v;
  };

  for (let c = 0; c < chunks; c++) {
    const p = Math.floor(c / (across * down));
    const k = c % (across * down);
    const cx = (k % across) * cw;
    const cy = Math.floor(k / across) * ch;
    const rows = Math.min(ch, height - cy);
    const at = offsets[c];
    const n = counts && counts[c] !== undefined ? counts[c] : b.length - at;
    if (!(at < b.length) || n <= 0) continue; // a chunk missing: left as it is
    let src = b.subarray(at, Math.min(b.length, at + n));
    if (fill === 2) src = reverseBits(src);
    let data;
    const want = ycc ? yccRowBytes * Math.ceil(ch / sv) : chunkBytes;
    switch (compression) {
      case 1: data = new Uint8Array(want); data.set(src.subarray(0, want)); break;
      case 32773: data = packbitsDecode(src, want); break;
      case 5: data = lzwDecode(src, want); break;
      case 8: case 32946: { data = new Uint8Array(want); data.set(inflate(src, { size: want, max: want }).subarray(0, want)); break; }
      case 2: case 3: case 4: data = ccittDecode(src, { width: cw, rows: tiled ? ch : rows, kind: compression, options: one(r, tags, compression === 4 ? 293 : 292, 0) }); break;
      case 7: {
        let stream = src;
        if (tables && tables.length > 4 && src.length > 2) {
          // The shared tables, less their end, then this strip, less its start.
          stream = new Uint8Array(tables.length - 2 + src.length - 2);
          stream.set(tables.subarray(0, tables.length - 2));
          stream.set(src.subarray(2), tables.length - 2);
        }
        const pic = jpeg(stream);
        if (!pic || !pic.width) bad('A JPEG-compressed strip of this TIFF picture is damaged.');
        // Straight into the page: a JPEG strip is already colour.
        for (let y = 0; y < rows && y < pic.height; y++) {
          for (let x = 0; x < cw && cx + x < width && x < pic.width; x++) {
            const s = (y * pic.width + x) * 4;
            const o = ((cy + y) * width + cx + x) * 4;
            rgba[o] = pic.rgba[s];
            rgba[o + 1] = pic.rgba[s + 1];
            rgba[o + 2] = pic.rgba[s + 2];
          }
        }
        continue;
      }
      default: break;
    }

    if (ycc) {
      // Blocks of sh × sv luma samples, then Cb and Cr.
      const blocksAcross = Math.ceil(cw / sh);
      for (let by = 0; by * sv < rows; by++) {
        for (let bx = 0; bx < blocksAcross; bx++) {
          const base = by * yccRowBytes + bx * (sh * sv + 2);
          const cb = data[base + sh * sv] - 128;
          const cr = data[base + sh * sv + 1] - 128;
          for (let j = 0; j < sv; j++) {
            for (let i = 0; i < sh; i++) {
              const x = cx + bx * sh + i;
              const y = cy + by * sv + j;
              if (x >= width || y >= height || x >= cx + cw) continue;
              const Y = data[base + j * sh + i];
              const o = (y * width + x) * 4;
              rgba[o] = Math.max(0, Math.min(255, Math.round(Y + 1.402 * cr)));
              rgba[o + 1] = Math.max(0, Math.min(255, Math.round(Y - 0.344136 * cb - 0.714136 * cr)));
              rgba[o + 2] = Math.max(0, Math.min(255, Math.round(Y + 1.772 * cb)));
            }
          }
        }
      }
      continue;
    }

    // Undo the predictor, a row at a time.
    if (predictor === 2) {
      for (let y = 0; y < rows; y++) {
        const row = y * rowBytes;
        if (bps === 8) {
          for (let i = chunkSpp; i < cw * chunkSpp; i++) data[row + i] = (data[row + i] + data[row + i - chunkSpp]) & 0xff;
        } else if (bps === 16) {
          for (let i = chunkSpp; i < cw * chunkSpp; i++) {
            const a = row + i * 2;
            const pa = a - chunkSpp * 2;
            const cur = little ? data[a] | (data[a + 1] << 8) : (data[a] << 8) | data[a + 1];
            const prv = little ? data[pa] | (data[pa + 1] << 8) : (data[pa] << 8) | data[pa + 1];
            const v = (cur + prv) & 0xffff;
            if (little) { data[a] = v & 0xff; data[a + 1] = v >> 8; } else { data[a] = v >> 8; data[a + 1] = v & 0xff; }
          }
        } else if (bps === 32 && !isFloat) {
          const v = new DataView(data.buffer, data.byteOffset, data.byteLength);
          for (let i = chunkSpp; i < cw * chunkSpp; i++) v.setUint32(row + i * 4, (v.getUint32(row + i * 4, little) + v.getUint32(row + (i - chunkSpp) * 4, little)) >>> 0, little);
        }
      }
    } else if (predictor === 3 && isFloat) {
      // Floating point: the bytes of each row differenced, then laid out a byte-plane at a time.
      const wc = cw * chunkSpp;
      const tmp = new Uint8Array(wc * 4);
      for (let y = 0; y < rows; y++) {
        const row = y * rowBytes;
        for (let i = chunkSpp; i < wc * 4; i++) data[row + i] = (data[row + i] + data[row + i - chunkSpp]) & 0xff;
        tmp.set(data.subarray(row, row + wc * 4));
        for (let i = 0; i < wc; i++) for (let byte = 0; byte < 4; byte++) data[row + i * 4 + byte] = tmp[(little ? 3 - byte : byte) * wc + i];
      }
    } else if (predictor !== 1 && predictor !== 2 && predictor !== 3) {
      bad('This TIFF picture uses a predictor Rutba Office does not know.');
    }

    // The samples, into the page.
    const dv = isFloat || bps === 32 ? new DataView(data.buffer, data.byteOffset, data.byteLength) : null;
    for (let y = 0; y < rows; y++) {
      const row = y * rowBytes;
      const py = cy + y;
      for (let x = 0; x < cw; x++) {
        const px = cx + x;
        if (px >= width) break;
        for (let s = 0; s < chunkSpp; s++) {
          const si = x * chunkSpp + s;
          const sample = planar === 2 ? p : s;
          let v;
          if (bps === 8) v = data[row + si];
          else if (bps === 16) v = little ? data[row + si * 2] | (data[row + si * 2 + 1] << 8) : (data[row + si * 2] << 8) | data[row + si * 2 + 1];
          else if (bps === 32) {
            if (isFloat) {
              samples[(py * width + px) * spp + sample] = dv.getFloat32(row + si * 4, little);
              continue;
            }
            v = dv.getUint32(row + si * 4, little) >>> 16;
          } else {
            const bit = si * bps;
            let acc = 0;
            for (let q = 0; q < bps; q++) {
              const at2 = bit + q;
              acc = (acc << 1) | ((data[row + (at2 >> 3)] >> (7 - (at2 & 7))) & 1);
            }
            v = acc;
          }
          if (format === 2) v = bps === 8 ? (v ^ 0x80) : bps === 16 ? (v ^ 0x8000) : v;
          putSample(px, py, sample, v);
        }
      }
    }
  }

  if (compression === 7) {
    return { width, height, rgba, pages: pages.length, orientation };
  }
  if (ycc) return { width, height, rgba, pages: pages.length, orientation };

  // Floats: 0..1 as they come; anything outside it (heights, measurements) stretched to fit.
  let lo = 0;
  let hi = 1;
  if (isFloat) {
    let min = Infinity;
    let max = -Infinity;
    for (let i = 0; i < samples.length; i++) {
      const v = samples[i];
      if (Number.isFinite(v)) {
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
    if (min < 0 || max > 1) {
      lo = min;
      hi = max > min ? max : min + 1;
    }
  }
  const level = (i) => {
    // A sample as 0..255.
    if (isFloat) {
      const v = samples[i];
      return Number.isFinite(v) ? Math.max(0, Math.min(255, Math.round(((v - lo) / (hi - lo)) * 255))) : 0;
    }
    if (wide) return plane[i * 2]; // the high byte
    if (bps === 8) return plane[i];
    return Math.round((plane[i] * 255) / maxVal);
  };
  const raw = (i) => (wide ? (plane[i * 2] << 8) | plane[i * 2 + 1] : plane[i]);
  const inkset = one(r, tags, 332, 1);
  for (let y = 0, px = 0; y < height; y++) {
    for (let x = 0; x < width; x++, px++) {
      const base = px * spp;
      const o = px * 4;
      if (photometric === 0 || photometric === 1) {
        const g = level(base);
        rgba[o] = rgba[o + 1] = rgba[o + 2] = photometric === 0 ? 255 - g : g;
      } else if (photometric === 3) {
        const k = Math.min(raw(base), paletteSize - 1);
        rgba[o] = cmap[k] >> 8;
        rgba[o + 1] = cmap[paletteSize + k] >> 8;
        rgba[o + 2] = cmap[2 * paletteSize + k] >> 8;
      } else if (photometric === 5) {
        if (spp >= 4 && inkset !== 2) {
          const kk = 255 - level(base + 3);
          rgba[o] = Math.round(((255 - level(base)) * kk) / 255);
          rgba[o + 1] = Math.round(((255 - level(base + 1)) * kk) / 255);
          rgba[o + 2] = Math.round(((255 - level(base + 2)) * kk) / 255);
        } else {
          const g = 255 - level(base);
          rgba[o] = rgba[o + 1] = rgba[o + 2] = g;
        }
      } else if (photometric === 8 || photometric === 9) {
        const L = (level(base) * 100) / 255;
        const sa = wide ? (raw(base + 1) << 16) >> 24 : (plane[base + 1] << 24) >> 24;
        const sb = wide ? (raw(base + 2) << 16) >> 24 : (plane[base + 2] << 24) >> 24;
        const [rr, gg, bb] = labToRgb(L, photometric === 9 ? (wide ? (raw(base + 1) >> 8) - 128 : plane[base + 1] - 128) : sa, photometric === 9 ? (wide ? (raw(base + 2) >> 8) - 128 : plane[base + 2] - 128) : sb);
        rgba[o] = rr;
        rgba[o + 1] = gg;
        rgba[o + 2] = bb;
      } else {
        // RGB, and YCbCr written planar (read as RGB would be wrong; converted)
        const r0 = level(base);
        const g0 = spp > 1 ? level(base + 1) : r0;
        const b0 = spp > 2 ? level(base + 2) : r0;
        if (photometric === 6) {
          const cb = g0 - 128;
          const cr = b0 - 128;
          rgba[o] = Math.max(0, Math.min(255, Math.round(r0 + 1.402 * cr)));
          rgba[o + 1] = Math.max(0, Math.min(255, Math.round(r0 - 0.344136 * cb - 0.714136 * cr)));
          rgba[o + 2] = Math.max(0, Math.min(255, Math.round(r0 + 1.772 * cb)));
        } else {
          rgba[o] = r0;
          rgba[o + 1] = g0;
          rgba[o + 2] = b0;
        }
      }
      if (alphaIndex >= 0) rgba[o + 3] = level(base + alphaIndex);
    }
  }
  if (associated) unpremultiply(rgba);
  return { width, height, rgba, pages: pages.length, orientation };
}

/* ── writing ───────────────────────────────────────────────────────────── */

/**
 * Pages of straight RGBA as a TIFF: RGB, or RGBA with unassociated alpha
 * where a page has see-through pixels; LZW with the horizontal predictor by
 * default (what Windows Photos, Photoshop and every fax viewer read), or
 * Deflate, or none.
 * @param {Array<{ rgba: Uint8Array, width: number, height: number }>} pages
 * @param {{ compression?: 'lzw' | 'deflate' | 'packbits' | 'none', software?: string }} [o]
 */
export function encodeTiff(pages, { compression = 'lzw', software = 'Rutba Office' } = {}) {
  const list = Array.isArray(pages) ? pages : [pages];
  const code = { none: 1, lzw: 5, deflate: 8, packbits: 32773 }[compression];
  if (!code) throw new Error(`There is no TIFF compression called ${compression}.`);
  const chunks = [];
  let size = 8;
  const push = (bytes) => {
    const at = size;
    chunks.push(bytes);
    size += bytes.length;
    if (size & 1) {
      chunks.push(new Uint8Array(1));
      size += 1;
    }
    return at;
  };
  const header = new Uint8Array(8);
  header.set([0x49, 0x49, 42, 0]);
  chunks.push(header);
  let linkAt = 4; // where the previous IFD's "next" offset is written
  const links = [];
  list.forEach((page, index) => {
    const { rgba, width, height } = page;
    let alpha = false;
    for (let i = 3; i < rgba.length; i += 4) if (rgba[i] !== 255) { alpha = true; break; }
    const spp = alpha ? 4 : 3;
    const rowBytes = width * spp;
    const rowsPerStrip = Math.max(1, Math.min(height, Math.floor(65536 / Math.max(1, rowBytes))));
    const offsets = [];
    const counts = [];
    for (let y0 = 0; y0 < height; y0 += rowsPerStrip) {
      const rows = Math.min(rowsPerStrip, height - y0);
      const raw = new Uint8Array(rows * rowBytes);
      for (let y = 0; y < rows; y++) {
        for (let x = 0; x < width; x++) {
          const s = ((y0 + y) * width + x) * 4;
          const d = y * rowBytes + x * spp;
          raw[d] = rgba[s];
          raw[d + 1] = rgba[s + 1];
          raw[d + 2] = rgba[s + 2];
          if (alpha) raw[d + 3] = rgba[s + 3];
        }
        if (code === 5 || code === 8) {
          for (let i = rowBytes - 1; i >= spp; i--) raw[y * rowBytes + i] = (raw[y * rowBytes + i] - raw[y * rowBytes + i - spp]) & 0xff;
        }
      }
      const packed = code === 5 ? lzwEncode(raw) : code === 8 ? deflate(raw) : code === 32773 ? packbitsEncode(raw, rowBytes) : raw;
      offsets.push(push(packed));
      counts.push(packed.length);
    }
    const sw = new TextEncoder().encode(`${software}\0`);
    const swAt = push(sw);
    const bpsAt = push(Uint8Array.from(alpha ? [8, 0, 8, 0, 8, 0, 8, 0] : [8, 0, 8, 0, 8, 0]));
    const resAt = push(Uint8Array.from([72, 0, 0, 0, 1, 0, 0, 0]));
    const long = (v) => Uint8Array.from([v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, v >>> 24]);
    const listAt = (arr) => {
      if (arr.length === 1) return null;
      const out = new Uint8Array(arr.length * 4);
      arr.forEach((v, i) => out.set(long(v), i * 4));
      return push(out);
    };
    const offAt = listAt(offsets);
    const cntAt = listAt(counts);
    const entries = [
      [254, 4, 1, list.length > 1 ? 2 : 0],
      [256, 4, 1, width],
      [257, 4, 1, height],
      [258, 3, spp, bpsAt],
      [259, 3, 1, code],
      [262, 3, 1, 2],
      [273, 4, offsets.length, offAt ?? offsets[0]],
      [277, 3, 1, spp],
      [278, 4, 1, rowsPerStrip],
      [279, 4, counts.length, cntAt ?? counts[0]],
      [282, 5, 1, resAt],
      [283, 5, 1, resAt],
      [284, 3, 1, 1],
      [296, 3, 1, 2],
      ...(list.length > 1 ? [[297, 3, 2, index | (list.length << 16)]] : []),
      [305, 2, sw.length, sw.length <= 4 ? 0 : swAt],
      ...(code === 5 || code === 8 ? [[317, 3, 1, 2]] : []),
      ...(alpha ? [[338, 3, 1, 2]] : []),
    ];
    const ifd = new Uint8Array(2 + entries.length * 12 + 4);
    const v = new DataView(ifd.buffer);
    v.setUint16(0, entries.length, true);
    entries.forEach(([tag, type, count, value], i) => {
      const at = 2 + i * 12;
      v.setUint16(at, tag, true);
      v.setUint16(at + 2, type, true);
      v.setUint32(at + 4, count, true);
      if (type === 3 && count === 1) v.setUint16(at + 8, value, true);
      else v.setUint32(at + 8, value >>> 0, true);
    });
    const ifdAt = push(ifd);
    links.push([linkAt, ifdAt]);
    linkAt = ifdAt + 2 + entries.length * 12;
  });
  const out = new Uint8Array(size);
  let at = 0;
  for (const c of chunks) {
    out.set(c, at);
    at += c.length;
  }
  const dv = new DataView(out.buffer);
  for (const [where, value] of links) dv.setUint32(where, value, true);
  return out;
}
