// Photoshop documents (PSD, and the large kind, PSB), flattened: the picture
// Photoshop keeps of the whole document beside its layers (the "image data"
// section, saved when Maximize Compatibility is on, which it is by default).
// The layers themselves are not read: what is shown is what Photoshop shows.
//
// Header, colour-mode data (a palette), image resources, the layers (skipped,
// though their count says whether the picture has transparency), then each
// channel of the flattened picture in turn, raw, PackBits by row, or zipped.

import { bad, reader, checkSize, rgbaOf } from './bytes.js';
import { packbitsDecode } from './lzw.js';
import { inflate } from './zlib.js';
import { labToRgb } from './tiff.js';

export function isPsd(b) {
  return b.length >= 26 && b[0] === 0x38 && b[1] === 0x42 && b[2] === 0x50 && b[3] === 0x53 && b[4] === 0 && (b[5] === 1 || b[5] === 2);
}

const MODES = { 0: 'bitmap', 1: 'grey', 2: 'indexed', 3: 'rgb', 4: 'cmyk', 7: 'multichannel', 8: 'duotone', 9: 'lab' };

function header(b) {
  if (!isPsd(b)) bad('This is not a Photoshop document.');
  const r = reader(b, 'Photoshop', false);
  const big = b[5] === 2;
  const h = { big, channels: r.u16(12), height: r.u32(14), width: r.u32(18), depth: r.u16(22), mode: r.u16(24) };
  if (!(h.channels >= 1 && h.channels <= 56)) bad(`This Photoshop document says it has ${h.channels} channels, which none has: it is damaged.`);
  if (![1, 8, 16, 32].includes(h.depth)) bad(`This Photoshop document says its colours are ${h.depth}-bit, which none is: it is damaged.`);
  if (!MODES[h.mode]) bad('This Photoshop document is in a colour mode Rutba Office does not read.');
  return { ...h, r };
}

export function psdInfo(b) {
  const h = header(b);
  return { format: 'psd', width: h.width, height: h.height };
}

/** The sections after the header: where each begins. */
function sections(b, h) {
  const { r, big } = h;
  let p = 26;
  const modeLen = r.u32(p);
  const modeAt = p + 4;
  p = modeAt + modeLen;
  const resLen = r.u32(p);
  const resAt = p + 4;
  p = resAt + resLen;
  const layerLen = big ? r.u64(p) : r.u32(p);
  const layerAt = p + (big ? 8 : 4);
  p = layerAt + layerLen;
  r.need(p, 2);
  return { modeAt, modeLen, resAt, resLen, layerAt, layerLen, imageAt: p };
}

/** Image resources by id: { at, len }. */
function resources(b, r, at, len) {
  const out = new Map();
  let p = at;
  const end = at + len;
  while (p + 12 <= end) {
    const sig = String.fromCharCode(b[p], b[p + 1], b[p + 2], b[p + 3]);
    if (sig !== '8BIM' && sig !== 'MeSa' && sig !== 'AgHg' && sig !== 'PHUT' && sig !== 'DCSR') break;
    const id = r.u16(p + 4);
    const nameLen = b[p + 6];
    let q = p + 6 + 1 + nameLen;
    if ((1 + nameLen) & 1) q++; // the name is padded to an even length
    const size = r.u32(q);
    q += 4;
    if (q + size > end) break;
    out.set(id, { at: q, len: size });
    p = q + size + (size & 1);
  }
  return out;
}

export function decodePsd(b, { limits } = {}) {
  const h = header(b);
  const { r, big, channels, width, height, depth, mode } = h;
  checkSize(width, height, limits, 'Photoshop');
  const s = sections(b, h);

  // Saved without its flattened picture: what is there is blank.
  const res = resources(b, r, s.resAt, s.resLen);
  const version = res.get(1057);
  if (version && version.len >= 5 && b[version.at + 4] === 0) {
    bad('This Photoshop document was saved without its flattened picture (Maximize Compatibility was off), so Rutba Office cannot show it. Save it again from Photoshop with that option on.');
  }

  // A negative layer count: the first alpha channel is the flattened picture's transparency.
  let transparent = false;
  if (s.layerLen >= (big ? 10 : 6)) {
    const infoLen = big ? r.u64(s.layerAt) : r.u32(s.layerAt);
    if (infoLen >= 2) transparent = r.i16(s.layerAt + (big ? 8 : 4)) < 0;
  }

  const compression = r.u16(s.imageAt);
  if (compression > 3) bad('This Photoshop document\'s flattened picture is packed in a way Rutba Office does not know: it may be damaged.');
  const rowBytes = depth === 1 ? Math.ceil(width / 8) : width * (depth / 8);
  const planeBytes = rowBytes * height;
  // Only the channels the picture needs are unpacked.
  const colourChannels = { bitmap: 1, grey: 1, indexed: 1, rgb: 3, cmyk: 4, multichannel: 1, duotone: 1, lab: 3 }[MODES[mode]];
  if (channels < colourChannels) bad('This Photoshop document has fewer channels than its colours need: it is damaged.');
  // A document with no layers section at all (a flat file other programs write) keeps its alpha the same way;
  // where there are layers and their count is positive, an extra channel is a saved selection, not shown.
  const alphaChannel = (transparent || s.layerLen === 0) && channels > colourChannels && MODES[mode] !== 'bitmap' && MODES[mode] !== 'indexed' ? colourChannels : -1;
  const want = Math.max(colourChannels, alphaChannel + 1);
  const planes = [];
  let p = s.imageAt + 2;
  if (compression === 1) {
    // Every row's packed length first, for every channel; then the rows.
    const countBytes = big ? 4 : 2;
    const counts = new Array(channels * height);
    for (let i = 0; i < channels * height; i++) counts[i] = big ? r.u32(p + i * 4) : r.u16(p + i * 2);
    p += channels * height * countBytes;
    for (let c = 0; c < want; c++) {
      const plane = new Uint8Array(planeBytes);
      for (let y = 0; y < height; y++) {
        const n = counts[c * height + y];
        if (p >= b.length) break;
        plane.set(packbitsDecode(b.subarray(p, Math.min(b.length, p + n)), rowBytes), y * rowBytes);
        p += n;
      }
      planes.push(plane);
    }
  } else if (compression === 0) {
    for (let c = 0; c < want; c++) {
      const plane = new Uint8Array(planeBytes);
      plane.set(b.subarray(p, Math.min(b.length, p + planeBytes)));
      p += planeBytes;
      planes.push(plane);
    }
  } else {
    // Zipped: the whole of the picture's channels in one stream.
    const all = inflate(b.subarray(p), { size: planeBytes * channels, max: planeBytes * channels });
    for (let c = 0; c < want; c++) {
      const plane = new Uint8Array(planeBytes);
      plane.set(all.subarray(c * planeBytes, Math.min(all.length, (c + 1) * planeBytes)));
      if (compression === 3) unpredict(plane, width, height, depth);
      planes.push(plane);
    }
  }

  // A sample as 0..255.
  const dv = depth === 32 ? planes.map((pl) => new DataView(pl.buffer, pl.byteOffset, pl.byteLength)) : null;
  const level = (c, i) => {
    if (depth === 8) return planes[c][i];
    if (depth === 16) return planes[c][i * 2];
    // 32-bit documents hold linear light: shown with sRGB's curve.
    const v = dv[c].getFloat32(i * 4, false);
    if (!(v > 0)) return 0;
    const g = v <= 0.0031308 ? 12.92 * v : 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
    return Math.min(255, Math.round(g * 255));
  };
  const palette = MODES[mode] === 'indexed' ? b.subarray(s.modeAt, s.modeAt + Math.min(768, s.modeLen)) : null;
  const out = rgbaOf(width, height);
  for (let y = 0, i = 0; y < height; y++) {
    for (let x = 0; x < width; x++, i++) {
      const o = i * 4;
      switch (MODES[mode]) {
        case 'bitmap': {
          const v = (planes[0][y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1 ? 0 : 255;
          out[o] = out[o + 1] = out[o + 2] = v;
          break;
        }
        case 'indexed': {
          const k = planes[0][i];
          out[o] = palette?.[k] ?? 0;
          out[o + 1] = palette?.[256 + k] ?? 0;
          out[o + 2] = palette?.[512 + k] ?? 0;
          break;
        }
        case 'rgb':
          out[o] = level(0, i);
          out[o + 1] = level(1, i);
          out[o + 2] = level(2, i);
          break;
        case 'cmyk': {
          // Photoshop keeps ink inverted: 255 is no ink.
          const k = level(3, i);
          out[o] = Math.round((level(0, i) * k) / 255);
          out[o + 1] = Math.round((level(1, i) * k) / 255);
          out[o + 2] = Math.round((level(2, i) * k) / 255);
          break;
        }
        case 'lab': {
          const [rr, gg, bb] = labToRgb((level(0, i) * 100) / 255, level(1, i) - 128, level(2, i) - 128);
          out[o] = rr;
          out[o + 1] = gg;
          out[o + 2] = bb;
          break;
        }
        default:
          out[o] = out[o + 1] = out[o + 2] = level(0, i);
      }
      if (alphaChannel >= 0) {
        const a = level(alphaChannel, i);
        out[o + 3] = a;
        // The flattened picture is matted on white where it is see-through: the white taken back out.
        if (a > 0 && a < 255) {
          for (let c = 0; c < 3; c++) out[o + c] = Math.max(0, Math.min(255, Math.round(((out[o + c] - (255 - a)) * 255) / a)));
        }
      }
    }
  }
  return { width, height, rgba: out, pages: 1 };
}

/** ZIP with prediction: each row's samples as differences from the one before. */
function unpredict(plane, width, height, depth) {
  if (depth === 8) {
    for (let y = 0; y < height; y++) for (let x = 1; x < width; x++) plane[y * width + x] = (plane[y * width + x] + plane[y * width + x - 1]) & 255;
  } else if (depth === 16) {
    for (let y = 0; y < height; y++) {
      for (let x = 1; x < width; x++) {
        const a = (y * width + x) * 2;
        const v = (((plane[a] << 8) | plane[a + 1]) + ((plane[a - 2] << 8) | plane[a - 1])) & 0xffff;
        plane[a] = v >> 8;
        plane[a + 1] = v & 255;
      }
    }
  } else if (depth === 32) {
    // Bytes differenced across the row, laid out a byte-plane at a time.
    const row = width * 4;
    const tmp = new Uint8Array(row);
    for (let y = 0; y < height; y++) {
      const at = y * row;
      for (let i = 1; i < row; i++) plane[at + i] = (plane[at + i] + plane[at + i - 1]) & 255;
      tmp.set(plane.subarray(at, at + row));
      for (let x = 0; x < width; x++) for (let k = 0; k < 4; k++) plane[at + x * 4 + k] = tmp[k * width + x];
    }
  }
}
