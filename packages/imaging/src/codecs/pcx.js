// ZSoft PCX: a 128-byte header, then each row as one or more planes of
// bytes, run-length packed (a byte of 0xC0 or more is a count; the next byte
// is repeated that many times). 1, 2, 4 or 8 bits a pixel in 1 to 4 planes:
// black and white, CGA and EGA palettes, 256 colours with the palette at the
// end, and 24-bit colour (and alpha) as three or four planes.

import { bad, reader, checkSize, rgbaOf } from './bytes.js';

export function isPcx(b) {
  return b.length >= 128 && b[0] === 0x0a && [0, 2, 3, 4, 5].includes(b[1]) && b[2] <= 1 && [1, 2, 4, 8].includes(b[3]) && b[65] >= 1 && b[65] <= 4;
}

function header(b) {
  if (!isPcx(b)) bad('This is not a PCX picture.');
  const r = reader(b, 'PCX');
  const xmin = r.u16(4);
  const ymin = r.u16(6);
  const xmax = r.u16(8);
  const ymax = r.u16(10);
  const h = { bits: b[3], planes: b[65], line: r.u16(66), packed: b[2] === 1, width: xmax - xmin + 1, height: ymax - ymin + 1 };
  if (h.width < 1 || h.height < 1) bad(`This PCX picture says it is ${h.width} × ${h.height} pixels, which no picture is: it is damaged.`);
  if (h.line < Math.ceil((h.width * h.bits) / 8) || h.line > 65535) bad('This PCX picture\'s rows are damaged.');
  const kinds = { '1:1': 1, '1:2': 1, '1:3': 1, '1:4': 1, '2:1': 1, '4:1': 1, '8:1': 1, '8:3': 1, '8:4': 1 };
  if (!kinds[`${h.bits}:${h.planes}`]) bad(`This PCX picture has ${h.planes} planes of ${h.bits} bits, which Rutba Office does not read.`);
  return h;
}

export function pcxInfo(b) {
  const h = header(b);
  return { format: 'pcx', width: h.width, height: h.height };
}

export function decodePcx(b, { limits } = {}) {
  const h = header(b);
  const { width, height, bits, planes, line } = h;
  checkSize(width, height, limits, 'PCX');
  const rowBytes = line * planes;
  const data = new Uint8Array(rowBytes * height);
  // 256 colours: the palette is the last 768 bytes, after a 12.
  const vga = bits === 8 && planes === 1 && b.length >= 128 + 769 && b[b.length - 769] === 0x0c;
  const end = vga ? b.length - 769 : b.length;
  let p = 128;
  if (h.packed) {
    let o = 0;
    while (o < data.length && p < end) {
      const c = b[p++];
      if (c >= 0xc0) {
        if (p >= end) break;
        const v = b[p++];
        const n = Math.min(c & 0x3f, data.length - o);
        data.fill(v, o, o + n);
        o += n;
      } else data[o++] = c;
    }
  } else {
    data.set(b.subarray(p, Math.min(end, p + data.length)));
  }
  const out = rgbaOf(width, height);
  // The 16-colour palette in the header (EGA), or CGA's, or black and white.
  const ega = (k) => [b[16 + k * 3], b[17 + k * 3], b[18 + k * 3]];
  for (let y = 0; y < height; y++) {
    const row = y * rowBytes;
    for (let x = 0; x < width; x++) {
      const o = (y * width + x) * 4;
      if (bits === 8 && planes >= 3) {
        out[o] = data[row + x];
        out[o + 1] = data[row + line + x];
        out[o + 2] = data[row + 2 * line + x];
        if (planes === 4) out[o + 3] = data[row + 3 * line + x];
      } else if (bits === 8) {
        const k = data[row + x];
        if (vga) {
          const at = b.length - 768 + k * 3;
          out[o] = b[at]; out[o + 1] = b[at + 1]; out[o + 2] = b[at + 2];
        } else out[o] = out[o + 1] = out[o + 2] = k;
      } else if (planes === 1) {
        const per = 8 / bits;
        const v = (data[row + Math.floor(x / per)] >> ((per - 1 - (x % per)) * bits)) & ((1 << bits) - 1);
        if (bits === 1) out[o] = out[o + 1] = out[o + 2] = v ? 255 : 0;
        else {
          const [r, g, bl] = ega(v);
          out[o] = r; out[o + 1] = g; out[o + 2] = bl;
        }
      } else {
        // 1 bit in 2 to 4 planes: the planes' bits make the index into the header's palette.
        let k = 0;
        for (let pl = 0; pl < planes; pl++) k |= ((data[row + pl * line + (x >> 3)] >> (7 - (x & 7))) & 1) << pl;
        const [r, g, bl] = ega(k);
        out[o] = r; out[o + 1] = g; out[o + 2] = bl;
      }
    }
  }
  return { width, height, rgba: out, pages: 1 };
}

/** RGB(A) as a 24-bit (or 32-bit) PCX, packed: for the tests and the fuzzer's good files. */
export function encodePcx(rgba, width, height, { alpha = false } = {}) {
  const planes = alpha ? 4 : 3;
  const line = width + (width & 1);
  const head = new Uint8Array(128);
  const v = new DataView(head.buffer);
  head[0] = 0x0a; head[1] = 5; head[2] = 1; head[3] = 8;
  v.setUint16(8, width - 1, true);
  v.setUint16(10, height - 1, true);
  v.setUint16(12, 72, true);
  v.setUint16(14, 72, true);
  head[65] = planes;
  v.setUint16(66, line, true);
  head[68] = 1;
  const body = [];
  for (let y = 0; y < height; y++) {
    const row = new Uint8Array(line * planes);
    for (let pl = 0; pl < planes; pl++) for (let x = 0; x < width; x++) row[pl * line + x] = rgba[(y * width + x) * 4 + pl];
    for (let i = 0; i < row.length;) {
      let n = 1;
      while (i + n < row.length && n < 63 && row[i + n] === row[i]) n++;
      if (n > 1 || row[i] >= 0xc0) body.push(0xc0 | n, row[i]);
      else body.push(row[i]);
      i += n;
    }
  }
  const out = new Uint8Array(128 + body.length);
  out.set(head);
  out.set(body, 128);
  return out;
}
