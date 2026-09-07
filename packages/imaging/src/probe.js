// What an image is, without decoding it.
//
// A picture viewer has to lay out a grid of two thousand thumbnails before it
// has decoded any of them, and it needs each one's shape to do that. Every
// container puts its dimensions within the first few hundred bytes, so reading
// those is thousands of times cheaper than decoding — and it is the difference
// between a folder that opens instantly and one that hangs.

const dec = (b, at, len) => new TextDecoder('latin1').decode(b.subarray(at, at + len));

function u16be(b, at) { return (b[at] << 8) | b[at + 1]; }
function u32be(b, at) { return ((b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3]) >>> 0; }
function u16le(b, at) { return b[at] | (b[at + 1] << 8); }
function u32le(b, at) { return (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0; }

/**
 * @param {Uint8Array} bytes at least the first 64 KB
 * @returns {{ format: string, width: number, height: number, bitDepth?: number, frames?: number, animated?: boolean } | null}
 */
export function probeImage(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (b.length < 16) return null;

  // PNG: IHDR is always the first chunk.
  if (b[0] === 0x89 && dec(b, 1, 3) === 'PNG') {
    const result = { format: 'png', width: u32be(b, 16), height: u32be(b, 20), bitDepth: b[24], colourType: b[25] };
    // APNG announces itself with an acTL chunk before the first IDAT.
    for (let at = 8; at + 8 < Math.min(b.length, 4096); ) {
      const len = u32be(b, at);
      const type = dec(b, at + 4, 4);
      if (type === 'acTL') {
        result.animated = true;
        result.frames = u32be(b, at + 8);
        break;
      }
      if (type === 'IDAT') break;
      at += 12 + len;
    }
    return result;
  }

  // JPEG: walk the markers to SOFn, which carries the size.
  if (b[0] === 0xff && b[1] === 0xd8) {
    let at = 2;
    while (at + 9 < b.length) {
      if (b[at] !== 0xff) {
        at++;
        continue;
      }
      const marker = b[at + 1];
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        at += 2;
        continue;
      }
      const length = u16be(b, at + 2);
      // SOF0..SOF15, excluding the DHT/JPG/DAC markers interleaved among them.
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { format: 'jpeg', height: u16be(b, at + 5), width: u16be(b, at + 7), components: b[at + 9], progressive: marker === 0xc2 };
      }
      at += 2 + length;
    }
    return { format: 'jpeg', width: 0, height: 0 };
  }

  if (dec(b, 0, 4) === 'GIF8') {
    return { format: 'gif', width: u16le(b, 6), height: u16le(b, 8), animated: dec(b, 0, 6) === 'GIF89a' };
  }

  if (b[0] === 0x42 && b[1] === 0x4d) {
    return { format: 'bmp', width: u32le(b, 18), height: Math.abs(u32le(b, 22) | 0), bitDepth: u16le(b, 28) };
  }

  if (dec(b, 0, 4) === 'RIFF' && dec(b, 8, 4) === 'WEBP') {
    const chunk = dec(b, 12, 4);
    if (chunk === 'VP8X') {
      return {
        format: 'webp',
        width: 1 + (b[24] | (b[25] << 8) | (b[26] << 16)),
        height: 1 + (b[27] | (b[28] << 8) | (b[29] << 16)),
        animated: Boolean(b[20] & 0x02),
      };
    }
    if (chunk === 'VP8 ') return { format: 'webp', width: u16le(b, 26) & 0x3fff, height: u16le(b, 28) & 0x3fff };
    if (chunk === 'VP8L') {
      const bits = b[21] | (b[22] << 8) | (b[23] << 16) | (b[24] << 24);
      return { format: 'webp', width: (bits & 0x3fff) + 1, height: ((bits >> 14) & 0x3fff) + 1, lossless: true };
    }
    return { format: 'webp', width: 0, height: 0 };
  }

  // TIFF, and the raw formats built on it.
  const little = b[0] === 0x49 && b[1] === 0x49;
  if ((little || (b[0] === 0x4d && b[1] === 0x4d)) && (little ? u16le(b, 2) : u16be(b, 2)) === 42) {
    const u16 = little ? u16le : u16be;
    const u32 = little ? u32le : u32be;
    const ifd = u32(b, 4);
    if (ifd + 2 > b.length) return { format: 'tiff', width: 0, height: 0 };
    const count = u16(b, ifd);
    let width = 0;
    let height = 0;
    for (let i = 0; i < count; i++) {
      const entry = ifd + 2 + i * 12;
      if (entry + 12 > b.length) break;
      const tag = u16(b, entry);
      const type = u16(b, entry + 2);
      const value = type === 3 ? u16(b, entry + 8) : u32(b, entry + 8);
      if (tag === 0x0100) width = value;
      if (tag === 0x0101) height = value;
    }
    return { format: 'tiff', width, height };
  }

  if (dec(b, 4, 4) === 'ftyp') {
    const brand = dec(b, 8, 4);
    const format = /avif|avis/.test(brand) ? 'avif' : 'heic';
    // ISO-BMFF images carry the size in an ispe box.
    for (let at = 0; at + 12 < Math.min(b.length, 65536); at++) {
      if (dec(b, at, 4) === 'ispe') {
        return { format, width: u32be(b, at + 8), height: u32be(b, at + 12) };
      }
    }
    return { format, width: 0, height: 0 };
  }

  if (b[0] === 0 && b[1] === 0 && b[2] === 1 && b[3] === 0) {
    return { format: 'ico', width: b[6] || 256, height: b[7] || 256, frames: u16le(b, 4) };
  }

  if (dec(b, 0, 4) === '8BPS') {
    return { format: 'psd', width: u32be(b, 18), height: u32be(b, 14) };
  }

  const text = dec(b, 0, Math.min(b.length, 1024));
  if (/<svg[\s>]/i.test(text)) {
    const w = /\bwidth\s*=\s*["']?([\d.]+)/i.exec(text);
    const h = /\bheight\s*=\s*["']?([\d.]+)/i.exec(text);
    const box = /viewBox\s*=\s*["']\s*[\d.-]+\s+[\d.-]+\s+([\d.]+)\s+([\d.]+)/i.exec(text);
    return {
      format: 'svg',
      width: Math.round(Number(w?.[1] || box?.[1] || 0)),
      height: Math.round(Number(h?.[1] || box?.[2] || 0)),
      vector: true,
    };
  }

  return null;
}

/** Megapixels, rounded the way a person would say it. */
export function megapixels(width, height) {
  const mp = (width * height) / 1e6;
  return mp >= 10 ? Math.round(mp) : Math.round(mp * 10) / 10;
}

/** The nearest common aspect-ratio name, for the crop tool. */
export function aspectName(width, height) {
  if (!width || !height) return '';
  const ratio = width / height;
  const known = [
    ['1:1', 1], ['4:3', 4 / 3], ['3:2', 1.5], ['16:10', 1.6], ['16:9', 16 / 9], ['21:9', 21 / 9],
    ['3:4', 0.75], ['2:3', 2 / 3], ['9:16', 9 / 16],
  ];
  let best = '';
  let closest = Infinity;
  for (const [name, value] of known) {
    const d = Math.abs(ratio - value);
    if (d < closest) {
      closest = d;
      best = name;
    }
  }
  return closest / ratio < 0.02 ? best : `${Math.round(ratio * 100) / 100}:1`;
}
