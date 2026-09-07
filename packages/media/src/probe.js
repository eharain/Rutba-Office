// What a media file is, from its container.
//
// Enough to draw a file list and decide whether the built-in player will
// manage it, read from the first megabyte rather than by decoding: MP4 and
// QuickTime carry a movie header with the duration and timescale, WebM and
// Matroska carry theirs in EBML.

const dec = (b, at, len) => new TextDecoder('latin1').decode(b.subarray(at, at + len));
const u32 = (b, at) => ((b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3]) >>> 0;

/** @returns {{ format: string, duration: number|null, width?: number, height?: number }|null} */
export function probeMedia(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (b.length < 16) return null;

  if (dec(b, 4, 4) === 'ftyp') {
    const brand = dec(b, 8, 4);
    const format = brand.startsWith('qt') ? 'mov' : brand.startsWith('M4A') ? 'm4a' : 'mp4';
    // mvhd holds the timescale and the duration, wherever moov happens to sit.
    for (let at = 0; at + 32 < Math.min(b.length, 4 * 1024 * 1024); at++) {
      if (dec(b, at, 4) !== 'mvhd') continue;
      const version = b[at + 4];
      const base = at + 8 + (version === 1 ? 16 : 8);
      const timescale = u32(b, base);
      const duration = version === 1 ? u32(b, base + 8) : u32(b, base + 4);
      return { format, duration: timescale ? duration / timescale : null };
    }
    return { format, duration: null };
  }

  if (b[0] === 0x1a && b[1] === 0x45 && b[2] === 0xdf && b[3] === 0xa3) {
    const head = dec(b, 0, Math.min(b.length, 4096));
    return { format: head.includes('webm') ? 'webm' : 'mkv', duration: null };
  }

  if (dec(b, 0, 4) === 'RIFF' && dec(b, 8, 4) === 'WAVE') {
    const rate = b[24] | (b[25] << 8) | (b[26] << 16) | (b[27] << 24);
    const dataAt = dec(b, 0, Math.min(b.length, 4096)).indexOf('data');
    const size = dataAt > 0 ? (b[dataAt + 4] | (b[dataAt + 5] << 8) | (b[dataAt + 6] << 16) | (b[dataAt + 7] << 24)) : 0;
    const byteRate = b[28] | (b[29] << 8) | (b[30] << 16) | (b[31] << 24);
    return { format: 'wav', duration: byteRate ? size / byteRate : null, sampleRate: rate };
  }

  if (dec(b, 0, 4) === 'OggS') return { format: 'ogg', duration: null };
  if (dec(b, 0, 4) === 'fLaC') return { format: 'flac', duration: null };
  if (dec(b, 0, 3) === 'ID3' || (b[0] === 0xff && (b[1] & 0xe0) === 0xe0)) return { format: 'mp3', duration: null };
  return null;
}
