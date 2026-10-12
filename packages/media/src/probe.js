// What a media file is, from its container.
//
// Enough to draw a file list, decide whether the built-in player will manage
// it, and read back an export: the length, the picture's size, and which
// tracks there are (a picture, a sound), read from the bytes rather than by
// decoding. MP4 and QuickTime carry a movie header with the duration and a
// track box per track; a file written in fragments (as a recorder writes
// MP4) carries its length only in the fragments, so those are added up.
// WebM and Matroska carry theirs in EBML; a recorder's WebM has no Duration
// element at all, so the length is the last block's time plus one frame.

const dec = (b, at, len) => new TextDecoder('latin1').decode(b.subarray(at, at + len));
const u32 = (b, at) => ((b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3]) >>> 0;
const u64 = (b, at) => u32(b, at) * 2 ** 32 + u32(b, at + 4);

/**
 * `complete`: the bytes are the whole file, so a length that has to be
 * added up from its fragments or its last block can be trusted. A head or a
 * tail alone gives a length only where the file states one.
 *
 * @returns {{ format: string, duration: number|null, width?: number, height?: number,
 *   tracks?: Array<{ type: 'video'|'audio'|'other', codec: string|null }>, hasVideo?: boolean, hasAudio?: boolean,
 *   sampleRate?: number }|null}
 */
export function probeMedia(bytes, { complete = false } = {}) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (b.length < 16) return null;

  if (dec(b, 4, 4) === 'ftyp') return probeMp4(b, complete);
  if (b[0] === 0x1a && b[1] === 0x45 && b[2] === 0xdf && b[3] === 0xa3) return probeMatroska(b, complete);

  if (dec(b, 0, 4) === 'RIFF' && dec(b, 8, 4) === 'WAVE') {
    const rate = b[24] | (b[25] << 8) | (b[26] << 16) | (b[27] << 24);
    const dataAt = dec(b, 0, Math.min(b.length, 4096)).indexOf('data');
    const size = dataAt > 0 ? (b[dataAt + 4] | (b[dataAt + 5] << 8) | (b[dataAt + 6] << 16) | (b[dataAt + 7] << 24)) : 0;
    const byteRate = b[28] | (b[29] << 8) | (b[30] << 16) | (b[31] << 24);
    return { format: 'wav', duration: byteRate ? size / byteRate : null, sampleRate: rate, tracks: [{ type: 'audio', codec: 'pcm' }], hasVideo: false, hasAudio: true };
  }

  if (dec(b, 0, 4) === 'OggS') return { format: 'ogg', duration: null };
  if (dec(b, 0, 4) === 'fLaC') return { format: 'flac', duration: null };
  if (dec(b, 0, 3) === 'ID3' || (b[0] === 0xff && (b[1] & 0xe0) === 0xe0)) return { format: 'mp3', duration: null };
  return null;
}

/* ── MP4 and QuickTime ────────────────────────────────────────────────── */

const MP4_CONTAINERS = new Set(['moov', 'trak', 'mdia', 'minf', 'stbl', 'moof', 'traf', 'mvex', 'edts']);

function probeMp4(b, complete) {
  const brand = dec(b, 8, 4);
  const format = brand.startsWith('qt') ? 'mov' : brand.startsWith('M4A') ? 'm4a' : 'mp4';
  let movie = null; // { timescale, duration }
  let fragmentDuration = null;
  const tracks = new Map(); // track id -> { type, codec, timescale, width, height }
  let current = null; // the trak being read
  const fragmentEnds = new Map(); // track id -> furthest decode time reached
  let traf = null; // { id, base, defaultDuration }

  const walk = (start, end, depth) => {
    let at = start;
    while (at + 8 <= end) {
      let size = u32(b, at);
      const type = dec(b, at + 4, 4);
      let head = 8;
      if (size === 1) {
        if (at + 16 > end) return;
        size = u64(b, at + 8);
        head = 16;
      } else if (size === 0) {
        size = end - at;
      }
      if (size < head) return;
      const body = at + head;
      const stop = Math.min(at + size, b.length);
      if (type === 'trak') {
        current = { type: 'other', codec: null, timescale: null, width: null, height: null, id: null };
        walk(body, stop, depth + 1);
        if (current.id !== null) tracks.set(current.id, current);
        current = null;
      } else if (type === 'traf') {
        traf = { id: null, base: 0, defaultDuration: 0 };
        walk(body, stop, depth + 1);
        traf = null;
      } else if (MP4_CONTAINERS.has(type)) {
        walk(body, stop, depth + 1);
      } else if (type === 'mvhd' && body + 20 <= b.length) {
        const v = b[body];
        const ts = v === 1 ? u32(b, body + 20) : u32(b, body + 12);
        const du = v === 1 ? u64(b, body + 24) : u32(b, body + 16);
        movie = { timescale: ts, duration: du };
      } else if (type === 'mehd' && body + 8 <= b.length) {
        fragmentDuration = b[body] === 1 ? u64(b, body + 4) : u32(b, body + 4);
      } else if (type === 'tkhd' && current && body + 84 <= b.length) {
        const v = b[body];
        current.id = v === 1 ? u32(b, body + 20) : u32(b, body + 12);
        const wAt = body + (v === 1 ? 88 : 76);
        if (wAt + 8 <= b.length) {
          current.width = u32(b, wAt) / 65536 || null;
          current.height = u32(b, wAt + 4) / 65536 || null;
        }
      } else if (type === 'mdhd' && current && body + 24 <= b.length) {
        current.timescale = b[body] === 1 ? u32(b, body + 20) : u32(b, body + 12);
      } else if (type === 'hdlr' && current && body + 12 <= b.length) {
        const h = dec(b, body + 8, 4);
        current.type = h === 'vide' ? 'video' : h === 'soun' ? 'audio' : 'other';
      } else if (type === 'stsd' && current && body + 16 <= b.length) {
        current.codec = dec(b, body + 12, 4).trim() || null;
      } else if (type === 'tfhd' && traf && body + 8 <= b.length) {
        const flags = (b[body + 1] << 16) | (b[body + 2] << 8) | b[body + 3];
        traf.id = u32(b, body + 4);
        let p = body + 8;
        if (flags & 0x01) p += 8; // base data offset
        if (flags & 0x02) p += 4; // sample description index
        if (flags & 0x08 && p + 4 <= b.length) traf.defaultDuration = u32(b, p);
      } else if (type === 'tfdt' && traf && body + 8 <= b.length) {
        traf.base = b[body] === 1 ? u64(b, body + 4) : u32(b, body + 4);
      } else if (type === 'trun' && traf && body + 8 <= b.length) {
        const flags = (b[body + 1] << 16) | (b[body + 2] << 8) | b[body + 3];
        const count = u32(b, body + 4);
        let p = body + 8;
        if (flags & 0x001) p += 4;
        if (flags & 0x004) p += 4;
        const per = ((flags & 0x100) ? 4 : 0) + ((flags & 0x200) ? 4 : 0) + ((flags & 0x400) ? 4 : 0) + ((flags & 0x800) ? 4 : 0);
        let total = 0;
        for (let i = 0; i < count && p + per <= stop; i++) {
          total += flags & 0x100 ? u32(b, p) : traf.defaultDuration;
          p += per;
        }
        const reached = traf.base + total;
        if (traf.id !== null) fragmentEnds.set(traf.id, Math.max(fragmentEnds.get(traf.id) || 0, reached));
        traf.base = reached;
      }
      if (size <= 0) return;
      at += size;
    }
  };
  walk(0, b.length, 0);

  const list = [...tracks.values()];
  let duration = movie && movie.timescale && movie.duration ? movie.duration / movie.timescale : null;
  if (!duration && fragmentDuration && movie?.timescale) duration = fragmentDuration / movie.timescale;
  // Fragments add up to the length only when every one of them is here.
  if (!duration && complete && fragmentEnds.size) {
    let longest = 0;
    for (const [id, end] of fragmentEnds) {
      const ts = tracks.get(id)?.timescale;
      if (ts) longest = Math.max(longest, end / ts);
    }
    duration = longest || null;
  }
  const video = list.find((t) => t.type === 'video');
  return {
    format,
    duration,
    ...(video?.width ? { width: video.width, height: video.height } : {}),
    tracks: list.map((t) => ({ type: t.type, codec: t.codec })),
    hasVideo: list.some((t) => t.type === 'video'),
    hasAudio: list.some((t) => t.type === 'audio'),
  };
}

/* ── WebM and Matroska ────────────────────────────────────────────────── */

const EBML_MASTERS = new Set([
  0x1a45dfa3, // the EBML header, which holds the DocType
  0x18538067, // Segment
  0x1549a966, // Info
  0x1654ae6b, // Tracks
  0xae, // TrackEntry
  0xe0, // Video
  0xe1, // Audio
  0x1f43b675, // Cluster
  0xa0, // BlockGroup
]);

/** An EBML id at `at`: { id, len } or null. */
function ebmlId(b, at) {
  const first = b[at];
  if (first === undefined) return null;
  const len = first & 0x80 ? 1 : first & 0x40 ? 2 : first & 0x20 ? 3 : first & 0x10 ? 4 : 0;
  if (!len || at + len > b.length) return null;
  let id = 0;
  for (let i = 0; i < len; i++) id = id * 256 + b[at + i];
  return { id, len };
}

/** An EBML size at `at`: { size, len, unknown } or null. */
function ebmlSize(b, at) {
  const first = b[at];
  if (first === undefined) return null;
  let len = 1;
  while (len <= 8 && !(first & (0x80 >> (len - 1)))) len++;
  if (len > 8 || at + len > b.length) return null;
  let size = first & (0xff >> len);
  let allOnes = size === (0xff >> len);
  for (let i = 1; i < len; i++) {
    size = size * 256 + b[at + i];
    if (b[at + i] !== 0xff) allOnes = false;
  }
  return { size, len, unknown: allOnes };
}

const ebmlUint = (b, at, size) => {
  let n = 0;
  for (let i = 0; i < size; i++) n = n * 256 + b[at + i];
  return n;
};
const ebmlFloat = (b, at, size) => {
  const v = new DataView(b.buffer, b.byteOffset + at, size);
  return size === 4 ? v.getFloat32(0) : size === 8 ? v.getFloat64(0) : null;
};

function probeMatroska(b, complete) {
  let format = 'mkv';
  let scale = 1e6; // nanoseconds per timecode tick
  let declared = null;
  const tracks = [];
  let track = null;
  let clusterTime = 0;
  const lastBlock = new Map(); // track number -> { t, step }
  let at = 0;
  while (at < b.length) {
    const id = ebmlId(b, at);
    if (!id) break;
    const sz = ebmlSize(b, at + id.len);
    if (!sz) break;
    const body = at + id.len + sz.len;
    if (EBML_MASTERS.has(id.id)) {
      if (id.id === 0xae) {
        track = { number: null, type: 'other', codec: null, width: null, height: null, rate: null };
        tracks.push(track);
      }
      at = body; // read the children in place, whether or not the size is known
      continue;
    }
    if (sz.unknown || body + sz.size > b.length) break;
    const size = sz.size;
    switch (id.id) {
      case 0x4282: // DocType
        format = dec(b, body, size).replace(/\0+$/, '') === 'webm' ? 'webm' : 'mkv';
        break;
      case 0x2ad7b1:
        scale = ebmlUint(b, body, size) || 1e6;
        break;
      case 0x4489:
        declared = ebmlFloat(b, body, size);
        break;
      case 0xd7:
        if (track) track.number = ebmlUint(b, body, size);
        break;
      case 0x83:
        if (track) {
          const v = ebmlUint(b, body, size);
          track.type = v === 1 ? 'video' : v === 2 ? 'audio' : 'other';
        }
        break;
      case 0x86:
        if (track) track.codec = dec(b, body, size).replace(/\0+$/, '');
        break;
      case 0xb0:
        if (track) track.width = ebmlUint(b, body, size);
        break;
      case 0xba:
        if (track) track.height = ebmlUint(b, body, size);
        break;
      case 0xb5:
        if (track) track.rate = ebmlFloat(b, body, size);
        break;
      case 0xe7: // the cluster's timecode
        clusterTime = ebmlUint(b, body, size);
        break;
      case 0xa3: // SimpleBlock
      case 0xa1: { // Block
        const tn = ebmlSize(b, body);
        if (tn && body + tn.len + 2 <= b.length) {
          let rel = (b[body + tn.len] << 8) | b[body + tn.len + 1];
          if (rel & 0x8000) rel -= 0x10000;
          const t = clusterTime + rel;
          const prev = lastBlock.get(tn.size);
          // A frame lasts as long as the shortest gap between two seen: a recording of a still
          // window has frames only when it changes, and its last gap is no frame's length.
          const step = prev && t > prev.t ? Math.min(prev.step || Infinity, t - prev.t) : prev?.step || 0;
          if (!prev || t >= prev.t) lastBlock.set(tn.size, { t, step });
        }
        break;
      }
      default:
        break;
    }
    at = body + size;
  }

  let duration = declared && declared > 0 ? (declared * scale) / 1e9 : null;
  // The last block is the end only when the bytes are the whole file.
  if (!duration && complete && lastBlock.size) {
    let end = 0;
    for (const { t, step } of lastBlock.values()) end = Math.max(end, t + step);
    duration = end > 0 ? (end * scale) / 1e9 : null;
  }
  const video = tracks.find((t) => t.type === 'video');
  const audio = tracks.find((t) => t.type === 'audio');
  return {
    format,
    duration,
    ...(video?.width ? { width: video.width, height: video.height } : {}),
    ...(audio?.rate ? { sampleRate: audio.rate } : {}),
    tracks: tracks.map((t) => ({ type: t.type, codec: t.codec })),
    hasVideo: Boolean(video),
    hasAudio: Boolean(audio),
  };
}
