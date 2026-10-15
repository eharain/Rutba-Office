// A WebM or Matroska file read down to its frames.
//
// For the export's decoders: each track's codec as WebCodecs names it, its
// configuration and size, and every frame (its bytes' place, its time in
// seconds, whether it is a key frame). A recorder's WebM is written live,
// so its Segment and Clusters have no size; the file is read element by
// element in order, descending into the masters that hold frames, which
// works whether a size is known or not.

const ID = {
  EBML: 0x1a45dfa3,
  Segment: 0x18538067,
  Info: 0x1549a966,
  TimecodeScale: 0x2ad7b1,
  Duration: 0x4489,
  Tracks: 0x1654ae6b,
  TrackEntry: 0xae,
  TrackNumber: 0xd7,
  TrackType: 0x83,
  CodecID: 0x86,
  CodecPrivate: 0x63a2,
  DefaultDuration: 0x23e383,
  CodecDelay: 0x56aa,
  Video: 0xe0,
  PixelWidth: 0xb0,
  PixelHeight: 0xba,
  Audio: 0xe1,
  SamplingFrequency: 0xb5,
  Channels: 0x9f,
  Cluster: 0x1f43b675,
  Timecode: 0xe7,
  SimpleBlock: 0xa3,
  BlockGroup: 0xa0,
  Block: 0xa1,
  ReferenceBlock: 0xfb,
  BlockDuration: 0x9b,
};

// Masters read inside rather than skipped: they hold what is wanted.
const DESCEND = new Set([ID.Segment, ID.Info, ID.Tracks, ID.TrackEntry, ID.Video, ID.Audio, ID.Cluster, ID.BlockGroup]);

/** A variable-length integer at `at`: { value, length, unknown }. `keepMarker` for element IDs. */
function vint(b, at, keepMarker = false) {
  const first = b[at];
  if (first === undefined || first === 0) return null;
  let length = 1;
  while (length <= 8 && !(first & (0x80 >> (length - 1)))) length++;
  if (length > 8 || at + length > b.length) return null;
  let value = keepMarker ? first : first & (0xff >> length);
  let allOnes = value === (0xff >> length);
  for (let i = 1; i < length; i++) {
    value = value * 256 + b[at + i];
    if (b[at + i] !== 0xff) allOnes = false;
  }
  return { value, length, unknown: !keepMarker && allOnes };
}

const uint = (b, at, len) => {
  let v = 0;
  for (let i = 0; i < len; i++) v = v * 256 + b[at + i];
  return v;
};
const float = (b, at, len) => {
  const dv = new DataView(b.buffer, b.byteOffset + at, len);
  return len === 4 ? dv.getFloat32(0) : len === 8 ? dv.getFloat64(0) : 0;
};
const text = (b, at, len) => new TextDecoder('latin1').decode(b.subarray(at, at + len)).replace(/\0+$/, '');

/** The codec string WebCodecs takes for a Matroska codec ID and its private data. */
function codecFor(id, priv) {
  switch (id) {
    case 'V_VP8': return 'vp8';
    case 'V_VP9': return 'vp09.00.10.08';
    case 'V_AV1': {
      if (priv && priv.length >= 4) {
        const profile = priv[1] >> 5;
        const level = priv[1] & 0x1f;
        const tier = priv[2] >> 7 ? 'H' : 'M';
        const depth = (priv[2] >> 6) & 1 ? ((priv[2] >> 5) & 1 ? 12 : 10) : 8;
        return `av01.${profile}.${String(level).padStart(2, '0')}${tier}.${String(depth).padStart(2, '0')}`;
      }
      return 'av01.0.04M.08';
    }
    case 'V_MPEG4/ISO/AVC': {
      if (priv && priv.length >= 4) return `avc1.${[1, 2, 3].map((i) => priv[i].toString(16).padStart(2, '0')).join('')}`;
      return 'avc1.42001f';
    }
    case 'A_OPUS': return 'opus';
    case 'A_VORBIS': return 'vorbis';
    case 'A_AAC': return 'mp4a.40.2';
    default: return null;
  }
}

/**
 * Read a WebM. `bytes` is the whole file.
 *
 * @returns {{ duration: number|null, tracks: Array<{ number, kind, codecId, codec, description,
 *   width, height, sampleRate, channels, frames: Array<{ offset, size, time, key }> }> }|null}
 */
export function readWebm(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (!(b[0] === 0x1a && b[1] === 0x45 && b[2] === 0xdf && b[3] === 0xa3)) return null;
  let scale = 1000000; // nanoseconds a tick
  let duration = null;
  const tracks = [];
  let track = null;
  let clusterTime = 0;
  let blockKeyless = false; // a BlockGroup's Block is a key frame unless a ReferenceBlock says otherwise
  let groupFrames = null;

  const addBlock = (at, size, simple) => {
    const num = vint(b, at);
    if (!num) return [];
    const t = tracks.find((x) => x.number === num.value);
    let p = at + num.length;
    const rel = ((b[p] << 8) | b[p + 1]) << 16 >> 16;
    const flags = b[p + 2];
    p += 3;
    const end = at + size;
    const time = ((clusterTime + rel) * scale) / 1e9;
    const key = simple ? Boolean(flags & 0x80) : true;
    const lacing = (flags >> 1) & 3;
    const parts = [];
    if (lacing === 0) {
      parts.push([p, end - p]);
    } else {
      const count = b[p] + 1;
      p += 1;
      const sizes = [];
      if (lacing === 1) { // Xiph
        for (let i = 0; i < count - 1; i++) {
          let s = 0;
          let c;
          do { c = b[p++]; s += c; } while (c === 255);
          sizes.push(s);
        }
      } else if (lacing === 3) { // EBML
        const first = vint(b, p);
        p += first.length;
        sizes.push(first.value);
        for (let i = 1; i < count - 1; i++) {
          const d = vint(b, p);
          p += d.length;
          const bias = 2 ** (7 * d.length - 1) - 1;
          sizes.push(sizes[i - 1] + (d.value - bias));
        }
      }
      if (lacing === 2) { // fixed
        const each = Math.floor((end - p) / count);
        for (let i = 0; i < count; i++) sizes.push(each);
      } else {
        sizes.push(end - p - sizes.reduce((a, c) => a + c, 0));
      }
      let q = p;
      for (const s of sizes) {
        parts.push([q, s]);
        q += s;
      }
    }
    if (!t) return [];
    // Laced frames follow each other at the track's frame length.
    const per = t.defaultDuration ? t.defaultDuration / 1e9 : 0;
    const frames = parts.map(([o, s], i) => ({ offset: o, size: s, time: time + i * per, key }));
    t.frames.push(...frames);
    return frames;
  };

  let at = 0;
  const end = b.length;
  while (at < end) {
    const id = vint(b, at, true);
    if (!id) break;
    const size = vint(b, at + id.length);
    if (!size) break;
    const body = at + id.length + size.length;
    const len = size.unknown ? null : size.value;
    const next = len === null ? null : body + len;
    if (id.value === ID.EBML) { at = next ?? end; continue; }
    if (DESCEND.has(id.value)) {
      if (id.value === ID.TrackEntry) {
        track = { number: 0, kind: 'other', codecId: null, codec: null, description: null, width: null, height: null, sampleRate: null, channels: null, defaultDuration: 0, codecDelay: 0, frames: [] };
        tracks.push(track);
      } else if (id.value === ID.BlockGroup) {
        blockKeyless = false;
        groupFrames = [];
        // A ReferenceBlock inside says the Block is not a key frame: settled when the group ends.
        const stop = next ?? end;
        let q = body;
        while (q < stop) {
          const cid = vint(b, q, true);
          const cs = cid && vint(b, q + cid.length);
          if (!cs || cs.unknown) break;
          const cb = q + cid.length + cs.length;
          if (cid.value === ID.Block) groupFrames.push(...addBlock(cb, cs.value, false));
          else if (cid.value === ID.ReferenceBlock) blockKeyless = true;
          q = cb + cs.value;
        }
        if (blockKeyless) for (const f of groupFrames) f.key = false;
        at = stop;
        continue;
      }
      at = body;
      continue;
    }
    if (len === null) break; // an unknown size on something not read inside: nothing more can be trusted
    switch (id.value) {
      case ID.TimecodeScale: scale = uint(b, body, len); break;
      case ID.Duration: duration = float(b, body, len); break;
      case ID.TrackNumber: if (track) track.number = uint(b, body, len); break;
      case ID.TrackType: if (track) { const v = uint(b, body, len); track.kind = v === 1 ? 'video' : v === 2 ? 'audio' : 'other'; } break;
      case ID.CodecID: if (track) track.codecId = text(b, body, len); break;
      case ID.CodecPrivate: if (track) track.description = b.slice(body, body + len); break;
      case ID.DefaultDuration: if (track) track.defaultDuration = uint(b, body, len); break;
      case ID.CodecDelay: if (track) track.codecDelay = uint(b, body, len); break;
      case ID.PixelWidth: if (track) track.width = uint(b, body, len); break;
      case ID.PixelHeight: if (track) track.height = uint(b, body, len); break;
      case ID.SamplingFrequency: if (track) track.sampleRate = float(b, body, len); break;
      case ID.Channels: if (track) track.channels = uint(b, body, len); break;
      case ID.Timecode: clusterTime = uint(b, body, len); break;
      case ID.SimpleBlock: addBlock(body, len, true); break;
      default: break;
    }
    at = next;
  }

  for (const t of tracks) {
    t.codec = codecFor(t.codecId, t.description);
    // A VP8 or VP9 track has no configuration to hand a decoder.
    if (t.codecId === 'V_VP8' || t.codecId === 'V_VP9') t.description = null;
    // Frames stay in the file's order, which is the order they decode in.
  }
  if (duration !== null) duration = (duration * scale) / 1e9;
  return { duration, tracks };
}
