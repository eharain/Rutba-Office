// An MP4 file written as it is made: our own muxer, for the Video export.
//
// The layout is the plain one every player reads: `ftyp`, then `mdat` with
// the samples as the encoders hand them over, then `moov` with the sample
// tables, written last because only then are they known. Nothing is held
// back but the tables: the samples go to the sink a few megabytes at a time,
// so an hour of video never sits in memory, and the one thing written twice
// is the mdat's size, patched at the end (a 64-bit size, so a file past
// 4 GB is still right). No fragments: a fragmented file is what a recorder
// writes because it may stop at any moment; an export ends, and a file with
// one sample table seeks at once in every player.
//
// Tracks: one picture (H.264 `avc1`, H.265 `hvc1`, VP9 `vp09`, AV1 `av01`)
// and one sound (AAC `mp4a`, Opus `Opus`). Times come in microseconds, as
// WebCodecs gives them; a picture track is kept at 90 kHz and a sound
// track at its sample rate. A picture encoder that reorders frames (B
// frames) gets composition offsets and an edit list, so the first frame
// shows at zero.

const enc = new TextEncoder();

/** A box: four bytes of size, four of type, then the body. */
function box(type, ...parts) {
  const body = concat(parts);
  const out = new Uint8Array(8 + body.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, out.length);
  out.set(enc.encode(type), 4);
  out.set(body, 8);
  return out;
}

/** A full box: a box whose body starts with a version and three bytes of flags. */
function fullBox(type, version, flags, ...parts) {
  return box(type, u8(version), u24(flags), ...parts);
}

function concat(parts) {
  const list = parts.flat(Infinity).filter(Boolean);
  let n = 0;
  for (const p of list) n += p.length;
  const out = new Uint8Array(n);
  let at = 0;
  for (const p of list) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

const u8 = (v) => Uint8Array.of(v & 0xff);
const u16 = (v) => Uint8Array.of((v >>> 8) & 0xff, v & 0xff);
const u24 = (v) => Uint8Array.of((v >>> 16) & 0xff, (v >>> 8) & 0xff, v & 0xff);
const u32 = (v) => {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, v >>> 0);
  return b;
};
const i32 = (v) => {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setInt32(0, v | 0);
  return b;
};
const u64 = (v) => {
  const b = new Uint8Array(8);
  new DataView(b.buffer).setBigUint64(0, BigInt(Math.max(0, Math.round(v))));
  return b;
};
const zeros = (n) => new Uint8Array(n);
const fourcc = (s) => enc.encode(String(s).padEnd(4, ' ').slice(0, 4));
const bytesOf = (d) => (d == null ? null : d instanceof Uint8Array ? d : ArrayBuffer.isView(d) ? new Uint8Array(d.buffer, d.byteOffset, d.byteLength) : new Uint8Array(d));

// The identity matrix every header carries.
const MATRIX = concat([u32(0x00010000), u32(0), u32(0), u32(0), u32(0x00010000), u32(0), u32(0), u32(0), u32(0x40000000)]);

const MAX32 = 0xffffffff;

/** The sample entry's four characters for a WebCodecs codec string. */
export function sampleEntryType(codec) {
  const c = String(codec || '');
  if (/^avc[13]\./.test(c)) return 'avc1';
  if (/^(hvc1|hev1)\./.test(c)) return 'hvc1';
  if (/^vp09\./.test(c)) return 'vp09';
  if (/^av01\./.test(c)) return 'av01';
  if (/^mp4a\./.test(c)) return 'mp4a';
  if (c === 'opus') return 'Opus';
  return null;
}

/** The VP9 codec configuration (vpcC) from a `vp09.PP.LL.DD` string. */
export function vpcC(codec, colour = {}) {
  const [, profile = '00', level = '10', depth = '08'] = String(codec).split('.');
  const chroma = 1; // 4:2:0, co-sited with luma
  return fullBox('vpcC', 1, 0,
    u8(Number(profile)), u8(Number(level)),
    u8(((Number(depth) & 0x0f) << 4) | (chroma << 1) | (colour.fullRange ? 1 : 0)),
    u8(colour.primaries ?? 1), u8(colour.transfer ?? 1), u8(colour.matrix ?? 1),
    u16(0));
}

/** An AAC AudioSpecificConfig for a sample rate and channel count, where the encoder gave none. */
export function audioSpecificConfig(sampleRate, channels, objectType = 2) {
  const RATES = [96000, 88200, 64000, 48000, 44100, 32000, 24000, 22050, 16000, 12000, 11025, 8000, 7350];
  const index = Math.max(0, RATES.indexOf(sampleRate));
  return Uint8Array.of((objectType << 3) | (index >> 1), ((index & 1) << 7) | ((channels & 0x0f) << 3));
}

/** An MPEG-4 descriptor: a tag, its length in the four-byte form, its body. */
function descriptor(tag, ...parts) {
  const body = concat(parts);
  const n = body.length;
  return concat([u8(tag), Uint8Array.of(0x80 | ((n >> 21) & 0x7f), 0x80 | ((n >> 14) & 0x7f), 0x80 | ((n >> 7) & 0x7f), n & 0x7f), body]);
}

function esds(asc, { bitrate = 128000 } = {}) {
  return fullBox('esds', 0, 0,
    descriptor(0x03, u16(0), u8(0),
      descriptor(0x04, u8(0x40), u8(0x15), u24(0), u32(bitrate), u32(bitrate),
        descriptor(0x05, asc)),
      descriptor(0x06, u8(0x02))));
}

/** Opus in MP4 (dOps), from the OpusHead an encoder may give (little-endian) or from the track. */
function dOps(head, { channels = 2, sampleRate = 48000 } = {}) {
  let ch = channels;
  let preSkip = 312;
  let rate = sampleRate;
  let gain = 0;
  if (head && head.length >= 19 && new TextDecoder('latin1').decode(head.subarray(0, 8)) === 'OpusHead') {
    const dv = new DataView(head.buffer, head.byteOffset, head.byteLength);
    ch = head[9];
    preSkip = dv.getUint16(10, true);
    rate = dv.getUint32(12, true);
    gain = dv.getInt16(16, true);
  }
  return box('dOps', u8(0), u8(ch), u16(preSkip), u32(rate), u16(gain & 0xffff), u8(0));
}

/** The colour box from a WebCodecs colour space, where the encoder named one. */
function colr(space) {
  if (!space) return null;
  const P = { bt709: 1, bt470bg: 5, smpte170m: 6, bt2020: 9, smpte432: 12 };
  const T = { bt709: 1, smpte170m: 6, iec61966_2_1: 13, linear: 8, pq: 16, hlg: 18 };
  const M = { rgb: 0, bt709: 1, bt470bg: 5, smpte170m: 6, bt2020_ncl: 9 };
  if (!(space.primaries in P) && !(space.transfer in T) && !(space.matrix in M)) return null;
  return box('colr', fourcc('nclx'), u16(P[space.primaries] ?? 2), u16(T[space.transfer] ?? 2), u16(M[space.matrix] ?? 2), u8(space.fullRange ? 0x80 : 0));
}

function visualEntry(track) {
  const type = sampleEntryType(track.codec);
  const desc = bytesOf(track.description);
  let config = null;
  if (type === 'avc1') config = box('avcC', desc);
  else if (type === 'hvc1') config = box('hvcC', desc);
  else if (type === 'vp09') config = vpcC(track.codec, track.colour);
  else if (type === 'av01') config = desc ? box('av1C', desc) : null;
  const name = zeros(32);
  return box(type,
    zeros(6), u16(1), // reserved, data reference index
    zeros(16), // pre-defined and reserved
    u16(track.width), u16(track.height),
    u32(0x00480000), u32(0x00480000), // 72 dpi
    u32(0), u16(1), // reserved, one frame a sample
    name, u16(0x0018), u16(0xffff),
    config,
    colr(track.colour),
    box('pasp', u32(1), u32(1)));
}

function audioEntry(track) {
  const type = sampleEntryType(track.codec);
  const desc = bytesOf(track.description);
  const config = type === 'mp4a'
    ? esds(desc && desc.length ? desc : audioSpecificConfig(track.sampleRate, track.channels), { bitrate: track.bitrate })
    : dOps(desc, track);
  return box(type,
    zeros(6), u16(1),
    zeros(8),
    u16(track.channels), u16(16), u16(0), u16(0),
    u32((type === 'Opus' ? 48000 : track.sampleRate) * 65536),
    config);
}

/** Runs of equal values: [[count, value], ...]. */
function runs(values) {
  const out = [];
  for (const v of values) {
    const last = out[out.length - 1];
    if (last && last[1] === v) last[0]++;
    else out.push([1, v]);
  }
  return out;
}

/**
 * The muxer. `sink` takes the file a piece at a time: `write(bytes)` adds
 * to the end, `patch(offset, bytes)` overwrites what is already written
 * (the mdat's size). Both may answer promises.
 *
 * @param {{ sink: { write: Function, patch: Function },
 *   video?: { codec: string, width: number, height: number, description?: BufferSource, colour?: object },
 *   audio?: { codec: string, sampleRate: number, channels: number, description?: BufferSource, bitrate?: number },
 *   flushBytes?: number }} options
 */
export class Mp4Writer {
  constructor({ sink, video = null, audio = null, flushBytes = 4 * 1024 * 1024 }) {
    if (!video && !audio) throw new Error('An MP4 needs a track.');
    this.sink = sink;
    this.flushBytes = flushBytes;
    this.tracks = [];
    if (video) this.tracks.push({ kind: 'video', id: this.tracks.length + 1, timescale: 90000, ...video, samples: [], chunks: [] });
    if (audio) this.tracks.push({ kind: 'audio', id: this.tracks.length + 1, timescale: audio.sampleRate, ...audio, samples: [], chunks: [] });
    this.position = 0;
    this.pending = [];
    this.pendingBytes = 0;
    this.lastTrack = null;
    this.mdatAt = 0;
    this.started = false;
    this.finished = false;
  }

  track(kind) {
    return this.tracks.find((t) => t.kind === kind) || null;
  }

  /** Set what a track learns from its encoder's first output: its description and colour. */
  configure(kind, { description, colour } = {}) {
    const t = this.track(kind);
    if (!t) return;
    if (description != null && t.description == null) t.description = bytesOf(description).slice();
    if (colour && !t.colour) t.colour = colour;
  }

  async start() {
    if (this.started) return;
    this.started = true;
    const brands = ['isom', 'iso2', 'mp41'];
    if (this.tracks.some((t) => sampleEntryType(t.codec) === 'avc1')) brands.push('avc1');
    const ftyp = box('ftyp', fourcc('isom'), u32(0x200), ...brands.map(fourcc));
    await this.emit(ftyp);
    this.mdatAt = this.position;
    // size 1: the 64-bit size follows, patched when the samples are all in.
    await this.emit(concat([u32(1), fourcc('mdat'), u64(0)]));
  }

  /**
   * One encoded sample: `data` its bytes, `timestamp` and `duration` in
   * microseconds (presentation time), `key` whether it is a sync sample.
   * Samples of a track come in decode order.
   */
  async add(kind, { data, timestamp, duration = 0, key = true }) {
    if (!this.started) await this.start();
    if (this.finished) throw new Error('The file is already finished.');
    const t = this.track(kind);
    if (!t) throw new Error(`This file has no ${kind} track.`);
    const bytes = bytesOf(data);
    if (this.lastTrack !== t || t.chunks[t.chunks.length - 1].count >= 256) {
      t.chunks.push({ offset: this.position, count: 0 });
      this.lastTrack = t;
    }
    t.chunks[t.chunks.length - 1].count++;
    t.samples.push({ size: bytes.length, pts: Number(timestamp) || 0, duration: Number(duration) || 0, key: Boolean(key) });
    await this.emit(bytes);
  }

  async emit(bytes) {
    this.pending.push(bytes);
    this.pendingBytes += bytes.length;
    this.position += bytes.length;
    if (this.pendingBytes >= this.flushBytes) await this.flush();
  }

  async flush() {
    if (!this.pending.length) return;
    const out = this.pending.length === 1 ? this.pending[0] : concat(this.pending);
    this.pending = [];
    this.pendingBytes = 0;
    await this.sink.write(out);
  }

  /** The tables, the mdat's size patched, the moov written. Answers { size, duration } (seconds). */
  async finish() {
    if (!this.started) await this.start();
    if (this.finished) throw new Error('The file is already finished.');
    this.finished = true;
    const mdatSize = this.position - this.mdatAt;
    await this.flush();
    await this.sink.patch(this.mdatAt + 8, u64(mdatSize));
    const tables = this.tracks.map((t) => this.timesOf(t));
    const duration = Math.max(0, ...tables.map((x) => x.seconds));
    const moov = this.moov(tables, duration);
    await this.emit(moov);
    await this.flush();
    return { size: this.position, duration };
  }

  /** Decode times, durations and composition offsets of a track, in its timescale. */
  timesOf(t) {
    const ts = t.timescale;
    const units = (us) => Math.round((us * ts) / 1e6);
    const pts = t.samples.map((s) => units(s.pts));
    const sorted = [...pts].sort((a, b) => a - b);
    let delay = 0;
    for (let i = 0; i < pts.length; i++) delay = Math.max(delay, sorted[i] - pts[i]);
    const dts = sorted.map((v) => v - delay);
    const durations = [];
    for (let i = 0; i < dts.length; i++) {
      if (i + 1 < dts.length) durations.push(Math.max(0, dts[i + 1] - dts[i]));
      else durations.push(Math.max(1, units(t.samples[i].duration) || (i > 0 ? durations[i - 1] : 1)));
    }
    const offsets = pts.map((p, i) => p - dts[i]);
    const reordered = offsets.some((o) => o !== 0);
    const first = sorted.length ? sorted[0] : 0;
    const total = durations.reduce((a, b) => a + b, 0);
    // What the track shows: from its first frame's time to its last's end.
    const shown = Math.max(0, (dts[dts.length - 1] ?? 0) + (durations[durations.length - 1] ?? 0) + delay - first);
    return { ts, dts, durations, offsets, reordered, delay, first, total, seconds: (first + shown) / ts, shown };
  }

  moov(tables, seconds) {
    const MOVIE_TS = 1000;
    const movieDur = Math.round(seconds * MOVIE_TS);
    const v1 = movieDur > MAX32;
    const mvhd = fullBox('mvhd', v1 ? 1 : 0, 0,
      v1 ? [u64(0), u64(0), u32(MOVIE_TS), u64(movieDur)] : [u32(0), u32(0), u32(MOVIE_TS), u32(movieDur)],
      u32(0x00010000), u16(0x0100), zeros(10), MATRIX, zeros(24), u32(this.tracks.length + 1));
    const traks = this.tracks.map((t, i) => this.trak(t, tables[i], MOVIE_TS));
    return box('moov', mvhd, ...traks);
  }

  trak(t, x, movieTs) {
    const video = t.kind === 'video';
    const trackDur = Math.round(((x.first + x.shown) / x.ts) * movieTs);
    const big = trackDur > MAX32 || x.total > MAX32;
    const tkhd = fullBox('tkhd', big ? 1 : 0, 0x3,
      big ? [u64(0), u64(0), u32(t.id), u32(0), u64(trackDur)] : [u32(0), u32(0), u32(t.id), u32(0), u32(trackDur)],
      zeros(8), u16(0), u16(video ? 0 : 1), u16(video ? 0 : 0x0100), u16(0), MATRIX,
      u32(video ? t.width * 65536 : 0), u32(video ? t.height * 65536 : 0));
    // The edit list: an empty stretch where the first sample starts late, and
    // the media's own start where frames were reordered.
    const edits = [];
    if (x.first > 0) edits.push([Math.round((x.first / x.ts) * movieTs), -1]);
    // The samples' times start at zero in the file; the first frame shown sits `delay` in.
    if (x.first > 0 || x.delay > 0) edits.push([Math.round((x.shown / x.ts) * movieTs), x.delay]);
    const edts = edits.length
      ? box('edts', fullBox('elst', 0, 0, u32(edits.length), ...edits.map(([d, m]) => [u32(d), i32(m), u16(1), u16(0)])))
      : null;
    const mdhd = fullBox('mdhd', big ? 1 : 0, 0,
      big ? [u64(0), u64(0), u32(x.ts), u64(x.total)] : [u32(0), u32(0), u32(x.ts), u32(x.total)],
      u16(0x55c4), u16(0)); // 'und'
    const hdlr = fullBox('hdlr', 0, 0, u32(0), fourcc(video ? 'vide' : 'soun'), zeros(12), enc.encode(video ? 'VideoHandler\0' : 'SoundHandler\0'));
    const head = video ? fullBox('vmhd', 0, 1, zeros(8)) : fullBox('smhd', 0, 0, zeros(4));
    const dinf = box('dinf', fullBox('dref', 0, 0, u32(1), fullBox('url ', 0, 1)));
    return box('trak', tkhd, edts, box('mdia', mdhd, hdlr, box('minf', head, dinf, this.stbl(t, x))));
  }

  stbl(t, x) {
    const video = t.kind === 'video';
    const stsd = fullBox('stsd', 0, 0, u32(1), video ? visualEntry(t) : audioEntry(t));
    const timeRuns = runs(x.durations);
    const stts = fullBox('stts', 0, 0, u32(timeRuns.length), ...timeRuns.map(([n, d]) => [u32(n), u32(d)]));
    let ctts = null;
    if (x.reordered) {
      const offRuns = runs(x.offsets);
      ctts = fullBox('ctts', 0, 0, u32(offRuns.length), ...offRuns.map(([n, o]) => [u32(n), u32(o)]));
    }
    let stss = null;
    if (t.samples.some((s) => !s.key)) {
      const keys = [];
      t.samples.forEach((s, i) => { if (s.key) keys.push(i + 1); });
      stss = fullBox('stss', 0, 0, u32(keys.length), ...keys.map(u32));
    }
    const sizes = t.samples.map((s) => s.size);
    const same = sizes.length > 0 && sizes.every((s) => s === sizes[0]);
    const stsz = same
      ? fullBox('stsz', 0, 0, u32(sizes[0]), u32(sizes.length))
      : fullBox('stsz', 0, 0, u32(0), u32(sizes.length), ...sizes.map(u32));
    const per = [];
    t.chunks.forEach((c, i) => {
      const last = per[per.length - 1];
      if (!last || last[1] !== c.count) per.push([i + 1, c.count]);
    });
    const stsc = fullBox('stsc', 0, 0, u32(per.length), ...per.map(([first, n]) => [u32(first), u32(n), u32(1)]));
    const wide = t.chunks.some((c) => c.offset > MAX32);
    const stco = wide
      ? fullBox('co64', 0, 0, u32(t.chunks.length), ...t.chunks.map((c) => u64(c.offset)))
      : fullBox('stco', 0, 0, u32(t.chunks.length), ...t.chunks.map((c) => u32(c.offset)));
    return box('stbl', stsd, stts, ctts, stss, stsz, stsc, stco);
  }
}

/** A sink that keeps the file in memory: for tests, and for a host with no file to stream to. */
export function memorySink() {
  const parts = [];
  let size = 0;
  let whole = null;
  const flat = () => {
    if (!whole || whole.length !== size) {
      whole = concat(parts);
      parts.length = 0;
      parts.push(whole);
    }
    return whole;
  };
  return {
    write(bytes) {
      parts.push(bytes.slice());
      size += bytes.length;
      whole = null;
    },
    patch(offset, bytes) {
      flat().set(bytes, offset);
    },
    bytes: () => flat(),
    get size() { return size; },
  };
}
