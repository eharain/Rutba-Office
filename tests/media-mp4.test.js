// The Video export's own MP4 (docs/STUDIO.md D3, Phase 4): the muxer writes
// ftyp, mdat and moov as a stream with the mdat's size patched at the end,
// and the reader takes it back to its box structure, durations, sample
// tables and key frames; a reordering encoder gets composition offsets and an
// edit list; Opus and VP9 carry their own configuration boxes; a WebM is read
// down to its frames, a recorder's unknown sizes and laced sound included;
// the frame reader answers the frame on screen at a time, decoding each
// sample once in order and starting again from a key frame when asked to go
// back; and the encode loop, given stand-in encoders, writes a file whose
// tables say exactly what was encoded.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Mp4Writer, memorySink, sampleEntryType, audioSpecificConfig } from '../packages/media/src/mp4-write.js';
import { readMp4, boxesIn, presentationTimes } from '../packages/media/src/mp4-read.js';
import { readWebm } from '../packages/media/src/webm-read.js';
import { FrameReader, demuxVideo } from '../packages/media/src/frames.js';
import { avcLevel, videoCandidates, encodeMovie, chooseVideo, chooseAudio } from '../packages/media/src/export.js';
import { probeMedia } from '../packages/media/src/probe.js';

// An avcC as an encoder hands it over: version 1, High profile, level 4.0.
const AVCC = Uint8Array.of(1, 0x64, 0x00, 0x28, 0xff, 0xe1, 0x00, 0x04, 0x67, 0x64, 0x00, 0x28, 0x01, 0x00, 0x04, 0x68, 0xee, 0x3c, 0x80);
const ASC = audioSpecificConfig(48000, 2);

const sample = (n, fill) => new Uint8Array(n).fill(fill);

async function writeFile({ seconds = 3, fps = 30, keyEvery = 30, audio = true, flushBytes } = {}) {
  const sink = memorySink();
  const writes = [];
  const patches = [];
  const watched = {
    write: (b) => { writes.push(b.length); return sink.write(b); },
    patch: (o, b) => { patches.push([o, b.length]); return sink.patch(o, b); },
  };
  const w = new Mp4Writer({
    sink: watched,
    video: { codec: 'avc1.640028', width: 1920, height: 1080 },
    audio: audio ? { codec: 'mp4a.40.2', sampleRate: 48000, channels: 2, bitrate: 128000 } : null,
    flushBytes,
  });
  await w.start();
  w.configure('video', { description: AVCC, colour: { primaries: 'bt709', transfer: 'bt709', matrix: 'bt709', fullRange: false } });
  const frames = Math.round(seconds * fps);
  const aac = Math.floor((seconds * 48000) / 1024);
  let a = 0;
  const written = [];
  for (let i = 0; i < frames; i++) {
    const t = Math.round((i * 1e6) / fps);
    while (audio && a < aac && Math.round((a * 1024 * 1e6) / 48000) <= t) {
      await w.add('audio', { data: sample(300 + (a % 7), 0xa0), timestamp: Math.round((a * 1024 * 1e6) / 48000), duration: Math.round((1024 * 1e6) / 48000) });
      a++;
    }
    const data = sample(1000 + i, i & 0xff);
    written.push(data);
    await w.add('video', { data, timestamp: t, duration: Math.round(1e6 / fps), key: i % keyEvery === 0 });
  }
  while (audio && a < aac) {
    await w.add('audio', { data: sample(300, 0xa0), timestamp: Math.round((a * 1024 * 1e6) / 48000), duration: Math.round((1024 * 1e6) / 48000) });
    a++;
  }
  const done = await w.finish();
  return { bytes: sink.bytes(), done, writes, patches, frames, written, aac };
}

test('the file is ftyp, then mdat with its size patched, then moov', async () => {
  const { bytes, patches } = await writeFile();
  const top = boxesIn(bytes);
  assert.deepEqual(top.map((b) => b.type), ['ftyp', 'mdat', 'moov']);
  assert.equal(top[1].at + top[1].size, top[2].at, 'the mdat runs exactly to the moov');
  assert.equal(top[2].at + top[2].size, bytes.length, 'nothing after the moov');
  assert.equal(patches.length, 1, 'one patch: the mdat size');
  assert.equal(patches[0][0], top[1].at + 8, 'the 64-bit size after "mdat"');
});

test('the picture track reads back: codec, size, length, every sample in place, key frames', async () => {
  const { bytes, frames, written } = await writeFile({ seconds: 3, fps: 30, keyEvery: 30 });
  const file = readMp4(bytes);
  assert.equal(file.brand, 'isom');
  assert.equal(file.fragmented, false);
  assert.ok(Math.abs(file.duration - 3) < 0.002, `movie ${file.duration}`);
  const v = file.tracks.find((t) => t.kind === 'video');
  assert.equal(v.entry, 'avc1');
  assert.equal(v.codec, 'avc1.640028');
  assert.deepEqual([...v.description], [...AVCC]);
  assert.equal(v.width, 1920);
  assert.equal(v.height, 1080);
  assert.equal(v.timescale, 90000);
  assert.equal(v.samples.length, frames);
  assert.ok(Math.abs(v.duration - 3) < 0.002);
  assert.deepEqual(v.tables.stts, [[90, 3000]], 'one run of 3000 ticks (a 30th at 90 kHz)');
  assert.deepEqual(v.tables.stss, [1, 31, 61], 'key frames every second');
  assert.deepEqual(v.samples.filter((s) => s.key).map((s) => s.dts / 90000), [0, 1, 2]);
  assert.equal(v.tables.ctts, null, 'no reordering, no composition offsets');
  assert.deepEqual(v.colour, { primaries: 1, transfer: 1, matrix: 1, fullRange: false });
  const mdat = file.boxes.find((b) => b.type === 'mdat');
  for (let i = 0; i < frames; i++) {
    const s = v.samples[i];
    assert.ok(s.offset >= mdat.at + 16 && s.offset + s.size <= mdat.at + mdat.size, `sample ${i} inside the mdat`);
    assert.deepEqual(bytes.subarray(s.offset, s.offset + s.size), written[i], `sample ${i}'s bytes`);
  }
});

test('the sound track reads back, interleaved with the picture in chunks', async () => {
  const { bytes, aac } = await writeFile({ seconds: 3 });
  const file = readMp4(bytes);
  const a = file.tracks.find((t) => t.kind === 'audio');
  assert.equal(a.entry, 'mp4a');
  assert.equal(a.codec, 'mp4a.40.2');
  assert.deepEqual([...a.description], [...ASC]);
  assert.equal(a.sampleRate, 48000);
  assert.equal(a.channels, 2);
  assert.equal(a.timescale, 48000);
  assert.equal(a.samples.length, aac);
  assert.deepEqual(a.tables.stts, [[aac, 1024]]);
  assert.equal(a.tables.stss, null, 'every sound sample is a sync sample');
  const v = file.tracks.find((t) => t.kind === 'video');
  assert.ok(a.tables.chunkOffsets.length > 10 && v.tables.chunkOffsets.length > 10, 'many chunks each: interleaved, not one block per track');
  const firstV = v.samples[0].offset;
  const lastA = a.samples[a.samples.length - 1].offset;
  assert.ok(lastA > firstV, 'sound continues after the first picture');
});

test('the probe the earlier checks use reads our file the same way', async () => {
  const { bytes } = await writeFile({ seconds: 2 });
  const p = probeMedia(bytes, { complete: true });
  assert.equal(p.format, 'mp4');
  assert.ok(Math.abs(p.duration - 2) < 0.01);
  assert.equal(p.width, 1920);
  assert.equal(p.height, 1080);
  assert.ok(p.hasVideo && p.hasAudio);
  assert.deepEqual(p.tracks.map((t) => `${t.type} ${t.codec}`), ['video avc1', 'audio mp4a']);
});

test('samples go to the disk as they come, a few megabytes at a time, never the whole file at the end', async () => {
  const { writes, bytes } = await writeFile({ seconds: 10, flushBytes: 64 * 1024 });
  assert.ok(writes.length > 5, `${writes.length} writes`);
  assert.ok(Math.max(...writes) < bytes.length / 2, 'no one write is most of the file');
});

test('a reordering encoder (B frames) gets composition offsets and an edit list; frames show from zero in order', async () => {
  const sink = memorySink();
  const w = new Mp4Writer({ sink, video: { codec: 'avc1.640028', width: 640, height: 360, description: AVCC } });
  // Decode order I P B B P B B, presentation 0 3 1 2 6 4 5 (in 30ths).
  const order = [0, 3, 1, 2, 6, 4, 5];
  for (const n of order) await w.add('video', { data: sample(50, n), timestamp: Math.round((n * 1e6) / 30), duration: Math.round(1e6 / 30), key: n === 0 });
  await w.finish();
  const v = readMp4(sink.bytes()).tracks[0];
  assert.ok(v.tables.ctts, 'ctts written');
  assert.ok(v.samples.every((s) => s.pts >= s.dts), 'no frame shows before it decodes');
  assert.equal(v.edits.length, 1);
  assert.ok(v.edits[0].mediaTime > 0, 'the edit list skips the decoder\'s delay');
  const shown = presentationTimes(v).map((s) => Math.round(s.time * 30)).sort((a, b) => a - b);
  assert.deepEqual(shown, [0, 1, 2, 3, 4, 5, 6]);
  assert.equal(presentationTimes(v)[1].time.toFixed(4), (3 / 30).toFixed(4), 'the second decoded frame shows third');
});

test('a track that starts late gets an empty edit; a file past 4 GB gets 64-bit chunk offsets', async () => {
  const sink = memorySink();
  const w = new Mp4Writer({ sink, audio: { codec: 'mp4a.40.2', sampleRate: 48000, channels: 1 } });
  await w.start();
  // Pretend the mdat already holds 4 GB: the next chunk lands past it.
  w.position += 2 ** 32;
  await w.add('audio', { data: sample(10, 1), timestamp: 500000, duration: 21333 });
  const moov = w.moov([w.timesOf(w.tracks[0])], 1);
  const text = new TextDecoder('latin1').decode(moov);
  assert.ok(text.includes('co64') && !text.includes('stco'), 'co64');
  assert.ok(text.includes('elst'), 'an edit list for the half second of nothing first');
});

test('Opus and VP9 carry their configuration boxes and read back as WebCodecs names them', async () => {
  const head = new Uint8Array(19);
  head.set(new TextEncoder().encode('OpusHead'));
  head[8] = 1;
  head[9] = 2;
  new DataView(head.buffer).setUint16(10, 3840, true);
  new DataView(head.buffer).setUint32(12, 48000, true);
  const sink = memorySink();
  const w = new Mp4Writer({ sink, video: { codec: 'vp09.00.40.08', width: 1280, height: 720 }, audio: { codec: 'opus', sampleRate: 48000, channels: 2, description: head } });
  for (let i = 0; i < 10; i++) {
    await w.add('video', { data: sample(40, i), timestamp: i * 33333, duration: 33333, key: i === 0 });
    await w.add('audio', { data: sample(20, i), timestamp: i * 20000, duration: 20000 });
  }
  await w.finish();
  const file = readMp4(sink.bytes());
  const v = file.tracks.find((t) => t.kind === 'video');
  const a = file.tracks.find((t) => t.kind === 'audio');
  assert.equal(v.entry, 'vp09');
  assert.equal(v.configBox, 'vpcC');
  assert.equal(v.codec, 'vp09.00.40.08');
  assert.equal(a.entry, 'Opus');
  assert.equal(a.configBox, 'dOps');
  assert.equal(a.codec, 'opus');
  assert.deepEqual([...a.description], [...head], 'dOps comes back as the OpusHead a decoder takes');
  assert.equal(sampleEntryType('avc1.4d0028'), 'avc1');
  assert.equal(sampleEntryType('hvc1.1.6.L120.B0'), 'hvc1');
});

test('the H.264 level follows the size and the rate', () => {
  assert.equal(avcLevel(1280, 720, 30), '1f');
  assert.equal(avcLevel(1920, 1080, 30), '28');
  assert.equal(avcLevel(1920, 1080, 60), '2a');
  assert.equal(avcLevel(3840, 2160, 30), '33');
  assert.equal(avcLevel(3840, 2160, 60), '34');
  assert.deepEqual(videoCandidates(1920, 1080, 30), ['avc1.640028', 'avc1.4d0028', 'avc1.420028', 'vp09.00.40.08']);
});

/* ── WebM ─────────────────────────────────────────────────────────────── */

const ebmlId = (id) => { const out = []; let v = id; while (v) { out.unshift(v & 0xff); v = Math.floor(v / 256); } return out; };
const vsize = (n) => [0x10, (n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff]; // four-byte size
const el = (id, ...body) => { const flat = body.flat(Infinity); return [...ebmlId(id), ...vsize(flat.length), ...flat]; };
const unknown = (id) => [...ebmlId(id), 0x01, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff];
const uintEl = (id, v) => el(id, [v >> 8 & 0xff, v & 0xff]);
const str = (s) => [...new TextEncoder().encode(s)];

function recorderWebm() {
  const header = el(0x1a45dfa3, el(0x4282, str('webm')));
  const info = el(0x1549a966, el(0x2ad7b1, [0x0f, 0x42, 0x40]));
  const tracks = el(0x1654ae6b,
    el(0xae, uintEl(0xd7, 1), uintEl(0x83, 1), el(0x86, str('V_VP8')), el(0xe0, uintEl(0xb0, 320), uintEl(0xba, 180))),
    el(0xae, uintEl(0xd7, 2), uintEl(0x83, 2), el(0x86, str('A_OPUS')), el(0x63a2, str('OpusHead')), el(0xe1, el(0xb5, [0x40, 0xe7, 0x70, 0, 0, 0, 0, 0]), uintEl(0x9f, 1))));
  const block = (track, ms, key, fill, n = 30) => el(0xa3, [0x80 | track, ms >> 8, ms & 0xff, key ? 0x80 : 0], new Array(n).fill(fill));
  const cluster1 = [...unknown(0x1f43b675), ...uintEl(0xe7, 0), ...block(1, 0, true, 1), ...block(2, 0, true, 9, 10), ...block(1, 40, false, 2), ...block(1, 80, false, 3)];
  // A BlockGroup whose ReferenceBlock says it is not a key frame, and a Xiph-laced sound block of three frames.
  const group = el(0xa0, el(0xa1, [0x81, 0, 0, 0], new Array(30).fill(5)), el(0xfb, [0xd8]));
  const laced = el(0xa3, [0x82, 0, 40, 0x80 | 0x02, 2, 10, 12], new Array(10).fill(7), new Array(12).fill(8), new Array(14).fill(6));
  const cluster2 = [...unknown(0x1f43b675), ...uintEl(0xe7, 1000), ...block(1, 0, true, 4), ...group, ...laced];
  return Uint8Array.from([...header, ...unknown(0x18538067), ...info, ...tracks, ...cluster1, ...cluster2]);
}

test('a recorder\'s WebM (no sizes on the segment and clusters) reads down to its frames', () => {
  const file = readWebm(recorderWebm());
  const v = file.tracks.find((t) => t.kind === 'video');
  const a = file.tracks.find((t) => t.kind === 'audio');
  assert.equal(v.codec, 'vp8');
  assert.equal(v.width, 320);
  assert.equal(v.height, 180);
  assert.equal(v.description, null, 'VP8 has no configuration to hand over');
  assert.deepEqual(v.frames.map((f) => [f.time, f.key]), [[0, true], [0.04, false], [0.08, false], [1, true], [1, false]]);
  const b = recorderWebm();
  assert.ok(v.frames.every((f, i) => b[f.offset] === [1, 2, 3, 4, 5][i] && f.size === 30), 'each frame\'s bytes where it says');
  assert.equal(a.codec, 'opus');
  assert.equal(a.sampleRate, 48000);
  assert.deepEqual(a.frames.map((f) => f.size), [10, 10, 12, 14], 'the laced block is three frames');
  const demuxed = demuxVideo(b);
  assert.equal(demuxed.container, 'webm');
  assert.equal(demuxed.samples.length, 5);
});

test('an MP4 picture track demuxes with its times after the edit list', async () => {
  const { bytes } = await writeFile({ seconds: 1, audio: false });
  const d = demuxVideo(bytes);
  assert.equal(d.container, 'mp4');
  assert.equal(d.codec, 'avc1.640028');
  assert.equal(d.samples.length, 30);
  assert.equal(d.samples[0].time, 0);
  assert.ok(Math.abs(d.samples[29].time - 29 / 30) < 1e-6);
});

/* ── the frame reader, with a stand-in decoder ───────────────────────── */

function fakeDecoderKit() {
  const log = { decoded: [], resets: 0, open: 0, closed: 0 };
  class Frame {
    constructor(timestamp) { this.timestamp = timestamp; log.open++; this.closed = false; }
    close() { if (!this.closed) { this.closed = true; log.closed++; } }
  }
  class Decoder {
    static async isConfigSupported(c) { return { supported: c.codec === 'avc1.640028' || c.codec === 'vp8' }; }
    constructor({ output }) { this.output = output; this.decodeQueueSize = 0; this.held = []; }
    addEventListener() {}
    configure() {}
    reset() { log.resets++; this.held = []; this.decodeQueueSize = 0; this.generation = (this.generation || 0) + 1; }
    decode(chunk) {
      log.decoded.push(chunk.timestamp);
      this.decodeQueueSize++;
      // Like a real decoder, a frame comes out a moment later.
      // and reset() drops what has not come out yet, as the real one does.
      const gen = this.generation;
      setTimeout(() => { if (gen !== this.generation) return; this.decodeQueueSize--; this.output(new Frame(chunk.timestamp)); }, 0);
    }
    async flush() { await new Promise((r) => setTimeout(r, 5)); }
    close() {}
  }
  class Chunk { constructor(o) { Object.assign(this, o); } }
  return { log, Decoder, Chunk };
}

test('the frame reader answers the frame on screen at each time, decoding each sample once going forward', async () => {
  const { bytes } = await writeFile({ seconds: 4, fps: 30, keyEvery: 30, audio: false });
  const kit = fakeDecoderKit();
  const r = await FrameReader.open(bytes, { Decoder: kit.Decoder, Chunk: kit.Chunk });
  assert.ok(r, 'opened');
  // An export at 25 frames a second over a 30 a second source.
  for (let i = 0; i < 100; i++) {
    const t = i / 25;
    const f = await r.frameAt(t);
    const want = Math.floor(t * 30 + 1e-6);
    assert.equal(Math.round((f.timestamp / 1e6) * 30), want, `at ${t.toFixed(2)} s`);
  }
  assert.equal(kit.log.resets, 0, 'never started again');
  assert.equal(new Set(kit.log.decoded).size, kit.log.decoded.length, 'no sample decoded twice');
  assert.ok(kit.log.open - kit.log.closed <= 6, `${kit.log.open - kit.log.closed} frames held at once`);
  r.close();
  assert.equal(kit.log.open, kit.log.closed, 'every frame closed at the end');
});

test('going back, or far ahead, starts again from the key frame before the time asked for', async () => {
  const { bytes } = await writeFile({ seconds: 6, fps: 30, keyEvery: 30, audio: false });
  const kit = fakeDecoderKit();
  const r = await FrameReader.open(bytes, { Decoder: kit.Decoder, Chunk: kit.Chunk });
  await r.frameAt(2.5);
  assert.equal(kit.log.resets, 1, 'from the start straight to 2.5 s: from the key frame at 2 s');
  assert.ok(kit.log.decoded.length <= 20, `${kit.log.decoded.length} decoded to reach 2.5 s`);
  const mark = kit.log.decoded.length;
  const f1 = await r.frameAt(1.2);
  assert.equal(Math.round((f1.timestamp / 1e6) * 30), 36);
  assert.equal(kit.log.resets, 2, 'back: one restart more');
  assert.equal(kit.log.decoded[mark], 1e6, 'decoding starts again at the key frame at 1 s');
  const before = kit.log.decoded.length;
  const f2 = await r.frameAt(5.1);
  assert.equal(Math.round((f2.timestamp / 1e6) * 30), 153);
  assert.equal(kit.log.resets, 3, 'far ahead: a restart rather than decoding all between');
  assert.ok(kit.log.decoded.length - before <= 30, `${kit.log.decoded.length - before} decoded to get there`);
  const last = await r.frameAt(99);
  assert.equal(Math.round((last.timestamp / 1e6) * 30), 179, 'past the end: the last frame');
  r.close();
});

test('a codec the decoder does not take gives no reader, so the caller falls back', async () => {
  const kit = fakeDecoderKit();
  const sink = memorySink();
  const w = new Mp4Writer({ sink, video: { codec: 'vp09.00.40.08', width: 64, height: 64 } });
  await w.add('video', { data: sample(10, 1), timestamp: 0, duration: 33333 });
  await w.finish();
  assert.equal(await FrameReader.open(sink.bytes(), { Decoder: kit.Decoder, Chunk: kit.Chunk }), null);
  assert.equal(await FrameReader.open(new Uint8Array(100), { Decoder: kit.Decoder, Chunk: kit.Chunk }), null);
});

/* ── the encode loop, with stand-in encoders ─────────────────────────── */

function fakeEncoders({ hardwareOk = true } = {}) {
  const seen = { configs: [], frames: 0, keys: [] };
  class VideoEncoder {
    static async isConfigSupported(c) { return { supported: c.codec.startsWith('avc1.64') && (hardwareOk || c.hardwareAcceleration !== 'prefer-hardware') }; }
    constructor({ output }) { this.output = output; this.encodeQueueSize = 0; this.first = true; }
    addEventListener() {}
    configure(c) { seen.configs.push(c); this.config = c; }
    encode(frame, { keyFrame }) {
      seen.frames++;
      if (keyFrame) seen.keys.push(frame.timestamp);
      const meta = this.first ? { decoderConfig: { codec: this.config.codec, description: AVCC.buffer.slice(0) } } : undefined;
      this.first = false;
      const ts = frame.timestamp;
      const dur = frame.duration;
      this.encodeQueueSize++;
      setTimeout(() => {
        this.encodeQueueSize--;
        this.output({ byteLength: 64, timestamp: ts, duration: dur, type: keyFrame ? 'key' : 'delta', copyTo: (u8) => u8.fill(ts & 0xff) }, meta);
      }, 0);
    }
    async flush() { await new Promise((r) => setTimeout(r, 10)); }
    close() {}
  }
  class VideoFrame { constructor(src, o) { this.timestamp = o.timestamp; this.duration = o.duration; } close() {} }
  class AudioEncoder {
    static async isConfigSupported(c) { return { supported: c.codec === 'mp4a.40.2' }; }
    constructor({ output }) { this.output = output; this.encodeQueueSize = 0; this.at = 0; }
    configure(c) { this.c = c; }
    encode(data) {
      // 1024 samples a chunk, as AAC.
      const total = data.numberOfFrames;
      for (let k = 0; k + 1024 <= total; k += 1024) {
        const ts = Math.round((this.at * 1e6) / 48000);
        this.at += 1024;
        this.output({ byteLength: 20, timestamp: ts, duration: Math.round(1024e6 / 48000), copyTo: (u8) => u8.fill(1) }, this.at === 1024 ? { decoderConfig: { description: ASC } } : undefined);
      }
    }
    async flush() {}
    close() {}
  }
  class AudioData { constructor(o) { Object.assign(this, o); } close() {} }
  return { seen, VideoEncoderCtor: VideoEncoder, VideoFrameCtor: VideoFrame, AudioEncoderCtor: AudioEncoder, AudioDataCtor: AudioData };
}

const fakeSound = (seconds) => {
  const n = Math.round(seconds * 48000);
  const ch = [new Float32Array(n), new Float32Array(n)];
  return { sampleRate: 48000, numberOfChannels: 2, length: n, duration: seconds, getChannelData: (c) => ch[c] };
};

test('encodeMovie paints every frame, asks for key frames every two seconds, and writes a file that says so', async () => {
  const kit = fakeEncoders();
  const sink = memorySink();
  const painted = [];
  const progress = [];
  const out = await encodeMovie({
    sink, width: 1920, height: 1080, fps: 30, duration: 5, bitrate: 8e6,
    paint: async (i, t) => { painted.push(t); return {}; },
    sound: fakeSound(5),
    onProgress: (p) => progress.push(p),
    ...kit,
  });
  assert.equal(painted.length, 150);
  assert.equal(out.frames, 150);
  assert.equal(out.video.codec, 'avc1.640028');
  assert.equal(out.video.hardware, true, 'hardware asked for first');
  assert.equal(kit.seen.configs[0].hardwareAcceleration, 'prefer-hardware');
  assert.deepEqual(kit.seen.keys, [0, 2e6, 4e6]);
  assert.equal(progress[progress.length - 1], 1);
  const file = readMp4(sink.bytes());
  const v = file.tracks.find((t) => t.kind === 'video');
  const a = file.tracks.find((t) => t.kind === 'audio');
  assert.equal(v.samples.length, 150);
  assert.deepEqual(v.tables.stss, [1, 61, 121]);
  assert.ok(Math.abs(v.duration - 5) < 0.002, `${v.duration}`);
  assert.equal(a.codec, 'mp4a.40.2');
  assert.ok(Math.abs(a.duration - 5) < 0.03, `${a.duration}`);
  assert.ok(Math.abs(file.duration - 5) < 0.03);
});

test('without the hardware encoder the software one is taken; a cancelled export stops and says so', async () => {
  const kit = fakeEncoders({ hardwareOk: false });
  assert.equal((await chooseVideo({ width: 1280, height: 720, Encoder: kit.VideoEncoderCtor })).hardware, false);
  assert.equal((await chooseAudio({ Encoder: kit.AudioEncoderCtor })).codec, 'mp4a.40.2');
  const abort = new AbortController();
  await assert.rejects(encodeMovie({
    sink: memorySink(), width: 1280, height: 720, fps: 30, duration: 10,
    paint: (i) => { if (i === 20) abort.abort(); return {}; },
    signal: abort.signal,
    ...kit,
  }), (err) => err.name === 'AbortError');
});

test('a movie is not taken for a picture: the picture probe leaves an MP4 to the media probe, and still knows HEIC and AVIF', async () => {
  const { probeImage } = await import('../packages/imaging/src/probe.js');
  const { bytes } = await writeFile({ seconds: 1, audio: false });
  assert.equal(probeImage(bytes), null);
  const ftyp = (major, ...compat) => {
    const names = [major, '\0\0\0\0', ...compat];
    const b = new Uint8Array(8 + names.length * 4 + 16);
    new DataView(b.buffer).setUint32(0, 8 + names.length * 4);
    b.set(new TextEncoder().encode('ftyp'), 4);
    names.forEach((n, i) => b.set(new TextEncoder().encode(n.padEnd(4, ' ').slice(0, 4)), 8 + i * 4));
    return b;
  };
  assert.equal(probeImage(ftyp('heic', 'mif1', 'heic')).format, 'heic');
  assert.equal(probeImage(ftyp('avif', 'mif1', 'avif')).format, 'avif');
  assert.equal(probeImage(ftyp('mif1', 'avif')).format, 'avif');
  assert.equal(probeImage(ftyp('qt  ')), null);
});
