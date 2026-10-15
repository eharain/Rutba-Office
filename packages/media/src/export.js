// A video encoded as fast as it can be painted, into an MP4 of our own.
//
// The caller paints each frame (a canvas, or anything a VideoFrame can be
// made from) and hands over the sound already mixed (an AudioBuffer).
// WebCodecs' VideoEncoder and AudioEncoder encode them, on the graphics
// hardware where the machine has an encoder there, and Mp4Writer writes the
// file through the caller's sink as the samples come. Nothing here plays in
// real time: a minute of video takes as long as its frames take to paint
// and encode, which on a laptop is a fraction of a minute.
//
// Codecs, in order: H.264 (High, then Main, then Baseline, at the level the
// size and rate need), then VP9 in MP4 where H.264 is missing; AAC, then
// Opus. Where none of them is there, `exportSupport` says so and the caller
// records in real time instead (MediaRecorder).

import { Mp4Writer } from './mp4-write.js';

/** The H.264 level, as two hex digits, that a size and rate need. */
export function avcLevel(width, height, fps) {
  const mbs = Math.ceil(width / 16) * Math.ceil(height / 16);
  const rate = mbs * fps;
  // [level, max macroblocks a frame, max macroblocks a second]
  const LEVELS = [[0x1f, 3600, 108000], [0x20, 5120, 216000], [0x28, 8192, 245760], [0x2a, 8704, 522240], [0x32, 22080, 589824], [0x33, 36864, 983040], [0x34, 36864, 2073600]];
  for (const [level, frame, second] of LEVELS) if (mbs <= frame && rate <= second) return level.toString(16).padStart(2, '0');
  return '3c';
}

/** The picture codecs to try for a size and rate, best first. */
export function videoCandidates(width, height, fps) {
  const level = avcLevel(width, height, fps);
  return [`avc1.6400${level}`, `avc1.4d00${level}`, `avc1.4200${level}`, 'vp09.00.40.08'];
}

/** A copy of a description, whatever kind of buffer the encoder gave. */
const bytesOf = (d) => (d == null ? null : ArrayBuffer.isView(d) ? new Uint8Array(d.buffer.slice(d.byteOffset, d.byteOffset + d.byteLength)) : new Uint8Array(d.slice(0)));

const even = (n) => Math.max(2, Math.round(n / 2) * 2);

/**
 * The picture encoder this machine has for a size and rate:
 * { config, codec, hardware } or null. Hardware is asked for first.
 */
export async function chooseVideo({ width, height, fps = 30, bitrate = 8e6, hardware = true, Encoder = globalThis.VideoEncoder } = {}) {
  if (typeof Encoder === 'undefined' || !Encoder) return null;
  const w = even(width);
  const h = even(height);
  for (const codec of videoCandidates(w, h, fps)) {
    for (const accel of hardware ? ['prefer-hardware', 'no-preference'] : ['no-preference']) {
      const config = {
        codec, width: w, height: h, bitrate: Math.round(bitrate), framerate: fps,
        hardwareAcceleration: accel, latencyMode: 'quality', bitrateMode: 'variable',
        ...(codec.startsWith('avc1') ? { avc: { format: 'avc' } } : {}),
      };
      try {
        const res = await Encoder.isConfigSupported(config);
        if (res?.supported) return { config, codec, hardware: accel === 'prefer-hardware' };
      } catch { /* not this one */ }
    }
  }
  return null;
}

/** The sound encoder this machine has: { config, codec } or null. AAC first. */
export async function chooseAudio({ sampleRate = 48000, channels = 2, bitrate = 128000, Encoder = globalThis.AudioEncoder } = {}) {
  if (typeof Encoder === 'undefined' || !Encoder) return null;
  for (const codec of ['mp4a.40.2', 'opus']) {
    const config = { codec, sampleRate: codec === 'opus' ? 48000 : sampleRate, numberOfChannels: channels, bitrate: Math.round(bitrate) };
    try {
      const res = await Encoder.isConfigSupported(config);
      if (res?.supported) return { config, codec };
    } catch { /* not this one */ }
  }
  return null;
}

/** What an export can be here: { video, audio, label } or null when there is no picture encoder. */
export async function exportSupport({ width = 1920, height = 1080, fps = 30, withAudio = true } = {}) {
  const video = await chooseVideo({ width, height, fps });
  if (!video) return null;
  const audio = withAudio ? await chooseAudio({}) : null;
  return { video, audio, videoName: video.codec.startsWith('avc1') ? 'H.264' : 'VP9', audioName: audio ? (audio.codec === 'opus' ? 'Opus' : 'AAC') : null };
}

/** Encode a mixed sound whole: [{ data, timestamp, duration }] in microseconds, and the decoder's description. */
export async function encodeSound(buffer, { config, signal, Encoder = globalThis.AudioEncoder, Data = globalThis.AudioData } = {}) {
  const chunks = [];
  let description = null;
  let failed = null;
  const enc = new Encoder({
    output: (chunk, meta) => {
      const data = new Uint8Array(chunk.byteLength);
      chunk.copyTo(data);
      chunks.push({ data, timestamp: chunk.timestamp, duration: chunk.duration || 0 });
      if (meta?.decoderConfig?.description && !description) description = bytesOf(meta.decoderConfig.description);
    },
    error: (err) => { failed = err || new Error('the sound encoder stopped'); },
  });
  enc.configure(config);
  const rate = buffer.sampleRate;
  const channels = Math.min(config.numberOfChannels, buffer.numberOfChannels) || 1;
  const planes = [];
  for (let c = 0; c < config.numberOfChannels; c++) planes.push(buffer.getChannelData(Math.min(c, channels - 1)));
  const per = 1024 * 8;
  for (let at = 0; at < buffer.length; at += per) {
    if (signal?.aborted) break;
    if (failed) break;
    const n = Math.min(per, buffer.length - at);
    const data = new Float32Array(n * planes.length);
    planes.forEach((p, c) => data.set(p.subarray(at, at + n), c * n));
    const ad = new Data({ format: 'f32-planar', sampleRate: rate, numberOfFrames: n, numberOfChannels: planes.length, timestamp: Math.round((at * 1e6) / rate), data });
    enc.encode(ad);
    ad.close();
    while (enc.encodeQueueSize > 16) await new Promise((r) => setTimeout(r, 1));
  }
  if (!failed && !signal?.aborted) await enc.flush().catch((err) => { failed = err; });
  try { enc.close(); } catch { /* closed */ }
  if (signal?.aborted) throw new DOMException('Export cancelled', 'AbortError');
  if (failed) throw failed;
  return { chunks, description };
}

/**
 * Paint, encode and write a movie.
 *
 * `paint(index, seconds)` answers the frame to encode (it may be async);
 * `sound` is the mixed AudioBuffer or null; `sink` takes the file (see
 * Mp4Writer). `hardware: false` keeps to the software encoder (a retry after
 * the hardware one failed). Answers { size, duration, frames, seconds (how
 * long it took), video: { codec, hardware }, audio: { codec } | null }.
 */
export async function encodeMovie({
  sink, width, height, fps = 30, duration, bitrate = 8e6, audioBitrate = 128000,
  paint, sound = null, onProgress, signal, hardware = true, keySeconds = 2,
  VideoEncoderCtor = globalThis.VideoEncoder, VideoFrameCtor = globalThis.VideoFrame,
  AudioEncoderCtor = globalThis.AudioEncoder, AudioDataCtor = globalThis.AudioData,
  now = () => (globalThis.performance ? performance.now() : Date.now()),
}) {
  const started = now();
  const w = even(width);
  const h = even(height);
  const video = await chooseVideo({ width: w, height: h, fps, bitrate, hardware, Encoder: VideoEncoderCtor });
  if (!video) throw new Error('This computer has no video encoder for MP4.');
  let audio = null;
  let sounds = null;
  if (sound) {
    audio = await chooseAudio({ sampleRate: sound.sampleRate, channels: Math.min(2, sound.numberOfChannels), bitrate: audioBitrate, Encoder: AudioEncoderCtor });
    if (audio) sounds = await encodeSound(sound, { config: audio.config, signal, Encoder: AudioEncoderCtor, Data: AudioDataCtor });
  }

  const writer = new Mp4Writer({
    sink,
    video: { codec: video.codec, width: w, height: h },
    audio: audio ? { codec: audio.codec, sampleRate: audio.config.sampleRate, channels: audio.config.numberOfChannels, bitrate: audio.config.bitrate, description: sounds.description } : null,
  });
  await writer.start();

  const pending = [];
  let failed = null;
  let wake = null;
  const poke = () => { const w2 = wake; wake = null; w2?.(); };
  const enc = new VideoEncoderCtor({
    output: (chunk, meta) => {
      if (meta?.decoderConfig) writer.configure('video', { description: meta.decoderConfig.description ? bytesOf(meta.decoderConfig.description) : null, colour: meta.decoderConfig.colorSpace || null });
      const data = new Uint8Array(chunk.byteLength);
      chunk.copyTo(data);
      pending.push({ data, timestamp: chunk.timestamp, duration: chunk.duration || 0, key: chunk.type === 'key' });
      poke();
    },
    error: (err) => { failed = err || new Error('the video encoder stopped'); poke(); },
  });
  enc.addEventListener?.('dequeue', poke);
  enc.configure(video.config);

  let soundAt = 0;
  const writeSoundTo = async (us) => {
    if (!sounds) return;
    while (soundAt < sounds.chunks.length && sounds.chunks[soundAt].timestamp <= us) {
      await writer.add('audio', { ...sounds.chunks[soundAt], key: true });
      soundAt++;
    }
  };
  // Sound and picture interleaved in half seconds: when the sound falls
  // behind the picture, the next half second of it goes in at once.
  const drain = async () => {
    while (pending.length) {
      const c = pending.shift();
      if (sounds && soundAt < sounds.chunks.length && sounds.chunks[soundAt].timestamp <= c.timestamp) await writeSoundTo(c.timestamp + 500000);
      await writer.add('video', c);
    }
  };

  const frames = Math.max(1, Math.round(duration * fps));
  const frameUs = 1e6 / fps;
  const keyEvery = Math.max(1, Math.round(keySeconds * fps));
  // Where the time went, in milliseconds: painting (the caller's decoding
  // included), handing frames to the encoder, writing, and waiting on it.
  const spent = { setup: now() - started, paint: 0, encode: 0, write: 0, wait: 0 };
  try {
    for (let i = 0; i < frames; i++) {
      if (signal?.aborted) throw new DOMException('Export cancelled', 'AbortError');
      if (failed) throw failed;
      let t = now();
      const source = await paint(i, i / fps);
      spent.paint += now() - t;
      t = now();
      const frame = new VideoFrameCtor(source, { timestamp: Math.round(i * frameUs), duration: Math.round(frameUs) });
      try {
        enc.encode(frame, { keyFrame: i % keyEvery === 0 });
      } finally {
        frame.close();
      }
      spent.encode += now() - t;
      t = now();
      await drain();
      spent.write += now() - t;
      t = now();
      // Ten frames in hand keep a graphics encoder busy: three left it waiting (74 frames a second at 1080p here, against 100 to 107 with eight to sixteen).
      while (enc.encodeQueueSize > 10 && !failed) {
        await new Promise((r) => { wake = r; setTimeout(poke, 20); });
        await drain();
      }
      spent.wait += now() - t;
      onProgress?.((i + 1) / frames);
    }
    await enc.flush();
    if (failed) throw failed;
    await drain();
    await writeSoundTo(Infinity);
    const out = await writer.finish();
    const total = now() - started;
    const ms = Object.fromEntries(Object.entries(spent).map(([k, v]) => [k, Math.round(v)]));
    ms.finish = Math.round(total - Object.values(spent).reduce((a, b) => a + b, 0));
    return { size: out.size, duration: out.duration, frames, seconds: total / 1000, spent: ms, video: { codec: video.codec, hardware: video.hardware }, audio: audio ? { codec: audio.codec } : null };
  } finally {
    try { enc.close(); } catch { /* closed */ }
  }
}
