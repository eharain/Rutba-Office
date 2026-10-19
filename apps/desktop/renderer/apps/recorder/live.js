// Go Live's encoder: the program, as it is drawn, made into H.264 and AAC by
// the platform's own encoders (WebCodecs, on the graphics chip where it has
// one) and sent to the main process a piece at a time, where every
// destination is given the same bytes (main/live.js).
//
// The picture is encoded at a constant bitrate with a key frame every two
// seconds, as YouTube, Facebook and Twitch ask. A frame that comes while the
// encoder is still busy is dropped and counted rather than queued, so the
// stream stays live instead of falling behind. The sound is the program's
// mix (the microphone and the computer's sound), read from its track.

import { t } from '@rutba/office-ui';

const KEY_EVERY_US = 2_000_000;
const AUDIO_KBPS = 128;

/** The H.264 profiles tried, the level each needs for its size and rate. */
function avcCandidates(width, height, fps) {
  const big = width * height > 1280 * 720;
  const fast = fps > 30;
  const level = big ? (fast ? '2a' : '28') : (fast ? '20' : '1f');
  return [`avc1.6400${level}`, `avc1.4d00${level}`, `avc1.4200${level}`, 'avc1.640028', 'avc1.4d0028', 'avc1.42e01f'];
}

/** The picture encoder this computer has for a live stream: { config, hardware } or null. */
export async function chooseLiveVideo({ width, height, fps, kbps }) {
  if (typeof VideoEncoder === 'undefined') return null;
  for (const codec of avcCandidates(width, height, fps)) {
    for (const accel of ['prefer-hardware', 'no-preference']) {
      const config = {
        codec, width, height, framerate: fps, bitrate: kbps * 1000, bitrateMode: 'constant',
        latencyMode: 'realtime', hardwareAcceleration: accel, avc: { format: 'avc' },
      };
      try {
        if ((await VideoEncoder.isConfigSupported(config))?.supported) return { config, hardware: accel === 'prefer-hardware' };
      } catch { /* not this one */ }
    }
  }
  return null;
}

/** The sound encoder: AAC, which RTMP's platforms take, or null. */
export async function chooseLiveAudio({ sampleRate = 48000, channels = 2 } = {}) {
  if (typeof AudioEncoder === 'undefined') return null;
  const config = { codec: 'mp4a.40.2', sampleRate, numberOfChannels: channels, bitrate: AUDIO_KBPS * 1000, aac: { format: 'aac' } };
  try {
    return (await AudioEncoder.isConfigSupported(config))?.supported ? { config } : null;
  } catch {
    return null;
  }
}

export class LiveEncoder {
  /** @param {{ shell, program, rig }} o */
  constructor({ shell, program, rig }) {
    this.shell = shell;
    this.program = program;
    this.rig = rig;
    this.video = null;
    this.audio = null;
    this.reader = null;
    this.off = null;
    this.t0 = 0;
    this.lastKey = -Infinity;
    this.stats = { frames: 0, dropped: 0, bytes: 0, audioFrames: 0, hardware: false };
    this.running = false;
    this.sentVideoConfig = false;
    this.sentAudioConfig = false;
  }

  send(piece) {
    this.stats.bytes += piece.data.byteLength;
    this.shell.live.send(piece).catch(() => {});
  }

  /** Check what this computer can encode, before anything connects: { video, audio } or a thrown sentence. */
  static async check({ width, height, fps, kbps }) {
    const video = await chooseLiveVideo({ width, height, fps, kbps });
    if (!video) throw new Error(t('This computer cannot make the H.264 picture a live stream needs.'));
    const audio = await chooseLiveAudio({});
    if (!audio) throw new Error(t('This computer cannot make the AAC sound a live stream needs.'));
    return { video, audio };
  }

  /** Start encoding the program. `video` is from check(); answers once the first frame is asked for. */
  async start({ video }) {
    const canvas = this.program.canvas;
    this.running = true;
    this.t0 = performance.now();
    this.stats.hardware = video.hardware;

    this.video = new VideoEncoder({
      output: (chunk, meta) => {
        const description = meta?.decoderConfig?.description;
        if (description && !this.sentVideoConfig) {
          this.sentVideoConfig = true;
          this.send({ kind: 'video', timestamp: chunk.timestamp / 1000, data: new Uint8Array(description.buffer ? description.buffer.slice(description.byteOffset, description.byteOffset + description.byteLength) : description), config: true });
        }
        const data = new Uint8Array(chunk.byteLength);
        chunk.copyTo(data);
        this.send({ kind: 'video', timestamp: chunk.timestamp / 1000, data, key: chunk.type === 'key' });
      },
      error: (err) => { this.error = err; },
    });
    this.video.configure({ ...video.config, width: canvas.width, height: canvas.height });

    this.off = this.program.onFrame((cv) => {
      if (!this.running || this.video.state !== 'configured') return;
      if (this.video.encodeQueueSize > 2) { this.stats.dropped += 1; return; }
      const ts = Math.round((performance.now() - this.t0) * 1000);
      const keyFrame = ts - this.lastKey >= KEY_EVERY_US;
      if (keyFrame) this.lastKey = ts;
      const frame = new VideoFrame(cv, { timestamp: ts });
      try {
        this.video.encode(frame, { keyFrame });
        this.stats.frames += 1;
      } finally {
        frame.close();
      }
    });

    this.startSound();
  }

  /** The program's sound, read off its track and encoded as AAC on its own clock, lined up with the picture's. */
  async startSound() {
    const track = this.rig.mix(['mic', 'system']) || this.rig.silence();
    if (!track || typeof MediaStreamTrackProcessor === 'undefined') return;
    this.soundTrack = track;
    const processor = new MediaStreamTrackProcessor({ track });
    this.reader = processor.readable.getReader();
    let offset = null;
    for (;;) {
      let got;
      try {
        got = await this.reader.read();
      } catch {
        break;
      }
      if (got.done || !this.running) { got.value?.close(); break; }
      const data = got.value;
      if (!this.audio) {
        this.audio = new AudioEncoder({
          output: (chunk, meta) => {
            const description = meta?.decoderConfig?.description;
            if (description && !this.sentAudioConfig) {
              this.sentAudioConfig = true;
              this.send({ kind: 'audio', timestamp: 0, data: new Uint8Array(description.buffer ? description.buffer.slice(description.byteOffset, description.byteOffset + description.byteLength) : description), config: true });
            }
            const bytes = new Uint8Array(chunk.byteLength);
            chunk.copyTo(bytes);
            this.send({ kind: 'audio', timestamp: Math.max(0, (chunk.timestamp - (offset ?? 0)) / 1000), data: bytes });
          },
          error: (err) => { this.error = err; },
        });
        this.audio.configure({ codec: 'mp4a.40.2', sampleRate: data.sampleRate, numberOfChannels: Math.min(2, data.numberOfChannels), bitrate: AUDIO_KBPS * 1000, aac: { format: 'aac' } });
      }
      // The track's clock lined up with the picture's: its first sample at the moment it came.
      if (offset == null) offset = data.timestamp - Math.round((performance.now() - this.t0) * 1000);
      try {
        if (this.audio.state === 'configured') { this.audio.encode(data); this.stats.audioFrames += 1; }
      } finally {
        data.close();
      }
    }
  }

  /** Stop: the frames no longer taken, what is in the encoders let go. */
  async stop() {
    this.running = false;
    this.off?.();
    try { await this.reader?.cancel(); } catch { /* done */ }
    try { this.soundTrack?.stop(); } catch { /* ended */ }
    for (const enc of [this.video, this.audio]) {
      if (!enc || enc.state === 'closed') continue;
      try { await enc.flush(); } catch { /* closing */ }
      try { enc.close(); } catch { /* closed */ }
    }
  }
}

export { AUDIO_KBPS };
