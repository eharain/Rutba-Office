// A source's pictures, decoded on demand at any time asked for.
//
// The export paints a frame, then the next, as fast as the encoder takes
// them, so it cannot play a <video> element along in real time. Instead each
// source is demuxed here (MP4, QuickTime or WebM) and its samples handed to
// WebCodecs' VideoDecoder in decode order; `frameAt(t)` answers the frame on
// screen at `t`, the latest one that starts at or before it. Asked in order,
// as an export asks, every sample is decoded once; asked earlier, or much
// later, the decoder starts again from the key frame before `t`. Decoded
// frames are closed as soon as a later one is shown, so a long source never
// holds more than a handful.

import { readMp4, presentationTimes } from './mp4-read.js';
import { readWebm } from './webm-read.js';

/**
 * The picture track of a file, as the decoder needs it: { codec,
 * description, width, height, samples: [{ offset, size, time, key }] in
 * decode order, times in seconds }, or null when there is none to read.
 */
export function demuxVideo(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (b.length < 12) return null;
  if (new TextDecoder('latin1').decode(b.subarray(4, 8)) === 'ftyp') {
    const file = readMp4(b);
    const track = file?.tracks.find((t) => t.kind === 'video' && t.codec && t.samples.length);
    if (!track) return null;
    const times = presentationTimes(track);
    return {
      container: 'mp4',
      codec: track.codec,
      description: track.description || null,
      width: track.width,
      height: track.height,
      samples: times.map((s) => ({ offset: s.offset, size: s.size, time: s.time, key: s.key })),
    };
  }
  const file = readWebm(b);
  const track = file?.tracks.find((t) => t.kind === 'video' && t.codec && t.frames.length);
  if (!track) return null;
  return {
    container: 'webm',
    codec: track.codec,
    description: track.description || null,
    width: track.width,
    height: track.height,
    samples: track.frames.map((f) => ({ offset: f.offset, size: f.size, time: f.time, key: f.key })),
  };
}

const EPS = 0.0005;

/** Frames of one source, by time. Make one with `FrameReader.open`. */
export class FrameReader {
  constructor(bytes, track, { Decoder = globalThis.VideoDecoder, Chunk = globalThis.EncodedVideoChunk, ahead = 4, hardware = true, stallMs = 15000, now = () => (globalThis.performance ? performance.now() : Date.now()) } = {}) {
    this.bytes = bytes;
    this.track = track;
    this.samples = track.samples;
    this.Decoder = Decoder;
    this.Chunk = Chunk;
    this.ahead = ahead;
    // An export started again on the processor asks for the processor's decoder too.
    this.config = { codec: track.codec, codedWidth: track.width || undefined, codedHeight: track.height || undefined, ...(track.description ? { description: track.description } : {}), ...(hardware ? {} : { hardwareAcceleration: 'prefer-software' }) };
    this.stallMs = stallMs;
    this.now = now;
    this.queue = [];
    this.current = null;
    this.next = 0;
    this.flushed = false;
    this.failed = null;
    this.waiter = null;
    this.decoded = 0;
    this.restarts = 0;
    // The first frame shown ahead of its time (a source that starts late): not a reason to start again.
    this.early = false;
    this.start();
  }

  /**
   * A reader for `bytes`, or null when the file has no picture this
   * machine's decoder takes (the caller then falls back to seeking a
   * <video>).
   */
  static async open(bytes, options = {}) {
    const Decoder = options.Decoder || globalThis.VideoDecoder;
    if (typeof Decoder === 'undefined' || !Decoder) return null;
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let track;
    try {
      track = demuxVideo(b);
    } catch {
      return null;
    }
    if (!track || !track.samples.length) return null;
    const base = { codec: track.codec, codedWidth: track.width || undefined, codedHeight: track.height || undefined, ...(track.description ? { description: track.description } : {}) };
    const takes = async (config) => {
      try {
        return Boolean((await Decoder.isConfigSupported(config))?.supported);
      } catch {
        return false;
      }
    };
    // Asked with the config the reader will use (review of 11 Nov). A codec
    // with no decoder on the processor keeps the graphics one: drawing the
    // clip slowly beats drawing nothing and calling the export done.
    let hardware = options.hardware !== false;
    if (!hardware && !(await takes({ ...base, hardwareAcceleration: 'prefer-software' }))) hardware = true;
    if (hardware && !(await takes(base))) return null;
    return new FrameReader(b, track, { ...options, hardware });
  }

  get width() { return this.track.width; }
  get height() { return this.track.height; }

  start() {
    this.decoder = new this.Decoder({
      output: (frame) => {
        this.decoded++;
        this.lastOut = this.now();
        // Frames come out in the order they show.
        this.queue.push(frame);
        this.wake();
      },
      error: (err) => {
        this.failed = err || new Error('the decoder stopped');
        this.wake();
      },
    });
    this.decoder.addEventListener?.('dequeue', () => this.wake());
    this.decoder.configure(this.config);
  }

  wake() {
    const w = this.waiter;
    this.waiter = null;
    w?.();
  }

  waitEvent() {
    return new Promise((resolve) => {
      this.waiter = resolve;
      // A decoder that says nothing for a while is asked again rather than waited on for ever.
      setTimeout(() => { if (this.waiter === resolve) this.wake(); }, 50);
    });
  }

  /** Start again from the key frame at or before `t`. */
  restart(t) {
    this.restarts++;
    for (const f of this.queue) f.close();
    this.queue = [];
    try { this.decoder.reset(); } catch { /* closed */ }
    try { this.decoder.configure(this.config); } catch { this.start(); }
    this.next = this.keyBefore(t);
    this.flushed = false;
  }

  /** The index of the last key frame starting at or before `t` (the first sample when there is none). */
  keyBefore(t) {
    let k = 0;
    for (let i = 0; i < this.samples.length; i++) {
      const s = this.samples[i];
      if (s.key && s.time <= t + EPS) k = i;
      if (s.time > t + 5) break;
    }
    return k;
  }

  feed() {
    const s = this.samples[this.next++];
    this.decoder.decode(new this.Chunk({
      type: s.key ? 'key' : 'delta',
      timestamp: Math.round(s.time * 1e6),
      data: this.bytes.subarray(s.offset, s.offset + s.size),
    }));
  }

  /**
   * The frame on screen at `t` seconds of the source. Owned by the reader:
   * draw it before the next call, never close it.
   */
  async frameAt(t) {
    if (this.failed) return this.current;
    const time = (f) => f.timestamp / 1e6;
    // A decoder owed samples that sends no frame for `stallMs` has stalled: said, not waited on for ever (review of 11 Nov).
    const asked = this.now();
    // Earlier than what is shown, or far beyond what is decoded: start again near `t`.
    const shown = this.current ? time(this.current) : null;
    if (shown !== null && t < shown - EPS && !this.early) {
      this.current.close();
      this.current = null;
      this.restart(t);
    } else {
      const k = this.keyBefore(t);
      if (k > this.next && this.samples[k].time > (shown ?? -Infinity) + 1) {
        if (this.current) { this.current.close(); this.current = null; }
        this.restart(t);
      }
    }
    for (;;) {
      if (this.failed) return this.current;
      while (this.queue.length && time(this.queue[0]) <= t + EPS) {
        this.current?.close();
        this.current = this.queue.shift();
        this.early = false;
      }
      if (this.queue.length) {
        // The next frame starts after `t`: what is shown is right. Before the
        // first frame of a source that starts late, show that first frame.
        if (!this.current) {
          this.current = this.queue.shift();
          this.early = true;
        }
        return this.current;
      }
      if (this.next < this.samples.length) {
        if (this.decoder.decodeQueueSize < this.ahead) {
          this.feed();
          continue;
        }
        await this.waitEvent();
        if (this.now() - Math.max(asked, this.lastOut || 0) > this.stallMs) {
          const err = new Error(`the video decoder sent nothing for ${Math.round(this.stallMs / 1000)} seconds`);
          err.name = 'StalledError';
          err.stalled = true;
          err.what = 'decoder';
          throw err;
        }
        continue;
      }
      if (!this.flushed) {
        this.flushed = true;
        // The flush has the same deadline, so a decoder that never finishes it is said to have stalled.
        let timer;
        const late = new Promise((r) => { timer = setTimeout(() => r('late'), this.stallMs); });
        try {
          if ((await Promise.race([this.decoder.flush(), late])) === 'late') {
            const err = new Error(`the video decoder did not finish in ${Math.round(this.stallMs / 1000)} seconds`);
            err.name = 'StalledError';
            err.stalled = true;
            err.what = 'decoder';
            throw err;
          }
        } catch (err) {
          if (err?.stalled) throw err;
          this.failed = err;
        } finally {
          clearTimeout(timer);
        }
        continue;
      }
      return this.current;
    }
  }

  close() {
    for (const f of this.queue) f.close();
    this.queue = [];
    this.current?.close();
    this.current = null;
    try { this.decoder.close(); } catch { /* already closed */ }
  }
}
