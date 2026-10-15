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
  constructor(bytes, track, { Decoder = globalThis.VideoDecoder, Chunk = globalThis.EncodedVideoChunk, ahead = 4 } = {}) {
    this.bytes = bytes;
    this.track = track;
    this.samples = track.samples;
    this.Decoder = Decoder;
    this.Chunk = Chunk;
    this.ahead = ahead;
    this.config = { codec: track.codec, codedWidth: track.width || undefined, codedHeight: track.height || undefined, ...(track.description ? { description: track.description } : {}) };
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
    try {
      const probe = await Decoder.isConfigSupported({ codec: track.codec, codedWidth: track.width || undefined, codedHeight: track.height || undefined, ...(track.description ? { description: track.description } : {}) });
      if (!probe?.supported) return null;
    } catch {
      return null;
    }
    return new FrameReader(b, track, options);
  }

  get width() { return this.track.width; }
  get height() { return this.track.height; }

  start() {
    this.decoder = new this.Decoder({
      output: (frame) => {
        this.decoded++;
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
        continue;
      }
      if (!this.flushed) {
        this.flushed = true;
        try { await this.decoder.flush(); } catch (err) { this.failed = err; }
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
