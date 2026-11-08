// One broadcast as a window runs it: the encoders for the shapes in use, the
// one sound, and which streams are on.
//
// Go live opens the broadcast with the streams that are on; the icons change
// that while it runs (setOn), each stream joining or leaving alone. The first
// stream of a shape starts that shape's encoder, and the last to go stops it,
// so with one shape on only one encoder runs. End broadcast ends the lot.

import { shapesOn, onChanges } from '@rutba/stream/bar';
import { LiveEncoder, LiveSound, AUDIO_KBPS } from './live.js';

const SHAPES = ['wide', 'tall'];

export class LiveSession {
  /**
   * @param {object} o
   * @param {object} o.shell
   * @param {object} o.media  the studio's media: what the sound comes from
   * @param {object} o.program  the wide picture's Program
   * @param {number} o.fps
   * @param {number} o.kbps  the bitrate each picture is asked for
   * @param {(() => object|null)} [o.makeTall]  a tall picture's Program, made when the first tall stream is on
   */
  constructor({ shell, media, program, fps, kbps, makeTall = null }) {
    this.shell = shell;
    this.media = media;
    this.program = program;
    this.fps = fps;
    this.kbps = kbps;
    this.makeTall = makeTall;
    this.encoders = new Map();
    this.sound = null;
    this.t0 = 0;
    this.on = [];
    this.streams = [];
    this.tallProgram = null;
  }

  /** What each picture is, for the metadata its streams get first. */
  videoFor(shape) {
    const canvas = this.programFor(shape)?.canvas;
    if (!canvas) return null;
    return { width: canvas.width, height: canvas.height, fps: this.fps, kbps: this.kbps };
  }

  programFor(shape) {
    if (shape === 'wide') return this.program;
    if (!this.tallProgram) this.tallProgram = this.makeTall?.() || null;
    return this.tallProgram;
  }

  /** Check what this computer can encode, then open: throws a sentence when it cannot. */
  async open({ ids, streams }) {
    this.streams = streams;
    this.on = [...ids];
    // Asked first, so a computer that cannot encode says so before anything connects.
    await LiveEncoder.check(this.videoFor('wide'));
    this.t0 = performance.now();
    const shapes = shapesOn(streams, ids);
    const started = await this.shell.live.start({
      ids,
      video: this.videoFor('wide'),
      tall: this.makeTall ? this.videoFor('tall') : null,
      audio: { kbps: AUDIO_KBPS, sampleRate: 48000, channels: 2 },
    });
    this.sound = new LiveSound({ shell: this.shell, rig: this.media, t0: this.t0 });
    this.sound.start();
    for (const shape of shapes) await this.ensure(shape);
    return started;
  }

  /** The encoder for a shape, started if it was not running. */
  async ensure(shape) {
    if (this.encoders.has(shape)) return this.encoders.get(shape);
    const program = this.programFor(shape);
    if (!program) return null;
    program.attach?.();
    const video = this.videoFor(shape);
    const checked = await LiveEncoder.check(video);
    const enc = new LiveEncoder({ shell: this.shell, program, shape, t0: this.t0 });
    await enc.start({ video: checked.video });
    this.encoders.set(shape, enc);
    return enc;
  }

  async retire(shape) {
    const enc = this.encoders.get(shape);
    if (!enc) return;
    this.encoders.delete(shape);
    await enc.stop();
    this.programFor(shape)?.detach?.();
  }

  /** The icons changed: each stream that changed joins or leaves, and the encoders follow. */
  async setOn(ids, streams) {
    const { start, stop } = onChanges(this.on, ids);
    this.on = [...ids];
    this.streams = streams;
    for (const id of stop) await this.shell.live.stopOne({ id });
    const shapes = shapesOn(streams, ids);
    for (const shape of SHAPES) {
      if (shapes.includes(shape)) await this.ensure(shape);
      else await this.retire(shape);
    }
    for (const id of start) await this.shell.live.startOne({ id });
  }

  /** A stream's key changed while it was on: it starts again with the new one. */
  async restart(id) {
    await this.shell.live.stopOne({ id });
    await this.shell.live.startOne({ id });
  }

  /** What the broadcast said: the bitrates to follow, and any shape a late joiner is waiting on. */
  onStatus({ rates = {}, keyFor = [] } = {}) {
    for (const [shape, enc] of this.encoders) if (rates[shape]?.kbps) enc.setBitrate(rates[shape].kbps);
    for (const shape of keyFor) this.encoders.get(shape)?.askKey();
  }

  /** The end: encoders and sound stopped, every stream closed. */
  async end() {
    for (const shape of [...this.encoders.keys()]) await this.retire(shape).catch(() => {});
    await this.sound?.stop().catch(() => {});
    this.sound = null;
    await this.shell.live.stop().catch(() => {});
  }
}
