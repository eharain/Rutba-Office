// One broadcast: every stream going out, each joining and leaving on its own.
//
// A broadcast is open from Go live until End broadcast, whether or not anything
// is going out in between: turning every stream off leaves it open, and turning
// one back on sends to it again (owner, 8 Nov: "the icons disabling all
// destinations sets all offline but the exclusive button should provide
// eventual control").
//
// It carries one picture per shape, wide and tall, each from its own encoder at
// its own bitrate, and one sound shared by both. The same bytes go to every
// stream of a shape. The metadata and the two configurations are kept, so a
// stream that joins late, or comes back after its line dropped, is given them
// first and then waits for a key frame, and the encoder is asked for one at
// once rather than leaving it to the next of its two-second ones. A stream that
// drops is tried again, after 2, 4, 8, 15 and then 30 seconds. One the server
// refused is said once and not tried again.
//
// Node only, and no Electron: main/live.js holds the keys and the window.

import { videoBody, audioBody, metadataBody } from './flv.js';
import { BitrateGovernor } from './adapt.js';

const BACKOFF_MS = [2000, 4000, 8000, 15000, 30000];

export const STREAM_SHAPES = ['wide', 'tall'];

function shapeFrom(video, audio) {
  const asked = Number(video.kbps) || 4500;
  return {
    meta: metadataBody({ width: video.width, height: video.height, fps: video.fps, videoKbps: video.kbps, audioKbps: audio.kbps, sampleRate: audio.sampleRate || 48000, channels: audio.channels || 2 }),
    videoConfig: null,
    asked,
    governor: new BitrateGovernor({ target: asked }),
    kbps: asked,
  };
}

export class Broadcast {
  /**
   * @param {object} o
   * @param {{ width, height, fps, kbps }} o.video  the wide picture
   * @param {{ width, height, fps, kbps }} [o.tall]  the tall picture, when one can go out
   * @param {{ kbps, sampleRate, channels }} [o.audio]
   * @param {(o: { url, key }) => import('./rtmp.js').RtmpPublisher} o.makePublisher
   * @param {(status: object) => void} [o.onChange]  told whenever anything a window shows has changed
   * @param {number[]} [o.backoffMs]  the waits before a dropped stream is tried again
   * @param {number} [o.tickMs]  how often the slowest line is looked at
   */
  constructor({ video = {}, tall = null, audio = {}, makePublisher, onChange = () => {}, backoffMs = BACKOFF_MS, tickMs = 2000 }) {
    this.running = true;
    this.makePublisher = makePublisher;
    this.onChange = onChange;
    this.backoffMs = backoffMs;
    this.dest = new Map();
    this.shapes = new Map([['wide', shapeFrom(video, audio)]]);
    if (tall) this.shapes.set('tall', shapeFrom(tall, audio));
    this.audioConfig = null;
    this.last = 0;
    this.keyFor = new Set();
    this.ticker = setInterval(() => this.tick(), tickMs);
    this.ticker.unref?.();
  }

  /** Each shape's rate: { wide: { kbps, target }, tall: ... }. */
  rates() {
    const out = {};
    for (const [name, sh] of this.shapes) out[name] = { kbps: sh.kbps, target: sh.governor.target };
    return out;
  }

  /** What the windows are shown. */
  status() {
    const rates = this.rates();
    return {
      live: this.running,
      destinations: [...this.dest.values()].map((s) => ({
        id: s.id,
        name: s.name,
        shape: s.shape,
        state: s.state,
        detail: s.detail,
        refused: Boolean(s.refused),
        bytes: s.pub ? s.pub.sent : s.bytes,
        dropped: s.pub ? s.pub.dropped : s.dropped,
        since: s.since,
        reconnects: s.reconnects,
      })),
      rates,
      kbps: rates.wide?.kbps || 0,
      target: rates.wide?.target || 0,
    };
  }

  /** Told to the window: the status, and which shapes' encoders should make a key frame now. */
  tell() {
    const status = this.status();
    if (this.keyFor.size) { status.keyFor = [...this.keyFor]; this.keyFor.clear(); }
    this.onChange(status);
  }

  /** A stream's slowest allowance within its shape: a platform's ceiling lowers the encoder's. */
  applyCeiling(shape) {
    const sh = this.shapes.get(shape);
    if (!sh) return;
    let cap = sh.asked;
    for (const s of this.dest.values()) if (s.shape === shape && s.maxKbps) cap = Math.min(cap, s.maxKbps);
    sh.governor.target = cap;
    sh.governor.kbps = Math.min(sh.governor.kbps, cap);
    sh.kbps = sh.governor.kbps;
  }

  /**
   * A stream joins: { id, name, url, key, shape, maxKbps }. One already there
   * and still trying is left as it is; one that was refused starts afresh.
   */
  add({ id, name, url, key, shape = 'wide', maxKbps = 0 }) {
    if (!this.running) throw new Error('The broadcast has ended.');
    const was = this.dest.get(id);
    if (was && !was.stopped) return was;
    if (was) this.dest.delete(id);
    const s = {
      id, name, url, key,
      // A broadcast opened with no tall picture sends a tall stream the wide one rather than nothing.
      shape: this.shapes.has(shape) ? shape : 'wide',
      maxKbps,
      state: 'waiting', detail: '', refused: false, pub: null, bytes: 0, dropped: 0, since: 0, reconnects: 0, stopped: false, retry: null, seenDropped: 0,
    };
    this.dest.set(id, s);
    this.applyCeiling(s.shape);
    this.connect(s);
    this.tell();
    return s;
  }

  /** A stream leaves: its connection closed, nothing tried again, the others untouched. */
  remove(id) {
    const s = this.dest.get(id);
    if (!s) return false;
    this.dest.delete(id);
    s.stopped = true;
    clearTimeout(s.retry);
    s.pub?.close('stopped');
    this.applyCeiling(s.shape);
    this.tell();
    return true;
  }

  /** Bring one stream's publisher up, and give it what it needs before frames. */
  async connect(s) {
    if (!this.running || s.stopped) return;
    s.state = s.reconnects ? 'reconnecting' : 'connecting';
    s.detail = '';
    this.tell();
    let pub;
    try {
      pub = this.makePublisher({ url: s.url, key: s.key });
    } catch (err) {
      s.state = 'error';
      s.detail = err.message;
      this.tell();
      return;
    }
    s.pub = pub;
    s.seenDropped = 0;
    pub.on('error', (err) => { s.detail = err.message; });
    pub.on('close', () => {
      if (s.pub !== pub) return;
      s.bytes += pub.sent;
      s.dropped += pub.dropped;
      s.pub = null;
      if (!this.running || s.stopped) return;
      // Dropped while live: try again, waiting longer each time.
      const wait = this.backoffMs[Math.min(s.reconnects, this.backoffMs.length - 1)];
      s.reconnects += 1;
      s.state = 'reconnecting';
      s.detail = s.detail || 'The connection dropped.';
      this.tell();
      s.retry = setTimeout(() => this.connect(s), wait);
    });
    try {
      await pub.start();
      if (!this.running || s.stopped) { pub.close(); return; }
      const sh = this.shapes.get(s.shape);
      s.state = 'live';
      s.detail = '';
      s.since = s.since || Date.now();
      pub.sendMetadata(sh.meta);
      if (sh.videoConfig) pub.sendVideo(sh.videoConfig, this.last, { config: true });
      if (this.audioConfig) pub.sendAudio(this.audioConfig, this.last);
      // It waits for a key frame before any picture: ask for one now.
      this.keyFor.add(s.shape);
      this.tell();
    } catch (err) {
      // A refused key is said once and not tried again; anything else is tried again by the close above.
      s.detail = err.message;
      if (/refused/.test(err.message)) {
        s.stopped = true;
        s.refused = true;
        clearTimeout(s.retry);
        s.state = 'error';
        this.applyCeiling(s.shape);
        this.tell();
      }
    }
  }

  /** One encoded piece: { kind, shape, timestamp (ms), data, key, config }. A configuration is kept for those who join later. */
  send({ kind, shape = 'wide', timestamp = 0, data, key = false, config = false } = {}) {
    if (!this.running || !data) return { live: 0 };
    const bytes = Buffer.from(data.buffer ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength) : data);
    this.last = Math.max(this.last, Math.round(timestamp));
    let live = 0;
    if (kind === 'video') {
      const sh = this.shapes.get(shape);
      if (!sh) return { live: 0 };
      const body = videoBody({ data: bytes, key, config });
      if (config) sh.videoConfig = body;
      for (const s of this.dest.values()) if (s.shape === shape && s.pub?.state === 'live' && s.pub.sendVideo(body, timestamp, { config })) live += 1;
    } else if (kind === 'audio') {
      const body = audioBody({ data: bytes, config });
      if (config) this.audioConfig = body;
      for (const s of this.dest.values()) if (s.pub?.state === 'live' && s.pub.sendAudio(body, timestamp)) live += 1;
    }
    return { live };
  }

  /** The picture's bitrate follows each shape's slowest stream (adapt.js), so a slow tall stream leaves wide alone. */
  tick() {
    if (!this.running) return;
    for (const [name, sh] of this.shapes) {
      const looks = [...this.dest.values()].filter((s) => s.shape === name && s.pub?.state === 'live').map((s) => {
        const dropped = Math.max(0, s.pub.dropped - (s.seenDropped || 0));
        s.seenDropped = s.pub.dropped;
        return { dropped, queued: (s.pub.socket?.writableLength || 0) / (s.pub.maxQueueBytes || 1) };
      });
      if (looks.length) sh.kbps = sh.governor.tick(looks);
    }
    this.tell();
  }

  /** The end: every stream closed, nothing tried again. */
  close() {
    if (!this.running) return this.status();
    this.running = false;
    clearInterval(this.ticker);
    for (const s of this.dest.values()) {
      s.stopped = true;
      clearTimeout(s.retry);
      s.pub?.close('stopped');
      s.state = 'ended';
    }
    const out = this.status();
    this.onChange(out);
    return out;
  }
}
