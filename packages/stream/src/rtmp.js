// An RTMP publisher: one live stream to one server, over TCP or TLS.
//
// It does what a streaming encoder does with YouTube, Facebook, Twitch or any
// RTMP server: the plain handshake, a chunk size of 4096, connect to the app,
// releaseStream and FCPublish, createStream, publish the key as live, then the
// metadata and the media as messages on that stream. It answers what the
// server sends while it runs: its chunk size, the window it wants
// acknowledged, and its pings.
//
// Back-pressure: the socket's queue is watched, and while it holds more than
// a few seconds of the stream, picture frames are dropped up to the next key
// frame (sound is never dropped). `dropped` counts them, for the window.
//
// Nothing here knows about windows or encoders; the main process's live
// service drives one of these per destination. `connect` is the socket
// maker, so the tests can hand it a local server.

import net from 'node:net';
import tls from 'node:tls';
import crypto from 'node:crypto';
import { EventEmitter } from 'node:events';
import { encode, decode } from './amf0.js';
import { AUDIO, VIDEO, DATA, isKeyBody } from './flv.js';
import { parseRtmpUrl } from './endpoints.js';

const SET_CHUNK_SIZE = 1;
const ABORT = 2;
const ACK = 3;
const USER_CONTROL = 4;
const WINDOW_ACK = 5;
const PEER_BANDWIDTH = 6;
const COMMAND = 20;

const OUT_CHUNK = 4096;
const CSID = { control: 2, command: 3, audio: 4, video: 6, data: 5 };
const HANDSHAKE = 1536;

/** The default socket maker: TLS for rtmps, TCP for rtmp. */
function defaultConnect({ secure, host, port }) {
  return secure ? tls.connect({ host, port, servername: host, ALPNProtocols: [] }) : net.connect({ host, port });
}

/** One RTMP message as chunks: a full header first, then continuations. */
export function chunkMessage({ csid, type, streamId = 0, timestamp = 0, body, chunkSize = OUT_CHUNK }) {
  const ts = Math.max(0, Math.round(timestamp)) >>> 0;
  const extended = ts >= 0xffffff;
  const basic = (fmt) => (csid < 64 ? Buffer.from([(fmt << 6) | csid]) : Buffer.from([(fmt << 6) | 0, csid - 64]));
  const head = Buffer.alloc(11);
  head.writeUIntBE(extended ? 0xffffff : ts, 0, 3);
  head.writeUIntBE(body.length, 3, 3);
  head[6] = type;
  head.writeUInt32LE(streamId, 7);
  const ext = extended ? (() => { const b = Buffer.alloc(4); b.writeUInt32BE(ts); return b; })() : Buffer.alloc(0);
  const parts = [basic(0), head, ext];
  for (let at = 0; at < body.length || at === 0; at += chunkSize) {
    if (at > 0) parts.push(basic(3), ext);
    parts.push(body.subarray(at, Math.min(body.length, at + chunkSize)));
    if (body.length === 0) break;
  }
  return Buffer.concat(parts);
}

/**
 * Reads chunks as they arrive and hands back whole messages:
 * { type, streamId, timestamp, body }. The peer's chunk size changes when it
 * says so.
 */
export class ChunkReader {
  constructor() {
    this.buf = Buffer.alloc(0);
    this.chunkSize = 128;
    this.streams = new Map(); // csid -> { timestamp, length, type, streamId, delta, parts, got }
  }

  push(data) {
    this.buf = this.buf.length ? Buffer.concat([this.buf, data]) : Buffer.from(data);
    const out = [];
    for (;;) {
      const msg = this.next();
      if (msg === null) break;
      if (msg) out.push(msg);
    }
    return out;
  }

  /** One chunk off the front: a message when it completes one, false when it does not, null when there is not enough yet. */
  next() {
    const b = this.buf;
    if (b.length < 1) return null;
    let at = 0;
    const fmt = b[0] >> 6;
    let csid = b[0] & 0x3f;
    at = 1;
    if (csid === 0) { if (b.length < 2) return null; csid = 64 + b[1]; at = 2; }
    else if (csid === 1) { if (b.length < 3) return null; csid = 64 + b[1] + b[2] * 256; at = 3; }
    const sizes = [11, 7, 3, 0];
    if (b.length < at + sizes[fmt]) return null;
    const prev = this.streams.get(csid) || { timestamp: 0, length: 0, type: 0, streamId: 0, delta: 0, parts: [], got: 0 };
    const s = { ...prev };
    let ts = null;
    if (fmt <= 2) ts = b.readUIntBE(at, 3);
    if (fmt <= 1) { s.length = b.readUIntBE(at + 3, 3); s.type = b[at + 6]; }
    if (fmt === 0) s.streamId = b.readUInt32LE(at + 7);
    at += sizes[fmt];
    const extended = (fmt <= 2 && ts === 0xffffff) || (fmt === 3 && prev.extended);
    if (extended) { if (b.length < at + 4) return null; ts = b.readUInt32BE(at); at += 4; }
    s.extended = extended;
    const starting = prev.got === 0 || fmt !== 3;
    if (starting) {
      if (fmt === 0) { s.timestamp = ts; s.delta = 0; }
      else if (fmt === 1 || fmt === 2) { s.delta = ts; s.timestamp = prev.timestamp + ts; }
      else s.timestamp = prev.timestamp + (prev.delta || 0);
      s.parts = [];
      s.got = 0;
    }
    const want = Math.min(this.chunkSize, s.length - s.got);
    if (b.length < at + want) return null;
    s.parts = [...s.parts, b.subarray(at, at + want)];
    s.got += want;
    this.buf = b.subarray(at + want);
    if (s.got >= s.length) {
      const body = Buffer.concat(s.parts);
      this.streams.set(csid, { ...s, parts: [], got: 0 });
      // A new chunk size holds from the very next chunk, which may already be in this buffer.
      if (s.type === SET_CHUNK_SIZE && body.length >= 4) this.chunkSize = body.readUInt32BE(0) & 0x7fffffff;
      return { type: s.type, streamId: s.streamId, timestamp: s.timestamp, body, csid };
    }
    this.streams.set(csid, s);
    return false;
  }
}

/**
 * One live stream: `new RtmpPublisher({ url, key })`, then `await start()`.
 * Events: 'state' ({ state, detail }), 'error' (Error), 'close'.
 */
export class RtmpPublisher extends EventEmitter {
  constructor({ url, key, connect = defaultConnect, timeoutMs = 15000, maxQueueBytes = 4 * 1024 * 1024 }) {
    super();
    this.target = parseRtmpUrl(url);
    if (!this.target) throw new Error('That is not an RTMP address (rtmp:// or rtmps://, then a server and an app).');
    if (!key) throw new Error('A stream key is needed.');
    this.key = String(key);
    this.connectFn = connect;
    this.timeoutMs = timeoutMs;
    this.maxQueueBytes = maxQueueBytes;
    this.socket = null;
    this.reader = new ChunkReader();
    this.state = 'idle';
    this.txn = 0;
    this.pending = new Map();
    this.streamId = 0;
    this.received = 0;
    this.acked = 0;
    this.window = 2500000;
    this.sent = 0;
    this.dropped = 0;
    this.waitingForKey = false;
    this.closed = false;
  }

  setState(state, detail = '') {
    this.state = state;
    this.emit('state', { state, detail });
  }

  write(buf) {
    if (!this.socket || this.socket.destroyed) return false;
    this.sent += buf.length;
    return this.socket.write(buf);
  }

  send(csid, type, body, { timestamp = 0, streamId = 0 } = {}) {
    return this.write(chunkMessage({ csid, type, body, timestamp, streamId }));
  }

  command(name, args = [], { streamId = 0, csid = CSID.command, reply = true } = {}) {
    const txn = reply ? ++this.txn : 0;
    const body = encode(name, txn, ...args);
    const answered = reply ? new Promise((resolve, reject) => this.pending.set(txn, { resolve, reject, name })) : Promise.resolve();
    this.send(csid, COMMAND, body, { streamId });
    return answered;
  }

  /** Connect, handshake, connect to the app, create the stream and publish the key. Answers when the server says it is publishing. */
  async start() {
    this.setState('connecting', `${this.target.host}:${this.target.port}`);
    const socket = this.connectFn(this.target);
    this.socket = socket;
    socket.setNoDelay?.(true);
    const timer = setTimeout(() => this.fail(new Error('The server did not answer in time.')), this.timeoutMs);
    try {
      await new Promise((resolve, reject) => {
        socket.once(this.target.secure && socket instanceof tls.TLSSocket ? 'secureConnect' : 'connect', resolve);
        socket.once('error', reject);
      });
      socket.on('error', (err) => this.fail(err));
      socket.on('close', () => this.onClose());
      await this.handshake();
      socket.on('data', (d) => this.onData(d));
      // Our chunk size, then connect.
      const size = Buffer.alloc(4);
      size.writeUInt32BE(OUT_CHUNK);
      this.send(CSID.control, SET_CHUNK_SIZE, size);
      this.setState('connecting', 'connect');
      await this.command('connect', [{
        app: this.target.app,
        type: 'nonprivate',
        flashVer: 'FMLE/3.0 (compatible; Rutba Recorder)',
        tcUrl: this.target.tcUrl,
      }]);
      this.command('releaseStream', [null, this.key], { reply: false });
      this.command('FCPublish', [null, this.key], { reply: false });
      const created = await this.command('createStream', [null]);
      this.streamId = Number(created[1]) || 1;
      this.setState('connecting', 'publish');
      const publishing = new Promise((resolve, reject) => { this.onPublish = { resolve, reject }; });
      this.command('publish', [null, this.key, 'live'], { streamId: this.streamId, csid: 8, reply: false });
      await publishing;
      this.waitingForKey = true;
      this.setState('live');
    } catch (err) {
      // Whatever stopped it, the connection goes with it.
      this.fail(err);
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  /** The plain handshake: C0 and C1, then C2 echoing S1 once S0, S1 and S2 are in. */
  async handshake() {
    const c1 = Buffer.alloc(HANDSHAKE);
    c1.writeUInt32BE(0, 0);
    crypto.randomFillSync(c1, 8);
    this.write(Buffer.concat([Buffer.from([3]), c1]));
    const got = await new Promise((resolve, reject) => {
      let buf = Buffer.alloc(0);
      const on = (d) => {
        buf = Buffer.concat([buf, d]);
        if (buf.length >= 1 + 2 * HANDSHAKE) {
          this.socket.off('data', on);
          resolve(buf);
        }
      };
      this.socket.on('data', on);
      this.socket.once('close', () => reject(new Error('The server closed the connection during the handshake.')));
    });
    if (got[0] !== 3) throw new Error(`The server answered RTMP version ${got[0]}.`);
    const s1 = got.subarray(1, 1 + HANDSHAKE);
    this.write(s1);
    const rest = got.subarray(1 + 2 * HANDSHAKE);
    if (rest.length) setImmediate(() => this.onData(rest));
  }

  onData(d) {
    this.received += d.length;
    if (this.received - this.acked >= this.window / 2) {
      const b = Buffer.alloc(4);
      b.writeUInt32BE(this.received >>> 0);
      this.send(CSID.control, ACK, b);
      this.acked = this.received;
    }
    let messages;
    try {
      messages = this.reader.push(d);
    } catch (err) {
      this.fail(err);
      return;
    }
    for (const m of messages) this.onMessage(m);
  }

  onMessage(m) {
    switch (m.type) {
      case SET_CHUNK_SIZE: this.reader.chunkSize = m.body.readUInt32BE(0) & 0x7fffffff; break;
      case WINDOW_ACK: this.window = m.body.readUInt32BE(0); break;
      case PEER_BANDWIDTH: {
        const b = Buffer.alloc(4);
        b.writeUInt32BE(m.body.readUInt32BE(0));
        this.send(CSID.control, WINDOW_ACK, b);
        break;
      }
      case USER_CONTROL: {
        // A ping (6) is answered with a pong (7) carrying its time.
        if (m.body.readUInt16BE(0) === 6) {
          const b = Buffer.alloc(6);
          b.writeUInt16BE(7, 0);
          m.body.copy(b, 2, 2, 6);
          this.send(CSID.control, USER_CONTROL, b);
        }
        break;
      }
      case ABORT: case ACK: break;
      case COMMAND: this.onCommand(decode(m.body)); break;
      default: break;
    }
  }

  onCommand([name, txn, , info]) {
    if (name === '_result' || name === '_error') {
      const p = this.pending.get(txn);
      if (!p) return;
      this.pending.delete(txn);
      if (name === '_result') p.resolve([name, info]);
      else p.reject(new Error(`The server refused ${p.name}: ${info?.description || info?.code || 'no reason given'}`));
      return;
    }
    if (name === 'onStatus') {
      const code = info?.code || '';
      if (code === 'NetStream.Publish.Start') this.onPublish?.resolve();
      else if (/Publish\.(BadName|Rejected|Failed)|Failed|Rejected/.test(code)) {
        const err = new Error(`The server refused the stream: ${info?.description || code}`);
        this.onPublish?.reject(err);
        if (this.state === 'live') this.fail(err);
      }
    }
  }

  /** The metadata, as the first thing on the stream. */
  sendMetadata(body) {
    if (this.state !== 'live') return;
    this.send(CSID.data, DATA, body, { streamId: this.streamId });
  }

  /** A video message: dropped while the line is behind, until the next key frame. */
  sendVideo(body, timestamp, { config = false } = {}) {
    if (this.state !== 'live') return false;
    const key = isKeyBody(body);
    const behind = (this.socket?.writableLength || 0) > this.maxQueueBytes;
    if (!config && (behind || (this.waitingForKey && !key))) {
      this.waitingForKey = true;
      this.dropped += 1;
      return false;
    }
    // The configuration is marked as a key frame but is not one: the wait goes on to a real one.
    if (key && !config) this.waitingForKey = false;
    this.send(CSID.video, VIDEO, body, { timestamp, streamId: this.streamId });
    return true;
  }

  sendAudio(body, timestamp) {
    if (this.state !== 'live') return false;
    this.send(CSID.audio, AUDIO, body, { timestamp, streamId: this.streamId });
    return true;
  }

  fail(err) {
    if (this.closed) return;
    // An error nobody listens for must not throw out of the socket's handler.
    if (this.listenerCount('error')) this.emit('error', err);
    for (const p of this.pending.values()) p.reject(err);
    this.pending.clear();
    this.onPublish?.reject(err);
    this.close(err.message);
  }

  onClose() {
    if (this.closed) return;
    this.fail(new Error('The server closed the connection.'));
  }

  /** Stop: unpublish politely when live, then close. */
  close(detail = '') {
    if (this.closed) return;
    this.closed = true;
    try {
      if (this.state === 'live') {
        this.command('FCUnpublish', [null, this.key], { reply: false });
        this.command('deleteStream', [null, this.streamId], { reply: false });
      }
    } catch { /* closing anyway */ }
    try { this.socket?.end(); } catch { /* closed */ }
    setTimeout(() => { try { this.socket?.destroy(); } catch { /* closed */ } }, 500).unref?.();
    this.setState('closed', detail);
    this.emit('close');
  }
}
