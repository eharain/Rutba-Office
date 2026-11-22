// Reads one camera over RTSP (RFC 2326) and hands back its H.264 picture.
//
// One connection carries everything: the camera is asked to send its picture
// down the same TCP (or TLS) connection the questions go up, interleaved with
// the answers. Nothing is listened on and no second connection is opened, so
// a camera is read through any firewall that lets the first one out.
//
// The reader asks, in order: what the camera can do (OPTIONS), what it sends
// (DESCRIBE), for its H.264 video over this connection (SETUP), and to start
// (PLAY). It then keeps the session awake and reads until it is stopped or
// the camera goes quiet. Camera sound and H.265 are not read.
//
// A failure is a CameraError whose `code` says what went wrong, for the
// window to word: address, unreachable, timeout, closed, certificate,
// sign-in-needed, sign-in-refused, sign-in-unsupported, not-found, refused,
// redirect, no-video, h265, codec, no-tcp, protocol, stalled, stopped.

import { EventEmitter } from 'node:events';
import net from 'node:net';
import tls from 'node:tls';
import { parseCameraAddress } from './address.js';
import { authorization, chooseChallenge, parseChallenges } from './auth.js';
import { H264Assembler } from './h264.js';
import { parseRtp } from './rtp.js';
import { controlUrl, parseSdp, spropSets } from './sdp.js';

export class CameraError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'CameraError';
    this.code = code;
  }
}

const EMPTY = Buffer.alloc(0);
const MAX_HEAD = 64 * 1024;
const MAX_BODY = 1024 * 1024;
const MAX_HOPS = 3;

const defaultConnect = ({ secure, host, port }) => (secure ? tls.connect({ host, port, servername: net.isIP(host) ? undefined : host }) : net.connect({ host, port }));

const UNREACHABLE = new Set(['ECONNREFUSED', 'EHOSTUNREACH', 'ENETUNREACH', 'ETIMEDOUT', 'ENOTFOUND', 'EAI_AGAIN', 'EADDRNOTAVAIL', 'EHOSTDOWN']);

function socketError(err) {
  if (err instanceof CameraError) return err;
  const code = String(err?.code || '');
  if (UNREACHABLE.has(code)) return new CameraError('unreachable', 'The camera could not be reached at that address.');
  if (/CERT|SELF_SIGNED|UNABLE_TO_VERIFY|ALTNAME/.test(code)) return new CameraError('certificate', 'The camera’s certificate could not be checked.');
  return new CameraError('closed', 'The connection to the camera was lost.');
}

const first = (res, name) => res.headers[name]?.[0] || '';

export class RtspReader extends EventEmitter {
  #opt;
  #socket = null;
  #buf = EMPTY;
  #cseq = 0;
  #pending = new Map();
  #challenge = null;
  #nonceCount = 0;
  #session = '';
  #channel = 0;
  #payload = -1;
  #assembler = null;
  #started = false;
  #playing = false;
  #dead = null;
  #keepTimer = null;
  #quietTimer = null;
  #lastData = 0;
  #aggregate = '';
  #sessionSeconds = 60;

  /**
   * `url` is the camera's address (rtsp:// or rtsps://); a sign-in typed
   * inside it is used when `user` and `password` are not given. `connect` is
   * the socket maker ({ secure, host, port }) the tests replace. `timeoutMs`
   * is how long the camera has to connect and to answer each question;
   * `quietMs` is how long its picture may stop before the reader gives up.
   */
  constructor({ url, user, password, connect = defaultConnect, timeoutMs = 8000, quietMs = 10000, agent = 'Rutba Office' } = {}) {
    super();
    this.#opt = { url, user, password, connect, timeoutMs, quietMs, agent };
  }

  /**
   * Connects and starts the camera's picture. Resolves once the camera is
   * playing, with { codec: 'H264', sound } (`sound` says whether the camera
   * also offers sound, which is not read). Rejects with a CameraError.
   *
   * Events from then on: 'config' ({ codec, width, height }) before the
   * first frame and whenever the stream's parameters change; 'unit'
   * ({ timestamp, key, data }) for each frame, Annex B, in microseconds; and
   * 'end' once, with the CameraError that ended it or null after stop().
   */
  async start() {
    if (this.#started) throw new Error('A reader starts once.');
    this.#started = true;
    let addr = parseCameraAddress(this.#opt.url);
    if (!addr) throw new CameraError('address', 'That is not a camera’s address.');
    if (this.#opt.user === undefined && this.#opt.password === undefined) {
      this.#opt.user = addr.user;
      this.#opt.password = addr.password;
    }
    try {
      let found = null;
      for (let hop = 0; !found; hop += 1) {
        await this.#connect(addr);
        const step = await this.#describe(addr);
        if (step.to) {
          if (hop >= MAX_HOPS) throw new CameraError('redirect', 'The camera kept sending the reader to another address.');
          this.#leave();
          addr = step.to;
        } else found = step;
      }
      await this.#setup(found);
      const res = await this.#request('PLAY', found.aggregate, { Range: 'npt=0.000-' });
      if (res.status !== 200) throw new CameraError('protocol', `The camera would not start its picture (${res.status}).`);
      this.#play(found);
      return { codec: 'H264', sound: found.sound };
    } catch (err) {
      const error = this.#dead || socketError(err);
      this.#fail(error);
      throw error;
    }
  }

  /** Stops reading: tells the camera the session is over and closes the connection. */
  stop() {
    if (this.#dead) return;
    let goodbye = null;
    if (this.#session && this.#aggregate) goodbye = this.#lines('TEARDOWN', this.#aggregate, {}).text;
    this.#fail(new CameraError('stopped', 'The camera was closed.'), goodbye);
  }

  #connect(addr) {
    return new Promise((resolve, reject) => {
      if (this.#dead) {
        reject(this.#dead);
        return;
      }
      let socket;
      try {
        socket = this.#opt.connect({ secure: addr.secure, host: addr.host, port: addr.port });
      } catch (err) {
        reject(socketError(err));
        return;
      }
      this.#socket = socket;
      this.#buf = EMPTY;
      let open = false;
      const fail = (err) => {
        clearTimeout(timer);
        if (socket !== this.#socket) return;
        if (open) this.#fail(err);
        else {
          socket.destroy();
          reject(this.#dead || err);
        }
      };
      const timer = setTimeout(() => fail(new CameraError('unreachable', 'The camera did not answer at that address.')), this.#opt.timeoutMs);
      socket.once(addr.secure ? 'secureConnect' : 'connect', () => {
        open = true;
        clearTimeout(timer);
        socket.setNoDelay?.(true);
        resolve();
      });
      socket.on('data', (chunk) => {
        if (socket === this.#socket && !this.#dead) this.#take(chunk);
      });
      socket.on('error', (err) => fail(socketError(err)));
      socket.on('close', () => fail(new CameraError('closed', 'The camera closed the connection.')));
    });
  }

  /** Lets go of the connection without ending the reader (the camera sent it elsewhere). */
  #leave() {
    const socket = this.#socket;
    this.#socket = null;
    for (const [, p] of this.#pending) clearTimeout(p.timer);
    this.#pending.clear();
    this.#challenge = null;
    this.#session = '';
    socket?.destroy();
  }

  async #describe(addr) {
    const options = await this.#request('OPTIONS', addr.url);
    const can = options.status === 200 ? first(options, 'public').toUpperCase() : '';
    const res = await this.#request('DESCRIBE', addr.url, { Accept: 'application/sdp' });
    if ([301, 302, 303, 307, 308].includes(res.status)) {
      const to = parseCameraAddress(first(res, 'location'));
      // A camera may move its stream to another port or path. It may not hand the sign-in to another machine, nor take a protected connection to an open one.
      if (!to || to.host.toLowerCase() !== addr.host.toLowerCase() || (addr.secure && !to.secure)) throw new CameraError('redirect', 'The camera sent the reader to another address, which is not followed.');
      return { to };
    }
    if (res.status === 404) throw new CameraError('not-found', 'The camera has no stream at that address.');
    if (res.status === 403) throw new CameraError('refused', 'The camera refused to give that stream.');
    if (res.status !== 200) throw new CameraError('protocol', `The camera would not describe its stream (${res.status}).`);

    const sdp = parseSdp(res.body.toString('utf8'));
    const base = first(res, 'content-base') || first(res, 'content-location') || addr.url;
    const videos = sdp.tracks.filter((t) => t.kind === 'video' && t.transport.toUpperCase().startsWith('RTP/AVP'));
    const track = videos.find((t) => t.codec === 'H264');
    if (!track) {
      if (videos.some((t) => t.codec === 'H265' || t.codec === 'HEVC')) throw new CameraError('h265', 'The camera sends H.265 only at that address.');
      if (videos.length) throw new CameraError('codec', 'The camera sends a kind of video that is not read.');
      throw new CameraError('no-video', 'The camera offers no picture at that address.');
    }
    // Mode 2 sends frames out of order in packet kinds that are not put back together here.
    if (String(track.fmtp['packetization-mode'] || '0') === '2') throw new CameraError('codec', 'The camera sends a kind of video that is not read.');
    return { track, can, base, aggregate: controlUrl(base, sdp.control), sound: sdp.tracks.some((t) => t.kind === 'audio') };
  }

  async #setup({ track, base }) {
    const res = await this.#request('SETUP', controlUrl(base, track.control), { Transport: 'RTP/AVP/TCP;unicast;interleaved=0-1' });
    if (res.status === 461) throw new CameraError('no-tcp', 'The camera will not send its picture over this connection.');
    if (res.status !== 200) throw new CameraError('protocol', `The camera would not set its stream up (${res.status}).`);
    const transport = first(res, 'transport');
    const channel = /interleaved=(\d+)/i.exec(transport);
    if (transport && !channel) throw new CameraError('no-tcp', 'The camera will not send its picture over this connection.');
    this.#channel = channel ? Number(channel[1]) : 0;
    const session = first(res, 'session');
    this.#session = session.split(';')[0].trim();
    if (!this.#session) throw new CameraError('protocol', 'The camera gave no session.');
    const timeout = /timeout\s*=\s*(\d+)/i.exec(session);
    this.#sessionSeconds = timeout ? Number(timeout[1]) : 60;
    this.#payload = track.payload;
    this.#assembler = new H264Assembler({ sets: spropSets(track), clock: track.clock || 90000 });
  }

  #play({ aggregate, can }) {
    this.#aggregate = aggregate;
    this.#playing = true;
    this.#lastData = Date.now();
    // Asked well inside the time the camera keeps a silent session, whatever it says that is.
    const every = Math.max(1000, Math.min(25000, (this.#sessionSeconds * 1000) / 2));
    const method = can.includes('GET_PARAMETER') ? 'GET_PARAMETER' : 'OPTIONS';
    this.#keepTimer = setInterval(() => {
      this.#exchange(method, aggregate, {}, { soft: true }).catch(() => {});
    }, every);
    const quiet = this.#opt.quietMs;
    this.#quietTimer = setInterval(() => {
      if (Date.now() - this.#lastData > quiet) this.#fail(new CameraError('stalled', 'The camera stopped sending its picture.'));
    }, Math.max(50, Math.min(1000, quiet / 2)));
  }

  /** Ends the reader, once. `goodbye` is a last request to send before the connection closes. */
  #fail(error, goodbye = null) {
    if (this.#dead) return;
    this.#dead = error;
    clearInterval(this.#keepTimer);
    clearInterval(this.#quietTimer);
    for (const [, p] of this.#pending) {
      clearTimeout(p.timer);
      p.reject(error);
    }
    this.#pending.clear();
    const socket = this.#socket;
    if (socket && !socket.destroyed) {
      if (goodbye) {
        try {
          socket.end(goodbye, 'utf8');
        } catch {
          // The connection is already gone.
        }
        setTimeout(() => socket.destroy(), 1000).unref?.();
      } else socket.destroy();
    }
    if (this.#playing) {
      this.#playing = false;
      this.emit('end', error.code === 'stopped' ? null : error);
    }
  }

  #lines(method, uri, headers) {
    this.#cseq += 1;
    const lines = [`${method} ${uri} RTSP/1.0`, `CSeq: ${this.#cseq}`, `User-Agent: ${this.#opt.agent}`];
    if (this.#challenge) {
      this.#nonceCount += 1;
      lines.push(`Authorization: ${authorization({ challenge: this.#challenge, user: this.#opt.user || '', password: this.#opt.password || '', method, uri, count: this.#nonceCount })}`);
    }
    if (this.#session) lines.push(`Session: ${this.#session}`);
    for (const [name, value] of Object.entries(headers)) lines.push(`${name}: ${value}`);
    return { cseq: this.#cseq, text: `${lines.join('\r\n')}\r\n\r\n` };
  }

  /** One request and its answer. A `soft` request that goes unanswered is given up alone; any other ends the reader. */
  #exchange(method, uri, headers = {}, { soft = false } = {}) {
    if (this.#dead) return Promise.reject(this.#dead);
    const { cseq, text } = this.#lines(method, uri, headers);
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.#pending.delete(cseq);
        const err = new CameraError('timeout', 'The camera did not answer.');
        reject(err);
        if (!soft) this.#fail(err);
      }, this.#opt.timeoutMs);
      this.#pending.set(cseq, { resolve, reject, timer });
      this.#socket.write(text, 'utf8');
    });
  }

  /** A request, signed in when the camera asks for it. */
  async #request(method, uri, headers = {}) {
    for (let attempt = 0; ; attempt += 1) {
      const signed = Boolean(this.#challenge);
      const res = await this.#exchange(method, uri, headers);
      if (res.status !== 401) return res;
      if (!this.#opt.user && !this.#opt.password) throw new CameraError('sign-in-needed', 'The camera asks for a user name and password.');
      const challenge = chooseChallenge(parseChallenges(res.headers['www-authenticate']));
      if (!challenge) throw new CameraError('sign-in-unsupported', 'The camera asks for a kind of sign-in that is not known here.');
      if ((signed && !challenge.stale) || attempt >= 2) throw new CameraError('sign-in-refused', 'The camera did not accept the user name and password.');
      this.#challenge = challenge;
      this.#nonceCount = 0;
    }
  }

  #take(chunk) {
    let buf = this.#buf.length ? Buffer.concat([this.#buf, chunk]) : chunk;
    let at = 0;
    while (at < buf.length && !this.#dead) {
      const lead = buf[at];
      if (lead === 0x24) {
        if (buf.length - at < 4) break;
        const size = buf.readUInt16BE(at + 2);
        if (buf.length - at < 4 + size) break;
        this.#interleaved(buf[at + 1], buf.subarray(at + 4, at + 4 + size));
        at += 4 + size;
      } else if (lead === 0x0d || lead === 0x0a) at += 1;
      else if (lead < 0x41 || lead > 0x5a) {
        this.#fail(new CameraError('protocol', 'What answered at that address is not a camera stream.'));
        return;
      } else {
        const end = buf.indexOf('\r\n\r\n', at, 'latin1');
        if (end < 0) {
          if (buf.length - at > MAX_HEAD) {
            this.#fail(new CameraError('protocol', 'What answered at that address is not a camera stream.'));
            return;
          }
          break;
        }
        const lines = buf.toString('utf8', at, end).split('\r\n');
        const headers = {};
        for (const line of lines.slice(1)) {
          const colon = line.indexOf(':');
          if (colon > 0) (headers[line.slice(0, colon).trim().toLowerCase()] ||= []).push(line.slice(colon + 1).trim());
        }
        const size = Number(headers['content-length']?.[0] || 0);
        if (!Number.isInteger(size) || size < 0 || size > MAX_BODY) {
          this.#fail(new CameraError('protocol', 'What answered at that address is not a camera stream.'));
          return;
        }
        if (buf.length < end + 4 + size) break;
        const body = buf.subarray(end + 4, end + 4 + size);
        at = end + 4 + size;
        this.#message(lines[0], headers, body);
      }
    }
    if (this.#dead) {
      this.#buf = EMPTY;
      return;
    }
    this.#buf = at >= buf.length ? EMPTY : buf.subarray(at);
    buf = null;
  }

  #message(line, headers, body) {
    const answer = /^RTSP\/\d\.\d\s+(\d{3})\s*(.*)$/.exec(line);
    if (answer) {
      const cseq = Number(headers.cseq?.[0]);
      // A camera that leaves the number off its answer is taken to be answering the one question asked.
      const key = this.#pending.has(cseq) ? cseq : this.#pending.size === 1 && !headers.cseq ? [...this.#pending.keys()][0] : null;
      if (key === null) return;
      const p = this.#pending.get(key);
      this.#pending.delete(key);
      clearTimeout(p.timer);
      p.resolve({ status: Number(answer[1]), reason: answer[2], headers, body });
      return;
    }
    const asked = /^([A-Z_]+)\s+\S+\s+RTSP\/\d\.\d$/.exec(line);
    if (!asked) {
      this.#fail(new CameraError('protocol', 'What answered at that address is not a camera stream.'));
      return;
    }
    // The camera may ask things too. It is told yes to what changes nothing and no to the rest.
    const fine = ['OPTIONS', 'GET_PARAMETER', 'SET_PARAMETER'].includes(asked[1]);
    const cseq = /^\d+$/.test(headers.cseq?.[0] || '') ? `CSeq: ${headers.cseq[0]}\r\n` : '';
    this.#socket?.write(`RTSP/1.0 ${fine ? '200 OK' : '405 Method Not Allowed'}\r\n${cseq}\r\n`, 'latin1');
  }

  #interleaved(channel, data) {
    if (!this.#assembler || channel !== this.#channel) return;
    const packet = parseRtp(data);
    if (!packet || packet.payloadType !== this.#payload) return;
    this.#lastData = Date.now();
    for (const unit of this.#assembler.push(packet)) {
      const { config, ...frame } = unit;
      if (config) this.emit('config', config);
      this.emit('unit', frame);
    }
  }
}
