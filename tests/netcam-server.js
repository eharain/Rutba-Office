// A camera for the tests: an RTSP server on this computer's loopback, on a
// port the system picks, that answers as network cameras do and sends
// whatever H.264 it is handed, over the connection that asked (interleaved).
// The tests, and the proof that decodes a picture, read it with the reader
// under test. It is never part of Office: Office listens on nothing.

import net from 'node:net';
import crypto from 'node:crypto';

const md5 = (s) => crypto.createHash('md5').update(s).digest('hex');
const sha256 = (s) => crypto.createHash('sha256').update(s).digest('hex');

/** The RTP packets one frame's NAL units go out in: small ones whole, large ones cut (FU-A), the first two in one packet when `stap` is asked. */
export function packetise(nals, { timestamp, sequence, payloadType = 96, mtu = 1200, stap = false, ssrc = 0x1234abcd }) {
  const payloads = [];
  let list = nals;
  if (stap && nals.length >= 2) {
    const parts = [Buffer.from([24])];
    for (const nal of nals.slice(0, 2)) {
      const size = Buffer.alloc(2);
      size.writeUInt16BE(nal.length);
      parts.push(size, nal);
    }
    payloads.push(Buffer.concat(parts));
    list = nals.slice(2);
  }
  for (const nal of list) {
    if (nal.length <= mtu) payloads.push(nal);
    else {
      const body = nal.subarray(1);
      for (let at = 0; at < body.length; at += mtu) {
        const end = at + mtu >= body.length;
        const head = Buffer.from([(nal[0] & 0xe0) | 28, (at === 0 ? 0x80 : 0) | (end ? 0x40 : 0) | (nal[0] & 0x1f)]);
        payloads.push(Buffer.concat([head, body.subarray(at, at + mtu)]));
      }
    }
  }
  return payloads.map((payload, i) => {
    const head = Buffer.alloc(12);
    head[0] = 0x80;
    head[1] = payloadType | (i === payloads.length - 1 ? 0x80 : 0);
    head.writeUInt16BE((sequence + i) & 0xffff, 2);
    head.writeUInt32BE(timestamp >>> 0, 4);
    head.writeUInt32BE(ssrc, 8);
    return Buffer.concat([head, payload]);
  });
}

/** A packet as it travels inside the RTSP connection: `$`, the channel, the length, the packet. */
export function interleave(channel, packet) {
  const head = Buffer.alloc(4);
  head[0] = 0x24;
  head[1] = channel;
  head.writeUInt16BE(packet.length, 2);
  return Buffer.concat([head, packet]);
}

function paramsOf(text) {
  const out = {};
  const re = /([a-z0-9_-]+)\s*=\s*(?:"([^"]*)"|([^\s,]*))/gi;
  for (let m = re.exec(text); m; m = re.exec(text)) out[m[1].toLowerCase()] = m[2] ?? m[3];
  return out;
}

/**
 * Starts the camera. Options:
 *   path        where the stream is ('/stream1')
 *   auth        null, or { scheme: 'basic' | 'digest', user, password, algorithm: 'MD5' | 'SHA-256', qop: true | false, also: 'basic' }
 *   codec       'H264' (default), 'H265', 'JPEG'
 *   sprop       [sps, pps] to name in the SDP
 *   sound       true to offer a sound stream as well
 *   base        'header' (Content-Base, default), 'none'
 *   absolute    true for stream controls that are whole addresses
 *   channels    the first interleaved channel the camera chooses (0)
 *   setup       a status to answer SETUP with instead of 200 (461)
 *   moved       { from, to } answers DESCRIBE at `from` with 302 to `to` (a path, or a whole address)
 *   noCseq      true to leave CSeq off every answer
 *   timeout     the session timeout the camera states (60)
 *   chop        send every answer and packet in pieces of this many bytes
 *   publicLine  what OPTIONS says the camera can do
 *   mode        packetization-mode in the SDP ('1')
 */
export async function startCamera(options = {}) {
  const opt = { path: '/stream1', auth: null, codec: 'H264', sprop: null, sound: false, base: 'header', absolute: false, channels: 0, setup: 200, moved: null, noCseq: false, timeout: 60, chop: 0, publicLine: 'OPTIONS, DESCRIBE, SETUP, PLAY, TEARDOWN, GET_PARAMETER', mode: '1', ...options };
  const requests = [];
  const players = new Set();
  const sockets = new Set();
  const nonce = crypto.randomBytes(8).toString('hex');
  let sequence = 100;
  let port = 0;

  const write = (socket, bytes) => {
    if (socket.destroyed) return;
    if (!opt.chop) socket.write(bytes);
    else for (let at = 0; at < bytes.length; at += opt.chop) socket.write(bytes.subarray(at, at + opt.chop));
  };

  const signedIn = (req) => {
    const auth = opt.auth;
    if (!auth) return true;
    const line = req.headers.authorization || '';
    if (/^basic /i.test(line)) {
      if (auth.scheme !== 'basic' && auth.also !== 'basic') return false;
      return Buffer.from(line.slice(6), 'base64').toString('utf8') === `${auth.user}:${auth.password}`;
    }
    if (!/^digest /i.test(line) || auth.scheme !== 'digest') return false;
    const p = paramsOf(line.slice(7));
    const h = auth.algorithm === 'SHA-256' ? sha256 : md5;
    if (p.username !== auth.user || p.nonce !== nonce || p.uri !== req.uri) return false;
    const ha1 = h(`${auth.user}:cam:${auth.password}`);
    const ha2 = h(`${req.method}:${req.uri}`);
    const want = auth.qop === false ? h(`${ha1}:${nonce}:${ha2}`) : h(`${ha1}:${nonce}:${p.nc}:${p.cnonce}:auth:${ha2}`);
    if (auth.qop !== false && p.qop !== 'auth') return false;
    return p.response === want;
  };

  const challenge = () => {
    const auth = opt.auth;
    const lines = [];
    if (auth.scheme === 'digest') {
      lines.push(`WWW-Authenticate: Digest realm="cam", nonce="${nonce}"${auth.qop === false ? '' : ', qop="auth"'}${auth.algorithm === 'SHA-256' ? ', algorithm=SHA-256' : ''}`);
      if (auth.also === 'basic') lines.push('WWW-Authenticate: Basic realm="cam"');
    } else if (auth.scheme === 'basic') lines.push('WWW-Authenticate: Basic realm="cam"');
    else lines.push(`WWW-Authenticate: ${auth.scheme} realm="cam"`);
    return lines;
  };

  const sdp = () => {
    const self = `rtsp://127.0.0.1:${port}${opt.path}`;
    const control = (name) => (opt.absolute ? `${self}/${name}` : name);
    const lines = ['v=0', 'o=- 1 1 IN IP4 127.0.0.1', 's=Test camera', 't=0 0', 'a=control:*'];
    if (opt.codec === 'JPEG') lines.push('m=video 0 RTP/AVP 26', `a=control:${control('trackID=0')}`);
    else {
      lines.push('m=video 0 RTP/AVP 96', `a=rtpmap:96 ${opt.codec}/90000`);
      const fmtp = [`packetization-mode=${opt.mode}`];
      if (opt.sprop) fmtp.push(`sprop-parameter-sets=${opt.sprop.map((b) => Buffer.from(b).toString('base64')).join(',')}`);
      lines.push(`a=fmtp:96 ${fmtp.join(';')}`, `a=control:${control('trackID=0')}`);
    }
    if (opt.sound) lines.push('m=audio 0 RTP/AVP 0', `a=control:${control('trackID=1')}`);
    return `${lines.join('\r\n')}\r\n`;
  };

  const server = net.createServer((socket) => {
    sockets.add(socket);
    const state = { socket, playing: false, channel: opt.channels, session: '' };
    let buf = Buffer.alloc(0);
    socket.on('error', () => {});
    socket.on('close', () => {
      sockets.delete(socket);
      players.delete(state);
    });
    const answer = (req, status, reason, headers = [], body = '') => {
      const lines = [`RTSP/1.0 ${status} ${reason}`];
      if (!opt.noCseq) lines.push(`CSeq: ${req.headers.cseq}`);
      lines.push(...headers);
      if (body) lines.push(`Content-Length: ${Buffer.byteLength(body)}`);
      write(socket, Buffer.from(`${lines.join('\r\n')}\r\n\r\n${body}`, 'utf8'));
    };
    const handle = (req) => {
      requests.push(req);
      camera.onRequest?.(req, state);
      if (!req.method) return;
      if (req.method === 'OPTIONS' && !opt.auth?.options) {
        answer(req, 200, 'OK', [`Public: ${opt.publicLine}`]);
        return;
      }
      if (!signedIn(req)) {
        answer(req, 401, 'Unauthorized', challenge());
        return;
      }
      const path = req.uri.replace(/^rtsps?:\/\/[^/]+/i, '');
      if (req.method === 'OPTIONS') answer(req, 200, 'OK', [`Public: ${opt.publicLine}`]);
      else if (req.method === 'DESCRIBE') {
        if (opt.moved && path === opt.moved.from) {
          const to = /^rtsps?:/i.test(opt.moved.to) ? opt.moved.to : `rtsp://127.0.0.1:${port}${opt.moved.to}`;
          answer(req, 302, 'Moved Temporarily', [`Location: ${to}`]);
        } else if (path !== opt.path) answer(req, 404, 'Not Found');
        else answer(req, 200, 'OK', ['Content-Type: application/sdp', ...(opt.base === 'header' ? [`Content-Base: rtsp://127.0.0.1:${port}${opt.path}/`] : [])], sdp());
      } else if (req.method === 'SETUP') {
        if (opt.setup !== 200) {
          answer(req, opt.setup, opt.setup === 461 ? 'Unsupported Transport' : 'No');
          return;
        }
        state.session = crypto.randomBytes(4).toString('hex');
        state.channel = opt.channels;
        answer(req, 200, 'OK', [`Transport: RTP/AVP/TCP;unicast;interleaved=${opt.channels}-${opt.channels + 1}`, `Session: ${state.session};timeout=${opt.timeout}`]);
      } else if (req.method === 'PLAY') {
        if (req.headers.session !== state.session) {
          answer(req, 454, 'Session Not Found');
          return;
        }
        state.playing = true;
        players.add(state);
        answer(req, 200, 'OK', [`Session: ${state.session}`, 'RTP-Info: url=trackID=0;seq=100']);
        camera.onPlay?.(state);
      } else if (req.method === 'TEARDOWN') {
        players.delete(state);
        state.playing = false;
        answer(req, 200, 'OK');
        socket.end();
      } else answer(req, 200, 'OK');
    };
    socket.on('data', (chunk) => {
      buf = Buffer.concat([buf, chunk]);
      for (;;) {
        const end = buf.indexOf('\r\n\r\n');
        if (end < 0) break;
        const lines = buf.toString('utf8', 0, end).split('\r\n');
        buf = buf.subarray(end + 4);
        const headers = {};
        for (const line of lines.slice(1)) {
          const colon = line.indexOf(':');
          if (colon > 0) headers[line.slice(0, colon).trim().toLowerCase()] = line.slice(colon + 1).trim();
        }
        const first = /^([A-Z_]+)\s+(\S+)\s+RTSP\/1\.0$/.exec(lines[0]);
        if (first) handle({ method: first[1], uri: first[2], headers });
        else requests.push({ method: '', line: lines[0], headers });
      }
    });
  });

  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  port = server.address().port;

  const camera = {
    port,
    url: `rtsp://127.0.0.1:${port}${opt.path}`,
    options: opt,
    requests,
    players,
    onRequest: null,
    onPlay: null,
    /** Sends one frame (its NAL units) to everyone playing. `skip` leaves that many sequence numbers out before it, as a lost packet does. */
    send(nals, { timestamp, stap = false, mtu = 1200, skip = 0, marker = true } = {}) {
      sequence = (sequence + skip) & 0xffff;
      const packets = packetise(nals, { timestamp, sequence, stap, mtu });
      if (!marker) packets[packets.length - 1][1] &= 0x7f;
      sequence = (sequence + packets.length) & 0xffff;
      for (const state of players) for (const packet of packets) write(state.socket, interleave(state.channel, packet));
    },
    /** Sends raw bytes to everyone playing (a report on the other channel, a question from the camera). */
    raw(bytes) {
      for (const state of players) write(state.socket, Buffer.from(bytes));
    },
    /** Drops every connection, as a camera that loses power does. */
    cut() {
      for (const socket of sockets) socket.destroy();
    },
    async close() {
      for (const socket of sockets) socket.destroy();
      await new Promise((resolve) => server.close(resolve));
    },
  };
  return camera;
}

class BitWriter {
  constructor() {
    this.bits = [];
  }

  u(n, v) {
    for (let i = n - 1; i >= 0; i -= 1) this.bits.push(Math.floor(v / 2 ** i) % 2);
    return this;
  }

  ue(v) {
    const code = v + 1;
    const size = Math.floor(Math.log2(code));
    this.u(size, 0);
    return this.u(size + 1, code);
  }

  bytes() {
    const bits = [...this.bits, 1];
    while (bits.length % 8) bits.push(0);
    const raw = [];
    for (let i = 0; i < bits.length; i += 8) raw.push(parseInt(bits.slice(i, i + 8).join(''), 2));
    // The escape byte the standard asks for wherever two zero bytes would be followed by 0 to 3.
    const out = [];
    let zeros = 0;
    for (const b of raw) {
      if (zeros >= 2 && b <= 3) {
        out.push(3);
        zeros = 0;
      }
      out.push(b);
      zeros = b === 0 ? zeros + 1 : 0;
    }
    return out;
  }
}

/** A sequence parameter set for a picture of this size, written by the standard's own layout (the tests read it back). */
export function makeSps({ width, height, profile = 66, level = 31, constraints = 0xc0 }) {
  const w = new BitWriter();
  w.u(8, profile).u(8, constraints).u(8, level).ue(0);
  if (profile === 100) w.ue(1).ue(0).ue(0).u(1, 0).u(1, 0);
  w.ue(0).ue(0).ue(0).ue(1).u(1, 0);
  const mbsWide = Math.ceil(width / 16);
  const mbsHigh = Math.ceil(height / 16);
  w.ue(mbsWide - 1).ue(mbsHigh - 1).u(1, 1).u(1, 1);
  const cropRight = (mbsWide * 16 - width) / 2;
  const cropBottom = (mbsHigh * 16 - height) / 2;
  if (cropRight || cropBottom) w.u(1, 1).ue(0).ue(cropRight).ue(0).ue(cropBottom);
  else w.u(1, 0);
  w.u(1, 0);
  return Buffer.from([0x67, ...w.bytes()]);
}

/** A picture parameter set, enough to be told from the rest. */
export const makePps = () => Buffer.from([0x68, 0xce, 0x3c, 0x80]);

/** A slice of `size` bytes: an IDR when `key`, with bytes a test can tell apart. */
export function makeSlice({ key = false, size = 40, fill = 0x55 } = {}) {
  const nal = Buffer.alloc(size, fill);
  nal[0] = key ? 0x65 : 0x41;
  return nal;
}
