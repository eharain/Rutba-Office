// The door of a Talk Server, as Office reaches it (the Talk server contract,
// "A meeting from start to end", steps 1, 4 and 5).
//
// A door comes only from the introducer, in the answer to a request Office
// made itself: nothing here takes an address from a person, a link or a
// server. What is checked before anything is sent through it:
//
// - Office connects to the `address` the introducer tested, never to where a
//   name points today.
// - A server with no name shows a certificate it made itself. The introducer
//   sends the `pin` (the SHA-256 of that certificate) it saw from outside;
//   Office goes on only when the certificate shown has exactly that pin, and
//   checks nothing else about it.
// - A server with a name shows a certificate for that name, checked the way
//   any https certificate is, against the name the introducer gave.
// - A door that does not protect its connection at all is used only on this
//   computer itself (a server being tried out here). Anywhere else it would
//   put the pass and the join token on the network in the clear.
// - A door never sends Office anywhere else: no redirect is followed.
//
// The pass is exchanged here (`exchange`), a host asks the door who is in the
// room and mutes (`control`), and the meeting's signalling is a WebSocket
// made by hand on the same checked connection (`openSocket`), so the window
// never connects to a server itself.

import crypto from 'node:crypto';
import http from 'node:http';
import net from 'node:net';
import tls from 'node:tls';
import { OP, acceptKey, closePayload, encodeFrame, frameReader } from './talk-frames.js';

const CONNECT_MS = 8000;
const ANSWER_MS = 15000;
const ANSWER_MOST = 64 * 1024;
const PIN = /^[0-9a-f]{64}$/;

/** Why a door could not be used: `code` is one of the words talk/server-lines.js has a sentence for. */
export class DoorError extends Error {
  constructor(code, extra = {}) {
    super(`door: ${code}`);
    this.code = code;
    this.extra = extra;
  }
}

/** Whether an address is this computer itself. */
export function isLoopback(address) {
  const a = String(address || '').toLowerCase();
  if (net.isIPv4(a)) return a.startsWith('127.');
  return a === '::1' || /^::ffff:127\.\d+\.\d+\.\d+$/.test(a);
}

/** The pin of a certificate: the SHA-256 of it as it is sent, in hex. */
export const pinOf = (cert) => (cert?.fingerprint256 ? String(cert.fingerprint256).replace(/:/g, '').toLowerCase() : '');

/**
 * A door as the introducer handed it out, read strictly: where Office
 * connects, and what it must see there. Anything not understood is refused,
 * not guessed at.
 * @param {{ url: string, address: string, name?: string, family?: number, media?: string }} door
 * @param {string} [pin] the server's pin, when it shows its own certificate
 * @returns {{ secure: boolean, address: string, port: number, family: 4|6, name: string|null, pin: string|null, media: string, host: string }}
 */
export function readDoor(door, pin = null) {
  let url;
  try { url = new URL(String(door?.url || '')); } catch { throw new DoorError('door'); }
  if (url.protocol !== 'ws:' && url.protocol !== 'wss:') throw new DoorError('door');
  if (url.username || url.password || url.search || url.hash || (url.pathname && url.pathname !== '/')) throw new DoorError('door');
  const secure = url.protocol === 'wss:';
  const address = String(door.address || '');
  const family = net.isIP(address);
  if (!family) throw new DoorError('door');
  const port = Number(url.port || (secure ? 443 : 80));
  if (!Number.isInteger(port) || port < 1 || port > 65535) throw new DoorError('door');
  const name = door.name ? String(door.name).toLowerCase() : null;
  if (name && (!/^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/.test(name) || net.isIP(name))) throw new DoorError('door');
  const wantPin = pin == null || pin === '' ? null : String(pin).toLowerCase();
  if (wantPin && !PIN.test(wantPin)) throw new DoorError('door');
  // A pin is for a door that speaks TLS; a name without TLS proves nothing.
  if (!secure && (wantPin || name)) throw new DoorError('door');
  if (!secure && !isLoopback(address)) throw new DoorError('plain');
  const shown = name || (family === 6 ? `[${address}]` : address);
  return { secure, address, port, family, name, pin: wantPin, media: door.media === 'tcp' ? 'tcp' : 'udp', host: `${shown}:${port}` };
}

const certFault = (err) => /CERT|ALTNAME|SELF_SIGNED|UNABLE_TO_VERIFY|ERR_TLS|HOSTNAME/i.test(String(err?.code || ''));

/**
 * A connection to a door, checked: resolves with the socket only once the
 * certificate shown is the one expected. Refuses with a DoorError:
 * `unreachable`, `pin`, `certificate`.
 */
export function connectDoor(d, { timeoutMs = CONNECT_MS, connectTls = tls.connect, connectPlain = net.connect } = {}) {
  return new Promise((resolve, reject) => {
    let done = false;
    let socket = null;
    const end = (err, ok) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      if (err) { try { socket?.destroy(); } catch { /* gone */ } reject(err); } else resolve(ok);
    };
    const timer = setTimeout(() => end(new DoorError('unreachable')), timeoutMs);
    try {
      if (!d.secure) {
        socket = connectPlain({ host: d.address, port: d.port, family: d.family });
        socket.once('connect', () => end(null, socket));
      } else {
        socket = connectTls({
          host: d.address,
          port: d.port,
          family: d.family,
          ALPNProtocols: ['http/1.1'],
          // With a pin nothing about the certificate counts but the pin itself, looked at below.
          rejectUnauthorized: !d.pin,
          ...(d.name ? { servername: d.name } : {}),
          checkServerIdentity: (_host, cert) => tls.checkServerIdentity(d.name || d.address, cert),
        });
        socket.once('secureConnect', () => {
          if (d.pin) {
            const seen = pinOf(socket.getPeerX509Certificate?.());
            if (!seen || seen !== d.pin) return end(new DoorError('pin'));
          } else if (!socket.authorized) return end(new DoorError('certificate'));
          return end(null, socket);
        });
      }
      socket.once('error', (err) => end(new DoorError(certFault(err) ? 'certificate' : 'unreachable')));
    } catch (err) {
      end(err instanceof DoorError ? err : new DoorError('unreachable'));
    }
  });
}

/**
 * One request to a door and its answer, read as JSON. The answer of a door
 * that tries to send Office elsewhere is a refusal like any other.
 * @returns {Promise<{ status: number, body: object }>}
 */
export async function askDoor(d, { method = 'POST', path, body = null }, { connect = connectDoor, timeoutMs = ANSWER_MS } = {}) {
  const socket = await connect(d);
  return new Promise((resolve, reject) => {
    const text = body == null ? '' : JSON.stringify(body);
    let done = false;
    const end = (err, ok) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      try { socket.destroy(); } catch { /* gone */ }
      if (err) reject(err); else resolve(ok);
    };
    const timer = setTimeout(() => end(new DoorError('unreachable')), timeoutMs);
    const req = http.request({
      createConnection: () => socket,
      method,
      path,
      headers: { host: d.host, accept: 'application/json', connection: 'close', ...(text ? { 'content-type': 'application/json', 'content-length': Buffer.byteLength(text) } : {}) },
    });
    req.on('error', () => end(new DoorError('unreachable')));
    req.on('response', (res) => {
      const chunks = [];
      let size = 0;
      res.on('data', (c) => { size += c.length; if (size > ANSWER_MOST) end(new DoorError('refused')); else chunks.push(c); });
      res.on('error', () => end(new DoorError('unreachable')));
      res.on('end', () => {
        let json = {};
        try { json = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}'); } catch { json = {}; }
        end(null, { status: res.statusCode || 0, body: json && typeof json === 'object' ? json : {} });
      });
    });
    req.end(text);
  });
}

/**
 * The meeting's signalling: a WebSocket through the door, made on a checked
 * connection. Only what the door passes on to its engine is asked for
 * (`/rtc` and below), and what the engine answers instead of taking the
 * connection up is told as a close with its status.
 *
 * @param {object} d a door from readDoor
 * @param {string} target the path and query, `/rtc...`
 * @param {object} on
 * @param {() => void} on.open
 * @param {(data: Buffer, binary: boolean) => void} on.message
 * @param {(e: { code: number, reason: string, status?: number, why?: string }) => void} on.close told once
 * @returns {{ send: (data: Buffer|string, binary?: boolean) => boolean, close: (code?: number, reason?: string) => void }}
 */
export function openSocket(d, target, on, { connect = connectDoor, protocols = [] } = {}) {
  let socket = null;
  let state = 'connecting';
  const closed = (e) => {
    if (state === 'closed') return;
    state = 'closed';
    try { socket?.destroy(); } catch { /* gone */ }
    on.close?.(e);
  };
  // `/rtc` and plain steps below it, then a query a request line can carry as it is: no step back up
  // (`/rtc/../admit`), and no character http.request would throw on instead of answering.
  if (!/^\/rtc(?:\/[\w-]+)*(?:\?[!-"$-~]*)?$/.test(String(target ?? ''))) {
    queueMicrotask(() => closed({ code: 1006, reason: '', why: 'path' }));
    return { send: () => false, close: () => {} };
  }
  const key = crypto.randomBytes(16).toString('base64');
  const wanted = [].concat(protocols || []).map(String).filter((p) => /^[\w.-]{1,64}$/.test(p));

  connect(d).then((s) => {
    if (state === 'closed') { try { s.destroy(); } catch { /* gone */ } return; }
    socket = s;
    const req = http.request({
      createConnection: () => s,
      method: 'GET',
      path: target,
      headers: { host: d.host, connection: 'Upgrade', upgrade: 'websocket', 'sec-websocket-version': '13', 'sec-websocket-key': key, ...(wanted.length ? { 'sec-websocket-protocol': wanted.join(', ') } : {}) },
    });
    req.on('error', () => closed({ code: 1006, reason: '', why: 'unreachable' }));
    // Not taken up as a WebSocket: the door's or the engine's refusal, with what it said.
    req.on('response', (res) => {
      const chunks = [];
      let size = 0;
      res.on('data', (c) => { size += c.length; if (size <= 4096) chunks.push(c); });
      const tell = () => closed({ code: 1006, reason: Buffer.concat(chunks).toString('utf8').slice(0, 300), status: res.statusCode || 0, why: 'refused' });
      res.on('end', tell);
      res.on('error', tell);
    });
    req.on('upgrade', (res, sock, head) => {
      if (String(res.headers['sec-websocket-accept'] || '') !== acceptKey(key)) return closed({ code: 1006, reason: '', why: 'refused' });
      state = 'open';
      const reader = frameReader({
        onMessage: (data, binary) => on.message?.(data, binary),
        onPing: (data) => { try { sock.write(encodeFrame(OP.pong, data)); } catch { /* closing */ } },
        onClose: (code, reason) => {
          try { sock.write(encodeFrame(OP.close, closePayload(code === 1005 ? 1000 : code))); } catch { /* gone */ }
          closed({ code, reason });
        },
        onError: (why) => closed({ code: 1002, reason: '', why }),
      });
      sock.on('data', (c) => reader.push(c));
      sock.on('error', () => closed({ code: 1006, reason: '', why: 'unreachable' }));
      sock.on('close', () => closed({ code: 1006, reason: '' }));
      socket = sock;
      on.open?.({ protocol: String(res.headers['sec-websocket-protocol'] || '') });
      if (head?.length) reader.push(head);
    });
    req.end();
  }, (err) => closed({ code: 1006, reason: '', why: err instanceof DoorError ? err.code : 'unreachable' }));

  return {
    send(data, binary = true) {
      if (state !== 'open') return false;
      try { socket.write(encodeFrame(binary ? OP.binary : OP.text, data)); return true; } catch { return false; }
    },
    close(code = 1000, reason = '') {
      if (state === 'closed') return;
      if (state === 'open') { try { socket.write(encodeFrame(OP.close, closePayload(code, reason))); } catch { /* gone */ } }
      closed({ code, reason });
    },
  };
}

/** The door's refusals that mean "this pass will not do, a new one might" (the contract, step 4). */
export const AGAIN = new Set(['expired', 'stale', 'used']);

/**
 * A pass exchanged for the engine's join token at the first door that
 * answers. { door, admitted } or a DoorError: the door's own refusal
 * (`expired`, `used`, `stale`, `ended`, `removed`, `not-serving`, ...), or
 * why no door could be used (`pin`, `certificate`, `plain`, `unreachable`).
 */
export async function exchange({ doors, pin = null, pass }, { ask = askDoor } = {}) {
  let last = new DoorError('unreachable');
  // What says "this is not the server the introducer tested" outranks "nothing answered".
  const rank = { pin: 3, certificate: 3, plain: 2, door: 2, unreachable: 1 };
  for (const door of [].concat(doors || [])) {
    let d;
    try { d = readDoor(door, pin); } catch (err) { if ((rank[err.code] || 0) >= (rank[last.code] || 0)) last = err; continue; }
    let got;
    try { got = await ask(d, { method: 'POST', path: '/admit', body: { pass } }); } catch (err) {
      const e = err instanceof DoorError ? err : new DoorError('unreachable');
      if ((rank[e.code] || 0) >= (rank[last.code] || 0)) last = e;
      continue;
    }
    if (got.status === 200 && typeof got.body.token === 'string' && got.body.token) return { door: d, admitted: got.body };
    // The server answered and said no: another of its doors would say the same.
    const code = typeof got.body.error === 'string' && /^[a-z-]{1,40}$/.test(got.body.error) ? got.body.error : 'refused';
    throw new DoorError(code, got.body.state ? { state: String(got.body.state) } : {});
  }
  throw last;
}
