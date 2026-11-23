// A meeting on a Talk Server: the window's connection to the meeting's
// signalling. A window of Office connects to nothing outside Office, and only
// the main side can check a server's certificate against what the introducer
// said of it; so the connection is made there (apps/desktop/main/talk-door.js)
// and this is the WebSocket the meeting's engine is handed in its place.
//
// It is given only for the address Talk's service made up for a way in
// (`wss://s<number>.talk-door.invalid`, which leads nowhere): every other
// WebSocket in the window is the browser's own, untouched.

const DOOR_HOST = 'talk-door.invalid';

/** The way in an address names, or null for any other address. */
export function sessionOf(url) {
  let u;
  try { u = new URL(String(url)); } catch { return null; }
  const m = /^s(\d{1,9})\.talk-door\.invalid$/.exec(u.hostname);
  if (!m || (u.protocol !== 'wss:' && u.protocol !== 'ws:')) return null;
  return { session: Number(m[1]), target: `${u.pathname}${u.search}` };
}

const event = (type, fields = {}) => Object.assign(new Event(type), fields);

/**
 * What every door socket of this window hears from the main side: one
 * listener for the window, each socket finding its own by number.
 */
const hubs = new WeakMap();
export function doorsOf(shell) {
  if (hubs.has(shell)) return hubs.get(shell);
  const open = new Map();
  /** What arrived for a socket whose number is not known yet (its opening is still being answered). */
  const early = new Map();
  const doors = { open, early, waiting: 0, refused: null };
  shell.on('talk:door', (msg) => {
    const s = open.get(msg?.socket);
    if (s) s.heard(msg);
    // Kept only while some socket is still waiting for its number: what is said of a socket already closed is dropped.
    else if (msg && doors.waiting > 0 && early.size < 64) early.set(msg.socket, [...(early.get(msg.socket) || []), msg].slice(-256));
  });
  hubs.set(shell, doors);
  return doors;
}

/** A WebSocket as the meeting's engine uses one, carried through Talk's service. */
export class DoorSocket extends EventTarget {
  static CONNECTING = 0;
  static OPEN = 1;
  static CLOSING = 2;
  static CLOSED = 3;

  constructor(shell, { session, target }, protocols = []) {
    super();
    this.shell = shell;
    this.session = session;
    this.url = `wss://s${session}.${DOOR_HOST}${target}`;
    this.readyState = DoorSocket.CONNECTING;
    this.binaryType = 'blob';
    this.protocol = '';
    this.extensions = '';
    this.bufferedAmount = 0;
    this.onopen = null;
    this.onmessage = null;
    this.onerror = null;
    this.onclose = null;
    this.number = null;
    this.doors = doorsOf(shell);
    /** Asked to close before the main side had answered with its number. */
    this.closing = null;
    this.doors.waiting += 1;
    const answered = () => { this.doors.waiting -= 1; if (this.doors.waiting === 0) this.doors.early.clear(); };
    Promise.resolve()
      .then(() => shell.talk.doorOpen({ session, target, protocols: [].concat(protocols || []) }))
      .then((r) => {
        if (!r || r.error || r.socket == null) { answered(); this.ended({ code: 1006, reason: '', why: r?.error || '' }, true); return; }
        this.number = r.socket;
        const early = this.doors.early.get(r.socket) || [];
        this.doors.early.delete(r.socket);
        answered();
        if (this.closing) { Promise.resolve(shell.talk.doorClose({ socket: r.socket, ...this.closing })).catch(() => {}); this.ended({ code: this.closing.code, reason: this.closing.reason }, false); return; }
        this.doors.open.set(r.socket, this);
        for (const msg of early) this.heard(msg);
      }, () => { answered(); this.ended({ code: 1006, reason: '' }, true); });
  }

  fire(type, fields) {
    const e = event(type, fields);
    try { this[`on${type}`]?.(e); } catch (err) { console.error(err); }
    this.dispatchEvent(e);
  }

  /** What the main side says of this socket. */
  heard(msg) {
    if (this.readyState === DoorSocket.CLOSED) return;
    if (msg.type === 'open') {
      if (this.readyState !== DoorSocket.CONNECTING) return;
      this.readyState = DoorSocket.OPEN;
      this.protocol = msg.protocol || '';
      this.fire('open');
    } else if (msg.type === 'message') {
      if (this.readyState !== DoorSocket.OPEN) return;
      let data = msg.data;
      if (msg.binary) {
        const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
        const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
        data = this.binaryType === 'arraybuffer' ? buffer : new Blob([buffer]);
      }
      this.fire('message', { data });
    } else if (msg.type === 'close') {
      // Refused before it was open (the engine's answer in `status` and `reason`): an error first, as a browser tells it.
      this.ended({ code: Number(msg.code) || 1006, reason: String(msg.reason || ''), status: msg.status || 0, why: msg.why || '' }, this.readyState === DoorSocket.CONNECTING || (msg.code !== 1000 && msg.code !== 1001));
    }
  }

  ended({ code, reason, status = 0, why = '' }, fault) {
    if (this.readyState === DoorSocket.CLOSED) return;
    this.readyState = DoorSocket.CLOSED;
    if (this.number != null) this.doors.open.delete(this.number);
    /** Why the door or the engine did not take the connection up, for whoever says it in words. */
    this.refusal = status || why ? { status, why, reason } : null;
    // Kept for the window too: the engine that held this socket says only that it could not connect.
    if (this.refusal) this.doors.refused = { session: this.session, ...this.refusal };
    if (fault) this.fire('error', { message: reason || why || '' });
    this.fire('close', { code, reason, wasClean: !fault });
  }

  send(data) {
    if (this.readyState !== DoorSocket.OPEN) throw new DOMException('The connection is not open.', 'InvalidStateError');
    if (typeof data === 'string') { Promise.resolve(this.shell.talk.doorSend({ socket: this.number, data, binary: false })).catch(() => {}); return; }
    const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : ArrayBuffer.isView(data) ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength) : null;
    if (!bytes) throw new TypeError('A door socket sends text or bytes.');
    // A copy: what is handed over must not change under the engine's feet.
    Promise.resolve(this.shell.talk.doorSend({ socket: this.number, data: bytes.slice(), binary: true })).catch(() => {});
  }

  close(code = 1000, reason = '') {
    if (this.readyState === DoorSocket.CLOSED || this.readyState === DoorSocket.CLOSING) return;
    if (this.number == null) { this.closing = { code, reason }; this.readyState = DoorSocket.CLOSING; return; }
    this.readyState = DoorSocket.CLOSING;
    Promise.resolve(this.shell.talk.doorClose({ socket: this.number, code, reason })).catch(() => {});
    this.ended({ code, reason }, false);
  }
}

/** The ways in this window holds just now: number -> the shell that reaches Talk's service. */
const ways = new Map();
let native = null;

/**
 * While a meeting on a server is open in this window, `WebSocket` gives a
 * door socket for the address of its way in and the browser's own for any
 * other. Answers what puts it back; the browser's own returns with the last
 * meeting.
 */
export function useDoorSockets(shell, session, scope = globalThis) {
  if (ways.size === 0) {
    native = scope.WebSocket;
    const Native = native;
    function Carried(url, protocols) {
      const way = sessionOf(url);
      const through = way && ways.get(way.session);
      if (through) return new DoorSocket(through, way, protocols);
      return protocols === undefined ? new Native(url) : new Native(url, protocols);
    }
    Carried.prototype = Native?.prototype;
    Object.assign(Carried, { CONNECTING: 0, OPEN: 1, CLOSING: 2, CLOSED: 3 });
    scope.WebSocket = Carried;
  }
  ways.set(session, shell);
  let done = false;
  return () => {
    if (done) return;
    done = true;
    if (ways.get(session) === shell) ways.delete(session);
    if (ways.size === 0) { scope.WebSocket = native; native = null; }
  };
}

/** Why the last connection of a way in was not taken up, read once: { status, why, reason } or null. */
export function lastRefusal(shell, session) {
  const doors = hubs.get(shell);
  const r = doors?.refused;
  if (!r || r.session !== session) return null;
  doors.refused = null;
  return { status: r.status, why: r.why, reason: r.reason };
}
