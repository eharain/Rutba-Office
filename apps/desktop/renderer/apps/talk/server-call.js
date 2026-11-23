// A meeting held on a Talk Server, as the window that is in it holds it: the
// same seat as in a small meeting (mesh.js: the microphone, the camera, a
// shared screen, a raised hand, what is written), with the server passing
// everyone's sound and picture on in place of each computer connecting to
// each other. That is what lets such a meeting hold more than four.
//
// The server passes on what it cannot read. Sound and picture are encrypted
// in this window by the engine with the meeting's key, which the people of
// the meeting hand each other (meet-key.js) and which no server is given.
// The line beside them (who is muted, whose hand is up, what is written) is
// sealed with that key by this file itself (wireSeal), so that what is
// heeded does not rest on what the engine says of a message: only words that
// open with the meeting's key, as from whoever the server says sent them,
// are read. The meeting is said to be encrypted only while every sound and
// picture that arrives is.
//
// A window of Office connects to nothing outside Office: the way to the
// server's door is made by Talk's service, which checks the door is the one
// the introducer named (main/talk-door.js), and the meeting's signalling
// travels through it (door-socket.js). Sound and picture go straight from
// this window to the server's media address, as a call's go to the other
// computer.
//
// The engine of such a meeting is LiveKit's client, the open-source library
// the server itself is built around. It is loaded when a meeting on a server
// is entered and never before.

import { Mesh, readWire } from './mesh.js';
import { lastRefusal, useDoorSockets } from './door-socket.js';
import { wireSeal } from './meet-key.js';

/** The line beside sound and picture, by name: anything else that arrives in the room is not ours. */
export const WIRE_TOPIC = 'rutba-talk';

const loadEngine = () => import('livekit-client');
const newWorker = () => new Worker('talk-worker.js');

/**
 * The meeting's key as the engine is given it: the 32 bytes written out as
 * text (base64url, 43 characters). Every one of LiveKit's clients takes a
 * key as text and draws the same working key from it, so whoever holds the
 * meeting's key can be in the meeting from any of them.
 */
export function engineKey(key) {
  let s = '';
  for (const b of key) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

/**
 * Why a meeting on a server could not be come into, in one of the words a
 * screen has a sentence for: from what the door or the engine answered the
 * connection with (`refusal`, door-socket.js) and what the engine itself said.
 */
export function whyNotConnected(lib, err, refusal) {
  const text = `${refusal?.reason || ''} ${err?.message || ''}`.toLowerCase();
  if (refusal?.why === 'pin' || refusal?.why === 'certificate' || refusal?.why === 'plain') return refusal.why;
  if (/max participants|room is full|exceeded/.test(text) || refusal?.status === 429) return 'full';
  if (refusal?.status === 401 || refusal?.status === 403 || refusal?.status === 404) return 'not-admitted';
  if (refusal?.status === 503) return 'not-serving';
  // The door was reached and the pass taken, and the engine behind it would not take the connection up. That is
  // what a meeting that already holds as many as it may looks like from here (LiveKit 1.13 answers 500 and no
  // reason), and nothing tells it from another fault of the engine, so the sentence for it says both.
  if (refusal?.status >= 500) return 'turned-away';
  if (refusal?.why === 'unreachable') return 'unreachable';
  const R = lib?.ConnectionErrorReason;
  if (R && err && typeof err.reason === 'number') {
    if (err.reason === R.NotAllowed) return 'not-admitted';
    if (err.reason === R.LeaveRequest) return 'not-admitted';
    if (err.reason === R.Cancelled) return 'left';
    if (err.reason === R.ServerUnreachable || err.reason === R.Timeout || err.reason === R.WebSocket) return 'unreachable';
  }
  return refusal ? 'unreachable' : 'failed';
}

/** Why a meeting on a server ended for this person, from what the engine said as it closed. */
export function whyGone(lib, reason) {
  const D = lib?.DisconnectReason || {};
  if (reason === D.PARTICIPANT_REMOVED) return 'removed';
  if (reason === D.ROOM_DELETED || reason === D.ROOM_CLOSED) return 'ended';
  if (reason === D.DUPLICATE_IDENTITY) return 'elsewhere';
  if (reason === D.CLIENT_INITIATED) return 'left';
  return 'lost';
}

export class ServerCall extends Mesh {
  /**
   * @param {object} o what a Mesh takes (me, onChange, onSay, mayMute, onMuted, mayAir, rtc), and:
   * @param {object} o.shell
   * @param {(why: string) => void} [o.onGone] the meeting ended for this person, not by their leaving: 'removed', 'ended', 'elsewhere', 'lost'
   * @param {() => Promise<object>} [o.load] the engine (the checks hand in a stand-in)
   * @param {() => Worker} [o.worker] the worker that encrypts and decrypts
   */
  constructor({ shell, onGone = null, load = loadEngine, worker = newWorker, ...rest }) {
    super({ ...rest, setup: async () => ({ iceServers: [] }), signal: () => {} });
    Object.assign(this, { shell, onGone, load, worker });
    this.lib = null;
    this.room = null;
    /** The way in Talk's service gave this window: its number there. */
    this.session = null;
    this.restore = null;
    /** Whether this person may ask the server to mute someone (they set the meeting up, or were made a host). */
    this.host = false;
    /** What this seat sends just now: { audio: { source, pub }, video: { source, pub } }. */
    this.sent = { audio: null, video: null };
    /** The connection to the server is being made again. */
    this.again = false;
    this.crypt = null;
    /** What seals and opens the line beside sound and picture (meet-key.js), from the meeting's key. */
    this.wire = null;
    /** What is said and what is heard on that line, each in its order. */
    this.outbox = Promise.resolve();
    this.inbox = Promise.resolve();
    /** The engine said that something which arrived did not open with this person's key: they hold another key than the others. */
    this.unreadable = false;
    /** In the room: from the moment the server took this person in until they leave or are put out. */
    this.inside = false;
  }

  /** Whether everything this person receives of the meeting is encrypted by the others' apps. */
  get encrypted() {
    if (!this.room) return false;
    for (const part of this.room.remoteParticipants?.values?.() || []) {
      for (const pub of part.trackPublications?.values?.() || []) if (pub.track && !pub.isEncrypted) return false;
    }
    return true;
  }

  /**
   * Into the meeting the server holds under `id`, with the meeting's key:
   * { entered: true } | { error, state? }. The microphone and camera are
   * this seat's already (start); they are sent once the room is open.
   */
  async enter({ id, key }) {
    if (!(key instanceof Uint8Array) || key.length !== 32) return { error: 'failed' };
    const entered = await Promise.resolve(this.shell.talk.serverEnter({ id })).catch(() => ({ error: 'offline' }));
    if (entered.error) return entered;
    if (this.closed || this.room) { Promise.resolve(this.shell.talk.serverLeave({ session: entered.session })).catch(() => {}); return { error: this.closed ? 'left' : 'failed' }; }
    // Held from here on, so that leaving at any moment (close) gives back whatever has been taken so far.
    this.session = entered.session;
    this.host = entered.host === true || entered.role === 'host';
    try {
      const lib = await this.load();
      const wire = await wireSeal(key, { meeting: String(id), me: this.me });
      if (this.closed) return { error: 'left' };
      this.lib = lib;
      this.wire = wire;
      this.unreadable = false;
      // The engine's own account of what it does stays out of the window's log, but for what went wrong.
      try { lib.setLogLevel?.('warn'); } catch { /* an engine without one */ }
      this.restore = useDoorSockets(this.shell, entered.session);
      this.crypt = this.worker();
      const keyProvider = new lib.ExternalE2EEKeyProvider();
      const room = new lib.Room({
        // Every picture is shown by the screen itself; the engine is not left to guess which are looked at.
        adaptiveStream: false,
        dynacast: true,
        stopLocalTrackOnUnpublish: false,
        disconnectOnPageLeave: true,
        encryption: { keyProvider, worker: this.crypt },
      });
      this.room = room;
      await keyProvider.setKey(engineKey(key));
      await room.setE2EEEnabled(true);
      this.hear(lib, room);
      await room.connect(entered.url, entered.token, { autoSubscribe: true });
    } catch (err) {
      if (this.closed) return { error: 'left' };
      const why = whyNotConnected(this.lib, err, lastRefusal(this.shell, entered.session));
      this.release();
      return { error: why };
    }
    if (this.closed) return { error: 'left' };
    this.inside = true;
    for (const part of this.room.remoteParticipants?.values?.() || []) this.came(part);
    this.tellAll();
    return { entered: true };
  }

  /** The room left and the way in given back; the seat's microphone and camera stay as they are. */
  release() {
    const { room, restore, crypt, session } = this;
    Object.assign(this, { room: null, restore: null, crypt: null, wire: null, session: null, inside: false, again: false, sent: { audio: null, video: null } });
    Promise.resolve()
      // The microphone and camera are the seat's, not the room's: they are not stopped with it.
      .then(() => room?.disconnect(false))
      .catch(() => {})
      .finally(() => {
        try { crypt?.terminate?.(); } catch { /* gone */ }
        restore?.();
        if (session != null) Promise.resolve(this.shell.talk.serverLeave({ session })).catch(() => {});
      });
  }

  /** What the room says, heard. */
  hear(lib, room) {
    const E = lib.RoomEvent;
    const ours = () => this.room === room && !this.closed;
    room.on(E.ParticipantConnected, (part) => { if (ours()) this.came(part); });
    room.on(E.ParticipantDisconnected, (part) => { if (ours() && this.unlink(part.identity)) this.changed(); });
    room.on(E.TrackSubscribed, (track, pub, part) => { if (ours()) this.arrived(track, part); });
    room.on(E.TrackUnsubscribed, (track, pub, part) => { if (ours()) this.left(track, part); });
    room.on(E.DataReceived, (payload, part, kind, topic, encryption, identity) => {
      if (!ours() || topic !== WIRE_TOPIC) return;
      const from = part?.identity || identity;
      // Only words that open with the meeting's key as this person's are heeded: anything else could have been
      // written by the server. One after the other, so they are read in the order they were said.
      this.inbox = this.inbox.then(async () => {
        const peer = this.peers.get(from);
        const wire = this.wire;
        if (!peer || !wire || !ours()) return;
        const text = await wire.open(payload, from, peer.heard);
        const m = text && readWire(text);
        if (m && ours() && this.peers.get(from) === peer) this.read(peer, m);
      }).catch(() => {});
    });
    if (E.EncryptionError) room.on(E.EncryptionError, () => { if (ours() && !this.unreadable) { this.unreadable = true; this.changed(); } });
    room.on(E.TrackMuted, (pub, part) => {
      // Muted from the server's side, at the asking of whoever may: this person is told, as in a small meeting.
      if (!ours() || part !== room.localParticipant || pub !== this.sent.audio?.pub || this.muted) return;
      if (this.mic) this.mic.enabled = false;
      this.muted = true;
      this.onMuted?.();
      this.tellAll();
    });
    room.on(E.ParticipantEncryptionStatusChanged, () => { if (ours()) this.changed(); });
    room.on(E.Reconnecting, () => { if (ours()) { this.again = true; this.changed(); } });
    room.on(E.Reconnected, () => { if (ours()) { this.again = false; this.tellAll(); } });
    room.on(E.Disconnected, (reason) => {
      if (!ours() || !this.inside) return;
      const why = whyGone(lib, reason);
      this.inside = false;
      if (why !== 'left') this.onGone?.(why);
    });
  }

  /** Someone in the room: known from this moment, and told how this person is. */
  came(part) {
    const id = part?.identity;
    if (!id || id === this.me || !this.room) return;
    let peer = this.peers.get(id);
    if (!peer) {
      peer = { id, name: part.name || this.names.get(id) || '', part, stream: new this.rtc.MediaStream(), state: { muted: false, camera: false, sharing: false, hand: false }, connected: true, failed: false, heard: { run: null, seq: 0 } };
      this.names.delete(id);
      this.peers.set(id, peer);
    } else {
      peer.part = part;
    }
    this.tell(peer);
    this.changed();
  }

  arrived(track, part) {
    this.came(part);
    const peer = this.peers.get(part?.identity);
    const mst = track?.mediaStreamTrack;
    if (!peer || !mst) return;
    // One sound and one picture each, as in a small meeting: what arrives takes the place of what was there.
    for (const old of peer.stream.getTracks().filter((t) => t.kind === mst.kind && t !== mst)) peer.stream.removeTrack(old);
    if (!peer.stream.getTracks().includes(mst)) peer.stream.addTrack(mst);
    this.changed();
  }

  left(track, part) {
    const peer = this.peers.get(part?.identity);
    const mst = track?.mediaStreamTrack;
    if (!peer || !mst || !peer.stream.getTracks().includes(mst)) return;
    peer.stream.removeTrack(mst);
    this.changed();
  }

  read(peer, m) {
    if (m.t === 'state') { peer.state = { muted: m.muted, camera: m.camera, sharing: m.sharing, hand: m.hand }; this.changed(); }
    else if (m.t === 'say') this.onSay?.(peer.id, m.text, m.id);
    else if (m.t === 'mute' && this.mayMute(peer.id) && !this.muted) { this.setMuted(true); this.onMuted?.(); }
    else if (m.t === 'air' && this.mayAir(peer.id)) { this.air = { on: m.on, you: m.you, how: m.on ? m.how : null }; this.airFrom = m.on ? peer.id : null; this.changed(); }
  }

  /** The others as a screen lists them: none of them is reached while the connection to the server is being made again. */
  people() { return super.people().map((p) => ({ ...p, connected: !this.again })); }

  /** In a small meeting each pair has its own two lines; here what this seat sends goes to the room once (sync). */
  attach() { /* nothing per person */ }

  /**
   * What the room is sent, made to match this seat: the microphone, and the
   * shared screen or else the camera. Taking its turn like every set-up, so
   * two changes never cross.
   */
  sync() {
    return this.later(async () => {
      const room = this.room;
      const S = this.lib?.Track?.Source;
      if (!room || !S) return;
      const local = room.localParticipant;
      const want = {
        audio: this.mic ? { source: this.mic, as: S.Microphone } : null,
        video: this.screen ? { source: this.screen, as: S.ScreenShare } : this.camera ? { source: this.camera, as: S.Camera } : null,
      };
      for (const kind of ['audio', 'video']) {
        const has = this.sent[kind];
        if (has && has.source !== want[kind]?.source) {
          this.sent[kind] = null;
          try { await local.unpublishTrack(has.pub.track, false); } catch { /* gone with the room */ }
        }
        if (want[kind] && !this.sent[kind]) {
          try {
            const pub = await local.publishTrack(want[kind].source, { source: want[kind].as });
            // Changed or closed while it was being sent: taken back.
            if (this.room !== room || this.closed) return;
            this.sent[kind] = { source: want[kind].source, pub };
          } catch { /* said by the seat still showing it as not sent */ }
        }
      }
      const audio = this.sent.audio?.pub;
      if (audio) {
        try {
          if (this.muted && !audio.isMuted) await audio.mute();
          else if (!this.muted && audio.isMuted) await audio.unmute();
        } catch { /* the room is going */ }
      }
    });
  }

  /** One message on the line beside sound and picture, to one person or to everyone: whether it could be sent. */
  send(m, to = null) {
    const local = this.room?.localParticipant;
    const wire = this.wire;
    if (!local || !wire || this.again) return false;
    this.outbox = this.outbox.then(async () => {
      const bytes = await wire.seal(m, to);
      // Left or closed while it was being sealed: not said.
      if (this.room?.localParticipant !== local) return;
      await local.publishData(bytes, { reliable: true, topic: WIRE_TOPIC, ...(to ? { destinationIdentities: [to] } : {}) });
    }).catch(() => {});
    return true;
  }

  to(peer, m) { return this.send(m, peer.id); }

  tellAll() {
    this.sync();
    this.later(() => {
      this.send({ t: 'state', ...this.mine() });
      // Who is in the picture differs from person to person, so each is told their own.
      if (this.airing) for (const p of this.peers.values()) this.to(p, { t: 'air', on: this.airing.on, you: this.airing.on && this.airing.shown.has(p.id), how: this.airing.how });
    });
    this.changed();
  }

  /* Nobody offers to anybody here: the room says who is in it. */
  call(person) { this.expect(person); return this.line; }
  signalled() { return this.line; }
  /** A connection is not made person by person: there is nothing to try again with one of them. */
  redial() { return false; }

  unlink(id) {
    const peer = this.peers.get(id);
    if (!peer) return false;
    this.peers.delete(id);
    if (this.airFrom === id) { this.air = { on: false, you: false, how: null }; this.airFrom = null; }
    return true;
  }

  /** Something written to everyone in the meeting: { id, text, to } with how many are there to read it, or null. */
  say(text) {
    const words = String(text || '').slice(0, 2000);
    if (!words.trim()) return null;
    const id = `s${Array.from(crypto.getRandomValues(new Uint8Array(9)), (b) => b.toString(16).padStart(2, '0')).join('')}`;
    const sent = this.send({ t: 'say', id, text: words });
    return { id, text: words, to: sent ? this.peers.size : 0 };
  }

  /**
   * One person asked to mute: their app is asked, which does it because the
   * request comes from whoever may ask; and the server is, by a host, so the
   * sound stops even from an app that would not listen.
   */
  askMute(id) {
    const peer = this.peers.get(id);
    if (!peer) return false;
    const asked = this.to(peer, { t: 'mute' });
    if (this.host && this.session != null) Promise.resolve(this.shell.talk.serverControl({ session: this.session, do: 'mute', person: id })).catch(() => {});
    return asked || this.host;
  }

  /** Out of the meeting: the room left, the way in given back, the microphone and camera let go. */
  close() {
    if (this.closed) return;
    this.release();
    super.close();
  }
}
