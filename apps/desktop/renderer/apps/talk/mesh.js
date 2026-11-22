// A small group call: up to four, each computer connected straight to each
// other (WebRTC), or through a relay when two of them cannot be. Sound and
// picture are encrypted end to end between each pair by WebRTC itself, and so
// is the line beside them that carries what the people of the call tell each
// other: who is muted, whose hand is up, what is written. None of it passes
// through a server of ours except, for a pair that needs it, through the
// relay, which cannot read it.
//
// Built once, for a small meeting in Talk first, and with nothing of Talk's
// screen in it, so Live Broadcast and the Recorder can use it for guests:
// whoever uses it says how a set-up reaches one person (`signal`) and hands
// in what arrives from one (`signalled`). The rule that makes the set-up
// simple: whoever comes in offers to everyone already there, and those there
// wait for the offer.
//
// As in a call (call.js), both ends of each pair always carry an audio and a
// video line, a line without a camera sending nothing, so the camera or a
// shared screen can come and go without the set-up being made again.

/** A call carried between the computers themselves holds this many (D-2026-10-20-11). */
export const MESH_MOST = 4;
/** What one written message holds. */
export const SAY_MOST = 2000;

/**
 * What travels between two computers of the call beside sound and picture,
 * checked as it arrives: anything else is dropped.
 *   { t: 'state', muted, camera, sharing, hand }   how that person is right now
 *   { t: 'say', id, text }                          something written to everyone
 *   { t: 'mute' }                                   a request to mute, honoured only from whoever may ask
 */
export function readWire(data) {
  let m;
  try { m = JSON.parse(String(data)); } catch { return null; }
  if (!m || typeof m !== 'object') return null;
  if (m.t === 'state') return { t: 'state', muted: m.muted === true, camera: m.camera === true, sharing: m.sharing === true, hand: m.hand === true };
  if (m.t === 'say') {
    const text = typeof m.text === 'string' ? m.text.slice(0, SAY_MOST) : '';
    return text.trim() && typeof m.id === 'string' && m.id && m.id.length <= 64 ? { t: 'say', id: m.id, text } : null;
  }
  if (m.t === 'mute') return { t: 'mute' };
  return null;
}

const sayId = () => `s${Array.from(crypto.getRandomValues(new Uint8Array(9)), (b) => b.toString(16).padStart(2, '0')).join('')}`;
const browserRtc = () => ({
  PeerConnection: globalThis.RTCPeerConnection,
  MediaStream: globalThis.MediaStream,
  getUserMedia: (c) => navigator.mediaDevices.getUserMedia(c),
});

export class Mesh {
  /**
   * @param {object} o
   * @param {string} o.me this person's id
   * @param {() => Promise<{ iceServers: object[], offerRelay?: object|null }>} o.setup what a connection is made with: the relays, and this person's own to offer the other side
   * @param {(to: string, data: object) => void} o.signal a set-up handed to whatever passes it to one person
   * @param {(mesh: Mesh) => void} [o.onChange] something a screen shows has changed
   * @param {(from: string, text: string, id: string) => void} [o.onSay] something written by someone else
   * @param {(from: string) => boolean} [o.mayMute] whether a request to mute from this person is honoured
   * @param {() => void} [o.onMuted] this person was muted by such a request
   * @param {object} [o.rtc] the browser's WebRTC, or a stand-in (the checks hand in their own)
   */
  constructor({ me, setup, signal, onChange = null, onSay = null, mayMute = () => false, onMuted = null, rtc = null }) {
    Object.assign(this, { me, setup, signal, onChange, onSay, mayMute, onMuted });
    this.rtc = rtc || browserRtc();
    /** id -> { id, name, pc, channel, stream, state, connected, failed, pending } */
    this.peers = new Map();
    /** People known to be in the call who have not offered yet: id -> name. */
    this.names = new Map();
    this.mic = null;
    this.camera = null;
    this.screen = null;
    this.muted = false;
    this.hand = false;
    this.closed = false;
    // One thing at a time: a set-up is never handled before the one that came ahead of it.
    this.line = Promise.resolve();
  }

  changed() { if (!this.closed) this.onChange?.(this); }
  later(work) {
    this.line = this.line.then(() => (this.closed ? undefined : work())).catch(() => {});
    return this.line;
  }

  /** This person's microphone, and their camera if they start with it. Without a microphone they listen. */
  async start({ video = false } = {}) {
    try {
      this.mic = (await this.rtc.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true }, video: false })).getAudioTracks()[0] || null;
    } catch {
      this.mic = null;
    }
    if (!this.mic) this.muted = true;
    if (video) await this.openCamera();
    this.changed();
  }

  async openCamera() {
    try {
      this.camera = (await this.rtc.getUserMedia({ video: { width: { ideal: 960 }, height: { ideal: 540 } }, audio: false })).getVideoTracks()[0] || null;
    } catch {
      this.camera = null;
    }
  }

  /** How this person is right now, as the others are told. */
  mine() { return { muted: this.muted || !this.mic, camera: Boolean(this.camera), sharing: Boolean(this.screen), hand: this.hand }; }

  /** The others, as a screen lists them. */
  people() { return [...this.peers.values()].map((p) => ({ id: p.id, name: p.name, stream: p.stream, connected: p.connected, failed: p.failed, ...p.state })); }

  /** The two lines of a pair, given this side's tracks. */
  attach(peer) {
    for (const tr of peer.pc.getTransceivers()) {
      const kind = tr.receiver?.track?.kind;
      if (kind === 'audio') { Promise.resolve(tr.sender.replaceTrack(this.mic)).catch(() => {}); tr.direction = 'sendrecv'; }
      if (kind === 'video') { Promise.resolve(tr.sender.replaceTrack(this.screen || this.camera)).catch(() => {}); tr.direction = 'sendrecv'; }
    }
  }

  /** The connection to one person, made new. */
  async link(person, extraRelay = null) {
    this.unlink(person.id);
    const setup = await this.setup();
    const pc = new this.rtc.PeerConnection({ iceServers: [...(setup.iceServers || []), ...(extraRelay?.urls?.length ? [extraRelay] : [])] });
    const peer = { id: person.id, name: person.name || this.names.get(person.id) || '', pc, channel: null, stream: new this.rtc.MediaStream(), state: { muted: false, camera: false, sharing: false, hand: false }, connected: false, failed: false, pending: [] };
    this.names.delete(person.id);
    this.peers.set(peer.id, peer);
    pc.onicecandidate = (e) => { if (this.peers.get(peer.id) === peer) this.signal(peer.id, { kind: 'candidate', candidate: e.candidate ? (e.candidate.toJSON?.() ?? e.candidate) : null }); };
    pc.ontrack = (e) => {
      if (!peer.stream.getTracks().includes(e.track)) peer.stream.addTrack(e.track);
      this.changed();
    };
    pc.onconnectionstatechange = () => {
      if (this.peers.get(peer.id) !== peer) return;
      if (pc.connectionState === 'connected') { peer.connected = true; peer.failed = false; } else if (pc.connectionState === 'failed') { peer.connected = false; peer.failed = true; } else return;
      this.changed();
    };
    pc.ondatachannel = (e) => this.wire(peer, e.channel);
    return { peer, offerRelay: setup.offerRelay || null };
  }

  /** The line beside sound and picture, for one pair. */
  wire(peer, channel) {
    peer.channel = channel;
    const tell = () => { if (this.peers.get(peer.id) === peer) this.tell(peer); };
    channel.onopen = tell;
    channel.onmessage = (e) => {
      if (this.peers.get(peer.id) !== peer) return;
      const m = readWire(e.data);
      if (!m) return;
      if (m.t === 'state') { peer.state = { muted: m.muted, camera: m.camera, sharing: m.sharing, hand: m.hand }; this.changed(); }
      else if (m.t === 'say') this.onSay?.(peer.id, m.text, m.id);
      else if (m.t === 'mute' && this.mayMute(peer.id) && !this.muted) { this.setMuted(true); this.onMuted?.(); }
    };
    if (channel.readyState === 'open') tell();
  }

  to(peer, m) {
    if (peer.channel?.readyState !== 'open') return false;
    try { peer.channel.send(JSON.stringify(m)); return true; } catch { return false; }
  }
  tell(peer) { this.to(peer, { t: 'state', ...this.mine() }); }
  tellAll() { for (const p of this.peers.values()) this.tell(p); this.changed(); }

  /** Someone already in the call when this person came in: offered to. */
  call(person) {
    if (!person?.id || person.id === this.me) return this.line;
    return this.later(async () => {
      const { peer, offerRelay } = await this.link(person);
      peer.pc.addTransceiver('audio', { direction: 'sendrecv' });
      peer.pc.addTransceiver('video', { direction: 'sendrecv' });
      this.wire(peer, peer.pc.createDataChannel('call'));
      this.attach(peer);
      await peer.pc.setLocalDescription(await peer.pc.createOffer());
      this.signal(peer.id, { kind: 'offer', sdp: peer.pc.localDescription.sdp, ...(offerRelay ? { relay: offerRelay } : {}) });
      this.changed();
    });
  }

  /** Someone who came in after this person: their offer is waited for. */
  expect(person) {
    if (!person?.id || person.id === this.me) return;
    this.names.set(person.id, person.name || '');
    this.changed();
  }

  /** A set-up from one person: their offer, their answer, or a way to reach them. */
  signalled(from, data) {
    if (!from || from === this.me || !data || typeof data !== 'object') return this.line;
    return this.later(async () => {
      if (data.kind === 'offer' && typeof data.sdp === 'string') {
        const { peer } = await this.link({ id: from }, data.relay && typeof data.relay === 'object' ? data.relay : null);
        await peer.pc.setRemoteDescription({ type: 'offer', sdp: data.sdp });
        this.attach(peer);
        await this.flush(peer);
        await peer.pc.setLocalDescription(await peer.pc.createAnswer());
        this.signal(from, { kind: 'answer', sdp: peer.pc.localDescription.sdp });
        this.changed();
        return;
      }
      const peer = this.peers.get(from);
      if (!peer) return;
      if (data.kind === 'answer' && typeof data.sdp === 'string') {
        await peer.pc.setRemoteDescription({ type: 'answer', sdp: data.sdp });
        await this.flush(peer);
        this.changed();
      } else if (data.kind === 'candidate') {
        if (!peer.pc.remoteDescription) peer.pending.push(data.candidate || null);
        else await Promise.resolve(peer.pc.addIceCandidate(data.candidate || null)).catch(() => {});
      }
    });
  }

  async flush(peer) {
    for (const c of peer.pending.splice(0)) await Promise.resolve(peer.pc.addIceCandidate(c)).catch(() => {});
  }

  unlink(id) {
    const peer = this.peers.get(id);
    if (!peer) return false;
    this.peers.delete(id);
    try { peer.channel?.close(); } catch { /* closed */ }
    try { peer.pc.close(); } catch { /* closed */ }
    return true;
  }

  /** Someone gone from the call. */
  drop(id) {
    return this.later(() => {
      const known = this.names.delete(id);
      if (this.unlink(id) || known) this.changed();
    });
  }

  setMuted(on) {
    this.muted = Boolean(on) || !this.mic;
    if (this.mic) this.mic.enabled = !this.muted;
    this.tellAll();
  }

  async setCamera(on) {
    if (on && !this.camera) await this.openCamera();
    else if (!on && this.camera) { this.camera.stop(); this.camera = null; }
    for (const p of this.peers.values()) this.attach(p);
    this.tellAll();
  }

  /** A screen or window shown in place of the camera, and the camera back when it stops. */
  share(stream) {
    this.screen?.stop();
    this.screen = stream ? stream.getVideoTracks()[0] || null : null;
    if (this.screen) this.screen.onended = () => this.share(null);
    for (const p of this.peers.values()) this.attach(p);
    this.tellAll();
  }

  setHand(on) {
    this.hand = Boolean(on);
    this.tellAll();
  }

  /** Something written to everyone in the call: { id, text, to } with how many it went to, or null for nothing to say. */
  say(text) {
    const words = String(text || '').slice(0, SAY_MOST);
    if (!words.trim()) return null;
    const id = sayId();
    let to = 0;
    for (const p of this.peers.values()) if (this.to(p, { t: 'say', id, text: words })) to += 1;
    return { id, text: words, to };
  }

  /** One person asked to mute. Their computer honours it only from whoever may ask. */
  askMute(id) {
    const peer = this.peers.get(id);
    return peer ? this.to(peer, { t: 'mute' }) : false;
  }

  /** Out of the call: every connection closed, the microphone and camera let go. */
  close() {
    if (this.closed) return;
    this.closed = true;
    for (const id of [...this.peers.keys()]) this.unlink(id);
    this.names.clear();
    for (const tr of [this.mic, this.camera, this.screen]) { try { tr?.stop(); } catch { /* stopped */ } }
    this.mic = null;
    this.camera = null;
    this.screen = null;
  }
}
