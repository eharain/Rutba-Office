// One call: the two computers connected straight to each other (WebRTC),
// or through a relay when they cannot be. The introducer passes the set-up
// between them (the offer, the answer, the connection candidates) and
// nothing else; the voice and the picture are encrypted end to end by
// WebRTC itself, whichever way they travel.
//
// Both ends always carry an audio and a video line, a line without a camera
// sending nothing, so the camera or a shared screen can come and go
// mid-call without the set-up being made again.

/** The safety code both people see: the same digits only when nobody sits between them. */
export async function safetyCode(localSdp, remoteSdp) {
  const prints = [localSdp, remoteSdp].map((sdp) => (/a=fingerprint:\S+ ([0-9A-F:]+)/i.exec(String(sdp || ''))?.[1] || '').toUpperCase());
  if (prints.some((p) => !p)) return '';
  const data = new TextEncoder().encode(prints.sort().join('|'));
  const hash = new Uint8Array(await crypto.subtle.digest('SHA-256', data));
  // Twelve digits in three groups of four, read out loud or compared on screen.
  let n = 0n;
  for (let i = 0; i < 8; i++) n = (n << 8n) | BigInt(hash[i]);
  const digits = (n % 1000000000000n).toString().padStart(12, '0');
  return `${digits.slice(0, 4)} ${digits.slice(4, 8)} ${digits.slice(8)}`;
}

export class Call {
  /**
   * @param {object} o
   * @param {object} o.shell
   * @param {string} o.callId
   * @param {{ id, name }} o.peer
   * @param {boolean} o.video whether this side starts with its camera
   * @param {boolean} o.outgoing
   * @param {(call: Call) => void} o.onChange
   */
  constructor({ shell, callId, peer, video, outgoing, onChange }) {
    Object.assign(this, { shell, callId, peer, video, outgoing, onChange });
    this.state = outgoing ? 'calling' : 'ringing';
    this.pc = null;
    this.mic = null;
    this.camera = null;
    this.screen = null;
    this.remote = new MediaStream();
    this.pending = [];
    this.muted = false;
    this.code = '';
    this.reason = '';
    this.startedAt = 0;
  }

  changed() { this.onChange?.(this); }

  async media(video) {
    this.mic = (await navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true }, video: false })).getAudioTracks()[0] || null;
    if (video) {
      try {
        this.camera = (await navigator.mediaDevices.getUserMedia({ video: { width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false })).getVideoTracks()[0] || null;
      } catch {
        this.camera = null;
      }
    }
  }

  async connection(extraRelay = null) {
    const setup = await this.shell.talk.callSetup();
    const iceServers = [...setup.iceServers, ...(extraRelay?.urls?.length ? [extraRelay] : [])];
    const pc = new RTCPeerConnection({ iceServers });
    this.pc = pc;
    pc.onicecandidate = (e) => {
      this.shell.talk.send({ type: 'call.candidate', callId: this.callId, candidate: e.candidate ? e.candidate.toJSON() : null }).catch(() => {});
    };
    pc.ontrack = (e) => {
      if (!this.remote.getTracks().includes(e.track)) this.remote.addTrack(e.track);
      this.changed();
    };
    pc.onconnectionstatechange = async () => {
      if (pc.connectionState === 'connected' && this.state !== 'connected') {
        this.state = 'connected';
        this.startedAt = Date.now();
        this.code = await safetyCode(pc.localDescription?.sdp, pc.remoteDescription?.sdp);
        this.changed();
      } else if (pc.connectionState === 'failed') {
        this.end('failed');
      }
    };
    return { pc, offerRelay: setup.offerRelay };
  }

  /** The two lines, audio and video, given this side's tracks. */
  attach() {
    for (const t of this.pc.getTransceivers()) {
      const kind = t.receiver.track?.kind;
      if (kind === 'audio') { t.sender.replaceTrack(this.mic).catch(() => {}); t.direction = 'sendrecv'; }
      if (kind === 'video') { t.sender.replaceTrack(this.screen || this.camera).catch(() => {}); t.direction = 'sendrecv'; }
    }
  }

  /** Calling: the offer goes out with this side's relay offered, if it has one of its own. */
  async place() {
    await this.media(this.video);
    const { pc, offerRelay } = await this.connection();
    pc.addTransceiver('audio', { direction: 'sendrecv' });
    pc.addTransceiver('video', { direction: 'sendrecv' });
    this.attach();
    await pc.setLocalDescription(await pc.createOffer());
    const sent = await this.shell.talk.send({ type: 'call.offer', to: this.peer.id, callId: this.callId, sdp: pc.localDescription.sdp, video: this.video, relay: offerRelay || undefined });
    this.changed();
    return sent;
  }

  /** Answering: the caller's offer, this side's lines, and the answer back. */
  async answer({ sdp, relay = null, video }) {
    this.video = Boolean(video);
    this.state = 'connecting';
    this.changed();
    await this.media(this.video);
    const { pc } = await this.connection(relay);
    await pc.setRemoteDescription({ type: 'offer', sdp });
    this.attach();
    await this.flush();
    await pc.setLocalDescription(await pc.createAnswer());
    await this.shell.talk.send({ type: 'call.answer', callId: this.callId, sdp: pc.localDescription.sdp });
  }

  async answered(sdp) {
    this.state = 'connecting';
    await this.pc.setRemoteDescription({ type: 'answer', sdp });
    await this.flush();
    this.changed();
  }

  async candidate(c) {
    if (!this.pc || !this.pc.remoteDescription) { this.pending.push(c); return; }
    await this.pc.addIceCandidate(c || null).catch(() => {});
  }

  async flush() {
    const list = this.pending.splice(0);
    for (const c of list) await this.pc.addIceCandidate(c || null).catch(() => {});
  }

  setMuted(on) {
    this.muted = on;
    if (this.mic) this.mic.enabled = !on;
    this.changed();
  }

  async setCamera(on) {
    if (on && !this.camera) {
      try {
        this.camera = (await navigator.mediaDevices.getUserMedia({ video: { width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false })).getVideoTracks()[0] || null;
      } catch {
        this.camera = null;
      }
    } else if (!on && this.camera) {
      this.camera.stop();
      this.camera = null;
    }
    this.video = Boolean(this.camera);
    if (this.pc) this.attach();
    this.changed();
  }

  /** A screen or window shared in place of the camera, and the camera back when it stops. */
  async share(stream) {
    this.screen?.stop();
    this.screen = stream ? stream.getVideoTracks()[0] || null : null;
    if (this.screen) this.screen.onended = () => this.share(null);
    if (this.pc) this.attach();
    this.changed();
  }

  /** Hang up, decline or let go: tells the other side unless it was the other side that ended it. */
  end(reason = 'hung-up', { tell = true } = {}) {
    if (this.state === 'ended') return;
    const was = this.state;
    this.state = 'ended';
    this.reason = reason;
    if (tell) {
      const type = was === 'ringing' && !this.outgoing ? 'call.decline' : 'call.end';
      this.shell.talk.send({ type, callId: this.callId, reason }).catch(() => {});
    }
    for (const t of [this.mic, this.camera, this.screen]) t?.stop();
    try { this.pc?.close(); } catch { /* closed */ }
    this.changed();
  }
}
