// A stand-in for the browser's WebRTC, for the checks of the group call
// (renderer/apps/talk/mesh.js) and of a meeting (talk/meet-room.js): peer
// connections that find each other through the set-up they exchange, with a
// line beside sound and picture that opens when the answer arrives. Nothing
// leaves the process.

export function fakeRtc({ microphone = true, camera = true } = {}) {
  const pcs = [];
  const asked = [];

  class Track {
    constructor(kind) { this.kind = kind; this.enabled = true; this.stopped = false; this.onended = null; }
    stop() { this.stopped = true; }
  }
  class Stream {
    constructor(tracks = []) { this.tracks = [...tracks]; }
    getTracks() { return this.tracks; }
    addTrack(t) { this.tracks.push(t); }
    removeTrack(t) { this.tracks = this.tracks.filter((x) => x !== t); }
    getAudioTracks() { return this.tracks.filter((t) => t.kind === 'audio'); }
    getVideoTracks() { return this.tracks.filter((t) => t.kind === 'video'); }
  }
  class Channel {
    constructor() { this.readyState = 'connecting'; this.other = null; this.sent = []; this.onopen = null; this.onmessage = null; }
    send(data) { this.sent.push(data); this.other?.onmessage?.({ data }); }
    close() { this.readyState = 'closed'; }
  }
  class PeerConnection {
    constructor(config) {
      Object.assign(this, { config, transceivers: [], localDescription: null, remoteDescription: null, candidates: [], closed: false, connectionState: 'new', channel: null });
      pcs.push(this);
      this.n = pcs.length;
    }
    addTransceiver(kind) {
      const tr = { receiver: { track: { kind } }, sender: { track: null, replaceTrack: async (t) => { tr.sender.track = t; } }, direction: 'recvonly' };
      this.transceivers.push(tr);
      return tr;
    }
    getTransceivers() { return this.transceivers; }
    createDataChannel(label) { this.channel = new Channel(); this.channel.label = label; return this.channel; }
    async createOffer() { return { type: 'offer', sdp: `offer:${this.n}` }; }
    async createAnswer() { return { type: 'answer', sdp: `answer:${this.n}` }; }
    async setLocalDescription(d) { this.localDescription = d; }
    async setRemoteDescription(d) {
      this.remoteDescription = d;
      const other = pcs[Number(String(d.sdp).split(':')[1]) - 1];
      if (!other) throw new Error('a set-up from nobody');
      if (d.type === 'offer') {
        // The side that answers: the same two lines, and its end of the line beside them.
        this.addTransceiver('audio');
        this.addTransceiver('video');
        const mine = new Channel();
        mine.other = other.channel;
        other.channel.other = mine;
        this.channel = mine;
        this.ondatachannel?.({ channel: mine });
      } else {
        other.open();
        this.open();
      }
    }
    open() {
      this.connectionState = 'connected';
      this.onconnectionstatechange?.();
      if (this.channel) { this.channel.readyState = 'open'; this.channel.onopen?.(); }
    }
    fail() { this.connectionState = 'failed'; this.onconnectionstatechange?.(); }
    async addIceCandidate(c) { this.candidates.push(c); }
    close() { this.closed = true; this.connectionState = 'closed'; }
  }

  return {
    PeerConnection,
    MediaStream: Stream,
    Track,
    getUserMedia: async (c) => {
      asked.push(c);
      if (c.audio && !microphone) throw new Error('no microphone');
      if (c.video && !camera) throw new Error('no camera');
      return new Stream([new Track(c.audio ? 'audio' : 'video')]);
    },
    /** Every connection made, in order. */
    pcs,
    /** What was asked of the microphone and camera, in order. */
    asked,
  };
}
