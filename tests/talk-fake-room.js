// A stand-in for the engine of a meeting on a Talk Server and for Talk's
// service, for the checks of the window's side (talk/server-call.js and
// talk/meet-room.js): rooms that find each other in one `world`, a line
// beside sound and picture that says whether what it carried was encrypted,
// and a shell that gives each person a way in and a door socket that opens.
// It answers by the names LiveKit's client uses and nothing of it is that
// library's code. Nothing leaves the process.

export const RoomEvent = {
  ParticipantConnected: 'participantConnected',
  ParticipantDisconnected: 'participantDisconnected',
  TrackSubscribed: 'trackSubscribed',
  TrackUnsubscribed: 'trackUnsubscribed',
  TrackMuted: 'trackMuted',
  DataReceived: 'dataReceived',
  EncryptionError: 'encryptionError',
  ParticipantEncryptionStatusChanged: 'participantEncryptionStatusChanged',
  Reconnecting: 'reconnecting',
  Reconnected: 'reconnected',
  Disconnected: 'disconnected',
};
export const Encryption_Type = { NONE: 0, GCM: 1, CUSTOM: 2 };
export const DisconnectReason = { UNKNOWN_REASON: 0, CLIENT_INITIATED: 1, DUPLICATE_IDENTITY: 2, SERVER_SHUTDOWN: 3, PARTICIPANT_REMOVED: 4, ROOM_DELETED: 5, ROOM_CLOSED: 10 };
export const ConnectionErrorReason = { NotAllowed: 0, ServerUnreachable: 1, InternalError: 2, Cancelled: 3, LeaveRequest: 4, Timeout: 5, WebSocket: 6 };
export const Track = { Source: { Microphone: 'microphone', Camera: 'camera', ScreenShare: 'screen_share' } };

class ConnectionError extends Error {
  constructor(message, reason) { super(message); this.reason = reason; }
}

/**
 * One server's rooms, and the people in them.
 * @param {object} [o]
 * @param {Record<string, string>} [o.names] id -> the name the server says of them
 */
export function fakeWorld({ names = {} } = {}) {
  /** Everyone in the room just now: identity -> Room. */
  const inside = new Map();
  /** What was asked of Talk's service, by everyone, in order. */
  const asked = [];
  /** Workers made, and whether each was stopped. */
  const workers = [];
  const keys = [];
  /** Every message put on the line beside sound and picture, as it travelled: { from, to, bytes }. */
  const carried = [];
  let sessions = 0;
  let sockets = 0;
  const world = {
    inside, asked, workers, keys, carried,
    /** What the next `serverEnter` answers in place of a way in. */
    enterRefusal: null,
    /** How the next connection to the door ends in place of opening: { status?, why?, reason? }. */
    doorRefusal: null,
    /** With it, the refusal above answers every connection, not the next one only. */
    doorRefusalStays: false,
    /** Made to wait at the door until `world.letIn()`. */
    holdDoor: false,
    waiting: [],
    letIn() { for (const go of world.waiting.splice(0)) go(); },
    /** Someone's sound stopped from the server's side. */
    serverMute(id) {
      const room = inside.get(id);
      const pub = [...room.localParticipant.pubs].find((p) => p.source === Track.Source.Microphone);
      pub.isMuted = true;
      pub.track.mediaStreamTrack.enabled = false;
      room.emit(RoomEvent.TrackMuted, pub, room.localParticipant);
    },
    /** Someone put out of the room, or the room closed on them. */
    putOut(id, reason) {
      const room = inside.get(id);
      room.gone();
      room.emit(RoomEvent.Disconnected, reason);
    },
    /** Something written on the line by the server itself, or by an app that does not encrypt. */
    inject(to, from, message, { topic = 'rutba-talk', encryption = Encryption_Type.NONE } = {}) {
      const room = inside.get(to);
      room.emit(RoomEvent.DataReceived, new TextEncoder().encode(JSON.stringify(message)), room.remoteParticipants.get(from), 0, topic, encryption, from);
    },
    /** Bytes put before someone as if `from` had sent them: what a server can do with what it passed on. */
    pass(to, from, bytes, { topic = 'rutba-talk' } = {}) {
      const room = inside.get(to);
      room.emit(RoomEvent.DataReceived, bytes, room.remoteParticipants.get(from), 0, topic, Encryption_Type.NONE, from);
    },
    /** A sound or picture that arrives without having been encrypted. */
    plainTrack(to, from, kind = 'audio') {
      const room = inside.get(to);
      const part = room.remoteParticipants.get(from);
      const track = { kind, mediaStreamTrack: { kind, enabled: true, plain: true } };
      const pub = { track, isEncrypted: false, source: kind === 'audio' ? Track.Source.Microphone : Track.Source.Camera };
      part.trackPublications.set(`plain-${kind}`, pub);
      room.emit(RoomEvent.TrackSubscribed, track, pub, part);
    },
  };

  class ExternalE2EEKeyProvider {
    async setKey(key) { this.key = key; keys.push(key); }
  }

  class Room {
    constructor(options) {
      this.options = options;
      this.hears = new Map();
      this.remoteParticipants = new Map();
      this.e2ee = false;
      this.identity = null;
      this.socket = null;
      const room = this;
      this.localParticipant = {
        pubs: new Set(),
        async publishTrack(mst, opts) {
          const pub = { source: opts.source, isMuted: false, isEncrypted: room.e2ee, track: { kind: mst.kind, mediaStreamTrack: mst }, async mute() { pub.isMuted = true; mst.enabled = false; room.emit(RoomEvent.TrackMuted, pub, room.localParticipant); }, async unmute() { pub.isMuted = false; mst.enabled = true; } };
          this.pubs.add(pub);
          for (const other of room.others()) {
            const me = other.remoteParticipants.get(room.identity);
            me.trackPublications.set(pub, pub);
            other.emit(RoomEvent.TrackSubscribed, pub.track, pub, me);
          }
          return pub;
        },
        async unpublishTrack(track) {
          const pub = [...this.pubs].find((p) => p.track === track);
          if (!pub) return;
          this.pubs.delete(pub);
          for (const other of room.others()) {
            const me = other.remoteParticipants.get(room.identity);
            me.trackPublications.delete(pub);
            other.emit(RoomEvent.TrackUnsubscribed, pub.track, pub, me);
          }
        },
        async publishData(bytes, { topic, destinationIdentities = null } = {}) {
          carried.push({ from: room.identity, to: destinationIdentities ? destinationIdentities[0] : null, bytes });
          for (const other of room.others()) {
            if (destinationIdentities && !destinationIdentities.includes(other.identity)) continue;
            // As LiveKit's client 2.25 does: what it says of a message's encryption is not to be gone by.
            other.emit(RoomEvent.DataReceived, bytes, other.remoteParticipants.get(room.identity), 0, topic, Encryption_Type.NONE, room.identity);
          }
        },
      };
    }

    on(event, fn) { this.hears.set(event, [...(this.hears.get(event) || []), fn]); return this; }
    emit(event, ...args) { for (const fn of this.hears.get(event) || []) fn(...args); }
    others() { return [...inside.values()].filter((r) => r !== this); }
    async setE2EEEnabled(on) { this.e2ee = Boolean(on); }

    /** Through the door first, as the real engine does: by the WebSocket the window gives it. */
    async connect(url, token) {
      this.identity = token;
      await new Promise((resolve, reject) => {
        const ws = new globalThis.WebSocket(`${url}/rtc/v1?access_token=${token}`);
        this.socket = ws;
        ws.onopen = () => resolve();
        ws.onerror = () => reject(new ConnectionError('could not connect', ConnectionErrorReason.ServerUnreachable));
        // Left while still connecting: the real engine gives the connection up the same way.
        ws.onclose = () => reject(new ConnectionError('cancelled', ConnectionErrorReason.Cancelled));
      });
      if (this.left) throw new ConnectionError('cancelled', ConnectionErrorReason.Cancelled);
      for (const other of this.others()) {
        this.remoteParticipants.set(other.identity, { identity: other.identity, name: names[other.identity] || '', trackPublications: new Map([...other.localParticipant.pubs].map((p) => [p, p])) });
      }
      inside.set(this.identity, this);
      for (const other of this.others()) {
        const me = { identity: this.identity, name: names[this.identity] || '', trackPublications: new Map() };
        other.remoteParticipants.set(this.identity, me);
        other.emit(RoomEvent.ParticipantConnected, me);
      }
      // What the others already send arrives once this person is in.
      for (const other of this.others()) {
        const them = this.remoteParticipants.get(other.identity);
        for (const pub of other.localParticipant.pubs) this.emit(RoomEvent.TrackSubscribed, pub.track, pub, them);
      }
    }

    gone() {
      if (inside.get(this.identity) !== this) return;
      inside.delete(this.identity);
      for (const other of this.others()) {
        const me = other.remoteParticipants.get(this.identity);
        other.remoteParticipants.delete(this.identity);
        other.emit(RoomEvent.ParticipantDisconnected, me);
      }
    }

    async disconnect(stopTracks = true) {
      this.left = true;
      // As the real engine: what was being sent is stopped with the room unless it is told not to.
      if (stopTracks) for (const pub of this.localParticipant.pubs) { try { pub.track.mediaStreamTrack.stop?.(); } catch { /* stopped */ } }
      this.gone();
      try { this.socket?.close(); } catch { /* closed */ }
    }
  }

  world.lib = { Room, RoomEvent, Track, Encryption_Type, DisconnectReason, ConnectionErrorReason, ConnectionError, ExternalE2EEKeyProvider };
  world.load = async () => world.lib;
  world.worker = () => { const w = { stopped: false, terminate() { w.stopped = true; } }; workers.push(w); return w; };

  /** Talk's service as one person's window reaches it. */
  world.shell = (id, { role = 'member' } = {}) => {
    const hears = [];
    const say = (msg) => { for (const fn of hears) fn(msg); };
    const mine = new Set();
    return {
      on: (event, fn) => { if (event === 'talk:door') hears.push(fn); return () => {}; },
      talk: {
        serverEnter: async (p) => {
          asked.push({ who: id, do: 'serverEnter', ...p });
          if (world.enterRefusal) { const r = world.enterRefusal; world.enterRefusal = null; return r; }
          sessions += 1;
          mine.add(sessions);
          return { session: sessions, url: `wss://s${sessions}.talk-door.invalid`, token: id, room: 'mtg_1', identity: id, role, host: role === 'host', e2ee: true, live: false };
        },
        serverLeave: async (p) => { asked.push({ who: id, do: 'serverLeave', ...p }); mine.delete(p.session); return { left: true }; },
        serverControl: async (p) => { asked.push({ who: id, ...p, do: 'serverControl', what: p.do }); return { done: true }; },
        doorOpen: async (p) => {
          asked.push({ who: id, do: 'doorOpen', ...p });
          if (!mine.has(p.session)) return { error: 'no-session' };
          sockets += 1;
          const socket = sockets;
          const answer = () => {
            if (world.doorRefusal) { const r = world.doorRefusal; if (!world.doorRefusalStays) world.doorRefusal = null; say({ socket, type: 'close', code: 1006, reason: r.reason || '', ...(r.status ? { status: r.status } : {}), ...(r.why ? { why: r.why } : {}) }); } else say({ socket, type: 'open', protocol: '' });
          };
          if (world.holdDoor) world.waiting.push(answer); else setTimeout(answer, 0);
          return { socket };
        },
        doorSend: async () => ({ sent: true }),
        doorClose: async (p) => { asked.push({ who: id, do: 'doorClose', ...p }); return { closed: true }; },
      },
    };
  };
  return world;
}
