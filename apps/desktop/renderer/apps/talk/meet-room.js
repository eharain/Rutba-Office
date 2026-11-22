// One small meeting, as the window that is in it holds it: Talk's service
// asks the introducer to let this person in and passes each pair's set-up;
// the group call itself is mesh.js. Sound, picture, what is written, a raised
// hand and who is muted go between the computers and nowhere else; what is
// written is kept only while the meeting is open in this window.
//
// The person who set the meeting up can mute someone (their computer honours
// that only from them), take someone out, lock the meeting, ask more of their
// contacts in, and end it. In a meeting set up to be streamed live or to be
// recorded, they also tell the others when it is on air or being recorded,
// and each of them whether they are in the picture.
//
// Live Broadcast and the Recorder hold a meeting the same way, with a `feed`:
// their seat sends the programme in place of a microphone and a camera.

import { Mesh } from './mesh.js';

const SAYS_CAP = 500;

export class MeetingRoom {
  /**
   * @param {object} o
   * @param {object} o.shell
   * @param {{ id, name }} o.me
   * @param {(room: MeetingRoom) => void} [o.onChange]
   * @param {(room: MeetingRoom) => void} [o.onMuted] this person was muted by the person who set the meeting up
   * @param {() => Promise<{ audio: object|null, video: object|null }>} [o.feed] this seat's sound and picture, given (a broadcast's programme) in place of a microphone and a camera
   * @param {object} [o.rtc] a stand-in for the browser's WebRTC (the checks)
   */
  constructor({ shell, me, onChange = null, onMuted = null, feed = null, rtc = null }) {
    Object.assign(this, { shell, me, onChange, onMuted, feed, rtc });
    /** What this person, having set the meeting up, tells the others of its being on air: kept so it is said again after the connection was made anew. */
    this.airing = null;
    /** 'joining', 'in', or 'gone' with `why`: 'left', 'refused', 'removed', 'elsewhere', 'ended', 'lost'. */
    this.state = 'joining';
    this.why = '';
    this.by = null;
    this.meeting = null;
    this.mesh = null;
    /** What was written here: [{ id, from, name, text, at, mine }]. */
    this.says = [];
    this.unseen = 0;
    this.startedAt = 0;
    /** Names of the people who came and went, for what they wrote. */
    this.names = new Map();
    /** What the introducer said of this meeting while this person was still coming in. */
    this.early = [];
    this.joining = null;
  }

  changed() { this.onChange?.(this); }
  get id() { return this.meeting?.id || null; }
  get mine() { return Boolean(this.meeting) && this.meeting.starter?.id === this.me.id; }
  /** Whether the meeting is on air or being recorded just now, and whether this person is in the picture: { on, you, how }. */
  get air() { return this.mesh?.air || { on: false, you: false, how: null }; }

  /** Into the meeting: { meeting } | { error, most?, starter? }. The microphone is asked for first, so nobody is in unheard by accident. */
  async join(id, { video = false, again = false } = {}) {
    const mesh = new Mesh({
      me: this.me.id,
      rtc: this.rtc,
      setup: () => this.shell.talk.callSetup(),
      signal: (to, data) => { this.shell.talk.send({ type: 'meet.signal', id: this.id, to, data }).catch(() => {}); },
      onChange: () => this.changed(),
      onSay: (from, text, sid) => this.heard(from, text, sid),
      // Only the person who set the meeting up can mute someone.
      mayMute: (from) => from === this.meeting?.starter?.id,
      onMuted: () => this.onMuted?.(this),
      // And only they say when it is on air.
      mayAir: (from) => from === this.meeting?.starter?.id,
    });
    this.mesh = mesh;
    this.joining = id;
    this.early = [];
    if (this.feed) mesh.provide(await this.feed().catch(() => ({})));
    else await mesh.start({ video });
    const r = await this.shell.talk.meetJoin({ id }).catch(() => ({ error: 'offline' }));
    this.joining = null;
    if (r.error) {
      mesh.close();
      this.mesh = null;
      this.early = [];
      this.state = 'gone';
      // Refused at the door, or, for someone who was in it, not reachable again.
      this.why = again ? 'lost' : 'refused';
      this.changed();
      return r;
    }
    this.meeting = r.meeting;
    this.state = 'in';
    this.startedAt = Date.now();
    for (const p of r.meeting.people || []) { this.names.set(p.id, p.name || ''); mesh.call(p); }
    for (const msg of this.early.splice(0)) this.handle(msg);
    if (this.airing && this.mine) mesh.setAir(this.airing);
    this.changed();
    return r;
  }

  /** The people in the meeting, this person first: [{ id, name, me, starter, stream?, connected, muted, camera, sharing, hand }]. */
  people() {
    if (!this.mesh || !this.meeting) return [];
    const starter = this.meeting.starter?.id;
    const others = this.mesh.people();
    const waited = [...this.mesh.names].filter(([id]) => !others.some((p) => p.id === id)).map(([id, name]) => ({ id, name, stream: null, connected: false, failed: false, muted: false, camera: false, sharing: false, hand: false }));
    return [
      { id: this.me.id, name: this.me.name || '', me: true, starter: starter === this.me.id, connected: true, ...this.mesh.mine() },
      ...[...others, ...waited].map((p) => ({ ...p, name: p.name || this.names.get(p.id) || '', me: false, starter: starter === p.id })),
    ];
  }

  /** What the introducer said, as every window hears it. */
  handle(msg) {
    if (!msg) return;
    // Still coming in: what is said of this meeting meanwhile is kept for the moment this person is in.
    if (this.state === 'joining' && this.joining && (msg.id === this.joining || msg.meeting?.id === this.joining) && msg.type !== 'meet.joined') { this.early.push(msg); return; }
    if (this.state !== 'in') return;
    if (msg.type === 'welcome') {
      // The connection was made again: the introducer has this copy in no meeting, so it comes back in.
      this.rejoin();
      return;
    }
    if (msg.id !== this.id && msg.meeting?.id !== this.id) return;
    switch (msg.type) {
      case 'meet.signal':
        this.mesh.signalled(msg.from, msg.data);
        break;
      case 'meet.person':
        if (!msg.person?.id || msg.person.id === this.me.id) break;
        this.names.set(msg.person.id, msg.person.name || '');
        if (msg.event === 'joined') this.mesh.expect(msg.person);
        else this.mesh.drop(msg.person.id);
        this.changed();
        break;
      case 'meet.changed':
        this.meeting = { ...this.meeting, ...msg.meeting };
        this.changed();
        break;
      case 'meet.removed':
        this.gone(msg.why === 'elsewhere' ? 'elsewhere' : 'removed', msg.by || null);
        break;
      case 'meet.ended':
        this.gone('ended', msg.by || null);
        break;
      default:
        break;
    }
  }

  async rejoin() {
    const id = this.id;
    const video = Boolean(this.mesh?.camera);
    const muted = Boolean(this.mesh?.muted);
    const says = this.says;
    this.mesh?.close();
    this.state = 'joining';
    this.changed();
    const r = await this.join(id, { video, again: true });
    this.says = says;
    if (r.error) return;
    if (muted) this.mesh.setMuted(true);
  }

  heard(from, text, id) {
    if (this.says.some((s) => s.id === id)) return;
    const name = this.mesh?.peers.get(from)?.name || this.names.get(from) || '';
    this.says = [...this.says, { id, from, name, text, at: Date.now(), mine: false }].slice(-SAYS_CAP);
    this.unseen += 1;
    this.changed();
  }

  /** Something written to everyone in the meeting. */
  say(text) {
    const sent = this.mesh?.say(text);
    if (!sent) return null;
    this.says = [...this.says, { id: sent.id, from: this.me.id, name: this.me.name || '', text: sent.text, at: Date.now(), mine: true }].slice(-SAYS_CAP);
    this.changed();
    return sent;
  }

  seen() { if (this.unseen) { this.unseen = 0; this.changed(); } }

  gone(why, by = null) {
    if (this.state === 'gone') return;
    this.state = 'gone';
    this.why = why;
    this.by = by;
    this.mesh?.close();
    this.changed();
  }

  /** Out of the meeting; it goes on for the others. */
  leave() {
    const id = this.id;
    if (this.state === 'gone') return;
    this.gone('left');
    if (id) this.shell.talk.meetLeave({ id }).catch(() => {});
  }

  /* What only the person who set the meeting up can do: each answers { error } when the introducer refuses. */

  async lock(locked) {
    const r = await this.shell.talk.meetLock({ id: this.id, locked }).catch(() => ({ error: 'offline' }));
    if (!r.error && r.meeting) { this.meeting = { ...this.meeting, ...r.meeting }; this.changed(); }
    return r;
  }

  remove(who) { return this.shell.talk.meetRemove({ id: this.id, who }).catch(() => ({ error: 'offline' })); }

  /** More contacts asked in while the meeting is on: { told } | { error }. */
  invite(ids) { return this.shell.talk.meetInvite({ id: this.id, invite: ids }).catch(() => ({ error: 'offline' })); }

  /**
   * The others told that the meeting is on air ('live') or being recorded
   * ('record'), or no longer, and who of them is in the picture. Only in a
   * meeting that was set up for it, and only by the person who set it up:
   * answers whether it was said.
   */
  setAir({ on = false, how = 'live', shown = [] } = {}) {
    if (!this.mine || !this.meeting || !(how === 'record' ? this.meeting.recorded : this.meeting.live)) return false;
    this.airing = { on: Boolean(on), how, shown: [...shown] };
    this.mesh?.setAir(this.airing);
    return true;
  }

  /** One person muted: asked of their computer, which does it because the request comes from the person who set the meeting up. */
  mute(who) { return this.mine ? Boolean(this.mesh?.askMute(who)) : false; }

  async end() {
    const id = this.id;
    const r = await this.shell.talk.meetEnd({ id }).catch(() => ({ error: 'offline' }));
    if (!r.error) this.gone('ended');
    return r;
  }
}
