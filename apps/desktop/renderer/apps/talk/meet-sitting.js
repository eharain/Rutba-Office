// A meeting held on a Talk Server, from the moment this person is in the
// introducer's room for it (meet-room.js) until they leave: which server
// holds this sitting, who is let into it, and how the meeting's key reaches
// each person (the Talk server contract, "A meeting from start to end").
//
// The person who set the meeting up brings it to a server (Talk's service
// asks the introducer for one), makes the key, and comes in. Everybody else
// asks that person's computer, through the introducer's room, to be let in:
// it lets them in at the server and answers with the key sealed for them
// alone (meet-key.js). Nobody types or is shown a server's address, and the
// key goes to no server.
//
// So a person can come in only while the one who set the meeting up is in
// it. Until then they wait, and are told so.
//
// A sitting has an id of its own (the meeting's with a short suffix): the
// same meeting can sit again later, on another server, with another key.

import { askForKey, newMeetingKey, readSetUp, sealKeyFor } from './meet-key.js';

/** Asked again when no answer came in this long. */
const ASK_AGAIN_MS = 15000;
/** How often coming in is tried again by itself before the person is told. */
const TRIES = 2;
/**
 * What is worth another try: a pass that ran out, a server that was busy or has moved, a moment off the network,
 * and an engine that turned the connection away after the door let it in (asked again, the door says if the
 * meeting is full). A meeting that is over (`ended`) and one that is full are not: they stop, with their sentence.
 */
const AGAIN = new Set(['expired', 'used', 'stale', 'server-gone', 'not-serving', 'no-pass', 'no-meeting', 'removed', 'unreachable', 'turned-away']);
/** Refusals of the introducer that mean this sitting is over on its server: the meeting is brought to one anew. */
const OVER = new Set(['no-meeting', 'meeting-taken', 'server-gone']);
/** The words one app tells another when it does not let them in (meet-key.js reads the same). */
const NO = new Set(['refused', 'no-server', 'server-gone', 'offline', 'failed', 'full']);

/** An id for one sitting of a meeting: the meeting's own, with a short suffix. */
export const sittingId = (id) => `${id}-${Array.from(globalThis.crypto.getRandomValues(new Uint8Array(6)), (b) => b.toString(16).padStart(2, '0')).join('')}`;

export class Sitting {
  /**
   * @param {import('./meet-room.js').MeetingRoom} room the meeting as this window holds it; its `mesh` is a ServerCall
   * @param {{ wait?: number }} [o]
   */
  constructor(room, { wait = ASK_AGAIN_MS } = {}) {
    Object.assign(this, { room, wait });
    /** The sitting's id on its server, and the meeting's key. */
    this.mid = null;
    this.key = null;
    /** 'starting' (a server is asked for), 'waiting' (for the person who lets people in), 'asking' (they were asked), 'entering', 'in'. */
    this.step = room.mine ? 'starting' : 'waiting';
    /** What this person asked with, until it is answered. */
    this.asking = null;
    this.timer = null;
    this.stopped = false;
    /** Comings-in that did not hold, since the last that did. */
    this.tries = 0;
    /** Asked once more because nothing the others send could be opened. */
    this.looked = false;
    /** Coming into a room, one at a time. */
    this.turn = Promise.resolve();
  }

  get talk() { return this.room.shell.talk; }
  get call() { return this.room.mesh; }
  get mine() { return this.room.mine; }
  get starter() { return this.room.meeting?.starter?.id || null; }
  /** Whether a person is in the introducer's room for the meeting just now. */
  here(id) { return Boolean(this.call?.names.has(id)); }

  set(step) { if (this.step !== step) { this.step = step; this.room.changed(); } }
  tell(to, data) { Promise.resolve(this.talk.send({ type: 'meet.signal', id: this.room.id, to, data })).catch(() => {}); }
  later(work) {
    this.turn = this.turn.then(() => (this.stopped ? undefined : work())).catch(() => {});
    return this.turn;
  }

  /** This person is in the introducer's room: the sitting begins for them. */
  begin() { return this.mine ? this.later(() => this.bring()) : this.ask(); }

  stop() {
    this.stopped = true;
    clearTimeout(this.timer);
    this.asking = null;
  }

  /* ── the person who set the meeting up ──────────────────────────── */

  /**
   * The meeting brought to a server, and this person into it. Asked again
   * for a sitting that is on, the introducer answers the same server, or
   * moves the sitting, whose id then changes; the others are told to ask
   * again either way.
   */
  async bring(again = true) {
    if (this.call.room) this.call.leaveRoom();
    this.set('starting');
    const r = await Promise.resolve(this.talk.serverCarry({ id: this.mid || sittingId(this.room.id) })).catch(() => ({ error: 'offline' }));
    if (this.stopped) return undefined;
    if (r.error === 'meeting-taken' && this.mid && again) {
      // That sitting is over on its server (everybody had left it): a new one, with a new key.
      this.mid = null;
      this.key = null;
      return this.bring(false);
    }
    if (r.error || !r.meeting?.id) return this.fail(r.error || 'failed', r);
    this.mid = r.meeting.id;
    if (!this.key) this.key = newMeetingKey();
    // Whoever is in the meeting already asks now, and is let in while this person comes in.
    for (const id of this.call.names.keys()) this.tell(id, { t: 'srv-moved' });
    return this.enter();
  }

  /** Somebody in the meeting asked to be let in: let in at the server, and handed the key sealed for them. */
  async admit(from, pub) {
    // Not brought to a server yet: they are told to ask again when it is.
    if (!this.mid || !this.key || !this.here(from)) return;
    const { mid, key } = this;
    const r = await Promise.resolve(this.talk.serverAdmit({ id: mid, person: from })).catch(() => ({ error: 'offline' }));
    if (this.stopped || mid !== this.mid) return;
    if (r.error) {
      // The sitting is over on its server: brought to one anew, and everybody asks again.
      if (OVER.has(r.error)) { this.later(() => (mid === this.mid ? this.bring() : undefined)); return; }
      this.tell(from, { t: 'srv-no', why: NO.has(r.error) ? r.error : 'refused' });
      return;
    }
    const sealed = await sealKeyFor(pub, key, { meeting: this.room.id, from: this.room.me.id, to: from });
    if (sealed && !this.stopped) this.tell(from, { t: 'srv-let', mid, ...sealed });
  }

  /** Somebody taken out of the meeting by this person: out of the server's room too, with no new pass. */
  removed(who) {
    this.call?.names.delete(who);
    if (!this.mid) return Promise.resolve({});
    return Promise.resolve(this.talk.serverRemove({ id: this.mid, person: who })).catch(() => ({ error: 'offline' }));
  }

  /** The sitting ended on its server, for everyone. */
  end() {
    if (!this.mid) return Promise.resolve({});
    return Promise.resolve(this.talk.serverEnd({ id: this.mid })).catch(() => ({ error: 'offline' }));
  }

  /* ── everybody else ─────────────────────────────────────────────── */

  /** The person who set the meeting up is asked to let this one in, and asked again while no answer comes. */
  async ask() {
    clearTimeout(this.timer);
    if (this.stopped || this.mine) return;
    const starter = this.starter;
    if (!starter || !this.here(starter)) {
      // They are not in the meeting: nobody can let this person in until they are.
      this.asking = null;
      if (!this.call.inside) this.set('waiting');
      return;
    }
    const asking = await askForKey({ meeting: this.room.id, from: starter, to: this.room.me.id });
    if (this.stopped) return;
    // They left meanwhile: nobody to ask.
    if (!this.here(starter)) { this.asking = null; if (!this.call.inside) this.set('waiting'); return; }
    this.asking = asking;
    if (!this.call.inside) this.set('asking');
    this.tell(starter, { t: 'srv-ask', pub: asking.pub });
    this.timer = setTimeout(() => { if (this.asking === asking) this.ask(); }, this.wait);
    this.timer.unref?.();
  }

  /** Let in: the key opened, and into the server's room with it. */
  async let(m) {
    const asking = this.asking;
    if (!asking) return;
    const key = await asking.open({ pub: m.pub, iv: m.iv, box: m.box });
    // Not sealed for this asking: an answer to an earlier one, or not from who it says.
    if (!key || this.stopped || this.asking !== asking) return;
    this.asking = null;
    clearTimeout(this.timer);
    this.later(() => {
      this.mid = m.mid;
      this.key = key;
      return this.enter();
    });
  }

  /* ── both ───────────────────────────────────────────────────────── */

  /** Into the server's room for this sitting, with the meeting's key. */
  async enter() {
    const call = this.call;
    if (call.room) call.leaveRoom();
    this.set('entering');
    const r = await call.enter({ id: this.mid, key: this.key });
    if (this.stopped || r.error === 'left') return undefined;
    if (r.error) return this.notIn(r);
    this.tries = 0;
    this.set('in');
    return undefined;
  }

  /** Coming in did not hold: once or twice more by itself where that can help, then the person is told why. */
  notIn(r) {
    if (r.error === 'ended') {
      // It went on under a new id (on another server), which the person who set it up tells: asked again.
      if (r.moved === true && (this.tries += 1) <= TRIES) return this.mine ? this.bring() : this.ask();
      // Asked again meanwhile, because the meeting sits anew: the answer to that is waited for.
      if (!this.mine && this.asking) return undefined;
      // Over. For the person who set it up the meeting ends for everyone, so nobody waits in it.
      if (this.mine) { this.room.end().then((x) => { if (x?.error) this.out('ended'); }); return undefined; }
      return this.out('ended');
    }
    this.tries += 1;
    if (this.tries > TRIES || !AGAIN.has(r.error)) return this.fail(r.error, r);
    return this.mine ? this.bring() : this.ask();
  }

  /** The server's room closed for this person, not by their leaving. */
  dropped(why) {
    if (this.stopped) return;
    this.call.leaveRoom();
    // On another computer of theirs now: this one is out, as in a small meeting.
    if (why === 'elsewhere') { this.room.gone('elsewhere'); return; }
    // The person who set it up, with the sitting ended under them (on its server's own page): the meeting is over.
    if (why === 'ended' && this.mine) { this.room.end().then((r) => { if (r?.error) this.out('ended'); }); return; }
    this.tries += 1;
    if (this.tries > TRIES) { this.out(why === 'lost' ? 'lost' : why); return; }
    // The connection was lost: in again with a new pass. Put out or ended: the person who lets people in is asked,
    // who may have moved the meeting or changed this person's place in it; if the meeting is over, the introducer says so.
    if (why === 'lost') this.later(() => this.enter());
    else if (this.mine) this.later(() => this.bring());
    else this.ask();
  }

  /** What the introducer's room said: a set-up word from another person's app. */
  signalled(from, data) {
    if (this.stopped) return;
    const m = readSetUp(data);
    if (!m) return;
    if (this.mine) { if (m.t === 'srv-ask') this.admit(from, m.pub); return; }
    // Only the person who set the meeting up lets people in.
    if (from !== this.starter) return;
    if (m.t === 'srv-let') this.let(m);
    else if (m.t === 'srv-moved') this.ask();
    else if (m.t === 'srv-no' && this.asking) { clearTimeout(this.timer); this.asking = null; this.fail(m.why); }
  }

  /** Somebody came into the introducer's room. */
  joined(id) {
    if (!this.mine && id === this.starter && !this.call.inside && !this.asking) this.ask();
  }

  /** Somebody left the introducer's room. While they are in the server's room they are still heard, and still shown. */
  left(id) {
    const known = this.call.names.delete(id);
    // The person who lets people in has left: whoever was asking them waits for them to come back.
    if (!this.mine && id === this.starter) { clearTimeout(this.timer); this.asking = null; if (this.step === 'asking') this.set('waiting'); }
    if (known) this.room.changed();
  }

  /** The connection to the introducer was made again: back into its room, with the server's room left as it is. */
  async back() {
    const room = this.room;
    const r = await Promise.resolve(this.talk.meetJoin({ id: room.id })).catch(() => ({ error: 'offline' }));
    if (this.stopped) return;
    if (r.error) { this.out('lost', false); return; }
    room.meeting = r.meeting;
    this.call.names.clear();
    for (const p of r.meeting.people || []) { room.names.set(p.id, p.name || ''); this.call.expect(p); }
    room.changed();
    if (this.mine) {
      // Whoever asked meanwhile got no answer: they are told to ask again.
      if (this.mid) for (const id of this.call.names.keys()) this.tell(id, { t: 'srv-moved' });
      else this.later(() => this.bring());
    } else if (!this.call.inside) this.ask();
  }

  /** Nothing the others send opens here: the key may be an older one, so it is asked for once more. */
  look() {
    if (this.stopped || this.mine || this.looked || this.step !== 'in' || !this.call?.unreadableAll) return;
    this.looked = true;
    this.ask();
  }

  /** Out of the meeting for a reason the screens already have words for. */
  out(why, tell = true) {
    const id = this.room.id;
    this.room.gone(why);
    if (tell && id) Promise.resolve(this.talk.meetLeave({ id })).catch(() => {});
  }

  /** Out of the meeting because its server could not be come into: `room.fault` says why. */
  fail(why, more = {}) {
    const room = this.room;
    const id = room.id;
    // 'removed' and 'ended' have their own words already.
    if (why === 'removed' || why === 'ended') { this.out(why); return; }
    // A meeting just set up that no server can hold is ended, so nobody who was asked waits in it.
    const alone = this.mine && !this.mid && this.call.names.size === 0;
    room.fault = { why: why || 'failed', mine: this.mine, ...(Array.isArray(more?.own) ? { own: more.own } : {}), ...(Number.isInteger(more?.max) ? { max: more.max } : {}) };
    room.gone('server');
    if (!id) return;
    Promise.resolve(alone ? this.talk.meetEnd({ id }) : this.talk.meetLeave({ id })).catch(() => {});
  }
}
