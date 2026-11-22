// The guests of one Live Broadcast or Recorder window, for as long as it is
// open (Office 2.27.0, the Talk, Meet and Live plan's phase 3): the small
// Talk meeting the window holds, who sits in which seat, and what everyone
// in it is told about being on air. No drawing and no React here, so the
// whole of it is tested without a window; guests.js draws it.
//
// What it holds to:
// - The meeting is set up to be streamed live (Live Broadcast) or to be
//   recorded (the Recorder), which is chosen at set-up and said to everyone
//   before they join. A meeting set up for neither is never held here.
// - Only the person who set a meeting up holds it here.
// - Someone who comes in is never put on the frame by it: a seat that is
//   shown is taken off the frame first, and they sit once the frame changed.
// - Everyone in the meeting is told when it is on air or being recorded, and
//   each whether they are in the picture.
// - This computer's seat sends the programme, not a microphone or a camera.

import { withGuests, seatPeople, seatsShown, guestsShown, hideSeats } from '@rutba/studio/guests';
import { MeetingRoom } from '../apps/talk/meet-room.js';

const EMPTY = Object.freeze({ items: [], scenes: [], sound: 'on' });

export class GuestSession {
  /**
   * @param {object} o
   * @param {object} o.shell
   * @param {'live'|'record'} o.how  Live Broadcast's, or the Recorder's
   * @param {object} o.guests  the studio's guest sources (guest-source.js): set(seats, people), sync(stack)
   * @param {() => Promise<{ audio, video }>} o.feed  what this computer's seat sends: the programme
   * @param {(fn: (stack) => stack) => void} o.change  the setup's stack, and what is held on air, changed by `fn`
   * @param {(what: string, data: object) => void} [o.say]  something to tell the person: 'came' { name }, 'gone' { why, by }
   * @param {() => void} [o.onChange]
   * @param {() => object} [o.sceneNames]  the ready-made scenes' names, in the person's language
   * @param {object} [o.rtc]  a stand-in for the browser's WebRTC (the checks)
   */
  constructor({ shell, how, guests, feed, change, say = null, onChange = null, sceneNames = () => ({}), rtc = null }) {
    Object.assign(this, { shell, how, guests, feed, change, say, onChange, sceneNames, rtc });
    this.room = null;
    /** { [seat]: { id, name } } */
    this.seats = {};
    /** Talk as it is: { phase, me, contacts }. */
    this.talk = null;
    this.stack = EMPTY;
    this.airStack = EMPTY;
    this.onAir = false;
    /** The seats being taken off the frame for someone coming in, until the frame has changed. */
    this.cleared = '';
    this.told = '';
    this.ending = false;
    this.closed = false;
  }

  changed() { if (!this.closed) this.onChange?.(); }
  get online() { return this.talk?.phase === 'online' && Boolean(this.talk?.me?.id); }
  /** Whether a meeting was set up for what this window does. */
  fits(meeting) { return Boolean(this.how === 'live' ? meeting?.live : meeting?.recorded); }

  setTalk(state) { this.talk = state || null; this.changed(); }

  /** What the introducer said, as every window hears it. */
  handle(msg) {
    if (msg?.type === 'welcome' || String(msg?.type || '').startsWith('meet.')) this.room?.handle(msg);
  }

  /** The guests in the meeting now, this computer's own seat left out. */
  people() { return this.room && this.room.state !== 'gone' ? this.room.people().filter((p) => !p.me) : []; }
  person(id) { return this.people().find((p) => p.id === id) || null; }
  /** The people whose seat is shown in the scene on air. */
  shown() { return guestsShown(this.airStack, this.seats); }
  seated(id) { return Object.values(this.seats).some((w) => w.id === id); }

  /** Where a person is: 'connecting', 'failed', 'unseated' (every seat was taken), 'on-air' or 'off-air'. */
  where(id) {
    const p = this.person(id);
    if (!p?.connected) return p?.failed ? 'failed' : 'connecting';
    if (!this.seated(id)) return 'unseated';
    return this.onAir && this.shown().includes(id) ? 'on-air' : 'off-air';
  }

  /** The three seats and their scenes, in the setup and in what is held on air. */
  makeSeats() {
    const names = this.sceneNames();
    this.change((s) => withGuests(s, names));
  }

  /**
   * This window into a meeting, its seat the programme: { meeting } |
   * { error }. 'not-live' and 'not-recorded' say the meeting was not set up
   * for what this window does; 'not-yours' that someone else set it up.
   */
  async hold(id) {
    if (this.room) return { error: 'busy' };
    if (!this.online) return { error: 'offline' };
    const room = new MeetingRoom({
      shell: this.shell,
      me: { id: this.talk.me.id, name: this.talk.me.name || '' },
      feed: this.feed,
      rtc: this.rtc,
      onChange: (r) => {
        if (this.room !== r) return;
        if (r.state === 'gone') {
          this.room = null;
          this.told = '';
          this.cleared = '';
          this.step();
          // Leaving, being refused at the door and ending it oneself need no telling: the person did it, or is told why.
          if (r.why !== 'left' && r.why !== 'refused' && !this.ending) this.say?.('gone', { why: r.why, by: r.by });
          this.ending = false;
          this.changed();
          return;
        }
        this.step();
        this.changed();
      },
    });
    this.room = room;
    this.changed();
    const res = await room.join(id);
    if (res.error) {
      if (this.room === room) this.room = null;
      this.changed();
      return res;
    }
    if (!room.mine || !this.fits(res.meeting)) {
      room.leave();
      return { error: !room.mine ? 'not-yours' : this.how === 'live' ? 'not-live' : 'not-recorded' };
    }
    this.makeSeats();
    this.step();
    this.changed();
    return res;
  }

  /** Contacts asked in; the meeting is set up when there is none yet. { ok } | { error }. */
  async invite(ids, { name = '' } = {}) {
    const asked = [...new Set(ids || [])];
    if (!asked.length) return { error: 'nobody' };
    if (this.room) {
      const r = await this.room.invite(asked);
      return r.error ? r : { ok: true, told: r.told };
    }
    if (!this.online) return { error: 'offline' };
    const made = await this.shell.talk.meetCreate({ name, live: this.how === 'live', recorded: this.how === 'record', invite: asked, now: true }).catch(() => ({ error: 'offline' }));
    if (made.error) return made;
    const res = await this.hold(made.meeting.id);
    return res.error ? res : { ok: true, meeting: res.meeting };
  }

  /** The frames noted and nothing done about them yet: what where() and shown() answer from while the window draws. */
  look({ stack, airStack, onAir }) {
    this.stack = stack || EMPTY;
    this.airStack = airStack || this.stack;
    this.onAir = Boolean(onAir);
  }

  /**
   * The window's frames as they are now: the setup's stack, what is on air,
   * and whether the broadcast or the recording is running. Whose sound goes
   * out follows, and so does what everyone in the meeting is told.
   */
  sync(frames) {
    this.look(frames);
    this.guests.sync(this.airStack);
    // The frame is as the window has it now: a seat still on it is taken off again.
    this.cleared = '';
    // Another setup opened while guests are here has no seats yet: it is given them.
    if (this.room?.state === 'in' && !this.stack.items.some((it) => it.kind === 'guest')) { this.makeSeats(); return; }
    this.step();
  }

  /**
   * Who sits where, after people came or went or the frame changed. Answers
   * false while a seat is being taken off the frame for someone coming in:
   * nobody sits until the window says the frame has changed (sync), so no
   * picture and no sound of theirs is ever on a frame they did not wait off.
   */
  step() {
    // The connection being made again: everyone keeps their seat until it is known who is still there.
    if (this.room?.state === 'joining') return true;
    const people = this.people();
    const busy = new Set([...seatsShown(this.stack), ...seatsShown(this.airStack)]);
    const r = seatPeople(this.seats, people);
    const taken = r.came.map((c) => c.seat).filter((ref) => busy.has(ref));
    if (taken.length) {
      if (this.cleared !== taken.join()) {
        this.cleared = taken.join();
        this.change((s) => hideSeats(s, taken));
      }
      return false;
    }
    const moved = JSON.stringify(r.seats) !== JSON.stringify(this.seats);
    this.seats = r.seats;
    this.guests.set(this.seats, people);
    for (const c of r.came) this.say?.('came', c);
    this.tell();
    if (moved) this.changed();
    return true;
  }

  /** Everyone in the meeting told whether it is on air, and each whether they are in the picture; said once per change. */
  tell() {
    const room = this.room;
    if (!room?.id) return;
    const shown = this.shown();
    const key = JSON.stringify([room.id, this.onAir, shown]);
    if (key === this.told) return;
    // Noted before it is said: saying it changes the room, which looks here again.
    const was = this.told;
    this.told = key;
    if (!room.setAir({ on: this.onAir, how: this.how, shown })) this.told = was;
  }

  mute(id) { return this.room?.mute(id) || false; }
  remove(id) { return this.room ? this.room.remove(id) : Promise.resolve({ error: 'no-meeting' }); }
  /** The meeting ended for everyone: { ended } | { error }. */
  async end() {
    if (!this.room) return { error: 'no-meeting' };
    this.ending = true;
    const r = await this.room.end();
    this.ending = false;
    return r;
  }

  /** This window out of the meeting, which goes on for the others; answers its code, for Talk to come in with. */
  leave() {
    const id = this.room?.id || null;
    this.room?.leave();
    return id;
  }

  close() {
    this.leave();
    this.closed = true;
  }
}
