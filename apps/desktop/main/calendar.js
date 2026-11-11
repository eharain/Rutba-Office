// The calendar.
//
// One JSON file in the profile, calendar.json: calendars, each with a name,
// a colour and its events as the iCalendar reader gives them. Nothing here
// reaches a server: a calendar an account brings (dav-sync.js) is marked
// with where it lives, and every change made here to one of its events is
// queued — the event's UID in `dirty`, a removed one's address in
// `deleted` — for the next sync to send. A window asks for the occurrences in the range it shows
// and gets every calendar's, expanded, with the recurring ones' exceptions
// applied; a file opened from disk is shown without being kept until the
// person says so; an invitation gets its answer written as the reply the
// organizer's calendar reads, and the answer is remembered (`answers`, by
// UID), so the invitation opened again says what was answered.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { readCalendar, writeCalendar, writeReply, isInvitation, occurrences, exceptionsOf, dateTimeAt, describeRule, localZone } from '@rutba/calendar';
import { createMailLink } from './calendar-mail.js';
import { createInvites } from './calendar-invites.js';
import { roomsFor, busyInCalendars } from './calendar-rooms.js';
import { createRoomDesk } from './calendar-desk.js';
import { createAvailability } from './calendar-availability.js';
import { wordsFrom } from './calendar-text.js';

const readPlain = (file, fallback) => {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return fallback;
  }
};
const writePlain = (file, value) => {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(value));
  fs.renameSync(tmp, file);
};
const writePlainLater = async (file, value) => {
  await fs.promises.mkdir(path.dirname(file), { recursive: true });
  const tmp = `${file}.${process.pid}.tmp`;
  await fs.promises.writeFile(tmp, JSON.stringify(value));
  await fs.promises.rename(tmp, file);
};

const COLOURS = ['#2b5fd9', '#1a9f7a', '#d9534f', '#e08b2b', '#7b5cd6', '#c2408f', '#3aa0b5', '#8c6d1f'];

export function createCalendarService({ stores, broadcast, mail = null, contacts = null, freeBusyTokens = null }) {
  // Mail as calendar-mail.js speaks of it: handed over so (the pack's), or Mail's own service wrapped.
  const link = mail && typeof mail.sendFrom === 'function' && typeof mail.inboxOf === 'function' ? mail : createMailLink(() => mail);
  const file = path.join(stores.dir, 'calendar.json');
  // Sealed with the profile's key while it has a password (stores.vault, profiles.js).
  const vault = stores.vault || null;
  vault?.cover(file);
  const readJson = vault ? vault.readJson : readPlain;
  const writeJson = vault ? vault.writeJson : writePlain;
  const writeJsonLater = vault ? vault.writeJsonLater : writePlainLater;
  const state = readJson(file, null) || { calendars: [] };
  if (!Array.isArray(state.calendars)) state.calendars = [];
  if (!state.calendars.length) state.calendars.push({ id: crypto.randomUUID(), name: 'My calendar', colour: COLOURS[0], visible: true, events: [] });

  const hooks = { local: null, freeBusy: null };
  // Writes coalesced: a burst of changes — an import of hundreds, a sync —
  // was a rewrite of the whole file for each one, on the main process; now
  // it is one write a moment after the last, off its way, and one at once
  // when the app quits.
  let pending = null;
  const writeNow = () => {
    if (!pending) return;
    clearTimeout(pending);
    pending = null;
    writeJson(file, state);
  };
  process.once('exit', writeNow);
  const save = () => {
    if (!pending) {
      pending = setTimeout(() => {
        pending = null;
        writeJsonLater(file, state).catch(() => writeJson(file, state));
      }, 250);
      pending.unref?.();
    }
    broadcast?.('calendar:changed', {});
  };
  /** An event of a server's calendar changed here: queued for the server. */
  const touch = (cal, uid) => {
    if (!cal?.remote || !uid) return;
    cal.dirty = [...new Set([...(cal.dirty || []), uid])];
    hooks.local?.();
  };
  /** An event of a server's calendar gone from here: its address queued for removal, or what is left of its series queued. */
  const forget = (cal, event) => {
    if (!cal?.remote || !event) return;
    if (cal.events.some((e) => e.uid && e.uid === event.uid)) return touch(cal, event.uid);
    if (event.remote?.href) cal.deleted = [...(cal.deleted || []), { href: event.remote.href, etag: event.remote.etag || null }];
    cal.dirty = (cal.dirty || []).filter((u) => u !== event.uid);
    hooks.local?.();
  };
  // `ai`: what the AI tools may do with the calendar (mcp/reach.js), kept with it; Off until chosen.
  const listed = () => state.calendars.map((c) => ({ id: c.id, name: c.name, colour: c.colour, visible: c.visible !== false, count: c.events.length, account: c.remote?.account || null, ai: c.ai || 'off' }));
  /**
   * This person's answer to an invitation already given: the one remembered
   * when it was answered here, or their line in the event kept from it (an
   * answer from before answers were remembered, or from another device).
   * Not for a later version of the invitation (a higher SEQUENCE): that one
   * asks again, as a changed meeting does. Answers { partstat, email }, the
   * email being the line answered for (the first, when no mail account here
   * is among the people invited), or null.
   */
  const answeredBefore = (invitation, email) => {
    const sequence = invitation.sequence || 0;
    const said = invitation.uid ? state.answers?.[invitation.uid] : null;
    if (said && sequence <= (said.sequence || 0)) return { partstat: said.partstat, email: said.email || null };
    if (!email) return null;
    for (const cal of state.calendars) {
      const kept = cal.events.find((e) => e.uid && e.uid === invitation.uid && !e.recurrenceId);
      const line = kept?.attendees?.find((a) => String(a.email || '').toLowerCase() === String(email).toLowerCase());
      if (line?.partstat && line.partstat !== 'NEEDS-ACTION' && sequence <= (kept.sequence || 0)) return { partstat: line.partstat, email: line.email };
    }
    return null;
  };
  /** An answer kept by the invitation's UID; the oldest go past a thousand. */
  const remember = (event, partstat, email) => {
    if (!event.uid) return;
    const answers = { ...(state.answers || {}), [event.uid]: { partstat, email: email || null, sequence: event.sequence || 0, at: Date.now() } };
    const uids = Object.keys(answers);
    if (uids.length > 1000) for (const uid of uids.sort((a, b) => answers[a].at - answers[b].at).slice(0, uids.length - 1000)) delete answers[uid];
    state.answers = answers;
  };
  const calendarOf = (id) => state.calendars.find((c) => c.id === id) || null;
  const find = (eventId) => {
    for (const cal of state.calendars) {
      const event = cal.events.find((e) => e.id === eventId);
      if (event) return { cal, event };
    }
    return null;
  };

  /** The addresses this person is known by: the mail accounts' own, but for those run here as a room's mailbox. */
  const myAddresses = () => {
    try {
      return (link.accounts() || []).map((a) => String(a.email || '').toLowerCase()).filter((e) => e && !desk?.isRoom(e));
    } catch {
      return [];
    }
  };
  let desk = null;
  const invites = createInvites({ stores, state: () => state, find, save, touch, mail: link, myAddresses });
  desk = createRoomDesk({ state: () => state, save, link, invites });
  /** The rooms and resources Contacts has (contacts.bookable): none while Contacts is absent. */
  const bookable = () => {
    try {
      return (typeof contacts === 'function' ? contacts() : contacts)?.bookable?.() || [];
    } catch {
      return [];
    }
  };
  // What is known of a room's time (calendar-rooms.js): the meetings here that have it on, to begin with.
  // And the rooms whose mailbox is run here, which know their whole day (calendar-desk.js).
  const roomSources = [{ busy: (email, from, to, except) => ({ busy: busyInCalendars(state.calendars, email, from, to, except), whole: false, source: 'calendar' }) }, desk.source];
  // Free/busy on request (calendar-availability.js): a server's, else asked by mail; what comes back is a source too.
  const availability = createAvailability({ state: () => state, save, link, invites, sources: [...roomSources], hooks, tokens: freeBusyTokens || (() => null) });
  roomSources.push(availability.source);

  /** One occurrence, as the window draws it. */
  function slot(cal, o, fileName = null) {
    const e = o.event;
    return {
      id: e.id || null,
      calendarId: cal.id,
      calendar: cal.name,
      colour: cal.colour,
      file: fileName,
      start: o.start,
      end: o.end,
      original: o.original,
      allDay: Boolean(e.allDay),
      summary: e.summary || '(no title)',
      location: e.location || '',
      recurring: Boolean(e.rrule) || Boolean(e.recurrenceId),
      status: e.status || null,
      organizer: e.organizer || null,
      attendees: e.attendees?.length || 0,
    };
  }

  function expand(cal, from, to, fileName = null) {
    const out = [];
    const events = cal.events || [];
    const siblingsOf = exceptionsOf(events);
    for (const e of events) {
      if (e.recurrenceId) continue; // drawn through its parent
      for (const o of occurrences(e, { from, to, siblings: siblingsOf(e), limit: 500 })) out.push(slot(cal, o, fileName));
    }
    return out;
  }

  /** A stored event, whole, with what the editor wants to show beside it. */
  function present(cal, e) {
    // Whether this person organizes it (no organizer yet, or one of their addresses): only then are its invitations sent from here.
    const organizedHere = !e.organizer?.email || myAddresses().includes(String(e.organizer.email).toLowerCase());
    return { ...e, calendarId: cal.id, calendar: cal.name, colour: cal.colour, rule: e.rrule ? describeRule(e.rrule, e.start) : 'Does not repeat', organizedHere };
  }

  const service = {
    calendars: () => listed(),

    saveCalendar: ({ calendar }) => {
      const kept = calendar.id ? calendarOf(calendar.id) : null;
      if (kept) Object.assign(kept, { name: calendar.name ?? kept.name, colour: calendar.colour ?? kept.colour, visible: calendar.visible ?? kept.visible, ai: calendar.ai ?? kept.ai });
      else state.calendars.push({ id: crypto.randomUUID(), name: calendar.name || 'Calendar', colour: calendar.colour || COLOURS[state.calendars.length % COLOURS.length], visible: true, events: [] });
      save();
      return listed();
    },

    removeCalendar: ({ id }) => {
      if (state.calendars.length <= 1) throw new Error('The last calendar stays; delete its events instead.');
      if (calendarOf(id)?.remote) throw new Error('This calendar comes from an account — remove the account to remove its calendars.');
      state.calendars = state.calendars.filter((c) => c.id !== id);
      save();
      return { removed: 1 };
    },

    /** Every occurrence between two instants, across the visible calendars. */
    events: ({ from, to, all = false }) => {
      const out = [];
      for (const cal of state.calendars) {
        if (!all && cal.visible === false) continue;
        out.push(...expand(cal, from, to));
      }
      out.sort((a, b) => a.start - b.start || Number(b.allDay) - Number(a.allDay));
      return out;
    },

    get: ({ id }) => {
      const hit = find(id);
      return hit ? present(hit.cal, hit.event) : null;
    },

    /**
     * Keep an event. `scope` for a recurring one: 'all' changes the series,
     * 'this' detaches the occurrence at `original` as an exception.
     */
    save: ({ calendarId, event, scope = 'all', original = null }) => {
      const cal = calendarOf(calendarId) || (event.id && find(event.id)?.cal) || state.calendars[0];
      const now = Date.now();
      const zone = event.start?.tzid || localZone();
      const clean = {
        ...event,
        uid: event.uid || `${crypto.randomUUID()}@rutba.io`,
        dtstamp: dateTimeAt(now, null),
        start: event.start,
        end: event.end,
        allDay: Boolean(event.allDay),
        exdates: event.exdates || [],
        rdates: event.rdates || [],
        alarms: event.alarms || [],
        attendees: event.attendees || [],
        extra: event.extra || [],
        calendarId: undefined, calendar: undefined, colour: undefined, rule: undefined, organizedHere: undefined,
      };
      if (event.id && scope === 'this' && original != null) {
        // An exception: a copy with a RECURRENCE-ID naming the instance it replaces.
        const parent = find(event.id);
        if (parent) {
          const exception = { ...clean, id: crypto.randomUUID(), rrule: null, recurrenceId: dateTimeAt(original, parent.event.start.tzid || zone, parent.event.allDay), updatedAt: now };
          cal.events = cal.events.filter((e) => !(e.uid === exception.uid && e.recurrenceId && Math.abs(e.recurrenceId.at - original) < 1000));
          cal.events.push(exception);
          touch(cal, exception.uid);
          save();
          return present(cal, exception);
        }
      }
      const hit = event.id ? find(event.id) : null;
      if (hit) {
        // Moved to another calendar, or changed in place.
        const moved = hit.cal !== cal;
        if (moved) { hit.cal.events = hit.cal.events.filter((e) => e !== hit.event); forget(hit.cal, hit.event); }
        const kept = { ...hit.event, ...clean, id: hit.event.id, sequence: (hit.event.sequence || 0) + 1, updatedAt: now, ...(moved ? { remote: undefined } : {}) };
        const at = cal.events.indexOf(hit.event);
        if (at >= 0) cal.events[at] = kept;
        else cal.events.push(kept);
        touch(cal, kept.uid);
        save();
        return present(cal, kept);
      }
      const fresh = { ...clean, id: crypto.randomUUID(), createdAt: now, updatedAt: now, remote: undefined };
      cal.events.push(fresh);
      touch(cal, fresh.uid);
      save();
      return present(cal, fresh);
    },

    /** Remove an event; for a recurring one, 'this' removes only the occurrence at `original`. */
    remove: ({ id, scope = 'all', original = null }) => {
      const hit = find(id);
      if (!hit) return { removed: 0 };
      if (scope === 'this' && hit.event.rrule && original != null) {
        hit.event.exdates = [...(hit.event.exdates || []), dateTimeAt(original, hit.event.start.tzid || null, hit.event.allDay)];
        hit.cal.events = hit.cal.events.filter((e) => !(e.uid === hit.event.uid && e.recurrenceId && Math.abs(e.recurrenceId.at - original) < 1000));
        touch(hit.cal, hit.event.uid);
        save();
        return { removed: 1, scope: 'this' };
      }
      hit.cal.events = hit.cal.events.filter((e) => e !== hit.event && !(e.uid && e.uid === hit.event.uid && e.recurrenceId));
      forget(hit.cal, hit.event);
      save();
      return { removed: 1, scope: 'all' };
    },

    /**
     * What a file holds, without keeping it: its events expanded in a range,
     * and whether it is an invitation to answer.
     */
    openFile: ({ path: target = null, text = null, name: given = null, from, to }) => {
      const cal = readCalendar(text != null ? text : fs.readFileSync(target));
      const name = cal.name || given || (target ? path.basename(target) : 'Invitation');
      const temp = { id: `file:${target || name}`, name, colour: '#8c6d1f', events: cal.events.map((e, i) => ({ ...e, id: `file:${i}` })) };
      const mine = myAddresses();
      const invitation = isInvitation(cal) ? cal.events[0] : null;
      const me = invitation ? invitation.attendees.find((a) => mine.includes(String(a.email || '').toLowerCase())) || null : null;
      // The file still says NEEDS-ACTION once it is answered: the answer given is what it says now.
      const answered = invitation ? answeredBefore(invitation, me?.email) : null;
      return {
        path: target,
        name,
        method: cal.method,
        count: cal.events.length,
        events: from != null && to != null ? expand(temp, from, to, name) : [],
        first: cal.events[0] ? present(temp, temp.events[0]) : null,
        invitation: invitation ? { summary: invitation.summary, organizer: invitation.organizer, start: invitation.start, end: invitation.end, allDay: invitation.allDay, me: me ? { email: me.email, partstat: answered?.partstat || me.partstat } : answered ? { email: answered.email, partstat: answered.partstat } : null } : null,
      };
    },

    /** Bring a file's events into a calendar; an event already there by UID is replaced when the file's is newer. */
    importFile: ({ path: target = null, text = null, calendarId = null }) => {
      const cal = calendarOf(calendarId) || state.calendars[0];
      const read = readCalendar(text != null ? text : fs.readFileSync(target));
      const now = Date.now();
      let added = 0;
      let updated = 0;
      for (const e of read.events) {
        const existing = e.uid ? cal.events.find((x) => x.uid === e.uid && Boolean(x.recurrenceId) === Boolean(e.recurrenceId) && (!e.recurrenceId || Math.abs((x.recurrenceId?.at ?? 0) - e.recurrenceId.at) < 1000)) : null;
        if (existing) {
          if ((e.sequence || 0) >= (existing.sequence || 0)) {
            Object.assign(existing, e, { id: existing.id, updatedAt: now, remote: existing.remote });
            touch(cal, existing.uid);
            updated += 1;
          }
        } else {
          cal.events.push({ ...e, id: crypto.randomUUID(), createdAt: now, updatedAt: now });
          touch(cal, e.uid);
          added += 1;
        }
      }
      if (added || updated) save();
      return { added, updated, total: read.events.length, calendar: cal.name };
    },

    /** A calendar, or all of them, as a .ics. */
    exportFile: ({ path: target, calendarId = null }) => {
      const chosen = calendarId ? [calendarOf(calendarId)].filter(Boolean) : state.calendars;
      const events = chosen.flatMap((c) => c.events);
      fs.writeFileSync(target, writeCalendar({ events, name: chosen.length === 1 ? chosen[0].name : 'Rutba Office' }), 'utf8');
      return { path: target, count: events.length };
    },

    /**
     * Answer an invitation: keep the event in a calendar with the answer on
     * this person's line, and hand back the reply file and whom to send it
     * to. Sending is mail's; the window opens a message with the file on it.
     */
    respond: ({ path: target = null, text = null, id = null, partstat = 'ACCEPTED', calendarId = null, comment = null }) => {
      const cal = calendarOf(calendarId) || state.calendars[0];
      let event = null;
      if (text != null || target) {
        const read = readCalendar(text != null ? text : fs.readFileSync(target));
        event = read.events[0] || null;
      } else if (id) event = find(id)?.event || null;
      if (!event) throw new Error('There is no invitation to answer.');
      const mine = myAddresses();
      const me = event.attendees.find((a) => mine.includes(String(a.email || '').toLowerCase())) || event.attendees[0] || null;
      const answer = partstat.toUpperCase();
      const attendees = event.attendees.map((a) => (a === me ? { ...a, partstat: answer, rsvp: false } : a));
      const kept = { ...event, attendees, status: answer === 'DECLINED' ? 'CANCELLED' : event.status || 'CONFIRMED' };
      const existing = cal.events.find((x) => x.uid === kept.uid && !x.recurrenceId);
      const now = Date.now();
      if (answer === 'DECLINED') {
        if (existing) { cal.events = cal.events.filter((x) => x !== existing); forget(cal, existing); }
      } else if (existing) { Object.assign(existing, kept, { id: existing.id, updatedAt: now, remote: existing.remote }); touch(cal, existing.uid); }
      else { cal.events.push({ ...kept, id: crypto.randomUUID(), createdAt: now, updatedAt: now }); touch(cal, kept.uid); }
      remember(event, answer, me?.email);
      save();
      const reply = me ? writeReply(kept, { email: me.email, name: me.name, partstat: answer, type: me.type || null, comment }) : null;
      let replyPath = null;
      if (reply) {
        replyPath = path.join(stores.dir, 'outgoing', `reply-${Date.now().toString(36)}.ics`);
        fs.mkdirSync(path.dirname(replyPath), { recursive: true });
        fs.writeFileSync(replyPath, reply, 'utf8');
      }
      const verb = { ACCEPTED: 'Accepted', TENTATIVE: 'Tentative', DECLINED: 'Declined' }[answer] || answer;
      return {
        kept: answer !== 'DECLINED',
        organizer: event.organizer,
        subject: `${verb}: ${event.summary || 'Invitation'}`,
        replyPath,
        // The file's size, for the message it goes on to show.
        replySize: reply ? Buffer.byteLength(reply, 'utf8') : 0,
        me: me ? me.email : null,
        summary: event.summary || '',
      };
    },

    /** The account an invitation goes from: the one Mail sends new mail from; null with none. */
    sendingAccount: () => invites.sendingAccount(),

    /** Send an event's invitation, or the change to one sent before (calendar-invites.js). */
    sendInvitation: (args) => invites.sendInvitation(args),

    /** Tell the people invited that a meeting, or one occurrence of it, is off; the window deletes it after. */
    sendCancel: (args) => invites.sendCancel(args),

    /**
     * Answer an invitation and send the answer: `respond`, then its REPLY
     * mailed to the organizer from the address invited. With no account of
     * that address here, `sent` is false and the window opens it in Mail.
     */
    answer: async ({ words = null, ...args }) => {
      const partstat = String(args.partstat || 'ACCEPTED').toUpperCase();
      const r = service.respond({ ...args, partstat });
      return invites.sendAnswer(r, { partstat, comment: args.comment || null, words, summary: r.summary });
    },

    /** A scheduling message that arrived, taken in: a REPLY, a CANCEL, a change (calendar-invites.js). */
    takeMail: (args) => invites.takeMail(args),

    /** The Inboxes read now for answers, changes and cancellations; a room's Inbox run here is answered for the room. */
    checkMail: () => invites.scan({ also: async (ctx) => (await desk.take(ctx)) || availability.take(ctx), rooms: desk.accounts() }),

    /**
     * Whether people and rooms are free for a time: what is known, else their
     * server's free/busy, else an enquiry by mail whose answer comes later
     * (`asked`); the editor asks again with `ask` false when the calendar changes.
     */
    availability: (args) => availability.check(args),

    /** Mail's accounts, each with whether this copy of Office runs it as a room's mailbox (calendar-desk.js). */
    roomMailboxes: () => desk.list(),

    /** Mark a Mail account as a room's mailbox, or not; `words` are those its answers are written in. */
    setRoomMailbox: ({ accountId, room, words = null }) => {
      if (words && typeof words === 'object') state.words = wordsFrom(words);
      return desk.set({ accountId, room });
    },

    /**
     * The rooms and resources for a time, for the editor's "Add a room": each
     * with whether it is free then, the free first, then those that seat
     * `people`. `uid` is the meeting being edited, left out of the busy times.
     */
    rooms: ({ start, end, uid = null, people = 0 }) =>
      roomsFor({ bookable: bookable(), sources: roomSources.map((s) => ({ busy: (email, from, to) => s.busy(email, from, to, uid) })), start, end, people }),

    /** An invitation for an event with attendees, written to a file a message can carry. */
    invitationFile: ({ id }) => {
      const hit = find(id);
      if (!hit) throw new Error('That event is no longer there.');
      const organizer = hit.event.organizer?.email ? hit.event.organizer : { email: myAddresses()[0] || null, name: null };
      const event = { ...hit.event, organizer, attendees: hit.event.attendees.map((a) => ({ ...a, rsvp: true, partstat: a.partstat || 'NEEDS-ACTION' })) };
      const text = writeCalendar({ events: [event], method: 'REQUEST' });
      const file = path.join(stores.dir, 'outgoing', `invite-${Date.now().toString(36)}.ics`);
      fs.mkdirSync(path.dirname(file), { recursive: true });
      fs.writeFileSync(file, text, 'utf8');
      return { path: file, size: Buffer.byteLength(text, 'utf8'), to: event.attendees.map((a) => a.email).filter(Boolean), subject: `Invitation: ${event.summary || 'Event'}` };
    },

    // The sync's own way in (dav-sync.js): not a window's, so not in the shell's contract.
    _dav: { state: () => state, save, hooks, colours: COLOURS },
  };
  return service;
}
