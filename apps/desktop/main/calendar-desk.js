// A room's mailbox, run by this copy of Office.
//
// A Mail account here can be marked "this mailbox is a room". While Office is
// open on this computer, the calendar reads that Inbox with the others
// (calendar-invites.js, scan) and answers for the room:
//
// - an invitation that clashes with nothing the room is booked for is
//   accepted and kept as the room's booking, and a REPLY goes back to the
//   organizer from the room's own address;
// - one that clashes is declined, the REPLY and the message saying what is
//   free that day;
// - a change from the organizer is checked again the same way, and a
//   cancellation takes the booking off;
// - an enquiry is answered: a VFREEBUSY REQUEST with the room's busy times
//   (never what fills them), and a plain message asking whether it is "free
//   on" a date with the free times that day.
//
// It answers only what arrives after it was marked, and only while this copy
// of Office is open: closed, the mail waits in the room's Inbox and is
// answered when it next opens. The room's bookings are kept here, apart from
// the person's own calendars; they are what the room's time is, for the
// answers and for "Add a room" (a source with the room's whole day).

import { readScheduling, writeReply, writeFreeBusyReply, freeBetween, clashes, occurrences, exceptionsOf, sameAddress } from '@rutba/calendar';
import { busyInCalendars } from './calendar-rooms.js';
import { wordsFrom, fill, whenText, subjectOf } from './calendar-text.js';

const DAY = 86400000;
/** How far ahead a repeating booking is checked for clashes. */
const AHEAD = 366 * DAY;
/** The hours a room's free times are given in, on this computer's clock. */
const HOURS = [8, 18];
/** The longest span a free/busy enquiry is answered for. */
const LONGEST = 62 * DAY;
/** Mail that is never answered: the room's own, and that of systems that do not read answers. */
const MACHINE = /^(?:no-?reply|do-?not-?reply|mailer-daemon|postmaster|bounces?)[@+.-]/i;

const lower = (s) => String(s || '').trim().toLowerCase();
const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const WEEKDAYS = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];

/**
 * The day an enquiry asks about, as local midnight, or null: "free on 12
 * March?", "is it available on Tuesday", "free 2026-11-17?", "free tomorrow".
 * A date with no year is the next one from today. It reads English words, as
 * the subject or the first lines of the message carry them.
 */
export function askedDay(text, now = Date.now()) {
  const s = String(text || '').toLowerCase();
  if (!/\b(free|available|availability|booked)\b/.test(s)) return null;
  const today = new Date(now);
  today.setHours(0, 0, 0, 0);
  const day = (y, m, d) => {
    const at = new Date(y, m, d);
    return at.getMonth() === m && at.getDate() === d ? at.getTime() : null;
  };
  const next = (m, d) => {
    const y = today.getFullYear();
    const at = day(y, m, d);
    return at !== null && at < today.getTime() ? day(y + 1, m, d) : at;
  };
  let m;
  if ((m = /\b(\d{4})-(\d{1,2})-(\d{1,2})\b/.exec(s))) return day(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const month = `(${MONTHS.join('|')})[a-z]*\\.?`;
  if ((m = new RegExp(`\\b(\\d{1,2})(?:st|nd|rd|th)?\\s+(?:of\\s+)?${month}(?:,?\\s+(\\d{4}))?`).exec(s))) {
    const mo = MONTHS.indexOf(m[2]);
    return m[3] ? day(Number(m[3]), mo, Number(m[1])) : next(mo, Number(m[1]));
  }
  if ((m = new RegExp(`\\b${month}\\s+(\\d{1,2})(?:st|nd|rd|th)?(?:,?\\s+(\\d{4}))?`).exec(s))) {
    const mo = MONTHS.indexOf(m[1]);
    return m[3] ? day(Number(m[3]), mo, Number(m[2])) : next(mo, Number(m[2]));
  }
  // Day first, as written in Britain and most of the world.
  if ((m = /\b(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b/.exec(s))) {
    const y = m[3] ? (m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3])) : null;
    return y ? day(y, Number(m[2]) - 1, Number(m[1])) : next(Number(m[2]) - 1, Number(m[1]));
  }
  if (/\btoday\b/.test(s)) return today.getTime();
  if (/\btomorrow\b/.test(s)) return today.getTime() + DAY;
  if ((m = new RegExp(`\\b(${WEEKDAYS.join('|')})[a-z]*\\b`).exec(s))) {
    const want = WEEKDAYS.indexOf(m[1]);
    const ahead = (want - today.getDay() + 7) % 7;
    const at = new Date(today);
    at.setDate(at.getDate() + ahead);
    return at.getTime();
  }
  return null;
}

/**
 * @param {object} o
 * @param {() => object} o.state the calendar's state (its `roomDesk`)
 * @param {() => void} o.save
 * @param {object} o.link the Mail link (calendar-mail.js)
 * @param {object} o.invites the invitations (calendar-invites.js), for the files Mail attaches
 * @param {() => number} [o.now]
 */
export function createRoomDesk({ state, save, link, invites, now = () => Date.now() }) {
  const desks = () => (state().roomDesk ||= {});
  const words = () => wordsFrom(state().words);
  const deskOf = (accountId) => (desks()[accountId]?.on ? desks()[accountId] : null);
  const deskFor = (email) => {
    const a = link.accounts().find((x) => sameAddress(x.email, email));
    return a ? deskOf(a.id) : null;
  };
  const roomName = (account) => account.name || account.email;

  /** The room's busy periods in [from, to), from its bookings here, leaving out one meeting (`except`, its UID). */
  const busyOf = (desk, email, from, to, except = null) => busyInCalendars([{ events: desk.bookings || [] }], email, from, to, except);

  /** The free times of a day, as the answer says them. */
  const freeText = (busy, dayStart, w, least = 30 * 60000) => {
    const from = dayStart + HOURS[0] * 3600000;
    const to = dayStart + HOURS[1] * 3600000;
    const time = new Intl.DateTimeFormat(w.language || 'en', { hour: 'numeric', minute: '2-digit' });
    return freeBetween(busy, Math.max(from, now()), to, { least }).map((p) => `${time.format(p.start)}–${time.format(p.end)}`);
  };
  const dayText = (at, w) => new Intl.DateTimeFormat(w.language || 'en', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }).format(at);
  const startOfDay = (at) => {
    const d = new Date(at);
    d.setHours(0, 0, 0, 0);
    return d.getTime();
  };

  const reply = async (account, { to, subject, text, ics = null }) => {
    const attachments = ics ? [invites.attach('REPLY', ics)] : [];
    const sent = await link.sendFrom({ account: account.id, to: [to], subject, text, attachments, headers: { 'Auto-Submitted': 'auto-replied' } });
    return { refused: sent?.refused || [] };
  };

  /** Which of the incoming meeting's times clash with the room's bookings: the first, or null. */
  const firstClash = (desk, email, events) => {
    const uid = events[0]?.uid || null;
    const from = now();
    const to = from + AHEAD;
    const siblingsOf = exceptionsOf(events);
    for (const e of events) {
      if (e.status === 'CANCELLED') continue;
      const times = e.recurrenceId ? [{ start: e.start.at, end: e.end?.at ?? e.start.at }]
        : occurrences(e, { from, to, siblings: siblingsOf(e), limit: 500 }).filter((o) => o.event?.status !== 'CANCELLED').map((o) => ({ start: o.start, end: o.end }));
      for (const t of times) {
        if (t.end <= from) continue;
        const busy = busyOf(desk, email, startOfDay(t.start), startOfDay(t.end) + DAY, uid);
        if (clashes(busy, t.start, t.end)) return { ...t, busy };
      }
    }
    return null;
  };

  /** An invitation, or a change to one, for the room. */
  const takeRequest = async (account, desk, read, w) => {
    const events = read.events.filter((e) => e.uid);
    const master = events.find((e) => !e.recurrenceId) || events[0];
    const organizer = master?.organizer;
    if (!master || !organizer?.email || sameAddress(organizer.email, account.email)) return null;
    const kept = (desk.bookings || []).filter((e) => e.uid === master.uid);
    const keptMaster = kept.find((e) => !e.recurrenceId) || kept[0];
    if (keptMaster && (master.sequence || 0) < (keptMaster.sequence || 0)) return null;
    const line = (master.attendees || []).find((a) => sameAddress(a.email, account.email));
    const type = line?.type && /^(ROOM|RESOURCE)$/i.test(line.type) ? line.type.toUpperCase() : 'ROOM';
    const room = { email: account.email, name: roomName(account) };
    const clash = firstClash(desk, account.email, events);
    const when = whenText(master, w);
    if (clash) {
      const times = freeText(clash.busy, startOfDay(clash.start), w, Math.max(15 * 60000, clash.end - clash.start));
      const free = times.length ? fill(w.roomFreeThatDay, { times: times.join(', ') }) : w.roomNoneFree;
      const comment = `${fill(w.roomClash, { room: room.name, when: whenText({ ...master, start: { ...master.start, at: clash.start }, end: { ...(master.end || master.start), at: clash.end } }, w) })} ${free}`;
      const ics = writeReply(master, { ...room, partstat: 'DECLINED', type, comment });
      // A change that now clashes gives the time up: the booking kept for the old time goes.
      if (kept.length) desk.bookings = desk.bookings.filter((e) => e.uid !== master.uid);
      save();
      const sent = await reply(account, { to: organizer.email, subject: subjectOf('DECLINED', master, w), text: [comment, '', fill(w.roomAsk, { room: room.name }), '', w.roomOffice].join('\n'), ics });
      return { kind: 'room-declined', room: room.email, uid: master.uid, summary: master.summary || '', clash: { start: clash.start, end: clash.end }, free: times, ...sent };
    }
    const mark = (e) => ({ ...e, attendees: (e.attendees || []).some((a) => sameAddress(a.email, room.email))
      ? e.attendees.map((a) => (sameAddress(a.email, room.email) ? { ...a, partstat: 'ACCEPTED', type } : a))
      : [...(e.attendees || []), { email: room.email, name: room.name, partstat: 'ACCEPTED', type, role: 'NON-PARTICIPANT' }] });
    // The whole series replaces what was kept; one occurrence replaces only itself.
    const keep = master.recurrenceId
      ? (desk.bookings || []).filter((e) => !(e.uid === master.uid && events.some((x) => x.recurrenceId && e.recurrenceId && Math.abs(x.recurrenceId.at - e.recurrenceId.at) < 1000)))
      : (desk.bookings || []).filter((e) => e.uid !== master.uid);
    desk.bookings = [...keep, ...events.map(mark)];
    save();
    const text = fill(w.roomBooked, { room: room.name, when });
    const ics = writeReply(master, { ...room, partstat: 'ACCEPTED', type });
    const sent = await reply(account, { to: organizer.email, subject: subjectOf('ACCEPTED', master, w), text: [text, '', w.roomOffice].join('\n'), ics });
    return { kind: 'room-booked', room: room.email, uid: master.uid, summary: master.summary || '', ...sent };
  };

  /** A cancellation from the organizer: the booking, or one occurrence of it, off. */
  const takeCancel = (account, desk, read, from) => {
    let changed = 0;
    for (const e of read.events.filter((x) => x.uid)) {
      const kept = (desk.bookings || []).filter((b) => b.uid === e.uid);
      if (!kept.length) continue;
      const organizer = (kept.find((b) => !b.recurrenceId) || kept[0]).organizer;
      if (from && organizer?.email && !sameAddress(from, organizer.email) && !sameAddress(e.organizer?.email, from)) continue;
      if (!e.recurrenceId) {
        desk.bookings = desk.bookings.filter((b) => b.uid !== e.uid);
      } else {
        desk.bookings = desk.bookings.filter((b) => !(b.uid === e.uid && b.recurrenceId && Math.abs(b.recurrenceId.at - e.recurrenceId.at) < 1000));
        desk.bookings.push({ ...e, status: 'CANCELLED' });
      }
      changed += 1;
    }
    if (changed) save();
    return changed ? { kind: 'room-cancelled', room: account.email, uid: read.events[0]?.uid || null } : null;
  };

  /** A VFREEBUSY REQUEST: the room's busy times in the span asked about. */
  const takeFreeBusy = async (account, desk, ask, w) => {
    const asker = ask.organizer?.email;
    if (!asker || !ask.start || !ask.end) return null;
    const start = ask.start;
    const end = Math.min(ask.end, start + LONGEST);
    const busy = busyOf(desk, account.email, start, end);
    const ics = writeFreeBusyReply({ uid: ask.uid || `fb-${start}`, organizer: ask.organizer, attendee: { email: account.email, name: roomName(account) }, start, end, busy });
    const sent = await reply(account, { to: asker, subject: fill(w.roomFreeSubject, { room: roomName(account) }), text: w.roomOffice, ics });
    return { kind: 'room-freebusy', room: account.email, busy: busy.length, ...sent };
  };

  /** A plain message asking whether the room is free on a day. */
  const takeEnquiry = async (account, desk, folder, row, w) => {
    const message = link.message({ accountId: account.id, folder, id: row.id });
    const subject = message?.subject ?? row.subject ?? '';
    const body = String(message?.text || row.preview || '').split(/\r?\n/).filter((l) => !/^\s*>/.test(l)).slice(0, 12).join('\n');
    if (/^\s*(re|aw|sv|antw?)\s*:/i.test(subject) && !/\b(free|available)\b/i.test(body)) return null;
    const asked = askedDay(subject, now()) ?? askedDay(body, now());
    if (asked === null) return null;
    const busy = busyOf(desk, account.email, asked, asked + DAY);
    const times = freeText(busy, asked, w);
    const date = dayText(asked, w);
    const room = roomName(account);
    const line = times.length ? fill(w.roomFreeOn, { room, date, times: times.join(', ') }) : fill(w.roomBusyAllDay, { room, date });
    const sent = await reply(account, { to: row.from.address, subject: fill(w.roomFreeSubject, { room }), text: [line, '', fill(w.roomAsk, { room }), '', w.roomOffice].join('\n') });
    return { kind: 'room-enquiry', room: account.email, day: asked, free: times, ...sent };
  };

  return {
    /** Mail's accounts with an address, each with whether it is run here as a room and how many bookings it has. */
    list: () => link.accounts().filter((a) => /@/.test(a.email || '')).map((a) => {
      const d = deskOf(a.id);
      return { id: a.id, email: a.email, name: a.name || null, room: Boolean(d), since: d?.since || null, bookings: (d?.bookings || []).filter((b) => !b.recurrenceId && b.status !== 'CANCELLED').length };
    }),

    /** Mark an account as a room's mailbox, or not. Its bookings are kept, for when it is marked again. */
    set: ({ accountId, room }) => {
      const a = link.accounts().find((x) => x.id === accountId);
      if (!a) throw new Error('There is no such Mail account.');
      const d = (desks()[accountId] ||= { bookings: [] });
      d.on = Boolean(room);
      d.email = a.email;
      if (room) d.since = now();
      save();
      return { id: a.id, email: a.email, room: d.on, since: d.since };
    },

    /** Whether an address is a room's mailbox run here. */
    isRoom: (email) => Boolean(deskFor(email)),

    /** The account ids run as rooms: their Inbox is the room's, not the person's. */
    accounts: () => Object.keys(desks()).filter((id) => desks()[id].on),

    /** What the room's time is, as a source for "Add a room" (calendar-rooms.js): its whole day. */
    source: {
      busy: (email, from, to, except) => {
        const d = deskFor(email);
        return d ? { busy: busyOf(d, email, from, to, except), whole: true, source: 'room-mailbox' } : null;
      },
    },

    /** A message that arrived in an Inbox, answered when that Inbox is a room's (the scan's `also`). */
    take: async ({ account, folder, row, part }) => {
      const desk = deskOf(account.id);
      if (!desk) return null;
      const from = row.from?.address || null;
      if (!from || sameAddress(from, account.email) || MACHINE.test(from) || link.accounts().some((a) => sameAddress(a.email, from) && deskOf(a.id))) return null;
      const arrived = row.date ? Date.parse(row.date) : NaN;
      if (desk.since && Number.isFinite(arrived) && arrived < desk.since - 60000) return null;
      const w = words();
      if (part?.text) {
        let read;
        try {
          read = readScheduling(part.text);
        } catch {
          return null;
        }
        const method = lower(read.method).toUpperCase();
        if (method === 'REQUEST' && read.freebusy?.length) return takeFreeBusy(account, desk, read.freebusy[0], w);
        if (method === 'REQUEST' && read.events?.length) return takeRequest(account, desk, read, w);
        if (method === 'CANCEL') return takeCancel(account, desk, read, from);
        return null;
      }
      return takeEnquiry(account, desk, folder, row, w);
    },
  };
}
