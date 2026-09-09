// The calendar.
//
// One JSON file in the profile, calendar.json: calendars, each with a name,
// a colour and its events as the iCalendar reader gives them. Nothing here
// reaches a server. A window asks for the occurrences in the range it shows
// and gets every calendar's, expanded, with the recurring ones' exceptions
// applied; a file opened from disk is shown without being kept until the
// person says so; an invitation gets its answer written as the reply the
// organizer's calendar reads.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { readCalendar, writeCalendar, writeReply, isInvitation, occurrences, dateTimeAt, describeRule, localZone } from '@rutba/calendar';

const readJson = (file, fallback) => {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return fallback;
  }
};
const writeJson = (file, value) => {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(value));
  fs.renameSync(tmp, file);
};

const COLOURS = ['#2b5fd9', '#1a9f7a', '#d9534f', '#e08b2b', '#7b5cd6', '#c2408f', '#3aa0b5', '#8c6d1f'];

export function createCalendarService({ stores, broadcast, mail = null }) {
  const file = path.join(stores.dir, 'calendar.json');
  const state = readJson(file, null) || { calendars: [] };
  if (!Array.isArray(state.calendars)) state.calendars = [];
  if (!state.calendars.length) state.calendars.push({ id: crypto.randomUUID(), name: 'My calendar', colour: COLOURS[0], visible: true, events: [] });

  const save = () => {
    writeJson(file, state);
    broadcast?.('calendar:changed', {});
  };
  const calendarOf = (id) => state.calendars.find((c) => c.id === id) || null;
  const find = (eventId) => {
    for (const cal of state.calendars) {
      const event = cal.events.find((e) => e.id === eventId);
      if (event) return { cal, event };
    }
    return null;
  };

  /** The addresses this person is known by: the mail accounts' own. */
  const myAddresses = () => {
    try {
      return (mail?.accounts?.() || []).map((a) => String(a.email || '').toLowerCase()).filter(Boolean);
    } catch {
      return [];
    }
  };

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
    for (const e of events) {
      if (e.recurrenceId) continue; // drawn through its parent
      const siblings = events.filter((s) => s.uid && s.uid === e.uid && s !== e);
      for (const o of occurrences(e, { from, to, siblings, limit: 500 })) out.push(slot(cal, o, fileName));
    }
    return out;
  }

  /** A stored event, whole, with what the editor wants to show beside it. */
  function present(cal, e) {
    return { ...e, calendarId: cal.id, calendar: cal.name, colour: cal.colour, rule: e.rrule ? describeRule(e.rrule, e.start) : 'Does not repeat' };
  }

  return {
    calendars: () => state.calendars.map((c) => ({ id: c.id, name: c.name, colour: c.colour, visible: c.visible !== false, count: c.events.length })),

    saveCalendar: ({ calendar }) => {
      const kept = calendar.id ? calendarOf(calendar.id) : null;
      if (kept) Object.assign(kept, { name: calendar.name ?? kept.name, colour: calendar.colour ?? kept.colour, visible: calendar.visible ?? kept.visible });
      else state.calendars.push({ id: crypto.randomUUID(), name: calendar.name || 'Calendar', colour: calendar.colour || COLOURS[state.calendars.length % COLOURS.length], visible: true, events: [] });
      save();
      return state.calendars.map((c) => ({ id: c.id, name: c.name, colour: c.colour, visible: c.visible !== false, count: c.events.length }));
    },

    removeCalendar: ({ id }) => {
      if (state.calendars.length <= 1) throw new Error('The last calendar stays; delete its events instead.');
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
        calendarId: undefined, calendar: undefined, colour: undefined, rule: undefined,
      };
      if (event.id && scope === 'this' && original != null) {
        // An exception: a copy with a RECURRENCE-ID naming the instance it replaces.
        const parent = find(event.id);
        if (parent) {
          const exception = { ...clean, id: crypto.randomUUID(), rrule: null, recurrenceId: dateTimeAt(original, parent.event.start.tzid || zone, parent.event.allDay), updatedAt: now };
          cal.events = cal.events.filter((e) => !(e.uid === exception.uid && e.recurrenceId && Math.abs(e.recurrenceId.at - original) < 1000));
          cal.events.push(exception);
          save();
          return present(cal, exception);
        }
      }
      const hit = event.id ? find(event.id) : null;
      if (hit) {
        // Moved to another calendar, or changed in place.
        if (hit.cal !== cal) hit.cal.events = hit.cal.events.filter((e) => e !== hit.event);
        const kept = { ...hit.event, ...clean, id: hit.event.id, sequence: (hit.event.sequence || 0) + 1, updatedAt: now };
        const at = cal.events.indexOf(hit.event);
        if (at >= 0) cal.events[at] = kept;
        else cal.events.push(kept);
        save();
        return present(cal, kept);
      }
      const fresh = { ...clean, id: crypto.randomUUID(), createdAt: now, updatedAt: now };
      cal.events.push(fresh);
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
        save();
        return { removed: 1, scope: 'this' };
      }
      hit.cal.events = hit.cal.events.filter((e) => e !== hit.event && !(e.uid && e.uid === hit.event.uid && e.recurrenceId));
      save();
      return { removed: 1, scope: 'all' };
    },

    /**
     * What a file holds, without keeping it: its events expanded in a range,
     * and whether it is an invitation to answer.
     */
    openFile: ({ path: target, from, to }) => {
      const cal = readCalendar(fs.readFileSync(target));
      const name = cal.name || path.basename(target);
      const temp = { id: `file:${target}`, name, colour: '#8c6d1f', events: cal.events.map((e, i) => ({ ...e, id: `file:${i}` })) };
      const mine = myAddresses();
      const invitation = isInvitation(cal) ? cal.events[0] : null;
      const me = invitation ? invitation.attendees.find((a) => mine.includes(String(a.email || '').toLowerCase())) || null : null;
      return {
        path: target,
        name,
        method: cal.method,
        count: cal.events.length,
        events: from != null && to != null ? expand(temp, from, to, name) : [],
        first: cal.events[0] ? present(temp, temp.events[0]) : null,
        invitation: invitation ? { summary: invitation.summary, organizer: invitation.organizer, start: invitation.start, end: invitation.end, allDay: invitation.allDay, me: me ? { email: me.email, partstat: me.partstat } : null } : null,
      };
    },

    /** Bring a file's events into a calendar; an event already there by UID is replaced when the file's is newer. */
    importFile: ({ path: target, calendarId = null }) => {
      const cal = calendarOf(calendarId) || state.calendars[0];
      const read = readCalendar(fs.readFileSync(target));
      const now = Date.now();
      let added = 0;
      let updated = 0;
      for (const e of read.events) {
        const existing = e.uid ? cal.events.find((x) => x.uid === e.uid && Boolean(x.recurrenceId) === Boolean(e.recurrenceId) && (!e.recurrenceId || Math.abs((x.recurrenceId?.at ?? 0) - e.recurrenceId.at) < 1000)) : null;
        if (existing) {
          if ((e.sequence || 0) >= (existing.sequence || 0)) {
            Object.assign(existing, e, { id: existing.id, updatedAt: now });
            updated += 1;
          }
        } else {
          cal.events.push({ ...e, id: crypto.randomUUID(), createdAt: now, updatedAt: now });
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
    respond: ({ path: target = null, id = null, partstat = 'ACCEPTED', calendarId = null }) => {
      const cal = calendarOf(calendarId) || state.calendars[0];
      let event = null;
      if (target) {
        const read = readCalendar(fs.readFileSync(target));
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
        if (existing) cal.events = cal.events.filter((x) => x !== existing);
      } else if (existing) Object.assign(existing, kept, { id: existing.id, updatedAt: now });
      else cal.events.push({ ...kept, id: crypto.randomUUID(), createdAt: now, updatedAt: now });
      save();
      const reply = me ? writeReply(kept, { email: me.email, name: me.name, partstat: answer }) : null;
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
        me: me ? me.email : null,
      };
    },

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
      return { path: file, to: event.attendees.map((a) => a.email).filter(Boolean), subject: `Invitation: ${event.summary || 'Event'}` };
    },
  };
}
