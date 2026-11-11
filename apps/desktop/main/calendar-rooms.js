// Rooms and resources, as the calendar books them.
//
// A room is a card in Contacts marked as a room or a resource, with its own
// mailbox (contacts.bookable). It is booked by inviting that mailbox, as an
// attendee with CUTYPE=ROOM or RESOURCE. The event editor's "Add a room"
// lists them for the time chosen: those free then first, then those that seat
// the people invited (the smallest that does first), then the rest.
//
// What is known of a room's time comes from `sources`, each answering
// `{ busy, whole, source }` for an address in a span (`whole` when it knows
// the room's whole day, not only some of it), or null when it knows nothing:
// the meetings in this person's own calendars that have the room on
// them; a room mailbox this copy of Office runs (calendar-desk.js); and the
// free/busy a server or an answer by mail gave (calendar-availability.js).

import { occurrences, exceptionsOf, mergeBusy, clashes, sameAddress } from '@rutba/calendar';

/**
 * The busy periods the calendars here know for an address: every occurrence
 * in [from, to) of a meeting with it on, not declined by it and not cancelled,
 * leaving out the meeting being edited (`except`, its UID).
 */
export function busyInCalendars(calendars, email, from, to, except = null) {
  const busy = [];
  for (const cal of calendars) {
    const events = cal.events || [];
    const siblingsOf = exceptionsOf(events);
    for (const e of events) {
      if (e.recurrenceId || (except && e.uid === except) || e.status === 'CANCELLED' || e.transparency === 'TRANSPARENT') continue;
      const line = (e.attendees || []).find((a) => sameAddress(a.email, email));
      if (!line || line.partstat === 'DECLINED') continue;
      for (const o of occurrences(e, { from, to, siblings: siblingsOf(e), limit: 500 })) {
        if (o.event?.status === 'CANCELLED') continue;
        busy.push({ start: o.start, end: o.end });
      }
    }
  }
  return mergeBusy(busy);
}

/**
 * The rooms and resources for [start, end), each with whether it is free
 * then (true; false; null when no clash is known here but nothing that has
 * the room's whole day has answered), the busy periods known
 * that day, and where that came from; in the order the editor lists them.
 */
export function roomsFor({ bookable, sources, start, end, people = 0 }) {
  const dayStart = new Date(start);
  dayStart.setHours(0, 0, 0, 0);
  const from = dayStart.getTime();
  const to = from + 86400000;
  const rows = bookable.map((room) => {
    let busy = [];
    // Whether a source that has the room's whole day answered (its own mailbox, a server, its reply),
    // not only the meetings here that have it on.
    let whole = false;
    const named = [];
    for (const source of sources) {
      const got = source.busy(room.email, Math.min(from, start), Math.max(to, end));
      if (!got) continue;
      busy = busy.concat(got.busy || []);
      if (got.whole) whole = true;
      if (got.source) named.push(got.source);
    }
    busy = mergeBusy(busy);
    const free = clashes(busy, start, end) ? false : whole ? true : null;
    return { ...room, free, busy, sources: [...new Set(named)] };
  });
  const rank = (r) => (r.free === true ? 0 : r.free === null ? 1 : 2);
  const fits = (r) => !people || !r.capacity || r.capacity >= people;
  return rows.sort((a, b) => rank(a) - rank(b)
    || Number(!fits(a)) - Number(!fits(b))
    || (fits(a) ? (a.capacity || Infinity) - (b.capacity || Infinity) : (b.capacity || 0) - (a.capacity || 0))
    || (a.kind === b.kind ? 0 : a.kind === 'room' ? -1 : 1)
    || String(a.name).localeCompare(String(b.name)));
}
