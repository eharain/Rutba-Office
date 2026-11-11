// Scheduling by email: iTIP (RFC 5546), as mail carries it (iMIP, RFC 6047).
//
// The organizer sends a REQUEST with everyone invited on it; an edit sends it
// again with SEQUENCE one higher; a cancellation is a CANCEL; each person
// invited answers with a REPLY that has only their own line on it, and the
// organizer's copy takes the answer onto that line. A REQUEST names its time
// zones, so each one it uses goes with it as a VTIMEZONE: a TZID with no
// VTIMEZONE is read by some calendars as UTC and by others as their own clock.

import { writeCalendar, readCalendar, parseComponents, parseDateTime, formatRRule, writeLine } from './ical.js';
import { ianaZone, offsetMinutes } from './tz.js';

const lower = (s) => String(s || '').trim().replace(/^mailto:/i, '').toLowerCase();

/** Whether two addresses are the same mailbox, mailto: and case aside. */
export const sameAddress = (a, b) => Boolean(lower(a)) && lower(a) === lower(b);

/**
 * Whether a change asks the people invited to answer again (RFC 5546 2.1.4):
 * a new time, a new place, or a new rule for when it repeats. A new title or
 * a new note is sent without asking again.
 */
export function significantChange(before, after) {
  if (!before || !after) return true;
  if ((before.start?.at ?? null) !== (after.start?.at ?? null)) return true;
  if ((before.end?.at ?? null) !== (after.end?.at ?? null)) return true;
  if (Boolean(before.allDay) !== Boolean(after.allDay)) return true;
  if ((formatRRule(before.rrule) || '') !== (formatRRule(after.rrule) || '')) return true;
  return String(before.location || '').trim() !== String(after.location || '').trim();
}

/* ── time zones ────────────────────────────────────────────────────────── */

const pad = (n, w = 2) => String(n).padStart(w, '0');
const offsetText = (minutes) => `${minutes < 0 ? '-' : '+'}${pad(Math.floor(Math.abs(minutes) / 60))}${pad(Math.abs(minutes) % 60)}`;
const wallStamp = (ms) => {
  const d = new Date(ms);
  return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}`;
};
const HOUR = 3600000;
const DAY = 86400000;

/** The instants in [from, to) where the zone's offset changes, each to the minute. */
function transitions(zone, from, to) {
  const out = [];
  let before = offsetMinutes(zone, from);
  // Six hours at a time finds every change the world's zones make (none is two in a day).
  for (let at = from + 6 * HOUR; at <= to; at += 6 * HOUR) {
    const now = offsetMinutes(zone, at);
    if (now === before) continue;
    let lo = at - 6 * HOUR;
    let hi = at;
    while (hi - lo > 60000) {
      const mid = lo + Math.floor((hi - lo) / 120000) * 60000;
      if (offsetMinutes(zone, mid) === before) lo = mid;
      else hi = mid;
    }
    out.push({ at: hi, from: before, to: now });
    before = now;
  }
  return out;
}

/**
 * A VTIMEZONE for a TZID, as lines, covering the years from `from` to `to`:
 * each change of offset in them written out as its own observance, the one
 * in force at the start first. Null for a zone this computer does not know
 * or for UTC, which needs none.
 */
export function vtimezone(tzid, from, to) {
  const zone = ianaZone(tzid);
  if (!zone || /^(Etc\/)?(UTC|GMT|Zulu)$/i.test(zone)) return null;
  const startYear = new Date(from).getUTCFullYear();
  const endYear = new Date(Math.max(from, to)).getUTCFullYear();
  const first = Date.UTC(startYear, 0, 1);
  const last = Date.UTC(Math.min(endYear, startYear + 10) + 1, 0, 1);
  const changes = transitions(zone, first, last);
  const opening = offsetMinutes(zone, first);
  // Daylight is the larger offset of the pair a change moves between.
  const standard = changes.length ? Math.min(opening, ...changes.map((c) => c.to)) : opening;
  const kind = (offset) => (offset > standard ? 'DAYLIGHT' : 'STANDARD');
  const lines = ['BEGIN:VTIMEZONE', writeLine('TZID', {}, tzid)];
  const observance = (type, dtstart, offFrom, offTo) => {
    lines.push(`BEGIN:${type}`, `DTSTART:${dtstart}`, `TZOFFSETFROM:${offsetText(offFrom)}`, `TZOFFSETTO:${offsetText(offTo)}`, `END:${type}`);
  };
  observance(kind(opening), '19700101T000000', opening, opening);
  for (const c of changes) observance(kind(c.to), wallStamp(c.at + c.from * 60000), c.from, c.to);
  lines.push('END:VTIMEZONE');
  return lines;
}

/** The VTIMEZONEs the events name, over the span they run. */
function zonesFor(events) {
  const zones = new Map();
  for (const e of events) {
    for (const dt of [e.start, e.end, e.recurrenceId]) {
      if (!dt?.tzid || dt.utc || dt.allDay) continue;
      const span = zones.get(dt.tzid) || { from: Infinity, to: -Infinity };
      const until = e.rrule ? (e.rrule.until?.at ?? (e.start?.at || 0) + 2 * 365 * DAY) : dt.at;
      span.from = Math.min(span.from, dt.at - DAY);
      span.to = Math.max(span.to, until + DAY);
      zones.set(dt.tzid, span);
    }
  }
  return [...zones].map(([tzid, s]) => vtimezone(tzid, s.from, s.to)).filter(Boolean);
}

/* ── what the organizer sends ──────────────────────────────────────────── */

const asList = (events) => (Array.isArray(events) ? events : [events]).filter(Boolean);

/**
 * The invitation, or the change to it: the event (with the exceptions of its
 * series, the same UID) as METHOD:REQUEST, the organizer named, everyone
 * invited asked to answer, a line not answered yet saying NEEDS-ACTION.
 */
export function writeRequest(events, { organizer } = {}) {
  const list = asList(events).map((e) => ({
    ...e,
    organizer: organizer || e.organizer,
    status: e.status === 'CANCELLED' ? 'CONFIRMED' : e.status || 'CONFIRMED',
    attendees: (e.attendees || []).filter((a) => a.email && !sameAddress(a.email, (organizer || e.organizer)?.email)).map((a) => ({ ...a, partstat: a.partstat || 'NEEDS-ACTION', rsvp: a.partstat ? a.partstat === 'NEEDS-ACTION' : true })),
  }));
  return writeCalendar({ events: list, method: 'REQUEST', components: zonesFor(list) });
}

/**
 * A cancellation: METHOD:CANCEL with STATUS:CANCELLED, to everyone on it, or
 * to `attendees` alone (the people taken off the list of a meeting that goes
 * on). The SEQUENCE given is the one it goes with: the caller raises it.
 */
export function writeCancel(events, { organizer, attendees = null } = {}) {
  const list = asList(events).map((e) => ({
    ...e,
    organizer: organizer || e.organizer,
    status: 'CANCELLED',
    alarms: [],
    attendees: (attendees || e.attendees || []).filter((a) => a.email).map((a) => ({ email: a.email, name: a.name || null, type: a.type || null, role: a.role || null })),
  }));
  return writeCalendar({ events: list, method: 'CANCEL', components: zonesFor(list) });
}

/* ── what comes back ───────────────────────────────────────────────────── */

/** A scheduling message's calendar: its method, its events, and any free/busy it carries. */
export function readScheduling(text) {
  const cal = readCalendar(text);
  const roots = parseComponents(text);
  const vcal = roots.find((c) => c.name === 'VCALENDAR') || { children: roots };
  const freebusy = vcal.children.filter((c) => c.name === 'VFREEBUSY').map(readFreeBusyNode);
  return { ...cal, freebusy };
}

/**
 * The organizer's event with a REPLY taken in: each line the reply carries
 * put on the matching attendee, an answer older than the one already taken
 * (by its DTSTAMP) left. A person the organizer did not invite who answers
 * anyway is added, as calendars do. Answers { attendees, changed, answers }.
 */
export function applyReply(event, reply) {
  const attendees = (event.attendees || []).map((a) => ({ ...a }));
  const answers = [];
  const stamp = reply.dtstamp?.at ?? Date.now();
  for (const line of reply.attendees || []) {
    if (!line.email || !line.partstat) continue;
    const partstat = line.partstat.toUpperCase();
    const kept = attendees.find((a) => sameAddress(a.email, line.email));
    if (kept) {
      if (kept.answeredAt && kept.answeredAt > stamp) continue;
      if (kept.partstat === partstat && kept.answeredAt) continue;
      Object.assign(kept, { partstat, rsvp: false, answeredAt: stamp, name: kept.name || line.name || null });
    } else attendees.push({ email: line.email, name: line.name || null, role: 'OPT-PARTICIPANT', partstat, rsvp: false, type: line.type || null, answeredAt: stamp, uninvited: true });
    answers.push({ email: line.email, name: line.name || kept?.name || null, partstat });
  }
  return { attendees, changed: answers.length > 0, answers };
}

/* ── free/busy (RFC 5545 3.6.4, RFC 5546 3.3) ──────────────────────────── */

function readFreeBusyNode(node) {
  const prop = (name) => node.props.find((p) => p.name === name) || null;
  const when = (name) => {
    const p = prop(name);
    return p ? parseDateTime(p.value, p.params) : null;
  };
  const busy = [];
  for (const p of node.props.filter((x) => x.name === 'FREEBUSY')) {
    const type = ((p.params.FBTYPE || [])[0] || 'BUSY').toUpperCase();
    if (type === 'FREE') continue;
    for (const period of p.value.split(',')) {
      const [a, b] = period.split('/');
      const start = parseDateTime(a, {});
      if (!start) continue;
      let end = parseDateTime(b, {});
      if (!end && /^[+-]?P/.test(b || '')) {
        const m = /^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(b);
        const ms = m ? ((Number(m[2] || 0) * 7 + Number(m[3] || 0)) * 86400 + Number(m[4] || 0) * 3600 + Number(m[5] || 0) * 60 + Number(m[6] || 0)) * 1000 : 0;
        end = { at: start.at + ms };
      }
      if (end && end.at > start.at) busy.push({ start: start.at, end: end.at, type });
    }
  }
  const organizer = prop('ORGANIZER');
  return {
    uid: prop('UID') ? prop('UID').value : null,
    start: when('DTSTART')?.at ?? null,
    end: when('DTEND')?.at ?? null,
    organizer: organizer ? { email: lower(organizer.value), name: (organizer.params.CN || [])[0] || null } : null,
    attendees: node.props.filter((x) => x.name === 'ATTENDEE').map((a) => ({ email: lower(a.value), name: (a.params.CN || [])[0] || null })),
    busy: mergeBusy(busy),
  };
}

/** Busy periods, overlapping ones joined, in order. */
export function mergeBusy(periods) {
  const sorted = [...periods].filter((p) => p.end > p.start).sort((a, b) => a.start - b.start);
  const out = [];
  for (const p of sorted) {
    const last = out[out.length - 1];
    if (last && p.start <= last.end) last.end = Math.max(last.end, p.end);
    else out.push({ start: p.start, end: p.end, type: p.type || 'BUSY' });
  }
  return out;
}

const utc = (ms) => `${wallStamp(ms)}Z`;

/**
 * Asking someone's calendar when they are busy: METHOD:REQUEST with a
 * VFREEBUSY naming the span and the people asked (RFC 5546 3.3.2).
 */
export function writeFreeBusyRequest({ uid, organizer, attendees, start, end }) {
  const lines = [
    'BEGIN:VFREEBUSY', writeLine('UID', {}, uid), `DTSTAMP:${utc(Date.now())}`, `DTSTART:${utc(start)}`, `DTEND:${utc(end)}`,
    writeLine('ORGANIZER', organizer.name ? { CN: [organizer.name] } : {}, `mailto:${organizer.email}`),
    ...attendees.map((a) => writeLine('ATTENDEE', a.name ? { CN: [a.name] } : {}, `mailto:${a.email}`)),
    'END:VFREEBUSY',
  ];
  return writeCalendar({ method: 'REQUEST', components: [lines] });
}

/**
 * The answer to such a request: METHOD:REPLY with the busy periods in the
 * span, and nothing about what fills them (RFC 5546 3.3.3).
 */
export function writeFreeBusyReply({ uid, organizer, attendee, start, end, busy }) {
  const periods = mergeBusy(busy).filter((p) => p.end > start && p.start < end).map((p) => `${utc(Math.max(p.start, start))}/${utc(Math.min(p.end, end))}`);
  const lines = [
    'BEGIN:VFREEBUSY', writeLine('UID', {}, uid), `DTSTAMP:${utc(Date.now())}`, `DTSTART:${utc(start)}`, `DTEND:${utc(end)}`,
    writeLine('ORGANIZER', organizer?.name ? { CN: [organizer.name] } : {}, `mailto:${organizer?.email || ''}`),
    writeLine('ATTENDEE', attendee.name ? { CN: [attendee.name] } : {}, `mailto:${attendee.email}`),
    ...(periods.length ? [writeLine('FREEBUSY', { FBTYPE: ['BUSY'] }, periods.join(','))] : []),
    'END:VFREEBUSY',
  ];
  return writeCalendar({ method: 'REPLY', components: [lines] });
}

/**
 * The free stretches between busy periods, inside [from, to), each at least
 * `least` long: what a room says when it turns a time down.
 */
export function freeBetween(busy, from, to, { least = 15 * 60000 } = {}) {
  const out = [];
  let at = from;
  for (const p of mergeBusy(busy)) {
    if (p.end <= at) continue;
    if (p.start >= to) break;
    if (p.start - at >= least) out.push({ start: at, end: Math.min(p.start, to) });
    at = Math.max(at, p.end);
  }
  if (to - at >= least) out.push({ start: at, end: to });
  return out;
}

/** Whether [start, end) runs into any busy period. */
export const clashes = (busy, start, end) => busy.some((p) => p.start < end && p.end > start);

