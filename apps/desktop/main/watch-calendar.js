// Calendar, as the notifier hears of it (watch.js, D-2026-11-22-23): an
// event's reminder, a meeting about to start, and the day's events still to
// come for the menu of the icon by the clock.
//
// Calendar keeps an event's reminders (the editor's "Reminder", and those a
// calendar account or a file brought) and until now nothing told the person
// of them. This file does, in the main process, so they come whether or not
// a Calendar window is open. It reads the calendars through Calendar's own
// service and writes nothing to them.
//
// Only while Rutba Office runs. A reminder whose time passed while Office
// was closed is shown when Office next starts if its event has not ended
// and is on now or starts within a day; an older one is let go. What has
// been told is kept in the settings, so nothing is told twice.
//
// A meeting is an event with other people on it, or with a Talk meeting's
// link in its place or notes. It is told of a few minutes before it starts,
// whether or not it has a reminder, with Join when the link is Talk's and
// Talk is on. A link to somewhere else is never opened from here: the
// notification opens the event, where the link is.

const DAY_MS = 86_400_000;
/** How far ahead events are read: a reminder may be set two weeks before. */
export const AHEAD_MS = 15 * DAY_MS;
/** How often what is due is looked at; a computer back from sleep is late by this much at most. */
export const TICK_MS = 30_000;
/** How often the calendars are read again when nothing says they changed. */
export const RESCAN_MS = 10 * 60_000;
/** How long before its start a meeting is told of. */
export const MEETING_LEAD_MS = 5 * 60_000;
/** How long after a change the calendars are read: an account in step changes many events at once. */
export const SETTLE_MS = 1500;
/** How long after Office starts they are first read: its first window comes before a late reminder does. */
export const FIRST_MS = 4000;
/** What has been told: { key: when its event ends }. */
export const TOLD_KEY = 'notifications.calendarTold';

const CODE = '(?:live|meet)-[A-Za-z0-9_-]{20}';
const LINK = new RegExp(`https?://[^\\s/<>"']+/m/(${CODE})(?![A-Za-z0-9_-])`);

/** The code of the first Talk meeting linked in some words (an event's place and notes), or null. */
export function meetingCodeIn(...texts) {
  for (const text of texts) {
    const m = LINK.exec(String(text || ''));
    if (m) return m[1];
  }
  return null;
}

/** A whole day is kept as midnight UTC of its date: here it is that date's midnight on this computer. */
const localDay = (ms) => {
  const d = new Date(ms);
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()).getTime();
};

/** An occurrence's start and end as the person lives them. */
export function timesOf(occ) {
  if (!occ.allDay) return { start: occ.start, end: Math.max(occ.end, occ.start) };
  const start = localDay(occ.start);
  return { start, end: Math.max(localDay(occ.end), start + DAY_MS) };
}

const linkOnly = (text) => /^\s*https?:\/\/\S+\s*$/i.test(String(text || ''));

/**
 * What is to be told of one occurrence, and when: its reminders, and its
 * start when it is a meeting. `[{ key, at, kind, start, end, code }]`.
 *
 * @param {object} occ    an occurrence, as Calendar's `events` gives it
 * @param {object} event  its event, whole, as Calendar's `get` gives it (reminders, notes)
 */
export function duesOf(occ, event) {
  if (!occ || occ.id == null || occ.status === 'CANCELLED') return [];
  const { start, end } = timesOf(occ);
  const which = `${occ.id}:${occ.original ?? occ.start}`;
  const code = meetingCodeIn(occ.location, event?.location, event?.description);
  const out = [];
  const meeting = !occ.allDay && (Boolean(code) || Number(occ.attendees) > 0);
  const meetAt = start - MEETING_LEAD_MS;
  if (meeting) out.push({ key: `${which}:meeting:${start}`, at: meetAt, kind: 'meeting', start, end, code });
  const seen = new Set();
  for (const alarm of Array.isArray(event?.alarms) ? event.alarms : []) {
    if (!alarm) continue;
    let at = null;
    if (alarm.before != null && Number.isFinite(Number(alarm.before))) at = (alarm.related === 'END' ? end : start) - Number(alarm.before) * 60_000;
    else if (Number.isFinite(alarm.at)) at = alarm.at;
    if (at == null || at > end || seen.has(at)) continue;
    seen.add(at);
    // A reminder set for when the meeting is told of anyway says nothing more.
    if (meeting && Math.abs(at - meetAt) < 60_000) continue;
    out.push({ key: `${which}:reminder:${at}`, at, kind: 'reminder', start, end, code: null });
  }
  return out;
}

/**
 * @param {object} o
 * @param {object} o.notifier
 * @param {object} o.calendar   Calendar's service: events({ from, to }), get({ id })
 * @param {{ get: Function, set: Function }} o.settings   the profile's settings store
 * @param {(event: string, fn: Function) => Function} o.on   listens to what the main process tells its windows
 * @param {(target: { app: string, query?: object }) => void} [o.open]   opens an app at a thing (Join)
 * @param {() => boolean} [o.talk]   whether Talk is installed and on: only then is Join offered
 * @param {() => number} [o.now]
 * @returns {{ scan: () => void, tick: () => void, dues: () => object[], waiting: () => boolean, stop: () => void } | null}
 */
export function watchCalendar({ notifier, calendar, settings, on, open = () => {}, talk = () => false, now = () => Date.now(), setTimer = setTimeout, clearTimer = clearTimeout }) {
  if (typeof calendar?.events !== 'function' || typeof calendar?.get !== 'function') return null;

  let dues = [];
  let scannedAt = 0;
  let timer = null;
  let settle = null;
  let stopped = false;

  const told = () => {
    const kept = settings.get(TOLD_KEY, null);
    return kept && typeof kept === 'object' ? kept : {};
  };
  const keepTold = (next) => settings.set(TOLD_KEY, next);

  const tell = (due) => {
    // Calendar finds the occurrence again by its event and the instant it starts.
    const at = { app: 'calendar', query: { event: due.occ.id, at: due.occ.start } };
    const summary = due.occ.summary && due.occ.summary !== '(no title)' ? due.occ.summary : '';
    const location = linkOnly(due.occ.location) ? '' : due.occ.location || '';
    const join = due.kind === 'meeting' && Boolean(due.code) && Boolean(talk());
    notifier.report({
      kind: due.kind,
      id: `event:${due.key}`,
      summary,
      start: due.start,
      allDay: Boolean(due.occ.allDay),
      location,
      join,
      // A link to a meeting streamed live is said so, with Talk or without.
      live: due.kind === 'meeting' && String(due.code || '').startsWith('live-'),
      app: 'calendar',
      open: at,
      ...(join ? { on: { join: () => open({ app: 'talk', query: { meet: due.code } }) } } : {}),
    });
  };

  /** What is due is told, once; a late one only while its event is on or starts within a day. */
  const fire = () => {
    const at = now();
    const kept = told();
    let changed = false;
    const waiting = [];
    for (const due of dues) {
      if (kept[due.key] != null) continue;
      if (due.at > at) { waiting.push(due); continue; }
      kept[due.key] = due.end;
      changed = true;
      if (due.end > at && due.start - at <= DAY_MS) {
        try { tell(due); } catch { /* one event's fault is its own */ }
      }
    }
    dues = waiting;
    if (changed) keepTold(kept);
  };

  const arm = () => {
    clearTimer(timer);
    timer = null;
    if (stopped) return;
    const at = now();
    const next = dues.reduce((soonest, d) => Math.min(soonest, d.at), at + TICK_MS);
    timer = setTimer(() => { timer = null; tick(); }, Math.max(250, Math.min(TICK_MS, next - at)));
    timer?.unref?.();
  };

  /** The calendars read again: what is to be told and when, and the day's events still to come. */
  const scan = () => {
    if (stopped) return;
    const at = now();
    scannedAt = at;
    let occurrences = [];
    try {
      occurrences = calendar.events({ from: at - DAY_MS, to: at + AHEAD_MS }) || [];
    } catch {
      occurrences = [];
    }
    const events = new Map();
    const eventOf = (id) => {
      if (!events.has(id)) {
        let e = null;
        try { e = calendar.get({ id }); } catch { e = null; }
        events.set(id, e);
      }
      return events.get(id);
    };
    const kept = told();
    const next = [];
    const today = [];
    const d = new Date(at);
    const dayEnd = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1).getTime();
    for (const occ of occurrences) {
      if (!occ || occ.id == null || String(occ.id).startsWith('file:') || occ.status === 'CANCELLED') continue;
      const { start, end } = timesOf(occ);
      if (end <= at) continue;
      if (start < dayEnd) today.push({ id: occ.id, summary: occ.summary && occ.summary !== '(no title)' ? occ.summary : '', start, end, allDay: Boolean(occ.allDay), at: occ.start });
      for (const due of duesOf(occ, eventOf(occ.id))) if (kept[due.key] == null) next.push({ ...due, occ });
    }
    dues = next.sort((a, b) => a.at - b.at);
    // What was told of events long over is forgotten.
    let pruned = false;
    for (const [key, end] of Object.entries(kept)) {
      if (!(Number(end) > at - 2 * DAY_MS)) { delete kept[key]; pruned = true; }
    }
    if (pruned) keepTold(kept);
    try { notifier.setEvents(today); } catch { /* the menu keeps what it had */ }
    fire();
    arm();
  };

  function tick() {
    if (stopped) return;
    if (now() - scannedAt >= RESCAN_MS) { scan(); return; }
    fire();
    // The day's events drop off the menu as they end.
    arm();
  }

  const stopHearing = on('calendar:changed', () => {
    if (stopped || settle) return;
    settle = setTimer(() => { settle = null; scan(); }, SETTLE_MS);
    settle?.unref?.();
  });

  settle = setTimer(() => { settle = null; scan(); }, FIRST_MS);
  settle?.unref?.();

  return {
    scan,
    tick,
    dues: () => dues.map(({ key, at, kind }) => ({ key, at, kind })),
    /** Whether a reminder or a meeting is still to be told of (the calendars are read first if they have not been). */
    waiting: () => {
      if (!scannedAt) scan();
      return dues.length > 0;
    },
    stop() {
      stopped = true;
      clearTimer(timer);
      clearTimer(settle);
      timer = null;
      settle = null;
      try { stopHearing?.(); } catch { /* gone already */ }
    },
  };
}
