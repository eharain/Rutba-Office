// The Calendar watcher (apps/desktop/main/watch-calendar.js) against
// Calendar's own service, kept in a folder of its own: an event saved as the
// editor saves it is reminded of, a series each time, a whole day by its
// date, a meeting before its start, and a change is heard as Calendar tells
// it. The clock and the timers are stand-ins; nothing is shown.

import test from 'node:test';
import assert from 'node:assert/strict';
import { tempDir } from './temp-dirs.js';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { watchCalendar, FIRST_MS, SETTLE_MS } from '../apps/desktop/main/watch-calendar.js';
import { createNotifier } from '../packages/office-shell/src/electron/notifier.js';

const pad = (n) => String(n).padStart(2, '0');
const local = (day, h, m = 0) => new Date(2026, 10, day, h, m).getTime();
/** A time as the editor writes it: the wall clock, and its instant. */
const when = (day, h, m = 0) => ({ date: `2026-11-${pad(day)}`, time: `${pad(h)}:${pad(m)}:00`, tzid: null, utc: false, allDay: false, at: local(day, h, m) });
const wholeDay = (day) => ({ date: `2026-11-${pad(day)}`, time: null, tzid: null, utc: false, allDay: true, at: Date.UTC(2026, 10, day) });
const remind = (before) => ({ action: 'DISPLAY', before, at: null, related: 'START', description: null });
const CODE = 'meet-ABCDEFGHIJ0123456789';

function setup(now) {
  const clock = { at: now };
  const kept = new Map();
  const settings = { get: (k, f) => (kept.has(k) ? structuredClone(kept.get(k)) : f), set: (k, v) => kept.set(k, structuredClone(v)) };
  const listeners = new Set();
  const calendar = createCalendarService({ stores: { dir: tempDir('watch-calendar-') }, broadcast: (name) => { if (name === 'calendar:changed') for (const fn of [...listeners]) fn(); } });
  const shown = [];
  const opened = [];
  const notifier = createNotifier({
    settings,
    show: (note) => { shown.push(note); return { close: () => {} }; },
    open: (target) => opened.push(target),
    now: () => clock.at,
    timeOf: (ms) => `${pad(new Date(ms).getHours())}:${pad(new Date(ms).getMinutes())}`,
    setTimer: () => ({}),
    clearTimer: () => {},
  });
  const timers = [];
  const watch = watchCalendar({
    notifier,
    calendar,
    settings,
    on: (name, fn) => { if (name === 'calendar:changed') listeners.add(fn); return () => listeners.delete(fn); },
    talk: () => true,
    now: () => clock.at,
    setTimer: (fn, ms) => { const t = { fn, ms }; timers.push(t); return t; },
    clearTimer: (t) => { const i = timers.indexOf(t); if (i >= 0) timers.splice(i, 1); },
  });
  const run = (ms) => { const i = timers.findIndex((t) => t.ms === ms); assert.ok(i >= 0, `a timer of ${ms} ms waits`); timers.splice(i, 1)[0].fn(); };
  const go = (to) => { clock.at = to; watch.tick(); };
  return { calendar, watch, notifier, shown, opened, timers, run, go, clock };
}

test('with Calendar itself: a saved event is reminded of, a series each time, a whole day by its date, a meeting before it starts', () => {
  const s = setup(local(23, 8));
  const cal = s.calendar.calendars()[0];
  const one = s.calendar.save({ calendarId: cal.id, event: { summary: 'Planning', location: 'Room 2', description: '', allDay: false, start: when(23, 11), end: when(23, 12), alarms: [remind(15)], attendees: [], status: 'CONFIRMED' } });
  const daily = s.calendar.save({ calendarId: cal.id, event: { summary: 'Stand-up', location: '', description: '', allDay: false, start: when(23, 9), end: when(23, 9, 15), rrule: { freq: 'DAILY', interval: 1, count: null, until: null, byDay: null, byMonthDay: null, byMonth: null, bySetPos: null, wkst: 'MO' }, alarms: [remind(10)], attendees: [], status: 'CONFIRMED' } });
  const day = s.calendar.save({ calendarId: cal.id, event: { summary: 'Holiday', location: '', description: '', allDay: true, start: wholeDay(24), end: wholeDay(25), alarms: [remind(0)], attendees: [], status: 'CONFIRMED' } });
  const meeting = s.calendar.save({ calendarId: cal.id, event: { summary: 'Weekly', location: `https://talk.example.org/m/${CODE}`, description: '', allDay: false, start: when(23, 14), end: when(23, 15), alarms: [], attendees: [{ email: 'amina@example.org', name: 'Amina', role: 'REQ-PARTICIPANT', partstat: 'NEEDS-ACTION', rsvp: true }], status: 'CONFIRMED' } });
  s.run(FIRST_MS);

  // The day's events for the menu, as the service tells them.
  assert.deepEqual(s.notifier.state().events.map((e) => e.summary), ['Stand-up', 'Planning', 'Weekly']);

  s.go(local(23, 8, 50));
  assert.deepEqual(s.shown.map((n) => [n.kind, n.title, n.body]), [['reminder', 'Stand-up', 'In 10 minutes']]);
  s.go(local(23, 10, 45));
  assert.deepEqual([s.shown[1].title, s.shown[1].body], ['Planning', 'In 15 minutes · Room 2']);
  s.shown[1].onClick();
  assert.deepEqual(s.opened, [{ app: 'calendar', query: { event: one.id, at: local(23, 11) } }]);

  s.go(local(23, 13, 55));
  assert.deepEqual([s.shown[2].kind, s.shown[2].title, s.shown[2].body, s.shown[2].actions.map((a) => a.id)], ['meeting', 'Weekly', 'In 5 minutes', ['join']]);

  // Midnight: the whole day begins, by this computer's clock.
  s.go(local(24, 0, 0));
  const holiday = s.shown.find((n) => n.title === 'Holiday');
  assert.equal(holiday.body, 'Today');
  holiday.onClick();
  assert.deepEqual(s.opened[1], { app: 'calendar', query: { event: day.id, at: Date.UTC(2026, 10, 24) } });
  // What the window is sent finds the occurrence again, as the window asks for it.
  const asked = s.opened[1].query;
  assert.ok(s.calendar.events({ from: asked.at - 1, to: asked.at + 60000, all: true }).some((o) => o.id === asked.event && o.start === asked.at));

  // The series again the next morning.
  s.go(local(24, 8, 50));
  assert.equal(s.shown.filter((n) => n.title === 'Stand-up').length, 2);
  assert.ok(daily.id && meeting.id);
});

test('with Calendar itself: an event saved or removed is heard, and its reminder follows it', () => {
  const s = setup(local(23, 8));
  s.run(FIRST_MS);
  const cal = s.calendar.calendars()[0];
  const saved = s.calendar.save({ calendarId: cal.id, event: { summary: 'Call the bank', location: '', description: '', allDay: false, start: when(23, 9), end: when(23, 9, 30), alarms: [remind(15)], attendees: [], status: 'CONFIRMED' } });
  s.run(SETTLE_MS);
  assert.deepEqual(s.watch.dues().map((d) => d.at), [local(23, 8, 45)]);
  s.calendar.remove({ id: saved.id });
  s.run(SETTLE_MS);
  assert.deepEqual(s.watch.dues(), []);
  s.go(local(23, 8, 45));
  assert.deepEqual(s.shown, []);
});
