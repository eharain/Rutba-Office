// Calendar, as the notifier hears of it (apps/desktop/main/watch-calendar.js,
// D-2026-11-22-23): an event's reminder, a meeting about to start, and the
// day's events for the menu of the icon by the clock. Calendar's service,
// the clock and the timers are stand-ins: nothing is shown on this computer.

import test from 'node:test';
import assert from 'node:assert/strict';
import { watchCalendar, duesOf, timesOf, meetingCodeIn, TOLD_KEY, MEETING_LEAD_MS, RESCAN_MS, SETTLE_MS, FIRST_MS, TICK_MS } from '../apps/desktop/main/watch-calendar.js';
import { createNotifier } from '../packages/office-shell/src/electron/notifier.js';

const at = (h, m = 0, day = 23) => new Date(2026, 10, day, h, m).getTime();
const MIN = 60_000;
const CODE = 'meet-ABCDEFGHIJ0123456789';
const LINK = `https://talk.example.org/m/${CODE}`;

/** An event as Calendar keeps it, and the one occurrence of it the service answers with. */
const event = (id, start, more = {}) => ({ id, summary: 'Planning', location: '', description: '', start, end: start + 60 * MIN, allDay: false, status: null, attendees: 0, alarms: [], ...more });
const remind = (before) => ({ action: 'DISPLAY', before, at: null, related: 'START', description: null });

function setup({ events = [], now = at(10), talk = true, stored = {} } = {}) {
  const clock = { at: now };
  const kept = new Map(Object.entries(stored));
  const settings = { get: (k, f) => (kept.has(k) ? structuredClone(kept.get(k)) : f), set: (k, v) => kept.set(k, structuredClone(v)) };
  const shown = [];
  const opened = [];
  const notifier = createNotifier({
    settings,
    show: (note) => { const e = { ...note, closed: false }; shown.push(e); return { close: () => { e.closed = true; } }; },
    open: (target) => opened.push(target),
    now: () => clock.at,
    timeOf: (ms) => `${String(new Date(ms).getHours()).padStart(2, '0')}:${String(new Date(ms).getMinutes()).padStart(2, '0')}`,
    dayOf: (ms) => `day ${new Date(ms).getDate()}`,
    setTimer: () => ({}),
    clearTimer: () => {},
  });
  const book = { events: [...events] };
  const asked = { events: 0, get: 0 };
  const calendar = {
    events: ({ from, to }) => {
      asked.events += 1;
      return book.events
        .filter((e) => e.end > from && e.start < to)
        .map((e) => ({ id: e.id, start: e.start, end: e.end, original: e.original ?? e.start, allDay: e.allDay, summary: e.summary || '(no title)', location: e.location, status: e.status, attendees: e.attendees }))
        .sort((a, b) => a.start - b.start);
    },
    get: ({ id }) => { asked.get += 1; return book.events.find((e) => e.id === id) || null; },
  };
  const listeners = new Map();
  const on = (name, fn) => {
    if (!listeners.has(name)) listeners.set(name, new Set());
    listeners.get(name).add(fn);
    return () => listeners.get(name).delete(fn);
  };
  const emit = (name, payload) => { for (const fn of [...(listeners.get(name) || [])]) fn(payload); };
  const timers = [];
  const ctx = { talk };
  const joined = [];
  const watch = watchCalendar({
    notifier,
    calendar,
    settings,
    on,
    open: (target) => joined.push(target),
    talk: () => ctx.talk,
    now: () => clock.at,
    setTimer: (fn, ms) => { const t = { fn, ms }; timers.push(t); return t; },
    clearTimer: (t) => { const i = timers.indexOf(t); if (i >= 0) timers.splice(i, 1); },
  });
  /** The timer that waits `ms` runs, as the clock would run it. */
  const run = (ms) => {
    const i = timers.findIndex((t) => t.ms === ms);
    assert.ok(i >= 0, `a timer of ${ms} ms waits`);
    timers.splice(i, 1)[0].fn();
  };
  /** Office has been up a moment: the calendars are read for the first time. */
  const start = () => run(FIRST_MS);
  /** The clock moves on, and what is due is looked at. */
  const go = (to) => { clock.at = to; watch.tick(); };
  return { watch, notifier, shown, opened, joined, book, asked, clock, kept, timers, emit, listeners, ctx, start, go, run };
}

test('a reminder comes at its time, once, says when and where, and opens Calendar at the event', () => {
  const s = setup({ events: [event('e1', at(11), { location: 'Room 2', alarms: [remind(15)] })] });
  assert.deepEqual(s.shown, [], 'nothing is read until Office has been up a moment');
  s.start();
  assert.deepEqual(s.watch.dues(), [{ key: `e1:${at(11)}:reminder:${at(10, 45)}`, at: at(10, 45), kind: 'reminder' }]);
  assert.deepEqual(s.shown, []);
  s.go(at(10, 44));
  assert.deepEqual(s.shown, []);
  s.go(at(10, 45));
  assert.equal(s.shown.length, 1);
  assert.equal(s.shown[0].kind, 'reminder');
  assert.equal(s.shown[0].title, 'Planning');
  assert.equal(s.shown[0].body, 'In 15 minutes · Room 2');
  assert.deepEqual(s.shown[0].actions, []);
  s.shown[0].onClick();
  assert.deepEqual(s.opened, [{ app: 'calendar', query: { event: 'e1', at: at(11) } }]);
  s.go(at(10, 46));
  s.watch.scan();
  assert.equal(s.shown.length, 1, 'told once, however often the calendars are read');
  assert.equal(s.kept.get(TOLD_KEY)[`e1:${at(11)}:reminder:${at(10, 45)}`], at(12), 'kept until its event is over');
});

test('the next thing due sets the timer, never further off than the tick', () => {
  const s = setup({ events: [event('e1', at(11), { alarms: [remind(59), remind(15)] })], now: at(10, 0) });
  s.start();
  assert.equal(s.timers.length, 1);
  assert.equal(s.timers[0].ms, TICK_MS, 'a reminder a minute off is looked for at the next tick');
  s.clock.at = at(10, 0) + 50_000;
  s.watch.tick();
  assert.equal(s.timers[0].ms, 10_000, 'ten seconds off, in ten seconds');
  assert.deepEqual(s.shown, []);
  s.clock.at = at(10, 1);
  s.run(10_000);
  assert.equal(s.shown.length, 1);
  assert.equal(s.shown[0].body, 'In 59 minutes');
});

test('Office started late: a reminder whose time passed is shown while its event is on or starts within a day, and let go otherwise', () => {
  const s = setup({
    now: at(10, 50),
    events: [
      event('soon', at(11), { summary: 'Soon', alarms: [remind(15)] }),
      event('on', at(9, 30), { summary: 'On now', end: at(11, 30), alarms: [remind(15)] }),
      event('over', at(8), { summary: 'Over', alarms: [remind(15)] }),
      event('far', at(9, 0, 26), { summary: 'In three days', alarms: [remind(7 * 24 * 60)] }),
    ],
  });
  s.start();
  assert.deepEqual(s.shown.map((n) => [n.title, n.body]).sort(), [['On now', 'Started at 09:30'], ['Soon', 'In 10 minutes']]);
  const told = s.kept.get(TOLD_KEY);
  assert.ok(Object.keys(told).some((k) => k.startsWith('far:')), 'the one let go is not offered again');
  s.go(at(9, 0, 26) - 10 * MIN);
  assert.equal(s.shown.length, 2);
});

test('a whole day is its date on this computer: reminded before its own midnight, and said by its day', () => {
  const day = Date.UTC(2026, 10, 24);
  const s = setup({ now: at(23, 40), events: [event('d1', day, { summary: 'Holiday', end: day + 86_400_000, allDay: true, alarms: [remind(15)] })] });
  assert.deepEqual(timesOf({ allDay: true, start: day, end: day + 86_400_000 }), { start: at(0, 0, 24), end: at(0, 0, 25) });
  s.start();
  assert.deepEqual(s.shown, []);
  s.go(at(23, 45));
  assert.equal(s.shown.length, 1);
  assert.equal(s.shown[0].body, 'Tomorrow');
  s.shown[0].onClick();
  assert.deepEqual(s.opened, [{ app: 'calendar', query: { event: 'd1', at: day } }], 'Calendar is asked by the instant it knows the day by');
});

test('a meeting is told of a few minutes before it starts, reminder or none; with Join when its link is Talk\'s and Talk is on', () => {
  const s = setup({
    now: at(10),
    events: [
      event('people', at(11), { summary: 'Review', attendees: 3 }),
      event('talk', at(12), { summary: 'Weekly', description: `Agenda\n\nJoin: ${LINK}`, location: LINK, alarms: [remind(5), remind(15)] }),
      event('alone', at(13), { summary: 'Lunch' }),
    ],
  });
  s.start();
  assert.deepEqual(s.watch.dues().map((d) => [d.kind, d.at]), [
    ['meeting', at(11) - MEETING_LEAD_MS],
    ['reminder', at(11, 45)],
    ['meeting', at(12) - MEETING_LEAD_MS],
  ], 'the reminder set for when the meeting is told of anyway is left out; an event of one\'s own with no reminder says nothing');
  s.go(at(10, 55));
  assert.equal(s.shown.length, 1);
  assert.deepEqual([s.shown[0].kind, s.shown[0].title, s.shown[0].body], ['meeting', 'Review', 'In 5 minutes']);
  assert.deepEqual(s.shown[0].actions, [], 'no link of Talk\'s: nothing to join from here');
  s.go(at(11, 45));
  assert.deepEqual([s.shown[1].kind, s.shown[1].body], ['reminder', 'In 15 minutes'], 'a place that is only a link is left out of the words');
  s.go(at(11, 55));
  const meeting = s.shown[2];
  assert.deepEqual([meeting.kind, meeting.title, meeting.body], ['meeting', 'Weekly', 'In 5 minutes']);
  assert.deepEqual(meeting.actions.map((a) => a.id), ['join']);
  meeting.onAction('join');
  assert.deepEqual(s.joined, [{ app: 'talk', query: { meet: CODE } }]);
  meeting.onClick();
  assert.deepEqual(s.opened, [{ app: 'calendar', query: { event: 'talk', at: at(12) } }]);
});

test('with Talk off or not installed a meeting has no Join; a link to somewhere else is never one', () => {
  const s = setup({ talk: false, now: at(11, 55), events: [event('talk', at(12), { description: LINK })] });
  s.start();
  assert.equal(s.shown.length, 1);
  assert.deepEqual(s.shown[0].actions, []);
  s.shown[0].onClick();
  assert.deepEqual(s.opened, [{ app: 'calendar', query: { event: 'talk', at: at(12) } }]);
  assert.equal(meetingCodeIn('https://meet.example.com/abc-defg-hij', 'https://talk.example.org/m/nothing'), null);
  assert.equal(meetingCodeIn('', `see ${LINK}.`), CODE);
  assert.equal(meetingCodeIn(`https://talk.example.org/m/live-ABCDEFGHIJ0123456789`), 'live-ABCDEFGHIJ0123456789');
});

test('what is due of one occurrence: reminders before the start or the end, one at a set time, each time once, none for a cancelled event', () => {
  const occ = { id: 'e1', start: at(11), end: at(12), original: at(11), allDay: false, attendees: 0, status: null, location: '' };
  const dues = duesOf(occ, { alarms: [remind(15), remind(15), { before: 10, related: 'END' }, { before: null, at: at(9) }, { before: null, at: at(13) }, null, { before: 'soon' }] });
  assert.deepEqual(dues.map((d) => [d.kind, d.at]), [['reminder', at(10, 45)], ['reminder', at(11, 50)], ['reminder', at(9)]], 'one set for after the event is over says nothing');
  assert.deepEqual(duesOf({ ...occ, status: 'CANCELLED' }, { alarms: [remind(15)] }), []);
  assert.deepEqual(duesOf(occ, null), []);
  assert.deepEqual(duesOf({ ...occ, allDay: true, attendees: 4, start: Date.UTC(2026, 10, 24), end: Date.UTC(2026, 10, 25) }, { alarms: [] }), [], 'a whole day with people on it is not a meeting about to start');
  // A moved occurrence of a series is its own: told again for its new time.
  const moved = duesOf({ ...occ, start: at(14), end: at(15) }, { alarms: [remind(15)] });
  assert.notEqual(moved[0].key, duesOf(occ, { alarms: [remind(15)] })[0].key);
});

test('the day\'s events still to come go to the menu: today\'s only, in order, a whole day by its date here, cancelled ones left out', () => {
  const s = setup({
    now: at(10),
    events: [
      event('past', at(8)),
      event('on', at(9, 30), { summary: 'On now', end: at(10, 30) }),
      event('later', at(15), { summary: '' }),
      event('off', at(16), { status: 'CANCELLED' }),
      event('day', Date.UTC(2026, 10, 23), { summary: 'Holiday', end: Date.UTC(2026, 10, 24), allDay: true }),
      event('tomorrow', at(9, 0, 24)),
    ],
  });
  s.start();
  const events = s.notifier.state().events;
  assert.deepEqual(events.map((e) => e.id).sort(), ['day', 'later', 'on']);
  const day = events.find((e) => e.id === 'day');
  assert.deepEqual([day.start, day.end, day.allDay, day.at], [at(0), at(0, 0, 24), true, Date.UTC(2026, 10, 23)]);
  assert.equal(events.find((e) => e.id === 'later').summary, '', 'an event with no title is named by the menu, in the person\'s language');
});

test('told that the calendars changed, they are read again once things settle: an event moved is reminded of at its new time', () => {
  const s = setup({ now: at(10), events: [event('e1', at(11), { alarms: [remind(15)] })] });
  s.emit('calendar:changed', {});
  assert.equal(s.timers.filter((t) => t.ms === SETTLE_MS).length, 0, 'the first reading is still to come, and will see it');
  s.start();
  const reads = s.asked.events;
  s.book.events[0] = event('e1', at(14), { alarms: [remind(15)] });
  s.emit('calendar:changed', {});
  s.emit('calendar:changed', {});
  s.emit('calendar:changed', {});
  assert.equal(s.timers.filter((t) => t.ms === SETTLE_MS).length, 1, 'many changes at once are one reading');
  s.run(SETTLE_MS);
  assert.equal(s.asked.events, reads + 1);
  assert.deepEqual(s.watch.dues().map((d) => d.at), [at(13, 45)]);
  s.go(at(10, 45));
  assert.deepEqual(s.shown, []);
  s.go(at(13, 45));
  assert.equal(s.shown.length, 1);
});

test('with nothing said to have changed the calendars are still read again now and then; an event read once is asked for once', () => {
  const s = setup({ now: at(10), events: [event('e1', at(11), { alarms: [remind(15)] }), event('e1', at(11, 0, 24), { alarms: [remind(15)] })] });
  s.start();
  assert.equal(s.asked.get, 1, 'two occurrences of one event');
  const reads = s.asked.events;
  s.go(at(10) + TICK_MS);
  assert.equal(s.asked.events, reads);
  s.go(at(10) + RESCAN_MS);
  assert.equal(s.asked.events, reads + 1);
});

test('what was told of events long over is forgotten; what is still to come is kept', () => {
  const s = setup({ now: at(10), stored: { [TOLD_KEY]: { 'old:1:reminder:1': at(10) - 3 * 86_400_000, 'new:1:reminder:1': at(12) } } });
  s.start();
  assert.deepEqual(Object.keys(s.kept.get(TOLD_KEY)), ['new:1:reminder:1']);
});

test('reminders turned off are not shown and not kept for later; an event from a file opened is not the person\'s own', () => {
  const s = setup({ now: at(10, 45), stored: { notifications: { kinds: { reminder: false } } }, events: [event('e1', at(11), { alarms: [remind(15)] }), event('file:1', at(11), { alarms: [remind(15)] })] });
  s.start();
  assert.deepEqual(s.shown, []);
  s.notifier.setSettings({ kinds: { reminder: true } });
  s.watch.scan();
  assert.deepEqual(s.shown, [], 'its time has gone');
});

test('a Calendar with a fault is read as empty; one that cannot be read at all is not watched; stopped, nothing waits', () => {
  const s = setup({ now: at(10), events: [event('e1', at(11), { alarms: [remind(15)] })] });
  s.start();
  s.watch.stop();
  assert.deepEqual(s.timers, []);
  assert.equal(s.listeners.get('calendar:changed').size, 0);
  s.clock.at = at(10, 45);
  s.watch.tick();
  s.watch.scan();
  assert.deepEqual(s.shown, []);

  const kept = new Map();
  const faulty = watchCalendar({
    notifier: { report: () => assert.fail('nothing to report'), setEvents: () => {} },
    calendar: { events: () => { throw new Error('unreadable'); }, get: () => null },
    settings: { get: (k, f) => (kept.has(k) ? kept.get(k) : f), set: (k, v) => kept.set(k, v) },
    on: () => () => {},
    setTimer: () => ({}),
    clearTimer: () => {},
  });
  assert.doesNotThrow(() => faulty.scan());
  assert.equal(watchCalendar({ notifier: {}, calendar: {}, settings: {}, on: () => () => {} }), null);
  assert.equal(watchCalendar({ notifier: {}, calendar: null, settings: {}, on: () => () => {} }), null);
});
