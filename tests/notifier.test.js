// The notifier (office-shell notifier.js, D-2026-11-22-23): the one place
// that shows a Windows notification, flashes the taskbar and counts what
// waits. Its Electron parts are stand-ins here: nothing reaches the screen.

import test from 'node:test';
import assert from 'node:assert/strict';
import { createNotifier, readSettings, inQuietHours, wordsOf, KINDS, COUNTED, SETTINGS_KEY, PAUSE_MS } from '../packages/office-shell/src/electron/notifier.js';

function setup({ at = new Date(2026, 10, 23, 10, 0).getTime(), stored = undefined } = {}) {
  const kept = new Map();
  if (stored !== undefined) kept.set(SETTINGS_KEY, stored);
  const settings = { get: (k, f) => (kept.has(k) ? structuredClone(kept.get(k)) : f), set: (k, v) => kept.set(k, structuredClone(v)) };
  const shown = [];
  const flashed = [];
  const opened = [];
  const timers = [];
  const clock = { at };
  const notifier = createNotifier({
    settings,
    show: (note) => {
      const entry = { ...note, closed: false };
      shown.push(entry);
      return { close: () => { entry.closed = true; } };
    },
    flash: (app) => flashed.push(app),
    open: (target) => opened.push(target),
    now: () => clock.at,
    timeOf: (ms) => `${String(new Date(ms).getHours()).padStart(2, '0')}:${String(new Date(ms).getMinutes()).padStart(2, '0')}`,
    setTimer: (fn, ms) => { const t = { fn, ms }; timers.push(t); return t; },
    clearTimer: (t) => { const i = timers.indexOf(t); if (i >= 0) timers.splice(i, 1); },
  });
  return { notifier, shown, flashed, opened, timers, clock, kept };
}

test('the settings as kept are whole whatever was stored: every kind on, sound and text on, kept running, not started with Windows', () => {
  const s = readSettings(null);
  assert.deepEqual(Object.keys(s.kinds), KINDS);
  assert.ok(KINDS.every((k) => s.kinds[k] === true));
  assert.deepEqual({ sound: s.sound, text: s.text, keepRunning: s.keepRunning, startup: s.startup, pausedUntil: s.pausedUntil }, { sound: true, text: true, keepRunning: true, startup: false, pausedUntil: 0 });
  assert.deepEqual(s.quiet, { on: false, from: '22:00', to: '07:00' });
  const odd = readSettings({ kinds: { mail: false, other: true }, sound: 0, quiet: { on: 1, from: '25:00', to: '6:30' }, pausedUntil: 'x', startup: 'yes' });
  assert.equal(odd.kinds.mail, false);
  assert.equal(odd.kinds.other, undefined);
  assert.equal(odd.sound, true, 'only false turns sound off');
  assert.deepEqual(odd.quiet, { on: true, from: '22:00', to: '07:00' }, 'a time that is no time falls back');
  assert.equal(odd.pausedUntil, 0);
  assert.equal(odd.startup, true);
});

test('quiet hours: within a day, past midnight, and none when both times are the same', () => {
  const at = (h, m = 0) => new Date(2026, 10, 23, h, m).getTime();
  const night = { on: true, from: '22:00', to: '07:00' };
  assert.equal(inQuietHours(night, at(23)), true);
  assert.equal(inQuietHours(night, at(6, 59)), true);
  assert.equal(inQuietHours(night, at(7)), false);
  assert.equal(inQuietHours(night, at(21, 59)), false);
  const lunch = { on: true, from: '12:00', to: '13:30' };
  assert.equal(inQuietHours(lunch, at(12)), true);
  assert.equal(inQuietHours(lunch, at(13, 30)), false);
  assert.equal(inQuietHours({ on: true, from: '09:00', to: '09:00' }, at(9)), false);
  assert.equal(inQuietHours({ ...night, on: false }, at(23)), false);
});

test('a message is counted, shown with its words, flashes its app and opens at the person when pressed', () => {
  const { notifier, shown, flashed, opened } = setup();
  const r = notifier.report({ kind: 'chat', id: 'chat:u1', from: 'Amina Yusuf', text: 'Are you  free\nat three?', app: 'talk', open: { app: 'talk', query: { contact: 'u1' } } });
  assert.deepEqual(r, { shown: true, id: 'chat:u1' });
  assert.equal(shown.length, 1);
  assert.equal(shown[0].title, 'Amina Yusuf');
  assert.equal(shown[0].body, 'Are you free at three?');
  assert.equal(shown[0].silent, false);
  assert.equal(shown[0].sticky, false);
  assert.deepEqual(flashed, ['talk']);
  assert.equal(notifier.state().waiting.chat, 1);
  assert.equal(notifier.state().waiting.total, 1);
  shown[0].onClick();
  assert.deepEqual(opened, [{ app: 'talk', query: { contact: 'u1' } }]);
});

test('the same thing reported again takes the earlier notification down', () => {
  const { notifier, shown } = setup();
  notifier.report({ kind: 'chat', id: 'chat:u1', from: 'Amina', text: 'one' });
  notifier.report({ kind: 'chat', id: 'chat:u1', from: 'Amina', text: 'two', count: 2 });
  assert.equal(shown[0].closed, true);
  assert.equal(shown[1].body, 'two (2 new messages)');
  assert.equal(notifier.state().waiting.chat, 3);
});

test('with the text hidden a message says how many came and an email says only who it is from', () => {
  const { notifier, shown } = setup({ stored: { text: false } });
  notifier.report({ kind: 'chat', from: 'Amina', text: 'the secret plan' });
  notifier.report({ kind: 'mail', from: 'Bilal Khan', subject: 'Salaries for November' });
  assert.equal(shown[0].title, 'Amina');
  assert.equal(shown[0].body, '1 new message');
  assert.equal(shown[1].title, 'Bilal Khan');
  assert.equal(shown[1].body, 'New email');
  assert.ok(!shown.some((n) => /secret|Salaries/.test(`${n.title} ${n.body}`)));
});

test('an email shows who and the subject; several at once are counted in one notification', () => {
  const one = wordsOf({ kind: 'mail', from: 'Bilal Khan', subject: 'Quarterly numbers' });
  assert.deepEqual([one.title, one.body], ['Bilal Khan', 'Quarterly numbers']);
  assert.equal(wordsOf({ kind: 'mail', from: 'Bilal Khan', subject: '' }).body, '(no subject)');
  const many = wordsOf({ kind: 'mail', count: 4, account: 'sam@example.org' });
  assert.deepEqual([many.title, many.body], ['4 new emails', 'sam@example.org']);
});

test('a call stays up, offers Answer and Decline, and each does its own thing', () => {
  const { notifier, shown, opened } = setup();
  let declined = 0;
  notifier.report({ kind: 'call', id: 'call:c1', from: 'Amina', video: true, app: 'talk', open: { app: 'talk', query: { call: 'c1' } }, on: { decline: () => { declined += 1; } } });
  assert.equal(shown[0].title, 'Incoming video call');
  assert.equal(shown[0].body, 'Amina');
  assert.equal(shown[0].sticky, true);
  assert.deepEqual(shown[0].actions.map((a) => [a.id, a.label]), [['answer', 'Answer'], ['decline', 'Decline']]);
  shown[0].onAction('answer');
  assert.deepEqual(opened, [{ app: 'talk', query: { call: 'c1' } }], 'Answer opens Talk at the call');
  shown[0].onAction('decline');
  assert.equal(declined, 1);
  assert.equal(opened.length, 1, 'Decline opens nothing');
  assert.equal(notifier.state().waiting.total, 0, 'a ringing call is not something that waits');
  assert.equal(notifier.withdraw('call:c1'), true);
  assert.equal(shown[0].closed, true);
  assert.equal(notifier.withdraw('call:c1'), false);
  assert.equal(wordsOf({ kind: 'call' }).title, 'Incoming call');
});

test('a kind turned off is counted and not shown; so is anything while paused or in the quiet hours', () => {
  const off = setup({ stored: { kinds: { mail: false } } });
  assert.deepEqual(off.notifier.report({ kind: 'mail', from: 'A', subject: 's', app: 'mail' }), { shown: false, why: 'off' });
  assert.equal(off.shown.length, 0);
  assert.deepEqual(off.flashed, [], 'nothing flashes for a kind turned off');
  assert.equal(off.notifier.state().waiting.mail, 1);
  assert.equal(off.notifier.report({ kind: 'chat', from: 'A', text: 't' }).shown, true, 'the other kinds are still shown');

  const paused = setup();
  const until = paused.notifier.pause();
  assert.equal(until, paused.clock.at + PAUSE_MS);
  assert.deepEqual(paused.notifier.report({ kind: 'missed', from: 'A' }), { shown: false, why: 'paused' });
  assert.equal(paused.notifier.state().paused, true);
  assert.equal(paused.notifier.state().waiting.missed, 1);
  paused.clock.at += PAUSE_MS + 1000;
  assert.equal(paused.notifier.state().paused, false, 'the hour is up');
  assert.equal(paused.notifier.report({ kind: 'missed', from: 'A' }).shown, true);

  const quiet = setup({ at: new Date(2026, 10, 23, 23, 30).getTime(), stored: { quiet: { on: true, from: '22:00', to: '07:00' } } });
  assert.deepEqual(quiet.notifier.report({ kind: 'call', from: 'A' }), { shown: false, why: 'quiet' });
  assert.equal(quiet.notifier.state().quiet, true);
});

test('pausing tells its listeners again when the hour is up, and resuming ends it at once', () => {
  const { notifier, timers } = setup();
  let told = 0;
  notifier.onChange(() => { told += 1; });
  notifier.pause();
  assert.equal(timers.length, 1);
  assert.ok(timers[0].ms >= PAUSE_MS);
  const before = told;
  timers.shift().fn();
  assert.equal(told, before + 1);
  notifier.pause();
  notifier.resume();
  assert.equal(notifier.state().paused, false);
  assert.equal(timers.length, 0, 'no timer is left');
});

test('sound off shows the notification silent', () => {
  const { notifier, shown } = setup({ stored: { sound: false } });
  notifier.report({ kind: 'chat', from: 'A', text: 't' });
  assert.equal(shown[0].silent, true);
});

test('what an app window shows is seen when it comes to the front, unless the app counts for itself', () => {
  const { notifier } = setup();
  notifier.report({ kind: 'chat', from: 'A', text: 't' });
  notifier.report({ kind: 'missed', from: 'A' });
  notifier.report({ kind: 'mail', from: 'B', subject: 's' });
  assert.deepEqual(notifier.state().waiting, { missed: 1, chat: 1, mail: 1, total: 3 });
  notifier.seenApp('talk');
  assert.deepEqual(notifier.state().waiting, { missed: 0, chat: 0, mail: 1, total: 1 });
  notifier.seenApp('mail');
  assert.equal(notifier.state().waiting.total, 0);

  // Talk says how many are unread: its number stands, a report adds nothing to it, and a look at its window clears nothing.
  notifier.setWaiting('chat', 5);
  notifier.report({ kind: 'chat', from: 'A', text: 't' });
  assert.equal(notifier.state().waiting.chat, 5);
  notifier.seenApp('talk');
  assert.equal(notifier.state().waiting.chat, 5);
  notifier.setWaiting('chat', 0);
  assert.equal(notifier.state().waiting.chat, 0);
  assert.throws(() => notifier.setWaiting('call', 1), /wait/);
  assert.deepEqual(COUNTED, ['missed', 'chat', 'mail']);
});

test('a window of the app in front means the person is looking: nothing is shown or counted', () => {
  const { notifier, shown } = setup();
  assert.deepEqual(notifier.report({ kind: 'chat', from: 'A', text: 't', front: true }), { shown: false, why: 'front' });
  assert.equal(shown.length, 0);
  assert.equal(notifier.state().waiting.total, 0);
});

test('a reminder says when: now, in so many minutes, or at a time; a meeting that can be joined offers Join', () => {
  const now = new Date(2026, 10, 23, 10, 0).getTime();
  const timeOf = () => '14:30';
  assert.equal(wordsOf({ kind: 'reminder', summary: 'Planning', start: now, location: 'Room 2' }, { now, timeOf }).body, 'Starting now · Room 2');
  assert.equal(wordsOf({ kind: 'reminder', summary: 'Planning', start: now + 15 * 60_000 }, { now, timeOf }).body, 'In 15 minutes');
  assert.equal(wordsOf({ kind: 'reminder', summary: 'Planning', start: now + 60_000 }, { now, timeOf }).body, 'In 1 minute');
  assert.equal(wordsOf({ kind: 'reminder', summary: 'Planning', start: now + 5 * 3600_000 }, { now, timeOf }).body, 'At 14:30');
  assert.equal(wordsOf({ kind: 'reminder', summary: '', start: now }, { now, timeOf }).title, 'Event');
  const meeting = wordsOf({ kind: 'meeting', summary: 'Weekly', start: now + 5 * 60_000, join: true }, { now, timeOf });
  assert.deepEqual(meeting.actions.map((a) => [a.id, a.label]), [['join', 'Join']]);
  assert.deepEqual(wordsOf({ kind: 'meeting', summary: 'Weekly', start: now }, { now, timeOf }).actions, []);
});

test('the settings change part by part; starting with Windows needs Office kept running; a pause is not a setting', () => {
  const { notifier, kept } = setup();
  let s = notifier.setSettings({ kinds: { mail: false }, quiet: { on: true, to: '06:30' }, startup: true });
  assert.equal(s.kinds.mail, false);
  assert.equal(s.kinds.chat, true);
  assert.deepEqual(s.quiet, { on: true, from: '22:00', to: '06:30' });
  assert.equal(s.startup, true);
  s = notifier.setSettings({ keepRunning: false });
  assert.equal(s.keepRunning, false);
  assert.equal(s.startup, false, 'started with Windows, it would have nowhere to be');
  notifier.pause();
  const until = notifier.settings().pausedUntil;
  assert.ok(until > 0);
  assert.equal(notifier.setSettings({ pausedUntil: 0, sound: false }).pausedUntil, until, 'only pause and resume change the pause');
  assert.equal(kept.get(SETTINGS_KEY).sound, false);
});

test('the day events for the menu are kept, past ones left out of the state, and a change is told once', () => {
  const { notifier, clock } = setup();
  let told = 0;
  notifier.onChange(() => { told += 1; });
  const list = [
    { id: 'a', summary: 'Stand-up', start: clock.at - 3600_000, end: clock.at - 1800_000 },
    { id: 'b', summary: 'Planning', start: clock.at + 3600_000, end: clock.at + 7200_000 },
  ];
  notifier.setEvents(list);
  notifier.setEvents(list);
  assert.equal(told, 1);
  assert.deepEqual(notifier.state().events.map((e) => e.id), ['b']);
});

test('Office says its own word once asked, a test notification can be shown, and closing takes everything down', () => {
  const { notifier, shown, opened } = setup();
  const r = notifier.tell({ id: 'office:still-running', title: 'Rutba Office is still running', body: 'By the clock.', open: { app: 'home' } });
  assert.equal(r.shown, true);
  assert.equal(shown[0].silent, true);
  shown[0].onClick();
  assert.deepEqual(opened, [{ app: 'home' }]);
  notifier.test();
  assert.equal(shown[1].title, 'Notifications are on');
  notifier.close();
  assert.ok(shown.every((n) => n.closed));
});

test('a report with no kind is refused, and a show that throws does not stop the count', () => {
  const kept = new Map();
  const notifier = createNotifier({ settings: { get: (k, f) => kept.get(k) ?? f, set: (k, v) => kept.set(k, v) }, show: () => { throw new Error('no notifications here'); } });
  assert.throws(() => notifier.report({ from: 'A' }), /kind/);
  assert.equal(notifier.report({ kind: 'mail', from: 'A', subject: 's' }).shown, true);
  assert.equal(notifier.state().waiting.mail, 1);
  notifier.close();
});

test('a reminder for another day names the day; a whole day says Today or Tomorrow; an event already on says when it started', () => {
  const now = new Date(2026, 10, 23, 10, 0).getTime();
  const opts = { now, timeOf: () => '09:00', dayOf: () => 'Friday 27 November' };
  const body = (note) => wordsOf({ kind: 'reminder', summary: 'Planning', ...note }, opts).body;
  assert.equal(body({ start: new Date(2026, 10, 24, 9, 0).getTime() }), 'Tomorrow at 09:00');
  assert.equal(body({ start: new Date(2026, 10, 27, 9, 0).getTime() }), 'Friday 27 November at 09:00');
  assert.equal(body({ start: new Date(2026, 10, 23).getTime(), allDay: true }), 'Today');
  assert.equal(body({ start: new Date(2026, 10, 24).getTime(), allDay: true, location: 'Karachi' }), 'Tomorrow · Karachi');
  assert.equal(body({ start: new Date(2026, 10, 27).getTime(), allDay: true }), 'Friday 27 November');
  assert.equal(body({ start: now - 30 * 60_000 }), 'Started at 09:00');
  assert.equal(body({ start: now - 60_000 }), 'Starting now', 'a minute either way is now');
  // The day is the notifier's to word when nothing hands one in.
  assert.match(wordsOf({ kind: 'reminder', summary: 'x', start: new Date(2026, 10, 27).getTime(), allDay: true }, { now }).body, /27/);
});

test('a meeting someone asks the person into says who asks and offers Join; several emails name who they are from', () => {
  const asked = wordsOf({ kind: 'meeting', summary: 'Launch', by: 'Ben Ortiz', join: true });
  assert.deepEqual([asked.title, asked.body, asked.actions.map((a) => a.id)], ['Launch', 'Ben Ortiz asks you to join', ['join']]);
  assert.equal(wordsOf({ kind: 'meeting', summary: '', join: true }).body, '');
  assert.equal(wordsOf({ kind: 'mail', count: 3, from: 'Amina Yusuf, Bilal Khan', account: 'sam@example.org' }).body, 'Amina Yusuf, Bilal Khan');
  assert.equal(wordsOf({ kind: 'mail', count: 3, from: '', account: 'sam@example.org' }).body, 'sam@example.org');
});

test('the day events for the menu carry the instant Calendar knows each by', () => {
  const { notifier, clock } = setup();
  notifier.setEvents([{ id: 'd', summary: 'Holiday', start: clock.at - 3600_000, end: clock.at + 3600_000, allDay: true, at: 1234 }, { id: 'e', summary: 'Call', start: clock.at + 60_000, end: clock.at + 120_000 }]);
  assert.deepEqual(notifier.state().events.map((e) => [e.id, e.at]), [['d', 1234], ['e', null]]);
});

// B18: Windows shows nothing when its own switch is off, and until 2.28.0 Office said nothing of it.
test('a notification Windows did not show is remembered, and told once, until one is shown on purpose or pressed', () => {
  const { notifier, shown } = setup();
  let told = 0;
  notifier.onChange(() => { told += 1; });
  assert.equal(notifier.state().refused, false);

  notifier.test();
  assert.equal(shown.at(-1).id, 'office:test');
  shown.at(-1).onFail('notifications are off');
  assert.equal(notifier.state().refused, true);
  assert.equal(told, 1);
  shown.at(-1).onFail('again');
  assert.equal(told, 1, 'said once');

  // Another notification does not take the answer back: Windows has not shown it yet.
  notifier.report({ kind: 'chat', id: 'chat:u1', from: 'Amina', text: 'hello', app: 'talk' });
  assert.equal(notifier.state().refused, true);

  // Shown on purpose again: Windows is asked afresh.
  const before = told;
  notifier.test();
  assert.equal(notifier.state().refused, false);
  assert.ok(told > before, 'the settings window is told the answer no longer stands');
  shown.at(-1).onFail('still off');
  assert.equal(notifier.state().refused, true);

  // Pressed: Windows showed it after all.
  notifier.report({ kind: 'chat', id: 'chat:u2', from: 'Bilal', text: 'there?', app: 'talk' });
  shown.at(-1).onClick();
  assert.equal(notifier.state().refused, false);
});

test('a notification that is shown says nothing of Windows', () => {
  const { notifier, shown } = setup();
  notifier.test();
  shown.at(-1).onClose();
  assert.equal(notifier.state().refused, false);
});
