// Keeping watch (apps/desktop/main/watch.js, D-2026-11-22-23): the icon by
// the clock, Office kept running there, started with Windows, and the
// `notify` calls a window makes. Electron's parts are stand-ins: no icon, no
// notification and no window is raised on this computer.

import test from 'node:test';
import assert from 'node:assert/strict';
import { createWatch, TOLD_KEY, TRAY_FLAG, WATCHED_APPS } from '../apps/desktop/main/watch.js';

function fakeWindow(app, { focused = false, url = `app://office/index.html?app=${app}` } = {}) {
  return {
    app,
    webContents: { getURL: () => url },
    did: [],
    minimized: false,
    visible: true,
    focused,
    isDestroyed: () => false,
    isFocused() { return this.focused; },
    isMinimized() { return this.minimized; },
    isVisible() { return this.visible; },
    restore() { this.minimized = false; this.did.push('restore'); },
    show() { this.visible = true; this.did.push('show'); },
    focus() { this.focused = true; this.did.push('focus'); },
  };
}

function setup({ check = false, can = true, apps = ['talk', 'mail', 'calendar'], stored = {}, open = [], trayFails = false } = {}) {
  const kept = new Map(Object.entries(stored));
  const stores = { settings: { get: (k, f) => (kept.has(k) ? structuredClone(kept.get(k)) : f), set: (k, v) => kept.set(k, structuredClone(v)) } };
  const shown = [];
  const trays = [];
  const login = { on: false, sets: [] };
  const parts = {
    show: (note) => { const e = { ...note, closed: false }; shown.push(e); return { close: () => { e.closed = true; } }; },
    flash: () => {},
    supported: () => true,
    makeTray: (icon) => {
      if (trayFails) throw new Error('no place by the clock');
      const t = { icon, tip: '', menu: null, handlers: {}, destroyed: false, setImage(i) { this.icon = i; }, setToolTip(s) { this.tip = s; }, setContextMenu(m) { this.menu = m; }, on(e, fn) { this.handlers[e] = fn; }, destroy() { this.destroyed = true; } };
      trays.push(t);
      return t;
    },
    buildMenu: (rows) => rows,
    icons: () => ({ plain: 'plain', waiting: 'waiting' }),
    startup: { can, get: () => login.on, set: (on) => { login.on = on; login.sets.push(on); } },
  };
  const wins = [...open];
  const openedWith = [];
  const windows = {
    all: () => wins,
    infoFor: (w) => ({ app: w.app }),
    findByApp: (app) => wins.find((w) => w.app === app) || null,
    open: (o) => { openedWith.push(o); const w = fakeWindow(o.app); wins.push(w); return w; },
  };
  const sentTo = [];
  const told = [];
  const quits = [];
  const watch = createWatch({
    stores,
    parts,
    windows,
    sendEvent: (win, event, payload) => sentTo.push([win, event, payload]),
    broadcast: (event, payload) => told.push([event, payload]),
    name: 'Rutba Office',
    apps: () => apps,
    check,
    quit: () => quits.push(1),
    now: () => new Date(2026, 10, 23, 10, 0).getTime(),
  });
  return { watch, kept, shown, trays, login, wins, openedWith, sentTo, told, quits };
}

const row = (rows, id) => rows.find((r) => r.id === id);

test('started, the icon goes by the clock; in a check run there is none and nothing keeps Office running', () => {
  const a = setup();
  assert.equal(a.watch.service.state().tray, false);
  a.watch.start();
  a.watch.start();
  assert.equal(a.trays.length, 1);
  assert.equal(a.watch.service.state().tray, true);

  const b = setup({ check: true });
  b.watch.start();
  assert.equal(b.trays.length, 0);
  assert.equal(b.watch.lastWindowClosed(), false);
  assert.equal(b.watch.startsQuiet(['x', TRAY_FLAG]), false);
});

test('the last window closed: Office stays by the clock, and says so the first time only', () => {
  const { watch, shown, kept } = setup();
  assert.equal(watch.lastWindowClosed(), false, 'before the icon is there, closing the last window quits');
  watch.start();
  assert.equal(watch.lastWindowClosed(), true);
  assert.equal(shown.length, 1);
  assert.equal(shown[0].title, 'Rutba Office is still running');
  assert.equal(kept.get(TOLD_KEY), true);
  assert.equal(watch.lastWindowClosed(), true);
  assert.equal(shown.length, 1);
});

test('with "keep running" off, or no icon to open it from, the last window closed quits Office as before', () => {
  const off = setup({ stored: { notifications: { keepRunning: false } } });
  off.watch.start();
  assert.equal(off.watch.lastWindowClosed(), false);
  assert.equal(off.shown.length, 0);

  const none = setup({ trayFails: true });
  none.watch.start();
  assert.equal(none.watch.service.state().tray, false);
  assert.equal(none.watch.lastWindowClosed(), false, 'kept running with no icon, there would be no way to it');
});

test('a start with Windows goes to the clock with no window, when the setting holds and the icon is there', () => {
  const { watch } = setup();
  assert.equal(watch.startsQuiet([TRAY_FLAG]), false, 'not before the icon');
  watch.start();
  assert.equal(watch.startsQuiet(['office.exe', TRAY_FLAG]), true);
  assert.equal(watch.startsQuiet(['office.exe']), false);
  watch.service.set({ settings: { keepRunning: false } });
  assert.equal(watch.startsQuiet(['office.exe', TRAY_FLAG]), false);
});

test('"start with Windows" is off until turned on, is told to Windows only when it differs, and goes off with "keep running"', () => {
  const { watch, login } = setup();
  watch.start();
  assert.deepEqual(login.sets, [], 'nothing is told to Windows at a start that changes nothing');
  assert.deepEqual(watch.service.state().startup, { can: true, on: false });
  let s = watch.service.set({ settings: { startup: true } });
  assert.deepEqual(s.startup, { can: true, on: true });
  assert.deepEqual(login.sets, [true]);
  watch.service.set({ settings: { sound: false } });
  assert.deepEqual(login.sets, [true]);
  s = watch.service.set({ settings: { keepRunning: false } });
  assert.deepEqual(s.startup, { can: true, on: false });
  assert.deepEqual(login.sets, [true, false]);
});

test('a copy that is not installed never starts with Windows, whatever is asked', () => {
  const { watch, login } = setup({ can: false });
  watch.start();
  const s = watch.service.set({ settings: { startup: true } });
  assert.deepEqual(s.startup, { can: false, on: false });
  assert.equal(s.settings.startup, false);
  assert.deepEqual(login.sets, []);
});

test('a setting kept from before that Windows does not hold is put right at the start', () => {
  const { watch, login } = setup({ stored: { notifications: { startup: true } } });
  watch.start();
  assert.deepEqual(login.sets, [true]);
});

test('an app opened at a thing: its open window is brought forward and told where; else a window opens on it', () => {
  const mail = fakeWindow('mail');
  mail.minimized = true;
  const { watch, sentTo, openedWith } = setup({ open: [mail] });
  assert.equal(watch.goTo({ app: 'mail', query: { show: 'new' } }), mail);
  assert.deepEqual(mail.did, ['restore', 'focus']);
  assert.deepEqual(sentTo, [[mail, 'app:link', { show: 'new' }]], 'told as a link opened for it is');
  watch.goTo({ app: 'mail' });
  assert.equal(sentTo.length, 1, 'with nowhere to go it is only brought forward');
  watch.goTo({ app: 'talk', query: { contact: 'u1' } });
  watch.goTo({ app: 'calendar' });
  assert.deepEqual(openedWith, [{ app: 'talk', query: { contact: 'u1' } }, { app: 'calendar' }]);
});

test('the menu of the icon lists the apps on in this profile, opens the settings in the launcher, and quits', () => {
  const { watch, trays, openedWith, quits } = setup({ apps: ['mail', 'writer', 'calendar'] });
  watch.start();
  const menu = trays[0].menu;
  assert.deepEqual(menu.filter((r) => r.id?.startsWith('app:')).map((r) => r.id), ['app:mail', 'app:calendar']);
  row(menu, 'settings').click();
  assert.deepEqual(openedWith, [{ app: 'home', query: { notifications: 1 } }]);
  row(menu, 'quit').click();
  assert.equal(quits.length, 1);
  assert.deepEqual(WATCHED_APPS, ['talk', 'mail', 'calendar']);
});

test('what a service tells the windows is told them as before, and heard here too', () => {
  const { watch, told } = setup();
  const heard = [];
  const off = watch.on('mail:new', (p) => heard.push(p));
  watch.on('mail:new', () => { throw new Error('one listener at fault'); });
  watch.hear('mail:new', { count: 2 });
  assert.deepEqual(told, [['mail:new', { count: 2 }]]);
  assert.deepEqual(heard, [{ count: 2 }]);
  off();
  watch.hear('mail:new', { count: 1 });
  assert.equal(heard.length, 1);
  assert.equal(told.length, 2);
});

test('Talk is watched by what it tells is worth a notification; its service is left as it is and only turns a call away', async () => {
  const talkWin = fakeWindow('talk');
  const { watch, shown, sentTo, wins, openedWith } = setup();
  const calls = [];
  const send = (msg, win) => { calls.push(['send', msg, win]); return Promise.resolve('sent'); };
  const resume = async (args, win) => { calls.push(['resume', args, win]); return { phase: 'online' }; };
  const service = { send, resume, state: () => ({}) };
  assert.deepEqual(watch.service.state().watched, []);
  watch.watchTalk(service);
  watch.watchTalk(service);
  watch.start();
  assert.deepEqual(watch.service.state().watched, ['call', 'missed', 'chat', 'meeting'], 'a meeting someone asks the person into is told of too');
  assert.equal(service.send, send, 'what a Talk window sends goes to Talk alone');
  assert.equal(service.resume, resume, 'a Talk window that starts is told of a ringing call by Talk itself');

  watch.hear('talk:state', { phase: 'online', unreadTotal: 0, missed: 0 });
  watch.hear('talk:notify', { kind: 'call', id: 'c1', from: { id: 'u1', name: 'Amina' }, video: false });
  assert.deepEqual(shown.map((n) => [n.kind, n.body]), [['call', 'Amina']]);
  // What Talk tells its windows otherwise is not for the notifier.
  watch.hear('talk:event', { type: 'call.incoming', callId: 'c9', from: { id: 'u2' } });
  assert.equal(shown.length, 1);

  // Answer opens Talk at the caller; nothing is sent to the window that Talk does not send itself.
  shown[0].onAction('answer');
  assert.deepEqual(openedWith, [{ app: 'talk', query: { contact: 'u1' } }]);
  wins.push(talkWin);
  assert.deepEqual(sentTo, []);

  // Talk says the call no longer rings: its notification comes down.
  watch.hear('talk:notify', { kind: 'call-over', id: 'c1', why: 'answered' });
  assert.equal(shown[0].closed, true);

  // Declined from the notification, the call is turned away through Talk's own service.
  watch.hear('talk:notify', { kind: 'call', id: 'c2', from: { id: 'u1', name: 'Amina' }, video: false });
  shown[1].onAction('decline');
  assert.deepEqual(calls, [['send', { type: 'call.decline', callId: 'c2' }, undefined]]);
  assert.equal(shown.length, 2);
});

test('a message while a Talk window is in front is not shown; in any other window it is', () => {
  const talkWin = fakeWindow('talk', { focused: true });
  const { watch, shown } = setup({ open: [talkWin] });
  watch.watchTalk({ send: () => {}, resume: async () => ({}) });
  watch.hear('talk:state', { phase: 'online', unreadTotal: 0, missed: 0 });
  watch.hear('talk:notify', { kind: 'message', id: 'm1', from: { id: 'u1', name: 'Amina' }, group: null, text: 'hi', at: 1 });
  assert.equal(shown.length, 0);
  talkWin.focused = false;
  watch.hear('talk:notify', { kind: 'message', id: 'm2', from: { id: 'u1', name: 'Amina' }, group: null, text: 'hi again', at: 2 });
  assert.equal(shown.length, 1);
});

test('where Talk is not on in this profile, or has no service, none of its kinds is offered', () => {
  const off = setup({ apps: ['mail'] });
  off.watch.watchTalk({ send: () => {}, resume: async () => ({}) });
  assert.deepEqual(off.watch.service.state().watched, []);
  // Its service may still be running: what it tells is not passed on.
  off.watch.hear('talk:state', { phase: 'online', unreadTotal: 2, missed: 1 });
  off.watch.hear('talk:notify', { kind: 'message', id: 'm1', from: { id: 'u1', name: 'Amina' }, group: null, text: 'hi', at: 1 });
  assert.deepEqual(off.shown, []);
  assert.equal(off.watch.service.state().waiting.chat, 0);
  assert.equal(off.watch.service.state().waiting.missed, 0);
  const none = setup();
  none.watch.watchTalk(null);
  assert.deepEqual(none.watch.service.state().watched, []);
});

test('a window coming to the front has seen what its app held; the windows are told when the state changes', () => {
  const { watch, told } = setup();
  watch.start();
  watch.notifier.report({ kind: 'mail', from: 'A', subject: 's' });
  const states = told.filter(([e]) => e === 'notify:state');
  assert.equal(states.length, 1);
  assert.equal(states[0][1].waiting.mail, 1);
  assert.equal(states[0][1].tray, true);
  watch.focused('mail');
  assert.equal(watch.service.state().waiting.mail, 0);
  watch.focused('writer');
});

test('the calls a window makes: pause for an hour or as long as asked, resume, and one to see', () => {
  const { watch, shown } = setup();
  watch.start();
  const at = new Date(2026, 10, 23, 10, 0).getTime();
  assert.equal(watch.service.pause().pausedUntil, at + 3600_000);
  assert.equal(watch.service.pause({ minutes: 15 }).pausedUntil, at + 900_000);
  assert.equal(watch.service.state().paused, true);
  assert.equal(watch.service.resume().paused, false);
  watch.service.test();
  assert.equal(shown.length, 1);
  assert.deepEqual(Object.keys(watch.service).sort(), ['pause', 'resume', 'set', 'state', 'test']);
  assert.doesNotThrow(() => watch.service.set());
});

test('stopped, the icon leaves the clock and what was shown comes down', () => {
  const { watch, trays, shown, told } = setup();
  watch.start();
  watch.notifier.report({ kind: 'mail', from: 'A', subject: 's' });
  const before = told.length;
  watch.stop();
  assert.equal(trays[0].destroyed, true);
  assert.equal(shown[0].closed, true);
  assert.equal(told.length, before, 'the windows are told nothing more');
});

test('a window that holds one thing by itself is not where an app is opened at; "fresh" always opens a window of its own', () => {
  const one = fakeWindow('mail', { url: 'app://office/index.html?app=mail&message=%7B%22id%22%3A4%7D' });
  const writing = fakeWindow('mail', { url: 'app://office/index.html?app=mail&compose=k1' });
  const a = setup({ open: [one, writing] });
  a.watch.goTo({ app: 'mail', query: { account: 'a1' } });
  assert.deepEqual(a.openedWith, [{ app: 'mail', query: { account: 'a1' } }], 'no main window of Mail is open: one opens');
  assert.deepEqual(a.sentTo, []);

  const main = fakeWindow('mail');
  const b = setup({ open: [one, main] });
  assert.equal(b.watch.goTo({ app: 'mail', query: { account: 'a1' } }), main, 'the main window, though the other was used last');
  assert.deepEqual(b.sentTo, [[main, 'app:link', { account: 'a1' }]]);
  const message = JSON.stringify({ accountId: 'a1', folder: 'INBOX', id: 7 });
  b.watch.goTo({ app: 'mail', fresh: true, query: { message } });
  assert.deepEqual(b.openedWith, [{ app: 'mail', query: { message } }]);
  assert.equal(b.sentTo.length, 1);
});

function mailService() {
  const listeners = new Set();
  return { listeners, onArrived: (fn) => { listeners.add(fn); return () => listeners.delete(fn); }, accounts: () => [{ id: 'a1', email: 'sam@example.org' }], arrive: (a) => { for (const fn of [...listeners]) fn(a); } };
}
const NOW = new Date(2026, 10, 23, 10, 0).getTime();
const newMail = (id, name, subject) => ({ accountId: 'a1', folder: 'INBOX', role: 'inbox', ids: [id], count: 1, at: new Date(NOW).toISOString(), filedByRules: false, invitations: [], messages: [{ id, subject, from: { name, address: 'x@example.org' }, unread: true, date: new Date(NOW - 60_000).toISOString() }] });

test('Mail is watched through its own service: new mail is shown, counted, and pressed opens the message in a window of its own', () => {
  const { watch, shown, openedWith } = setup();
  const mail = mailService();
  assert.ok(!watch.service.state().watched.includes('mail'));
  watch.watchMail(mail);
  watch.watchMail(mail);
  assert.equal(mail.listeners.size, 1);
  assert.deepEqual(watch.service.state().watched, ['mail']);
  mail.arrive(newMail(7, 'Amina Yusuf', 'Quarterly numbers'));
  assert.deepEqual(shown.map((n) => [n.kind, n.title, n.body]), [['mail', 'Amina Yusuf', 'Quarterly numbers']]);
  assert.equal(watch.service.state().waiting.mail, 1);
  shown[0].onClick();
  assert.deepEqual(openedWith, [{ app: 'mail', query: { message: JSON.stringify({ accountId: 'a1', folder: 'INBOX', id: 7 }) } }]);
  // A Mail window brought to the front has seen it.
  watch.focused('mail');
  assert.equal(watch.service.state().waiting.mail, 0);
  watch.stop();
  assert.equal(mail.listeners.size, 0);
});

test('with a Mail window in front new mail is not shown; where Mail is off in this profile nothing of it is told', () => {
  const mailWin = fakeWindow('mail', { focused: true });
  const a = setup({ open: [mailWin] });
  const mail = mailService();
  a.watch.watchMail(mail);
  mail.arrive(newMail(8, 'Amina Yusuf', 'Seen as it comes'));
  assert.deepEqual(a.shown, []);
  assert.equal(a.watch.service.state().waiting.mail, 0);

  const off = setup({ apps: ['talk', 'calendar'] });
  const other = mailService();
  off.watch.watchMail(other);
  assert.ok(!off.watch.service.state().watched.includes('mail'));
  other.arrive(newMail(9, 'Amina Yusuf', 'Unasked'));
  assert.deepEqual(off.shown, []);
  assert.equal(off.watch.service.state().waiting.mail, 0);
  const none = setup();
  none.watch.watchMail(null);
  assert.deepEqual(none.watch.service.state().watched, []);
});

test('Calendar is watched through its own service: reminders and meetings are offered, each kind in the order the settings list them', () => {
  const { watch, kept } = setup();
  const reads = [];
  const calendar = { events: (range) => { reads.push(range); return []; }, get: () => null };
  watch.watchCalendar(calendar);
  watch.watchCalendar(calendar);
  assert.deepEqual(watch.service.state().watched, ['reminder', 'meeting']);
  watch.watchTalk({ send: () => {}, resume: async () => ({}) });
  watch.watchMail(mailService());
  assert.deepEqual(watch.service.state().watched, ['call', 'missed', 'chat', 'mail', 'reminder', 'meeting']);
  assert.deepEqual(reads, [], 'read a moment after the start, not during it');
  assert.equal(kept.has('notifications.calendarTold'), false);
  watch.stop();
  const none = setup({ apps: ['mail'] });
  none.watch.watchCalendar(calendar);
  assert.deepEqual(none.watch.service.state().watched, []);
  none.watch.stop();
});
