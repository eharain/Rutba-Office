// The ring for a call that comes while no Talk window is open
// (office-shell call-ring.js, notifier.js; apps/desktop/main/watch-talk.js,
// watch.js). With no window to ring in there was only the Windows
// notification's one short sound; now the notifier rings until the call is
// answered, turned away or stops ringing, or a Talk window opens and rings
// for itself. Never both. Nothing sounds here: the sound device and the
// notification are stand-ins.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { RING, RING_LIMIT_MS, RING_PAGE, ringScript } from '../packages/office-shell/src/electron/call-ring.js';
import { createNotifier, SETTINGS_KEY } from '../packages/office-shell/src/electron/notifier.js';
import { watchTalk, NOTIFY_EVENT } from '../apps/desktop/main/watch-talk.js';
import { createWatch } from '../apps/desktop/main/watch.js';
import { SOUNDS } from '../apps/desktop/renderer/apps/talk/ring.js';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const AMINA = { id: 'u1', name: 'Amina Yusuf' };

/** A notifier whose notification and ring are kept, not made. */
function setup({ stored = undefined } = {}) {
  const kept = new Map();
  if (stored !== undefined) kept.set(SETTINGS_KEY, stored);
  const shown = [];
  const rings = [];
  const opened = [];
  const notifier = createNotifier({
    settings: { get: (k, f) => (kept.has(k) ? structuredClone(kept.get(k)) : f), set: (k, v) => kept.set(k, structuredClone(v)) },
    show: (note) => {
      const entry = { ...note, closed: false };
      shown.push(entry);
      return { close: () => { entry.closed = true; } };
    },
    ring: () => {
      const r = { stopped: 0 };
      rings.push(r);
      return () => { r.stopped += 1; };
    },
    open: (target) => opened.push(target),
    now: () => new Date(2026, 10, 23, 10, 0).getTime(),
    setTimer: () => ({}),
    clearTimer: () => {},
  });
  const call = (id = 'c1', more = {}) => notifier.report({ kind: 'call', id: `call:${id}`, from: 'Amina Yusuf', app: 'talk', open: { app: 'talk', query: { contact: 'u1' } }, ring: true, ...more });
  return { notifier, shown, rings, opened, call };
}

test('the ring is Talk\'s own, so a call sounds the same with a Talk window and without one', () => {
  assert.deepEqual(RING, SOUNDS.ring);
});

test('what the ring\'s window runs lays the phrases out ahead on the sound device\'s clock, for as long as a ring may last', () => {
  const made = [];
  class Device {
    constructor() { this.currentTime = 10; this.destination = {}; this.state = 'running'; this.resumed = 0; made.push(this); this.notes = []; }
    resume() { this.resumed += 1; return Promise.resolve(); }
    createOscillator() {
      const note = { frequency: {}, starts: null, stops: null };
      this.notes.push(note);
      return { frequency: note.frequency, connect() {}, start: (at) => { note.starts = at; }, stop: (at) => { note.stops = at; }, set type(v) { note.type = v; } };
    }
    createGain() {
      const levels = [];
      return { gain: { setValueAtTime: (v, at) => levels.push([v, at]), linearRampToValueAtTime: (v, at) => levels.push([v, at]) }, connect() {}, levels };
    }
  }
  const page = {};
  const run = (src) => new Function('AudioContext', 'window', `return ${src};`)(Device, page);
  const told = run(ringScript());
  const phrases = Math.ceil(RING_LIMIT_MS / RING.every);
  assert.deepEqual(told, { state: 'running', phrases, notes: phrases * RING.notes.length });
  const [device] = made;
  assert.equal(device.resumed, 1, 'a device that holds its sound back is asked for it');
  assert.equal(page.__ring, device, 'the device is kept, so the sound lasts as long as the window');
  // The first phrase: the ring's own notes, in order, from just after now.
  const first = device.notes.slice(0, RING.notes.length);
  assert.deepEqual(first.map((n) => n.frequency.value), RING.notes.map(([, , hz]) => hz));
  assert.ok(first.every((n) => n.type === 'sine'));
  for (const [i, [at, lasts]] of RING.notes.entries()) {
    assert.ok(Math.abs(first[i].starts - (10.05 + at)) < 1e-9, `note ${i} starts`);
    assert.ok(first[i].stops > first[i].starts + lasts, `note ${i} stops after it has sounded`);
  }
  // Each phrase comes `every` after the one before, to the last.
  const starts = (k) => device.notes[k * RING.notes.length].starts;
  assert.ok(Math.abs(starts(1) - starts(0) - RING.every / 1000) < 1e-9);
  assert.ok(Math.abs(starts(phrases - 1) - starts(0) - ((phrases - 1) * RING.every) / 1000) < 1e-9);
  assert.ok(starts(phrases - 1) - starts(0) < RING_LIMIT_MS / 1000, 'nothing is laid out past the limit');
  // A shorter ring, asked for, is shorter; there is always one phrase.
  assert.equal(run(ringScript({ ms: 7000 })).phrases, 3);
  assert.equal(run(ringScript({ ms: 1 })).phrases, 1);
  assert.match(RING_PAGE, /^<!doctype html>/);
  assert.ok(!/<script|src=|href=/.test(RING_PAGE), 'the page holds nothing: no script of its own, nothing fetched');
});

test('a call no window rings for is rung for as long as its notification is up, and Windows\' short sound is not played over it', () => {
  const { notifier, shown, rings, call } = setup();
  assert.equal(call().shown, true);
  assert.equal(rings.length, 1, 'the ring starts with the notification');
  assert.equal(shown[0].silent, true, 'the ring is its sound');
  assert.equal(shown[0].sticky, true);
  assert.equal(rings[0].stopped, 0, 'and goes on');
  // Talk says the call no longer rings: the notification comes down and the ring with it.
  assert.equal(notifier.withdraw('call:c1'), true);
  assert.equal(shown[0].closed, true);
  assert.equal(rings[0].stopped, 1);
  assert.equal(notifier.withdraw('call:c1'), false);
  assert.equal(rings[0].stopped, 1, 'stopped once');
});

test('pressing the notification or either of its buttons stops the ring; so does putting it away', () => {
  for (const press of [(n) => n.onClick(), (n) => n.onAction('answer'), (n) => n.onAction('decline'), (n) => n.onClose()]) {
    const { shown, rings, call, notifier } = setup();
    call();
    press(shown[0]);
    assert.equal(rings[0].stopped, 1);
    // What comes after stops nothing twice.
    shown[0].onClose();
    notifier.withdraw('call:c1');
    assert.equal(rings[0].stopped, 1);
  }
});

test('a Talk window that opens rings for itself: the ring made here stops and the notification stays', () => {
  const { notifier, shown, rings, call } = setup();
  call('c1');
  assert.equal(notifier.quietApp('mail'), 0, 'another app\'s window is nothing to a call');
  assert.equal(rings[0].stopped, 0);
  assert.equal(notifier.quietApp('talk'), 1);
  assert.equal(rings[0].stopped, 1);
  assert.equal(shown[0].closed, false, 'the call still shows by the clock');
  assert.equal(notifier.quietApp('talk'), 0, 'nothing is left to stop');
  assert.equal(notifier.withdraw('call:c1'), true);
  assert.equal(rings[0].stopped, 1);
});

test('no ring is made when a window rings, when sounds are off, when the call is not shown, or for anything but a call', () => {
  // A Talk window is open: the report asks for no ring, and the notification keeps its own sound as before.
  const open = setup();
  open.call('c1', { ring: false });
  assert.equal(open.rings.length, 0);
  assert.equal(open.shown[0].silent, false);
  // Sound off in the notification settings: silent, and no ring.
  const off = setup({ stored: { sound: false } });
  off.call();
  assert.equal(off.rings.length, 0);
  assert.equal(off.shown[0].silent, true);
  // Calls turned off, paused, or in the quiet hours: nothing is shown and nothing rings.
  const kindOff = setup({ stored: { kinds: { call: false } } });
  assert.equal(kindOff.call().shown, false);
  assert.equal(kindOff.rings.length, 0);
  const paused = setup();
  paused.notifier.pause();
  assert.equal(paused.call().shown, false);
  assert.equal(paused.rings.length, 0);
  // A message that asks to ring is not a call.
  const chat = setup();
  chat.notifier.report({ kind: 'chat', id: 'chat:u1', from: 'Amina Yusuf', text: 'Hello', app: 'talk', ring: true });
  assert.equal(chat.rings.length, 0);
  assert.equal(chat.shown[0].silent, false);
});

test('the same call reported again rings once: the earlier ring stops as its notification comes down', () => {
  const { rings, shown, call } = setup();
  call();
  call();
  assert.equal(rings.length, 2);
  assert.equal(rings[0].stopped, 1);
  assert.equal(rings[1].stopped, 0);
  assert.equal(shown[0].closed, true);
});

test('a ring that cannot be made, or a notification that cannot be shown, is no fault to anyone', () => {
  const kept = new Map();
  const settings = { get: (k, f) => (kept.has(k) ? kept.get(k) : f), set: (k, v) => kept.set(k, v) };
  const shown = [];
  const broken = createNotifier({ settings, show: (n) => { shown.push(n); return { close() {} }; }, ring: () => { throw new Error('no sound device'); }, setTimer: () => ({}), clearTimer: () => {} });
  assert.equal(broken.report({ kind: 'call', id: 'call:c1', from: 'Amina', ring: true }).shown, true);
  assert.equal(broken.withdraw('call:c1'), true);
  // Windows shows nothing (notifications off for Office): nothing rings either.
  let rang = 0;
  const unseen = createNotifier({ settings, show: () => null, ring: () => { rang += 1; return () => {}; }, setTimer: () => ({}), clearTimer: () => {} });
  unseen.report({ kind: 'call', id: 'call:c1', from: 'Amina', ring: true });
  assert.equal(rang, 0);
  // A notifier handed no ring (a test, an older host) shows the call as it always did.
  const plain = createNotifier({ settings, show: (n) => { shown.push(n); return { close() {} }; }, setTimer: () => ({}), clearTimer: () => {} });
  assert.equal(plain.report({ kind: 'call', id: 'call:c2', from: 'Amina', ring: true }).shown, true);
});

/** Talk's watch over a notifier that keeps what it was reported. */
function talkSetup({ windowOpen = false, state = undefined, open = undefined } = {}) {
  const reported = [];
  const ctx = { windowOpen, state };
  const listeners = new Map();
  const on = (event, fn) => {
    if (!listeners.has(event)) listeners.set(event, new Set());
    listeners.get(event).add(fn);
    return () => listeners.get(event).delete(fn);
  };
  const emit = (event, payload) => { for (const fn of [...(listeners.get(event) || [])]) fn(payload); };
  const notifier = { report: (note) => { reported.push(note); return { shown: true, id: note.id }; }, withdraw: () => true, setWaiting: () => {} };
  watchTalk({ notifier, on, ...(open === null ? {} : { open: (app) => app === 'talk' && ctx.windowOpen }), ...(state === undefined ? {} : { state: () => ctx.state }) });
  return { reported, ctx, emit, tell: (n) => emit(NOTIFY_EVENT, n) };
}

test('Talk\'s watch asks for the ring only when no Talk window is open and the person has Talk\'s sounds on', () => {
  const none = talkSetup({ windowOpen: false });
  none.tell({ kind: 'call', id: 'c1', from: AMINA, video: false });
  assert.equal(none.reported[0].ring, true, 'no window to ring in: the notifier rings');
  assert.equal(none.reported[0].app, 'talk', 'so a Talk window that opens stops it');

  const open = talkSetup({ windowOpen: true });
  open.tell({ kind: 'call', id: 'c1', from: AMINA, video: false });
  assert.equal(open.reported[0].ring, false, 'the window rings: never both');

  // "Play a sound for calls" off, as Talk's service says now.
  const off = talkSetup({ windowOpen: false, state: { phase: 'online', ring: false } });
  off.tell({ kind: 'call', id: 'c1', from: AMINA, video: false });
  assert.equal(off.reported[0].ring, false);
  off.ctx.state = { phase: 'online', ring: true };
  off.tell({ kind: 'call', id: 'c2', from: AMINA, video: false });
  assert.equal(off.reported[1].ring, true, 'turned on again, the next call rings');

  // With no service to ask, what Talk last told its windows is used.
  const told = talkSetup({ windowOpen: false });
  told.emit('talk:state', { phase: 'online', ring: false, unreadTotal: 0, missed: 0 });
  told.tell({ kind: 'call', id: 'c1', from: AMINA, video: false });
  assert.equal(told.reported[0].ring, false);

  // A host that cannot say whether a window is open: Talk's window is taken to ring, as before.
  const old = talkSetup({ open: null });
  old.tell({ kind: 'call', id: 'c1', from: AMINA, video: false });
  assert.equal(old.reported[0].ring, false);

  // Only a call is rung for.
  none.tell({ kind: 'missed', id: 'c9', from: AMINA, video: false, count: 1 });
  none.tell({ kind: 'message', id: 'm1', from: AMINA, text: 'Hello' });
  assert.ok(none.reported.slice(1).every((r) => !r.ring));
});

/** The whole watch, as the application wires it, with stand-ins for its Electron parts. */
function watchSetup({ windows: wins = [] } = {}) {
  const kept = new Map();
  const stores = { settings: { get: (k, f) => (kept.has(k) ? structuredClone(kept.get(k)) : f), set: (k, v) => kept.set(k, structuredClone(v)) } };
  const shown = [];
  const rings = [];
  const parts = {
    show: (note) => { const e = { ...note, closed: false }; shown.push(e); return { close: () => { e.closed = true; } }; },
    flash: () => {},
    ring: () => { const r = { stopped: 0 }; rings.push(r); return () => { r.stopped += 1; }; },
    supported: () => true,
    makeTray: () => ({ setImage() {}, setToolTip() {}, setContextMenu() {}, on() {}, destroy() {} }),
    buildMenu: (rows) => rows,
    icons: () => ({ plain: 'plain', waiting: 'waiting' }),
    startup: { can: true, get: () => false, set: () => {} },
  };
  const windows = { all: () => wins, infoFor: (w) => ({ app: w.app }), findByApp: (app) => wins.find((w) => w.app === app) || null, open: () => null };
  const watch = createWatch({ stores, parts, windows, sendEvent: () => {}, broadcast: () => {}, apps: () => ['talk', 'mail', 'calendar'], quit: () => {}, now: () => new Date(2026, 10, 23, 10, 0).getTime() });
  watch.watchTalk({ send: () => {}, state: () => ({ phase: 'online', ring: true }) });
  watch.start();
  const ring = (id) => watch.hear('talk:notify', { kind: 'call', id, from: AMINA, video: false });
  const over = (id) => watch.hear('talk:notify', { kind: 'call-over', id, why: 'ended' });
  return { watch, shown, rings, wins, ring, over };
}

const fakeWindow = (app, { visible = true, focused = false } = {}) => ({ app, isDestroyed: () => false, isFocused: () => focused, isVisible: () => visible, isMinimized: () => false, webContents: { getURL: () => '' } });

test('in the application: a call with no Talk window rings by the clock until Talk says it is over', () => {
  const s = watchSetup();
  s.ring('c1');
  assert.equal(s.rings.length, 1);
  assert.equal(s.shown[0].kind, 'call');
  assert.equal(s.shown[0].silent, true);
  assert.equal(s.rings[0].stopped, 0);
  s.over('c1');
  assert.equal(s.rings[0].stopped, 1);
  assert.equal(s.shown[0].closed, true);
});

test('in the application: with a Talk window open, shown or hidden, the window rings and the notifier does not', () => {
  for (const visible of [true, false]) {
    const s = watchSetup({ windows: [fakeWindow('talk', { visible })] });
    s.ring('c1');
    assert.equal(s.rings.length, 0, visible ? 'shown' : 'hidden');
    assert.equal(s.shown[0].silent, false, 'the notification is as it was');
  }
  // A window of another app is not one that rings.
  const other = watchSetup({ windows: [fakeWindow('mail')] });
  other.ring('c1');
  assert.equal(other.rings.length, 1);
});

test('in the application: a Talk window coming to the front while the notifier rings takes the ring over', () => {
  const s = watchSetup();
  s.ring('c1');
  assert.equal(s.rings[0].stopped, 0);
  s.watch.focused('mail');
  assert.equal(s.rings[0].stopped, 0, 'Mail coming forward is nothing to the call');
  s.wins.push(fakeWindow('talk', { focused: true }));
  s.watch.focused('talk');
  assert.equal(s.rings[0].stopped, 1, 'the window rings now');
  assert.equal(s.shown[0].closed, false);
  s.over('c1');
  assert.equal(s.rings[0].stopped, 1);
});

test('Electron\'s side makes the ring in a window that is never shown, and makes none in a check run', () => {
  const src = fs.readFileSync(path.join(root, 'packages/office-shell/src/electron/notify-electron.js'), 'utf8');
  const at = src.indexOf('const ring = () => {');
  assert.ok(at > 0);
  const body = src.slice(at, src.indexOf('/** The taskbar button', at));
  assert.ok(body.indexOf('if (check)') < body.indexOf('new BrowserWindow'), 'a check run answers before any window is made');
  assert.match(body, /show: false/);
  assert.match(body, /focusable: false/);
  assert.match(body, /skipTaskbar: true/);
  assert.match(body, /sandbox: true/);
  assert.match(body, /nodeIntegration: false/);
  assert.ok(!/preload/.test(body), 'the window is handed nothing of Office');
  assert.match(body, /RING_LIMIT_MS/, 'it stops by itself in the end');
  assert.match(body, /win\.destroy\(\)/, 'stopping closes the window, and the sound with it');
  assert.match(src, /\n    ring,\r?\n/, 'the notifier is handed it');
  const watch = fs.readFileSync(path.join(root, 'apps/desktop/main/watch.js'), 'utf8');
  assert.match(watch, /ring: parts\.ring/);
});
