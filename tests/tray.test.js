// The icon by the clock (office-shell tray.js, D-2026-11-22-23): its menu,
// its tip and its mark. The icon and the menu are stand-ins here: nothing
// is put by the clock on this computer.

import test from 'node:test';
import assert from 'node:assert/strict';
import { createTray, menuRows, waitingWords, paintMark, stillRunningWords, EVENTS_MOST } from '../packages/office-shell/src/electron/tray.js';
import { createNotifier } from '../packages/office-shell/src/electron/notifier.js';

const AT = new Date(2026, 10, 23, 10, 0).getTime();
const timeOf = (ms) => `${String(new Date(ms).getHours()).padStart(2, '0')}:${String(new Date(ms).getMinutes()).padStart(2, '0')}`;

function setup({ apps = ['talk', 'mail', 'calendar'], update = () => null, menuOnClick = undefined } = {}) {
  const kept = new Map();
  const clock = { at: AT };
  const notifier = createNotifier({ settings: { get: (k, f) => (kept.has(k) ? kept.get(k) : f), set: (k, v) => kept.set(k, v) }, now: () => clock.at, setTimer: () => ({}), clearTimer: () => {} });
  const made = [];
  const did = [];
  const tray = createTray({
    make: (icon) => {
      const t = { icon, tip: '', menu: null, handlers: {}, popped: 0, popUpContextMenu() { this.popped += 1; }, destroyed: false, setImage(i) { this.icon = i; }, setToolTip(s) { this.tip = s; }, setContextMenu(m) { this.menu = m; }, on(e, fn) { this.handlers[e] = fn; }, destroy() { this.destroyed = true; } };
      made.push(t);
      return t;
    },
    buildMenu: (rows) => rows,
    icons: { plain: 'plain', waiting: 'waiting' },
    notifier,
    name: 'Rutba Office',
    apps: () => apps,
    timeOf,
    menuOnClick,
    does: {
      open: (target) => did.push(['open', target]),
      pause: () => { did.push(['pause']); notifier.pause(); },
      resume: () => { did.push(['resume']); notifier.resume(); },
      settings: () => did.push(['settings']),
      quit: () => did.push(['quit']),
      update,
    },
  });
  return { tray, notifier, made, did, clock };
}

const ids = (rows) => rows.map((r) => (r.type === 'separator' ? '-' : r.id));
const row = (rows, id) => rows.find((r) => r.id === id);

test('with nothing waiting the menu says so, opens Office and its apps, pauses, opens the settings and quits', () => {
  const { tray, made, did } = setup();
  assert.equal(tray.show(), true);
  assert.equal(made.length, 1);
  const t = made[0];
  assert.equal(t.icon, 'plain');
  assert.equal(t.tip, 'Rutba Office');
  assert.deepEqual(ids(t.menu), ['nothing', '-', 'home', 'app:talk', 'app:mail', 'app:calendar', '-', 'pause', 'settings', '-', 'quit']);
  assert.equal(row(t.menu, 'nothing').enabled, false);
  assert.deepEqual(['home', 'app:talk', 'app:mail', 'app:calendar', 'pause', 'settings', 'quit'].map((id) => row(t.menu, id).label), ['Show Rutba Office', 'Talk', 'Mail', 'Calendar', 'Pause notifications for an hour', 'Notification settings…', 'Quit Rutba Office']);
  row(t.menu, 'app:mail').click();
  row(t.menu, 'settings').click();
  row(t.menu, 'quit').click();
  t.handlers.click();
  assert.deepEqual(did, [['open', { app: 'mail' }], ['settings'], ['quit'], ['open', { app: 'home' }]]);
});

test('what waits is in the menu and the tip, each row opening its app there, and the icon wears its mark', () => {
  const { tray, notifier, made, did } = setup();
  tray.show();
  notifier.report({ kind: 'missed', from: 'A' });
  notifier.report({ kind: 'chat', from: 'A', text: 't', count: 3 });
  notifier.report({ kind: 'mail', from: 'B', subject: 's' });
  const t = made[0];
  assert.equal(t.icon, 'waiting');
  assert.equal(t.tip, 'Rutba Office: 1 missed call, 3 unread messages, 1 new email');
  assert.deepEqual(ids(t.menu).slice(0, 4), ['missed', 'chat', 'mail', '-']);
  assert.deepEqual(['missed', 'chat', 'mail'].map((id) => row(t.menu, id).label), ['1 missed call', '3 unread messages', '1 new email']);
  row(t.menu, 'missed').click();
  row(t.menu, 'chat').click();
  row(t.menu, 'mail').click();
  assert.deepEqual(did, [['open', { app: 'talk', query: { show: 'calls' } }], ['open', { app: 'talk', query: { show: 'unread' } }], ['open', { app: 'mail', query: { show: 'new' } }]]);
  notifier.seenApp('talk');
  notifier.seenApp('mail');
  assert.equal(t.icon, 'plain');
  assert.equal(t.tip, 'Rutba Office');
});

test('the day events still to come are listed under Today, a few of them, each opening Calendar at it', () => {
  const { tray, notifier, made, did, clock } = setup();
  tray.show();
  const events = Array.from({ length: EVENTS_MOST + 2 }, (_, i) => ({ id: `e${i}`, summary: `Event ${i}`, start: clock.at + (i + 1) * 3600_000, end: clock.at + (i + 2) * 3600_000 }));
  events[1] = { ...events[1], allDay: true };
  notifier.setEvents(events);
  const t = made[0];
  assert.deepEqual(ids(t.menu).slice(0, EVENTS_MOST + 3), ['today', 'event:e0', 'event:e1', 'event:e2', 'event:e3', 'more', '-']);
  assert.equal(row(t.menu, 'today').label, 'Today');
  assert.equal(row(t.menu, 'event:e0').label, '11:00  Event 0');
  assert.equal(row(t.menu, 'event:e1').label, 'All day  Event 1');
  assert.equal(row(t.menu, 'more').label, '2 more');
  row(t.menu, 'event:e0').click();
  assert.deepEqual(did, [['open', { app: 'calendar', query: { event: 'e0', at: events[0].start } }]]);
  assert.equal(t.icon, 'plain', 'events are not something that waits');
});

test('an app that is not on in this profile has no row and none of its things', () => {
  const state = { waiting: { missed: 2, chat: 1, mail: 4, total: 7 }, events: [{ id: 'e', summary: 'S', start: AT, end: AT + 1 }], paused: false, pausedUntil: 0 };
  const rows = menuRows({ state, name: 'Rutba Office', apps: ['mail'], does: {}, timeOf });
  assert.deepEqual(ids(rows), ['mail', '-', 'home', 'app:mail', '-', 'pause', 'settings', '-', 'quit']);
  assert.deepEqual(ids(menuRows({ state, name: 'Rutba Office', apps: [], does: {}, timeOf })), ['nothing', '-', 'home', '-', 'pause', 'settings', '-', 'quit']);
});

test('paused, the menu says until when and offers to turn notifications back on', () => {
  const { tray, made, did } = setup();
  tray.show();
  const t = made[0];
  row(t.menu, 'pause').click();
  assert.deepEqual(ids(t.menu).slice(-6), ['-', 'paused', 'resume', 'settings', '-', 'quit']);
  assert.equal(row(t.menu, 'paused').label, 'Notifications are paused until 11:00');
  assert.equal(row(t.menu, 'paused').enabled, false);
  row(t.menu, 'resume').click();
  assert.ok(row(t.menu, 'pause'));
  assert.deepEqual(did, [['pause'], ['resume']]);
});

test('an update that waits for a restart has a row of its own', () => {
  let ran = 0;
  const { tray, made } = setup({ update: () => ({ version: '2.27.0', run: () => { ran += 1; } }) });
  tray.show();
  const r = row(made[0].menu, 'update');
  assert.equal(r.label, 'Restart and update to 2.27.0');
  r.click();
  assert.equal(ran, 1);
  const none = setup();
  none.tray.show();
  assert.equal(row(none.made[0].menu, 'update'), undefined);
});

test('the icon is put by the clock once, and taken away with nothing left listening', () => {
  const { tray, notifier, made } = setup();
  assert.equal(tray.shown(), false);
  tray.show();
  tray.show();
  assert.equal(made.length, 1);
  assert.equal(tray.shown(), true);
  tray.hide();
  assert.equal(made[0].destroyed, true);
  assert.equal(tray.shown(), false);
  const menu = made[0].menu;
  notifier.report({ kind: 'mail', from: 'A', subject: 's' });
  assert.equal(made[0].menu, menu, 'an icon taken away is not drawn again');
});

test('an icon that cannot be made leaves Office without one, and says so', () => {
  const kept = new Map();
  const notifier = createNotifier({ settings: { get: (k, f) => kept.get(k) ?? f, set: (k, v) => kept.set(k, v) } });
  const tray = createTray({ make: () => { throw new Error('no clock here'); }, buildMenu: (r) => r, icons: {}, notifier, does: {} });
  assert.equal(tray.show(), false);
  assert.equal(tray.shown(), false);
});

test('the words for the tip, and for the first time Office stays by the clock', () => {
  assert.deepEqual(waitingWords({ missed: 0, chat: 0, mail: 0 }), []);
  assert.deepEqual(waitingWords({ missed: 2, chat: 0, mail: 1 }), ['2 missed calls', '1 new email']);
  const w = stillRunningWords('Rutba Office');
  assert.equal(w.title, 'Rutba Office is still running');
  assert.match(w.body, /icon by the clock/);
});

test('the mark is a red dot with a white rim in the lower right corner, and the picture handed in is left alone', () => {
  for (const size of [16, 20, 32, 48]) {
    const plain = Buffer.alloc(size * size * 4, 0);
    const marked = paintMark(plain, size, size);
    assert.ok(plain.every((b) => b === 0), 'the picture handed in is as it was');
    assert.equal(marked.length, plain.length);
    const px = (x, y) => [...marked.subarray((y * size + x) * 4, (y * size + x) * 4 + 4)];
    // The middle of the dot: red, and solid (blue, green, red, alpha).
    const r = size * 0.27;
    const cx = Math.floor(size - r - 1);
    assert.deepEqual(px(cx, cx), [77, 72, 229, 255], `the dot at ${size} pixels`);
    // The upper left is untouched.
    assert.deepEqual(px(0, 0), [0, 0, 0, 0]);
    assert.deepEqual(px(Math.floor(size / 4), Math.floor(size / 4)), [0, 0, 0, 0]);
    // Somewhere round the dot is white rim.
    let white = 0;
    for (let i = 0; i < marked.length; i += 4) if (marked[i] > 240 && marked[i + 1] > 240 && marked[i + 2] > 240 && marked[i + 3] > 200) white += 1;
    assert.ok(white > 0, `a rim at ${size} pixels`);
  }
});

// B18: on Windows the system opens an icon's menu on the right button only, and a person who presses the
// icon with the left one saw a window open in place of the menu they were after.
test('where the menu opens on a press: the left button asks for the icon\'s own menu, and pressed twice it opens Office', () => {
  const { tray, made, did } = setup({ menuOnClick: true });
  assert.equal(tray.show(), true);
  const t = made[0];
  assert.ok(t.menu, 'the menu is set on the icon, so the right button opens the same one');
  t.handlers.click();
  assert.equal(t.popped, 1);
  assert.deepEqual(did, [], 'no window is opened by a single press');
  t.handlers['double-click']();
  assert.deepEqual(did, [['open', { app: 'home' }]]);
});

test('an icon that cannot be asked for its menu is left to the right button, and nothing is thrown', () => {
  const { tray, made, did } = setup({ menuOnClick: true });
  tray.show();
  const t = made[0];
  t.popUpContextMenu = () => { throw new Error('not here'); };
  assert.doesNotThrow(() => t.handlers.click());
  delete t.popUpContextMenu;
  assert.doesNotThrow(() => t.handlers.click());
  assert.deepEqual(did, []);
});

test('elsewhere a press still opens Office, and the menu is the system\'s to open', () => {
  const { tray, made, did } = setup();
  tray.show();
  const t = made[0];
  t.handlers.click();
  assert.equal(t.popped, 0);
  assert.equal(t.handlers['double-click'], undefined);
  assert.deepEqual(did, [['open', { app: 'home' }]]);
});
