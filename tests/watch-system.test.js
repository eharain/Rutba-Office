// What the settings window is told of the system's own say over
// notifications (apps/desktop/main/watch.js, B18): why none is shown, when
// that is known, and the way to the system's own page. Until 2.28.0 a
// Windows with its notifications turned off showed nothing and Office said
// nothing of it. Electron's parts and the system are stand-ins: no icon and
// no notification is raised, and nothing of this computer is read.

import test from 'node:test';
import assert from 'node:assert/strict';
import { createWatch } from '../apps/desktop/main/watch.js';

const settled = () => new Promise((resolve) => { setImmediate(resolve); });

function setup({ check = false, toasts = undefined, menuOnClick = undefined, system = undefined } = {}) {
  const kept = new Map();
  const stores = { settings: { get: (k, f) => (kept.has(k) ? structuredClone(kept.get(k)) : f), set: (k, v) => kept.set(k, structuredClone(v)) } };
  const shown = [];
  const trays = [];
  const parts = {
    show: (note) => { const e = { ...note }; shown.push(e); return { close: () => {} }; },
    flash: () => {},
    supported: () => toasts !== false,
    makeTray: (icon) => {
      const t = { icon, menu: null, handlers: {}, popped: 0, setImage() {}, setToolTip() {}, setContextMenu(m) { this.menu = m; }, on(e, fn) { this.handlers[e] = fn; }, popUpContextMenu() { this.popped += 1; }, destroy() {} };
      trays.push(t);
      return t;
    },
    buildMenu: (rows) => rows,
    icons: () => ({ plain: 'plain', waiting: 'waiting' }),
    startup: { can: false, get: () => false, set: () => {} },
  };
  if (toasts !== undefined) parts.toasts = toasts;
  if (menuOnClick !== undefined) parts.menuOnClick = menuOnClick;
  const told = [];
  const opened = [];
  const watch = createWatch({
    stores,
    parts,
    windows: { all: () => [], infoFor: () => ({}), findByApp: () => null, open: (o) => { opened.push(o); return { webContents: { getURL: () => '' }, isDestroyed: () => false, isMinimized: () => false, isVisible: () => true, show() {}, focus() {} }; } },
    sendEvent: () => {},
    broadcast: (event, payload) => told.push([event, payload]),
    name: 'Rutba Office',
    apps: () => ['mail'],
    check,
    quit: () => {},
    system,
  });
  const states = () => told.filter(([event]) => event === 'notify:state').map(([, payload]) => payload);
  return { watch, shown, trays, told, states, opened };
}

/** A system whose answer the test sets, counting how often it is asked. */
function fakeSystem(first = null) {
  const system = { answer: first, asked: 0, opened: 0, block: async () => { system.asked += 1; return system.answer; }, openSettings: () => { system.opened += 1; } };
  return system;
}

test('with no system to ask nothing stands in the way, and there is no page of its to open', async () => {
  const { watch, states } = setup();
  const s = watch.service.state();
  assert.equal(s.block, null);
  assert.equal(s.systemPage, false);
  assert.equal(s.refused, false);
  assert.equal(watch.service.windowsSettings(), true, 'asked for anyway, it does nothing');
  await settled();
  assert.deepEqual(states(), []);
});

test('a copy that hands the system no notification says so, and the system is never asked', async () => {
  const system = fakeSystem('all');
  const { watch } = setup({ toasts: false, system });
  assert.equal(watch.service.state().block, 'copy');
  assert.equal(watch.service.state().supported, false);
  watch.service.test();
  await settled();
  assert.equal(system.asked, 0);
  assert.equal(watch.service.state().block, 'copy');
});

test('a check run is told of no block and asks nothing of the system', async () => {
  const system = fakeSystem('all');
  const { watch } = setup({ check: true, system });
  watch.service.state();
  watch.service.test();
  await settled();
  assert.equal(system.asked, 0);
  assert.equal(watch.service.state().block, null);
});

test('the system is asked when the settings window asks, and the windows are told when its answer changes', async () => {
  const system = fakeSystem('all');
  const { watch, states } = setup({ system });
  const first = watch.service.state();
  assert.equal(first.block, null, 'not known yet: the window is answered at once');
  assert.equal(first.systemPage, true);
  await settled();
  assert.equal(system.asked, 1);
  assert.deepEqual(states().map((s) => s.block), ['all']);
  assert.equal(watch.service.state().block, 'all');

  // The same answer again is not told again.
  await settled();
  assert.equal(system.asked, 2);
  assert.equal(states().length, 1);

  for (const [answer, block] of [['policy', 'policy'], ['app', 'app'], [null, null]]) {
    system.answer = answer;
    watch.service.state();
    await settled();
    assert.equal(states().at(-1).block, block);
    assert.equal(watch.service.state().block, block);
    await settled();
  }
  assert.equal(states().length, 4);
});

test('an answer that is none of the system\'s three, and a system that fails, are no block', async () => {
  const system = fakeSystem('something else');
  const { watch, states } = setup({ system });
  watch.service.state();
  await settled();
  assert.equal(watch.service.state().block, null);
  await settled();
  system.block = async () => { throw new Error('no answer'); };
  watch.service.state();
  await settled();
  system.block = () => { throw new Error('at once'); };
  watch.service.state();
  await settled();
  assert.equal(watch.service.state().block, null);
  assert.deepEqual(states(), []);
});

test('it is asked once at a time', async () => {
  const system = fakeSystem('all');
  const { watch } = setup({ system });
  watch.service.state();
  watch.service.state();
  watch.service.test();
  assert.equal(system.asked, 0, 'asked after the call has answered');
  await settled();
  assert.equal(system.asked, 1);
});

test('"Show one now" asks the system again; one the system did not show is said so, and asks it too', async () => {
  const system = fakeSystem(null);
  const { watch, shown, states } = setup({ system });
  watch.start();
  const r = watch.service.test();
  assert.equal(r.shown, true);
  await settled();
  assert.equal(system.asked, 1);
  assert.equal(watch.service.state().block, null);
  await settled();
  const before = system.asked;

  // Windows says it showed nothing, and its switches read as on (Do not disturb, or a Windows that keeps them elsewhere).
  shown.at(-1).onFail('not shown');
  assert.equal(states().at(-1).refused, true);
  assert.equal(states().at(-1).block, 'refused');
  await settled();
  assert.equal(system.asked, before + 1, 'the switches are read again when Windows says so');

  // Its switch is found off: that is the plainer thing to say.
  system.answer = 'all';
  watch.service.state();
  await settled();
  assert.equal(watch.service.state().block, 'all');
  await settled();

  // Turned on again in Windows, and shown on purpose: nothing stands in the way.
  system.answer = null;
  watch.service.test();
  await settled();
  const last = watch.service.state();
  assert.deepEqual([last.block, last.refused], [null, false]);
  watch.stop?.();
});

test('with no system to ask, one that was not shown is still not called a block', () => {
  const { watch, shown } = setup();
  watch.start();
  watch.service.test();
  shown.at(-1).onFail('not shown');
  const s = watch.service.state();
  assert.deepEqual([s.refused, s.block], [true, null]);
});

test('the system\'s own page is opened when asked, and one that cannot be opened throws nothing', () => {
  const system = fakeSystem(null);
  const { watch } = setup({ system });
  assert.equal(watch.service.windowsSettings(), true);
  assert.equal(system.opened, 1);
  system.openSettings = () => { throw new Error('no such page'); };
  assert.equal(watch.service.windowsSettings(), true);
});

test('the icon is told where its menu opens on a press of the left button', () => {
  const on = setup({ menuOnClick: true });
  on.watch.start();
  on.trays[0].handlers.click();
  assert.equal(on.trays[0].popped, 1);
  assert.deepEqual(on.opened, []);
  on.trays[0].handlers['double-click']();
  assert.equal(on.opened.length, 1);

  const off = setup();
  off.watch.start();
  off.trays[0].handlers.click();
  assert.equal(off.trays[0].popped, 0);
  assert.equal(off.opened.length, 1);
});
