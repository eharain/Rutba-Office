// The window-size wait that ends on what the window did (12 Nov pilot): the
// ribbon check asked 83 times for a width this PC's windows come out 16 px
// wider than, and each wait ran out at 15 s.
import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import vm from 'node:vm';
import { once, settled, sizeTo } from '../apps/desktop/main/verify-events.js';

/**
 * A window whose content comes out `extra` px bigger than asked (Windows'
 * resize border), never past `room`, and whose page reports its size at once.
 */
function fakeWin({ start = [1280, 900], extra = [16, 8], room = [1920, 1032], max = false, full = false } = {}) {
  const win = new EventEmitter();
  let size = [...start];
  const state = { asks: [], maximized: max, full };
  win.getContentSize = () => [...size];
  win.setContentSize = (w, h) => {
    state.asks.push([w, h]);
    size = [Math.min(w + extra[0], room[0]), Math.min(h + extra[1], room[1])];
  };
  win.isMaximized = () => state.maximized;
  win.isFullScreen = () => state.full;
  win.unmaximize = () => setTimeout(() => { state.maximized = false; win.emit('unmaximize'); }, 5);
  win.setFullScreen = (on) => setTimeout(() => { state.full = on; win.emit('leave-full-screen'); }, 5);
  win.webContents = { executeJavaScript: async () => [...size] };
  return { win, state };
}

test('a window that comes out wider than asked is asked again by the difference, and the wait ends at once', async () => {
  const { win, state } = fakeWin();
  const got = await sizeTo(win, 1680, 860, { timeout: 2000 });
  assert.deepEqual([got.width, got.height], [1680, 860]);
  assert.equal(got.reached, true);
  assert.equal(got.asks, 2);
  assert.deepEqual(state.asks, [[1680, 860], [1664, 852]]);
  assert.ok(got.ms < 1000, `${got.ms} ms`);
});

test('a window already at its size is not asked at all', async () => {
  const { win, state } = fakeWin({ start: [900, 860] });
  const got = await sizeTo(win, 900, 860, { timeout: 2000 });
  assert.equal(got.reached, true);
  assert.equal(state.asks.length, 0);
});

test('a window that comes out where it was on the first ask is corrected, not taken for one with no room', async () => {
  // At 1296 x 868 (asked 1280 x 860 before), asked 1280 x 860 again: Windows gives 1296 x 868 once more.
  const { win, state } = fakeWin({ start: [1296, 868] });
  const got = await sizeTo(win, 1280, 860, { timeout: 2000 });
  assert.equal(got.reached, true, JSON.stringify(state.asks));
  assert.deepEqual(state.asks, [[1280, 860], [1264, 852]]);
});

test('a screen without the room ends the asking and says the size was not reached', async () => {
  const { win, state } = fakeWin({ room: [1500, 1032] });
  const got = await sizeTo(win, 1680, 860, { timeout: 2000 });
  assert.equal(got.reached, false);
  assert.equal(got.width, 1500);
  assert.equal(state.asks.length, 2, JSON.stringify(state.asks));
});

test('a maximized or full-screen window is put back first, on its own event', async () => {
  const { win, state } = fakeWin({ max: true, full: true });
  const got = await sizeTo(win, 1280, 860, { timeout: 2000 });
  assert.equal(state.maximized, false);
  assert.equal(state.full, false);
  assert.equal(got.reached, true);
});

test('once answers true on the event and false when it never comes, leaving no listener behind', async () => {
  const e = new EventEmitter();
  setTimeout(() => e.emit('ready'), 5);
  assert.equal(await once(e, 'ready', 1000), true);
  assert.equal(await once(e, 'never', 20), false);
  assert.equal(e.listenerCount('never'), 0);
  assert.equal(e.listenerCount('ready'), 0);
});

/** A page for the in-page wait: its size, its resize listeners, and frames 16 ms apart. */
function fakePage(w, h) {
  const listeners = new Set();
  const page = {
    innerWidth: w,
    innerHeight: h,
    frames: 0,
    addEventListener: (e, f) => { if (e === 'resize') listeners.add(f); },
    removeEventListener: (e, f) => { if (e === 'resize') listeners.delete(f); },
    requestAnimationFrame: (f) => setTimeout(() => { page.frames += 1; f(); }, 16),
    setTimeout,
    clearTimeout,
    resize: (nw, nh) => { page.innerWidth = nw; page.innerHeight = nh; for (const f of [...listeners]) f(); },
    listeners,
  };
  return { page, run: (script) => vm.runInContext(script, vm.createContext(page)) };
}

test('in the page: a page already at the size is let go after its layout frames, not before', async () => {
  const { page, run } = fakePage(1280, 860);
  assert.deepEqual([...await run(settled(1280, 860, { frames: 3, idle: 2000, timeout: 3000 }))], [1280, 860]);
  assert.equal(page.frames, 3);
  assert.equal(page.listeners.size, 0);
});

test('in the page: a page that comes to the size later is waited for, to its resize, within a pixel', async () => {
  const { page, run } = fakePage(1296, 868);
  const p = run(settled(1280, 860, { frames: 2, idle: 2000, timeout: 3000 }));
  setTimeout(() => page.resize(1290, 862), 20); // not there yet
  setTimeout(() => page.resize(1281, 860), 60); // a pixel over: a scale that is not whole
  const started = Date.now();
  assert.deepEqual([...await p], [1281, 860]);
  assert.ok(Date.now() - started < 1000);
  assert.equal(page.listeners.size, 0);
});

test('in the page: a page never at the size answers null at its timeout and leaves no listener', async () => {
  const { page, run } = fakePage(1296, 868);
  assert.equal(await run(settled(1280, 860, { timeout: 120 })), null);
  assert.equal(page.listeners.size, 0);
});

test('a window that resizes a moment after it is asked is waited for on its resize event', async () => {
  const win = new EventEmitter();
  let size = [1280, 900];
  win.getContentSize = () => [...size];
  win.setContentSize = (w, h) => setTimeout(() => { size = [w, h]; win.emit('resize'); }, 30);
  win.isMaximized = () => false;
  win.isFullScreen = () => false;
  win.webContents = { executeJavaScript: async () => [...size] };
  const got = await sizeTo(win, 1680, 860, { timeout: 2000 });
  assert.equal(got.reached, true);
  assert.equal(got.asks, 1);
  assert.ok(got.ms < 400, `${got.ms} ms`);
});
