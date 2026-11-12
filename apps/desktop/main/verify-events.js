// Waits that end on what the window says happened, not on a time.
//
// A check that changed a window's size used to poll the page every 80 ms
// until it was the width asked for, then sleep a fixed time for the layout.
// On this PC a window asked for w came out w + 16 (the 8 px resize border
// Windows keeps on each side), so the poll never matched and every one of
// those waits ran to its floor: 83 waits and 1,249 s of the ribbon check's
// 1,436 s (12 Nov). Here the window is asked, what it became is read back and
// corrected, and the wait ends when the page reports that size and has drawn
// its layout frames: as soon as the machine has done it, however loaded.

import { UI_WAIT } from './verify-waits.js';

/**
 * Resolves true when `emitter` emits `event`, or false after `ms`. The
 * promise's `cancel()` lets go of the listener and the timer when the answer
 * is no longer wanted.
 */
export function once(emitter, event, ms = UI_WAIT) {
  let done = () => {};
  const p = new Promise((resolve) => {
    let timer = null;
    const on = () => done(true);
    done = (v) => {
      clearTimeout(timer);
      emitter.removeListener?.(event, on);
      resolve(v);
    };
    timer = setTimeout(() => done(false), ms);
    emitter.once(event, on);
  });
  p.cancel = () => done(false);
  return p;
}

/** Within a pixel: at a scale that is not whole, the window's size and the page's round apart. */
const near = (a, b) => Math.abs(a - b) <= 1;

/**
 * In the page: resolves with its size once it is `w` x `h` (now, or at a
 * resize event) and `frames` animation frames have run after (the ribbon fits
 * its groups one frame after its ResizeObserver). A page drawing no frames
 * (hidden) is let go after `idle` ms; one never at the size, with null after
 * `timeout`, its listener taken off.
 */
export const settled = (w, h, { frames = 3, idle = 400, timeout = UI_WAIT } = {}) => `new Promise((resolve) => {
  let over = false;
  const finish = (v) => { if (!over) { over = true; removeEventListener('resize', on); clearTimeout(cap); resolve(v); } };
  const laidOut = () => {
    let left = ${frames};
    const step = () => (--left <= 0 ? finish([innerWidth, innerHeight]) : requestAnimationFrame(step));
    requestAnimationFrame(step);
    setTimeout(() => finish([innerWidth, innerHeight]), ${idle});
  };
  const there = () => Math.abs(innerWidth - ${w}) <= 1 && Math.abs(innerHeight - ${h}) <= 1;
  const on = () => { if (there()) { removeEventListener('resize', on); laidOut(); } };
  const cap = setTimeout(() => finish(null), ${timeout});
  if (there()) laidOut();
  else addEventListener('resize', on);
})`;

const same = (a, b) => near(a[0], b[0]) && near(a[1], b[1]);

/**
 * The window's page at `width` x `height` CSS px, laid out. Not maximized or
 * full screen first. Answers what the page became, how long it took and how
 * many times the window was asked; `reached` is false when the screen does
 * not have the room (Windows keeps a window inside its screen) or the window
 * will not go so small, for the caller to say.
 *
 * @param {Electron.BrowserWindow} win
 * @param {number} width
 * @param {number} height
 */
export async function sizeTo(win, width, height, { timeout = UI_WAIT, frames = 3, idle = 400, tries = 3 } = {}) {
  const started = Date.now();
  if (win.isFullScreen()) {
    const left = once(win, 'leave-full-screen', timeout);
    win.setFullScreen(false);
    await left;
  }
  if (win.isMaximized()) {
    const back = once(win, 'unmaximize', timeout);
    win.unmaximize();
    await back;
  }
  const want = [Math.round(width), Math.round(height)];
  const ask = [...want];
  let asks = 0;
  // A side that did not move when asked is as near as it goes (the screen has no more room, or the window's least size).
  const stuck = [false, false];
  const open = (got) => [0, 1].some((i) => !near(got[i], want[i]) && !stuck[i]);
  // The window's own size, read back at once: what Windows added is taken off the next ask.
  while (open(win.getContentSize()) && asks < tries) {
    const before = win.getContentSize();
    const moved = once(win, 'resize', 500);
    win.setContentSize(Math.max(1, ask[0]), Math.max(1, ask[1]));
    asks += 1;
    // Read back at once as a rule; if it reads unchanged, the window is given its resize event (or half a second) first.
    let got = win.getContentSize();
    if (same(got, before)) {
      await moved;
      got = win.getContentSize();
    } else moved.cancel();
    for (const i of [0, 1]) {
      if (near(got[i], want[i])) continue;
      // Not on the first ask: a window 16 px over, asked for its size less 16, comes out where it was.
      if (asks > 1 && got[i] === before[i]) stuck[i] = true;
      else ask[i] += want[i] - got[i];
    }
  }
  const final = win.getContentSize();
  let backstop = null;
  const page = await Promise.race([
    win.webContents.executeJavaScript(settled(final[0], final[1], { frames, idle, timeout })),
    new Promise((r) => { backstop = setTimeout(() => r(null), timeout + 1000); }),
  ]).finally(() => clearTimeout(backstop));
  const [w, h] = page || await win.webContents.executeJavaScript('[innerWidth, innerHeight]');
  return { width: w, height: h, window: final, reached: near(w, want[0]) && near(h, want[1]), timedOut: !page, asks, ms: Date.now() - started };
}
