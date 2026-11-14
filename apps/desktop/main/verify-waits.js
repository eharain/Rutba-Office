// How long a check waits for a window to show something.
//
// Each wait used to carry its own figure, 3,000 to 8,000 ms for 887 waits on
// a menu, a dialog or a button in the shared lane, sized on a quiet machine.
// Beside another lane those figures were beaten (gate 8 closed on a 5,000 ms
// wait for an Add menu), and each was then raised by hand, one closed gate at
// a time. Now there is one floor for them all, sized for a loaded machine:
// a wait asked for less is given UI_WAIT. A wait for something that comes
// quickly still ends as soon as it comes, so a passing run is no slower; only
// a wait that gives up takes the longer time, and each one that does is said
// in the log with its name (review of 11 Nov).
//
// A wait whose length is itself what is checked says so with { exact: true }
// and keeps its figure; tests/gate-reform.test.js holds those to a short list.

import { callSite } from './verify-ids.js';
import { loadMark, loadSince, saidLoad } from './verify-load.js';

/**
 * The last wait that gave up, with where it was asked for and the load over it
 * (verify-load.js): a check that fails after it names it as the likely cause
 * (owner, 11 Nov: say exactly where it failed). verify-apps.js clears it as
 * each part begins.
 */
export const gaveUp = { last: null };

/** The least any wait for the window is given, in ms. */
export const UI_WAIT = 15000;

/**
 * The outer limit for a wait on the app's own signal that a thing happened (a
 * take in its scene, a frame painted, a source drawn), where a fixed wait or a
 * short limit gave up under load (D-2026-11-13-20, the owner: "Wait on signals
 * (Recommended)"). It ends as soon as the signal comes; only a signal that
 * never comes takes this long.
 */
export const SIGNAL_WAIT = 60000;

export const wait = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * The Video window's playhead moved to `t` seconds by its scrub bar, and
 * waited for until the preview says it painted that instant with each clip's
 * decoder on it (the stage's data-time and data-settled): what a fixed wait
 * after a seek used to guess at (D-2026-11-13-20). Answers whether it said so
 * within SIGNAL_WAIT; the caller reads the frame either way.
 */
export async function seekVideoPainted(js, t, { waitFor = until } = {}) {
  const was = await js(`document.querySelector('.vd-canvas')?.dataset.time ?? ''`);
  await js(`(() => { const s = document.querySelector('.vd-scrub'); if (!s) return 0; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, String(${Number(t)})); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
  // A seek to where the playhead already is paints nothing new: that instant, settled, is enough.
  const painted = `(() => { const c = document.querySelector('.vd-canvas'); if (!c) return false; const at = c.dataset.time; return (at !== ${JSON.stringify(was)} || Math.abs(Number(at) - ${Number(t)}) < 1e-6) && c.dataset.settled === '1'; })()`;
  return waitFor(() => js(painted), `the frame at ${t} s painted`, SIGNAL_WAIT).then(() => true, () => false);
}

/** The time a wait asked for `timeout` is given. */
export const waitFor = (timeout, { exact = false } = {}) => (exact ? timeout : Math.max(timeout, UI_WAIT));

/**
 * Wait for `condition` to hold, asking again every 80 ms; throws, naming
 * `what`, when it has not after the time given (at least UI_WAIT).
 */
export async function until(condition, what, timeout = 8000, { exact = false, log = (line) => console.log(line), tick = 80 } = {}) {
  const ms = waitFor(timeout, { exact });
  const mark = loadMark();
  const started = Date.now();
  // Where the wait was asked for, read now: once it gives up the stack is the timer's.
  const at = callSite(new Error().stack, 2, 2);
  const deadline = started + ms;
  for (;;) {
    let held = false;
    try {
      held = await condition();
    } catch {
      held = false;
    }
    if (held) return true;
    if (Date.now() > deadline) {
      const took = Date.now() - started;
      const load = loadSince(mark);
      gaveUp.last = { what, ms: took, at, load };
      log(`     [wait] gave up on ${what} after ${took} ms${at ? ` (${at})` : ''}, ${saidLoad(load)}`);
      throw new Error(`waited ${ms} ms for ${what} and it never happened`);
    }
    await wait(tick);
  }
}
