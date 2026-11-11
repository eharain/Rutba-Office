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

/**
 * The last wait that gave up, with where it was asked for: a check that fails
 * after it names it as the likely cause (owner, 11 Nov: say exactly where it
 * failed). verify-apps.js clears it as each part begins.
 */
export const gaveUp = { last: null };

/** The least any wait for the window is given, in ms. */
export const UI_WAIT = 15000;

export const wait = (ms) => new Promise((r) => setTimeout(r, ms));

/** The time a wait asked for `timeout` is given. */
export const waitFor = (timeout, { exact = false } = {}) => (exact ? timeout : Math.max(timeout, UI_WAIT));

/**
 * Wait for `condition` to hold, asking again every 80 ms; throws, naming
 * `what`, when it has not after the time given (at least UI_WAIT).
 */
export async function until(condition, what, timeout = 8000, { exact = false, log = (line) => console.log(line), tick = 80 } = {}) {
  const ms = waitFor(timeout, { exact });
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
      gaveUp.last = { what, ms: Date.now() - started, at };
      log(`     [wait] gave up on ${what} after ${Date.now() - started} ms${at ? ` (${at})` : ''}`);
      throw new Error(`waited ${ms} ms for ${what} and it never happened`);
    }
    await wait(tick);
  }
}
