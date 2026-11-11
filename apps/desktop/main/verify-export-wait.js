// Waiting for a video export in the shared lane.
//
// The checks used to give an export the project's length and forty seconds,
// a deadline of about a fifth of real time: a budget in the shared lane,
// where the graphics encoder is shared with the lanes beside it (review of
// 11 Nov). And an export that missed it was left running, so the next export
// in the same window found the box still open and never began: one slow
// export failed six checks in gate 9.
//
// Now the wait goes on while the export moves (its line or its note
// changes), up to `capMs` in all, and an export still open at the end is
// cancelled and its box seen to close before the check goes on. How fast it
// goes is the machine lane's to measure (verify-video-mp4.js).

import { UI_WAIT } from './verify-waits.js';

/** What the export box shows that moves when the export does. */
const SHOWN = `(() => {
  const b = document.querySelector('.vd-export-busy');
  if (!b) return null;
  return [b.querySelector('.vd-export-line')?.textContent || '', b.querySelector('.vd-export-again')?.textContent || ''].join(' | ');
})()`;

const CANCEL = `(() => {
  const b = document.querySelector('.vd-export-busy');
  const c = b && [...b.querySelectorAll('button')].find((x) => /^Cancel$/.test(x.textContent.trim()));
  if (c) c.click();
  return Boolean(c);
})()`;

/**
 * Wait for an export begun in `win` to end. `done()` answers truthy once it
 * has (the file written, or a refusal shown). Answers { done, seconds,
 * stopped, shown }: `stopped` says why it was given up (null when done),
 * `shown` what the box last said.
 */
export async function waitForExport(win, done, { quietMs = 30000, capMs = 300000, tickMs = 250, openMs = UI_WAIT } = {}) {
  const js = (code) => win.webContents.executeJavaScript(code);
  const started = Date.now();
  let shown;
  let movedAt = Date.now();
  let closedAt = null;
  let seen = false;
  for (;;) {
    if (await done()) return { done: true, seconds: (Date.now() - started) / 1000, stopped: null, shown };
    const now = await js(SHOWN).catch(() => null);
    if (now !== shown) { shown = now; movedAt = Date.now(); }
    if (now !== null) seen = true;
    // A box not drawn yet is not a box that closed (review of 11 Nov): beside
    // other lanes it can take seconds to come, and an export it then started
    // would run on into the next check. It is given UI_WAIT to come.
    if (now === null && !seen) {
      // Its quiet is counted from when it opens, not from the press.
      movedAt = Date.now();
      if (Date.now() - started > openMs) return { done: Boolean(await done()), seconds: (Date.now() - started) / 1000, stopped: `the export box did not open in ${Math.round(openMs / 1000)} seconds`, shown };
    } else if (now === null) {
      // The box closed and nothing came: the export ended without a file (its toast says why). A moment for the file to land first.
      closedAt ??= Date.now();
      if (Date.now() - closedAt > 3000) return { done: Boolean(await done()), seconds: (Date.now() - started) / 1000, stopped: 'the export box closed and no file came', shown };
    } else closedAt = null;
    const still = Date.now() - movedAt > quietMs;
    const long = Date.now() - started > capMs;
    if (still || long) {
      // Never left running: the next export in this window would find the box still open.
      await js(CANCEL).catch(() => false);
      const t0 = Date.now();
      while (Date.now() - t0 < 15000 && (await js(`Boolean(document.querySelector('.vd-export-busy'))`).catch(() => false))) await new Promise((r) => setTimeout(r, tickMs));
      return { done: false, seconds: (Date.now() - started) / 1000, stopped: still ? `it did not move for ${Math.round(quietMs / 1000)} seconds, and was cancelled` : `it ran past ${Math.round(capMs / 1000)} seconds, and was cancelled`, shown };
    }
    await new Promise((r) => setTimeout(r, tickMs));
  }
}
