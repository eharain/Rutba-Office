// What keeps a window check steady beside other lanes (2.28.0).
//
// Eleven checks failed once each in the 2.27.0 gate while other lanes ran, and
// passed alone (D-2026-11-24-01). In none of them was the window slow: the
// wait that gave up had its whole 15 seconds. The check's own press or key
// had no effect at all, in one of three ways:
//
//   - a menu or the ribbon's KeyTips, which the window puts away when it loses
//     the keyboard or is resized, were put away from outside the check in the
//     middle of its keys (02.0035, 02.0040, 01.0087 and the three after it);
//   - a handle was read where it stood a moment before the sheet came to rest,
//     and the drag began beside it (08.0010 and the two after it, 09.0237);
//   - a file being read took longer than a fixed wait (12.0135).
//
// So a check here says what it saw, and tries again only on that evidence:
//
//   - outside(): what reached the page from outside the check (the window lost
//     the keyboard, or was resized). steadily() runs a step again, at most
//     three times, only when one of those was seen during a try that failed. A
//     try that fails with nothing seen is the window's own fault and fails at
//     once, and the check's detail says how many tries it took and what was seen.
//   - still(): a point read until two readings a moment apart agree.
//   - pressWatch() and pressLanded(): what took the next press of the mouse, so
//     a drag is tried again only when its press landed beside the handle. A
//     press that landed on the handle and did nothing fails.
//
// None of this passes a window that does the wrong thing: every try is judged
// by the same test as before, and the tries are in the detail.

/** Begin noting, in a window's page, what reaches it from outside a check: 'blur' (the window, not a field) and 'resize'. */
export const outside = (js) => js(`(() => {
  if (window.__rutbaOutside) return window.__rutbaOutside.length;
  const seen = window.__rutbaOutside = [];
  window.addEventListener('blur', (e) => { if (e.target === window) seen.push('blur'); });
  window.addEventListener('resize', () => seen.push('resize'));
  return 0;
})()`);

/** What has been noted since the first `from` of them. */
export const outsideSince = (js, from) => js(`(window.__rutbaOutside || []).slice(${Number(from) || 0})`);

/**
 * Run `attempt` (async, answering an object with `ok`) and, when it fails with
 * something from outside seen in the page during it, run it again: at most
 * `tries` times. Answers the last try's answer with `tries` (how many were
 * made) and `seen` (what came from outside during the tries that failed).
 *
 * @param {(code: string) => Promise<any>} js runs code in the window's page
 * @param {(n: number) => Promise<{ ok: boolean }>} attempt
 */
export async function steadily(js, attempt, { tries = 3 } = {}) {
  const seen = [];
  let last = { ok: false };
  for (let n = 1; n <= tries; n++) {
    const from = await outside(js).catch(() => 0);
    last = await attempt(n);
    if (last?.ok) return { ...last, tries: n, seen };
    const during = await outsideSince(js, from).catch(() => []);
    if (!during.length) return { ...last, tries: n, seen };
    seen.push(...during);
  }
  return { ...last, tries, seen };
}

/**
 * A point read until it stands still: `read` (async, answering { x, y } or
 * null) asked every `gap` ms until two readings in a row agree within a pixel,
 * for at most `limit` ms. Answers the last reading either way.
 */
export async function still(read, wait, { gap = 120, limit = 4000 } = {}) {
  const started = Date.now();
  let was = await read();
  for (let n = 0; n < Math.ceil(limit / gap); n++) {
    await wait(gap);
    const now = await read();
    if (was && now && Math.abs(now.x - was.x) <= 1 && Math.abs(now.y - was.y) <= 1) return now;
    was = now;
    if (Date.now() - started > limit) break;
  }
  return was;
}

/**
 * Note what takes the next press of the mouse in the page: whether it was the
 * handle (`selector`, or something inside it), and the class of what it was.
 */
export const pressWatch = (js, selector) => js(`(() => {
  window.__rutbaPressed = null;
  window.addEventListener('mousedown', (e) => {
    const n = e.target;
    const el = n && typeof n.closest === 'function' ? n : null;
    window.__rutbaPressed = {
      handle: Boolean(el && el.closest(${JSON.stringify(selector)})),
      took: String((el && el.getAttribute('class')) || (n && n.tagName) || 'nothing').slice(0, 60),
    };
  }, { capture: true, once: true });
  return 1;
})()`);

/** What took the press noted by pressWatch, `{ handle, took }`: null when no press reached the page. */
export const pressLanded = (js) => js(`window.__rutbaPressed ?? null`);

/** Whether a drag that did nothing may be tried again: its press reached nothing, or landed beside the handle. */
export const missed = (landed) => !landed || landed.handle !== true;
