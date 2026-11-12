// When a camera the studio still wants is asked for again. openCameraPatiently
// gives a busy camera about nine seconds; after that the studio used to give
// up for good, so a camera that came free a moment later, or was plugged in,
// still said "No camera was found" until the setup changed. Now it is asked
// for again every few seconds for as long as an item shows it. A camera the
// person or Windows refused is not: asking again cannot change the answer,
// and each refusal can show a prompt.
//
// Kept apart from the window so the waiting can be tested without one.

/** How long after a failed open a camera is asked for again. */
export const CAMERA_RETRY_MS = 4000;

/** Whether a camera that could not be opened is worth asking for again. */
export const retryCamera = (err) => !['NotAllowedError', 'SecurityError'].includes(err?.name);

/**
 * How long a capture that opened but has given no frame is waited on before
 * it is opened again, by how many times it has been opened again already.
 * The wait grows, so a camera slow to give its first frame is not closed
 * each time just before it would have.
 */
export const framelessWaitMs = (reopened = 0) => Math.min(20000, 3000 * 2 ** Math.max(0, reopened));

/** How many times a capture that gives no frame is opened again: a camera while it is wanted, a screen or window once (a minimised window gives none, and opening it again does not help). */
export const reopenLimit = (camera) => (camera ? Infinity : 1);

/**
 * A wait that can be ended early: `wait.done` resolves after `ms`, or as soon
 * as `wait.cancel()` is called, to whether the time ran out. Its timer is
 * always cleared, so a capture let go leaves nothing running.
 */
export function cancellableWait(ms, { set = setTimeout, clear = clearTimeout } = {}) {
  let finish = null;
  let timer = null;
  const done = new Promise((resolve) => { finish = resolve; });
  timer = set(() => { timer = null; finish(true); }, ms);
  return {
    done,
    cancel() {
      if (timer != null) { clear(timer); timer = null; }
      finish(false);
    },
  };
}

/**
 * Opens with `open()` until it works, while `wanted()` still says so. Each
 * failure goes to `onFail(err)` (to show it) and, unless `retry(err)` says
 * it is not worth it, another try follows after `delayMs`, waited through
 * `pause(ms)`, which returns a cancellable wait. Resolves to what `open()`
 * gave, or to null once it is not wanted or not worth trying again.
 * `onWait(w)` is handed each wait, so the caller can cancel it.
 */
export async function openWhileWanted(open, {
  wanted = () => true, onFail = () => {}, retry = retryCamera, delayMs = CAMERA_RETRY_MS,
  pause = cancellableWait, onWait = () => {},
} = {}) {
  while (wanted()) {
    try {
      return await open();
    } catch (err) {
      if (!wanted()) return null;
      onFail(err);
      if (!retry(err)) return null;
      const w = pause(delayMs);
      onWait(w);
      await w.done;
    }
  }
  return null;
}
