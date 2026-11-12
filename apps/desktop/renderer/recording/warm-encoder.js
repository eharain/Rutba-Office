// A fresh window's video encoder is slow to come up: on some machines the
// H.264 recorder gives no data at all for its first 2 to 4 seconds, so a
// short first recording in a new window came out as an empty file. A second
// recording in the same window is fine. So before the recording that counts,
// a throwaway one is run on the same tracks until the encoder hands over its
// first piece, and that piece is thrown away. An MP4 recorder's "start" event
// only comes when it stops, so the first data is the one sign it is ready.
//
// Kept apart from the window so the waiting can be tested without one.

/** How long the throwaway recording may wait for the encoder before the real one starts anyway. */
export const WARM_CAP_MS = 10000;
/** How often the throwaway recording is asked for what it has. */
export const WARM_SLICE_MS = 100;

const pause = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Run `make()`'s recorder until it gives a piece of data, then stop it and
 * drop what it gave. `draw()` is called every frame meanwhile, since a canvas
 * stream sends nothing unless it is drawn on. It ends early when
 * `cancelled()` says so, and after `capMs` whatever happens. Resolves to
 * whether the encoder answered; it never throws, so a recorder that will not
 * start leaves the real recording to go ahead as it would have.
 */
export async function warmEncoder(make, {
  draw = () => {}, cancelled = () => false, capMs = WARM_CAP_MS, sliceMs = WARM_SLICE_MS,
  frameMs = 1000 / 30, now = () => Date.now(), wait = pause,
} = {}) {
  let rec = null;
  let answered = false;
  let stopped = null;
  try {
    rec = make();
    stopped = new Promise((resolve) => { rec.onstop = resolve; rec.onerror = resolve; });
    rec.ondataavailable = (e) => { if (e?.data?.size) answered = true; };
    rec.start(sliceMs);
    const from = now();
    while (!answered && !cancelled() && now() - from < capMs) {
      draw();
      await wait(frameMs);
    }
  } catch { /* the real recording is what matters */ } finally {
    if (rec && rec.state !== 'inactive') {
      // Its last piece arrives before its stop; waiting for the stop means
      // the real recorder never starts while this one still holds the tracks.
      let stopping = false;
      try { rec.stop(); stopping = true; } catch { /* stopped already */ }
      if (stopping) {
        let timer = null;
        await Promise.race([stopped, new Promise((r) => { timer = setTimeout(r, 1500); })]);
        clearTimeout(timer);
      }
    }
    if (rec) rec.ondataavailable = null;
  }
  return answered;
}

/** Whether a finished recording holds anything: an empty one is never written as a file. */
export const recordedSomething = (bytes) => Boolean(bytes && (bytes.length ?? bytes.byteLength ?? bytes.size));
