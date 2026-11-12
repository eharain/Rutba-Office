// A fresh window's video encoder is slow to come up: on some machines the
// H.264 recorder gives no data at all for its first 2 to 4 seconds, so a
// short first recording in a new window came out as an empty file. A second
// recording in the same window is fine. So before the recording that counts,
// a throwaway one is run until the encoder hands over its first piece, and
// that piece is thrown away. An MP4 recorder's "start" event only comes when
// it stops, so the first data is the one sign it is ready.
//
// The throwaway one records a small canvas of its own, never the tracks the
// real recording is made from: a track handed from a just-stopped recorder
// to a new one is itself a race, which in the studio's export gave the very
// empty file its primer was there to prevent (packages/studio, primeEncoder).
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
  // A recorder that fails stops of itself and will give nothing more: the
  // wait ends there rather than running on to its cap.
  let failed = false;
  let stopped = null;
  try {
    rec = make();
    stopped = new Promise((resolve) => { rec.onstop = resolve; rec.onerror = () => { failed = true; resolve(); }; });
    rec.ondataavailable = (e) => { if (e?.data?.size) answered = true; };
    rec.start(sliceMs);
    const from = now();
    while (!answered && !failed && !cancelled() && now() - from < capMs) {
      draw();
      await wait(frameMs);
    }
  } catch { /* the real recording is what matters */ } finally {
    if (rec && rec.state !== 'inactive') {
      // Its last piece arrives before its stop; waiting for the stop means
      // the real recorder never starts while this one is still encoding.
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

/**
 * Warm the encoder for `mimeType` on a 480 by 270 canvas of its own, drawn
 * a frame at a time, as warmEncoder does; the canvas's track is stopped
 * after, whatever happened. `cancelled()` and `capMs` as warmEncoder's.
 * `canvas` and `Recorder` are the window's own unless a test hands in others.
 */
export async function warmOnOwnCanvas(mimeType, {
  cancelled = () => false, capMs = WARM_CAP_MS,
  canvas = () => document.createElement('canvas'),
  Recorder = typeof MediaRecorder === 'undefined' ? null : MediaRecorder,
  ...rest
} = {}) {
  let stream = null;
  try {
    if (!Recorder) return false;
    const c = canvas();
    c.width = 480;
    c.height = 270;
    const g = c.getContext('2d');
    stream = c.captureStream(0);
    const track = stream.getVideoTracks()[0];
    let n = 0;
    // A canvas stream with no frame rate sends a frame only when asked, so
    // each frame drawn differs from the last and is asked for.
    const draw = () => {
      n += 1;
      g.fillStyle = n % 2 ? '#222' : '#444';
      g.fillRect(0, 0, c.width, c.height);
      track?.requestFrame?.();
    };
    draw();
    return await warmEncoder(() => new Recorder(stream, { mimeType }), { draw, cancelled, capMs, ...rest });
  } catch {
    return false;
  } finally {
    for (const k of stream?.getTracks?.() || []) { try { k.stop(); } catch { /* ended */ } }
  }
}

/** Whether a finished recording holds anything: an empty one is never written as a file. */
export const recordedSomething = (bytes) => Boolean(bytes && (bytes.length ?? bytes.byteLength ?? bytes.size));
