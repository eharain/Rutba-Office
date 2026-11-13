// A finished recording looked at as a player would show it, before it is
// saved. On a busy machine the H.264 recorder can come up so slowly that a
// short export gets the sound and none of the picture: a file of the right
// length with nothing to see (09.0275, 13 Nov 2026, with every processor busy).
// recordedSomething() cannot see that, since the sound alone is data. So two
// frames are drawn, half a second in and halfway through: a size and
// something drawn on either is a picture. A recording with no picture is made
// again once, the encoder having come up by then, and if the second has none
// either nothing is saved.
//
// Kept apart from the window so the deciding can be tested without one.

/** Whether what was read from a recording is a picture: a size, and something drawn on it. */
export const isPicture = (seen) => Boolean(seen && seen.width > 0 && seen.height > 0 && seen.alpha > 0);

/** The error a recording with no picture ends in, twice over: nothing is saved. */
export const NO_PICTURE = 'no-picture';

/**
 * Make a recording with `make()`, look at it with `look(result)`, and make it
 * again when it has no picture: `tries` in all. `look` answers what the
 * recording held ({ width, height, alpha }), or null when it could not be
 * opened at all, which is not taken for a fault: that recording is kept as
 * before. Cancelled, the recording made is handed back and not made again.
 * Throws an error whose `code` is NO_PICTURE when none of the tries had a
 * picture.
 */
export async function withPicture(make, look, { tries = 2, cancelled = () => false } = {}) {
  for (let i = 0; i < tries; i++) {
    const made = await make(i);
    if (cancelled()) return made;
    const seen = await look(made);
    // Cancelled while it was looked at: not made again.
    if (cancelled() || seen === null || isPicture(seen)) return made;
  }
  return Promise.reject(Object.assign(new Error('the recording has no picture'), { code: NO_PICTURE }));
}

/**
 * withPicture for a caller that saves what comes back: cancelled at any
 * point, while it was made or while it was looked at, it ends in an
 * AbortError and nothing is handed back to be saved.
 */
export async function pictureOrAbort(make, look, { cancelled = () => false, ...rest } = {}) {
  const made = await withPicture(make, look, { cancelled, ...rest });
  if (cancelled()) throw new DOMException('The export was cancelled', 'AbortError');
  return made;
}

/**
 * What a recording shows: { width, height, alpha }. `bytes` may be a Blob,
 * looked at where it is, with no copy. Two frames are drawn and
 * alpha is the higher of their average pixel alpha: 0 is nothing drawn, and a
 * picture drawn is opaque, a black slide as much as any. Null only
 * when it cannot be opened in `waitMs`, or there is no document to open it
 * in. Opened, a recording with no picture size has no picture, and a frame
 * that will not decode counts as nothing drawn: on a busy machine a missing
 * picture must not pass as one that was slow to read. `cancelled()` ends the
 * look early, with null.
 */
export async function readFrame(bytes, contentType, { at = 0.5, waitMs = 8000, doc = globalThis.document, cancelled = () => false } = {}) {
  if (!doc?.createElement) return null;
  const url = URL.createObjectURL(bytes instanceof Blob ? bytes : new Blob([bytes], { type: contentType }));
  const v = doc.createElement('video');
  v.muted = true;
  v.preload = 'auto';
  const event = (names, ms) => new Promise((resolve) => {
    let timer = null;
    const done = (ok) => { clearTimeout(timer); for (const n of names) v.removeEventListener(n, on); resolve(ok); };
    const on = () => done(true);
    for (const n of names) v.addEventListener(n, on);
    timer = setTimeout(() => done(false), ms);
  });
  // A seek counts once it has landed where it was sent: a late answer to the one before it is not this one.
  // Where the length is not known, the seek may stop short of the time asked for: there it counts once it has stopped.
  const landed = (time, known) => !v.seeking && (!known || Math.abs(v.currentTime - time) < 0.05);
  const seekTo = async (time, known) => {
    v.currentTime = time;
    const from = Date.now();
    while (Date.now() - from < waitMs && !cancelled()) {
      if (landed(time, known)) return true;
      // Looked at again every quarter second too: a seek that lands between the look and the listening is not missed.
      await event(['seeked'], 250);
    }
    return landed(time, known);
  };
  const drawn = async () => {
    for (const from = Date.now(); v.readyState < 2 && Date.now() - from < waitMs && !cancelled();) await event(['loadeddata', 'canplay'], 250);
    if (v.readyState < 2) return 0;
    const c = doc.createElement('canvas');
    c.width = 32;
    c.height = 18;
    const g = c.getContext('2d');
    g.drawImage(v, 0, 0, c.width, c.height);
    const px = g.getImageData(0, 0, c.width, c.height).data;
    let alpha = 0;
    for (let i = 3; i < px.length; i += 4) alpha += px[i];
    return alpha / (px.length / 4);
  };
  try {
    const loaded = event(['loadedmetadata', 'error'], waitMs);
    v.src = url;
    if (!(await loaded) || v.error) return null;
    const width = v.videoWidth;
    const height = v.videoHeight;
    // The sound alone: no picture to look for.
    if (!(width > 0 && height > 0)) return { width, height, alpha: 0 };
    let length = v.duration;
    // A recording written as it went may not say its length until the end is looked for.
    if (!Number.isFinite(length)) {
      const told = event(['durationchange'], waitMs);
      v.currentTime = 1e6;
      await told;
      length = v.duration;
    }
    const known = Number.isFinite(length) && length > 0;
    const times = known ? [Math.min(at, length / 2), length / 2] : [at];
    let alpha = 0;
    for (const time of [...new Set(times)]) {
      if (cancelled()) return null;
      if (await seekTo(time, known)) alpha = Math.max(alpha, await drawn());
      if (alpha > 0) break;
    }
    return cancelled() ? null : { width, height, alpha };
  } catch {
    return null;
  } finally {
    v.removeAttribute('src');
    try { v.load(); } catch { /* let go */ }
    URL.revokeObjectURL(url);
  }
}
