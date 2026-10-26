// The computer's camera, shared: one stream per way of asking for it (which
// camera, how tall a picture), counted, and stopped with the last holder.
// Presentations' cameos and the recorder both open it here.
//
// The open is patient: a camera that is there but busy (let go a moment ago
// by a preview, held by another program, or slow to start) is asked for
// again, waiting a little longer each time, for about nine seconds.

import { t } from '@rutba/office-ui';

const shared = new Map(); // key -> { stream, users }
const opening = new Map(); // key -> promise

/** Whether a stream still has its camera: a stream whose camera stopped is not handed out again. */
export const alive = (stream) => Boolean(stream?.active && stream.getVideoTracks().some((t) => t.readyState === 'live'));

/** A camera that is there but busy, or slow to start, is asked for again rather than given up on. */
export const BUSY = ['NotReadableError', 'AbortError', 'TrackStartError'];

/** What getUserMedia is asked for: the camera picked (or any), at about `height` lines, 16:9. */
export function cameraConstraints({ deviceId = '', height = 720 } = {}) {
  const h = Number(height) || 720;
  return {
    video: { ...(deviceId ? { deviceId: { exact: deviceId } } : {}), width: { ideal: Math.round((h * 16) / 9) }, height: { ideal: h } },
    audio: false,
  };
}

const keyOf = (want) => JSON.stringify({ d: want?.deviceId || '', h: Number(want?.height) || 720 });

/** The camera's stream, opened or shared; each call is one holder until closeCamera. */
export async function openCamera(want = null) {
  const key = keyOf(want);
  let entry = shared.get(key);
  if (entry && !alive(entry.stream)) {
    for (const t of entry.stream.getTracks()) t.stop();
    shared.delete(key);
    entry = null;
  }
  if (entry) { entry.users += 1; return entry.stream; }
  if (!opening.has(key)) {
    opening.set(key, navigator.mediaDevices.getUserMedia(cameraConstraints(want || {}))
      .then((stream) => { shared.set(key, { stream, users: 0 }); return stream; })
      .finally(() => { opening.delete(key); }));
  }
  const stream = await opening.get(key);
  let now = shared.get(key);
  if (!now || now.stream !== stream) { now = { stream, users: 0 }; shared.set(key, now); }
  now.users += 1;
  return stream;
}

/** The camera, asked for again while it is busy, waiting a little longer each time, for about nine seconds. */
export async function openCameraPatiently(want = null) {
  for (let tries = 1; ; tries += 1) {
    try {
      return await openCamera(want);
    } catch (err) {
      if (!BUSY.includes(err?.name) || tries > 5) throw err;
      await new Promise((resolve) => { setTimeout(resolve, 300 * 2 ** (tries - 1)); });
    }
  }
}

/** One holder of `stream` lets it go; the camera is closed with the last. A stream already replaced is left alone. */
export function closeCamera(stream) {
  for (const [key, entry] of shared) {
    if (stream && entry.stream !== stream) continue;
    entry.users = Math.max(0, entry.users - 1);
    if (!entry.users) {
      for (const t of entry.stream.getTracks()) t.stop();
      shared.delete(key);
    }
    return;
  }
}

/** Why a camera could not be had, in a sentence. */
export function cameraTrouble(err) {
  if (err?.name === 'NotAllowedError') return t('The camera was not allowed.');
  if (err?.name === 'NotReadableError') return t('The camera is in use by another program.');
  return t('No camera was found.');
}
