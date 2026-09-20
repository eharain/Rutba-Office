// A frame of a clip, drawn by the window.
//
// The platform makes most thumbnails (thumbs.js in the shell): the system
// knows the videos it has a codec for. For the ones it does not — a WebM on
// Windows without the codec pack, anything on Linux — the window can draw a
// frame itself, since Chromium plays what the platform will not. One clip at
// a time, well into the clip rather than its black first frame, and the
// result handed to the platform so the next visit finds it in the cache.

import { fileUrl } from './library.js';

const SIZE = 256;
const TIMEOUT = 8000;

const results = new Map(); // path -> Promise<string|null> (an object URL, or nothing)
let chain = Promise.resolve();

function grab(path) {
  return new Promise((resolve) => {
    const video = document.createElement('video');
    video.muted = true;
    video.playsInline = true;
    video.preload = 'metadata';
    // The clip comes from the file host, the page from the app host: a
    // canvas drawn from it can be read back only if it was asked for as a
    // cross-origin resource and the host allowed it (protocol.js does).
    video.crossOrigin = 'anonymous';
    let done = false;
    const finish = (value) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      video.removeAttribute('src');
      video.load();
      resolve(value);
    };
    const timer = setTimeout(() => finish(null), TIMEOUT);
    video.onerror = () => finish(null);
    video.onloadedmetadata = () => {
      const at = Number.isFinite(video.duration) && video.duration > 0 ? Math.min(1, video.duration * 0.1) : 0;
      video.currentTime = at;
    };
    video.onseeked = () => {
      try {
        const w = video.videoWidth;
        const h = video.videoHeight;
        if (!w || !h) return finish(null);
        const scale = Math.min(1, SIZE / Math.max(w, h));
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.round(w * scale));
        canvas.height = Math.max(1, Math.round(h * scale));
        canvas.getContext('2d').drawImage(video, 0, 0, canvas.width, canvas.height);
        canvas.toBlob((blob) => finish(blob || null), 'image/jpeg', 0.82);
      } catch {
        finish(null);
      }
      return undefined;
    };
    video.src = fileUrl(path);
  });
}

/**
 * The frame's object URL for a clip, drawn once per window; null when the
 * clip will not play here either. `store` is given the JPEG bytes to keep.
 */
export function frameOf(path, store) {
  if (results.has(path)) return results.get(path);
  const job = chain.then(() => grab(path)).then(async (blob) => {
    if (!blob) return null;
    if (store) {
      try {
        await store(new Uint8Array(await blob.arrayBuffer()));
      } catch {
        /* the platform keeps it if it can */
      }
    }
    return URL.createObjectURL(blob);
  });
  chain = job.catch(() => null);
  results.set(path, job);
  return job;
}
