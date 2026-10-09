// A frame of a clip, or a picture at a tile's size, drawn by the window.
//
// The platform makes most thumbnails (thumbs.js in the shell): the system
// knows the videos it has a codec for. For the ones it does not — a WebM on
// Windows without the codec pack, anything on Linux — the window can draw a
// frame itself, since Chromium plays what the platform will not. One clip at
// a time, well into the clip rather than its black first frame, and the
// result handed to the platform so the next visit finds it in the cache.
// A picture the platform had no tile for (an SVG, an AVIF, a large photo on
// Linux) is drawn the same way, decoded off the page's thread at the tile's
// size, so a 40-megapixel photograph is never held whole.

import { fileUrl } from './library.js';

const SIZE = 256;
const TIMEOUT = 8000;

const results = new Map(); // path -> Promise<string|null> (an object URL, or nothing)
/** How many clips' frames are kept; past it the least lately asked for is let go, its object URL with it. */
const KEPT = 300;
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

/** A canvas drawn at the tile's size, as JPEG. */
function tileOf(source, w, h) {
  const scale = Math.min(1, SIZE / Math.max(w, h));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(w * scale));
  canvas.height = Math.max(1, Math.round(h * scale));
  const g = canvas.getContext('2d');
  // A picture with see-through parts on white, as the system draws one.
  g.fillStyle = '#fff';
  g.fillRect(0, 0, canvas.width, canvas.height);
  g.drawImage(source, 0, 0, canvas.width, canvas.height);
  return new Promise((resolve) => canvas.toBlob((blob) => resolve(blob || null), 'image/jpeg', 0.82));
}

/** A picture at the tile's size: decoded at that size where the format lets it, and never shown whole. */
async function grabStill(path) {
  let timer;
  const late = new Promise((resolve) => { timer = setTimeout(() => resolve(null), TIMEOUT); });
  const work = (async () => {
    const res = await fetch(fileUrl(path));
    if (!res.ok) return null;
    const blob = await res.blob();
    if (/\.svg$/i.test(path)) {
      // A drawing has no pixels to shrink: it is drawn at the tile's size.
      const url = URL.createObjectURL(new Blob([blob], { type: 'image/svg+xml' }));
      try {
        const img = new Image();
        img.src = url;
        await img.decode();
        const w = img.naturalWidth || SIZE;
        const h = img.naturalHeight || SIZE;
        return await tileOf(img, w, h);
      } finally {
        URL.revokeObjectURL(url);
      }
    }
    const bitmap = await createImageBitmap(blob, { resizeWidth: SIZE, resizeQuality: 'medium' });
    try {
      return await tileOf(bitmap, bitmap.width, bitmap.height);
    } finally {
      bitmap.close();
    }
  })().catch(() => null);
  try {
    return await Promise.race([work, late]);
  } finally {
    clearTimeout(timer);
  }
}

/**
 * The frame's object URL for a clip, or a picture's at a tile's size, drawn
 * once per window; null when it will not decode here either. `store` is
 * given the JPEG bytes to keep; `kind` is the file's kind (library.js).
 */
export function frameOf(path, store, kind = 'video') {
  if (results.has(path)) {
    // Asked for again: the most lately used, so the last to go.
    const kept = results.get(path);
    results.delete(path);
    results.set(path, kept);
    return kept;
  }
  while (results.size >= KEPT) {
    const [oldest, job] = results.entries().next().value;
    results.delete(oldest);
    job.then((url) => { if (url) URL.revokeObjectURL(url); }).catch(() => {});
  }
  const job = chain.then(() => (kind === 'video' ? grab(path) : grabStill(path))).then(async (blob) => {
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
