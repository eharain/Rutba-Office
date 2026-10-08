// Thumbnails, made once and kept.
//
// A grid tile that shows a 40-megapixel photograph by decoding the whole
// photograph costs as much as opening it, and a folder of five hundred costs
// five hundred openings — which is what made the viewer crawl on a large
// folder and stutter as tiles came and went. A tile of a video was worse: a
// live <video> element per file, each holding a decoder, and Chromium
// suspending and resuming them as the grid scrolled.
//
// This makes a small JPEG of a file instead — the operating system's own
// thumbnail (the shell on Windows, QuickLook on macOS), which knows videos,
// HEIC and whatever else the machine has a codec for, and which the system
// caches itself — and keeps it on disk keyed by the file's path, size and
// modification time, so the second visit to a folder costs a directory read.
// A window can also hand a thumbnail in (`put`) for a file the system could
// not do: a WebM on Windows without the codec pack, or anything on Linux,
// where the window draws a frame of the clip itself.
//
// Requests are answered newest first: the tiles a person is looking at now
// are the ones that were asked for last, and the ones scrolled past can wait.

import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';

export const THUMB_SIZE = 256;
const QUALITY = 82;

/** The cache file for a path at a size: the same file, unchanged, is the same thumbnail. */
export function cacheKey(target, { size, mtimeMs, fileSize }) {
  return crypto.createHash('sha1').update(`${target}|${size}|${Math.round(mtimeMs)}|${fileSize}`).digest('hex');
}

/**
 * The system's thumbnail of a file, as JPEG bytes, or null when the system
 * has none for it. Windows and macOS answer for pictures and for the videos
 * they have a codec for; Linux answers for nothing, and a picture is then
 * read and shrunk by Electron's own decoder, which knows PNG and JPEG.
 */
async function systemThumbnail(target, size) {
  const { nativeImage } = await import('electron');
  let image = null;
  try {
    image = await nativeImage.createThumbnailFromPath(target, { width: size, height: size });
  } catch {
    image = null;
  }
  if (!image || image.isEmpty()) {
    if (!/\.(png|jpe?g|jpe)$/i.test(target)) return null;
    try {
      const whole = nativeImage.createFromPath(target);
      if (whole.isEmpty()) return null;
      const { width, height } = whole.getSize();
      const scale = Math.min(1, size / Math.max(width, height));
      image = scale < 1 ? whole.resize({ width: Math.max(1, Math.round(width * scale)), height: Math.max(1, Math.round(height * scale)), quality: 'good' }) : whole;
    } catch {
      return null;
    }
  }
  const bytes = image.toJPEG(QUALITY);
  return bytes.length ? bytes : null;
}

const CACHE_MAX = 500 * 1024 * 1024;
const CACHE_TRIM = 400 * 1024 * 1024;

/** The thumbnail cache brought under its size, the least lately made going first. Answers how many went. */
export async function trimCache(dir, { max = CACHE_MAX, trim = CACHE_TRIM } = {}) {
  const names = (await fsp.readdir(dir)).filter((n) => n.endsWith('.jpg'));
  const files = [];
  let total = 0;
  for (let i = 0; i < names.length; i += 64) {
    for (const f of await Promise.all(names.slice(i, i + 64).map(async (n) => {
      const s = await fsp.stat(path.join(dir, n)).catch(() => null);
      return s ? { file: path.join(dir, n), size: s.size, at: s.mtimeMs } : null;
    }))) if (f) { files.push(f); total += f.size; }
  }
  if (total <= max) return 0;
  files.sort((a, b) => a.at - b.at);
  let gone = 0;
  for (const f of files) {
    if (total <= trim) break;
    await fsp.rm(f.file, { force: true }).catch(() => {});
    total -= f.size;
    gone += 1;
  }
  return gone;
}

/**
 * A queue that runs the newest job first, a few at a time.
 *
 * Newest first because a request is made when a tile comes into view: the
 * newest requests are the tiles in view now. A job is run even when the
 * request that asked for it was abandoned (the tile scrolled away before
 * its turn): the thumbnail is kept on disk for the next time, and a second
 * request for the same file, sharing the job, must not inherit the first
 * one's abandonment. Only the newest `keep` wait, though: a folder of
 * thousands scrolled through quickly queued a job for every tile it ever
 * passed; the oldest — the tiles long gone from view — are let go, answered
 * with nothing, and asked again if their tile comes back.
 */
export function createQueue({ concurrency = 3, run, keep = 200 }) {
  const waiting = [];
  let running = 0;
  const pump = () => {
    while (running < concurrency && waiting.length) {
      const job = waiting.pop();
      running += 1;
      Promise.resolve()
        .then(() => run(job))
        .then(job.resolve, () => job.resolve(null))
        .finally(() => {
          running -= 1;
          pump();
        });
    }
  };
  return {
    push: (job) => new Promise((resolve) => {
      waiting.push({ ...job, resolve });
      while (waiting.length > keep) waiting.shift().resolve(null);
      pump();
    }),
    get pending() {
      return waiting.length;
    },
    get running() {
      return running;
    },
  };
}

/**
 * @param {object} o
 * @param {string} o.dir where the JPEGs are kept
 * @param {(target: string, size: number) => Promise<Buffer|null>} [o.make] the thumbnail maker; the system's by default
 * @param {number} [o.concurrency] how many are made at once
 */
export function createThumbnailer({ dir, make = systemThumbnail, concurrency = 3 }) {
  fs.mkdirSync(dir, { recursive: true });
  // Files the system said no to, so a grid of them is not asked again on
  // every scroll. Keyed like the cache: a changed file is asked afresh.
  const refused = new Set();
  const inflight = new Map();
  const queue = createQueue({
    concurrency,
    run: async ({ target, size, key }) => {
      const bytes = await make(target, size);
      if (!bytes) {
        refused.add(key);
        return null;
      }
      await fsp.writeFile(path.join(dir, `${key}.jpg`), bytes).catch(() => {});
      return bytes;
    },
  });

  const keyFor = async (target, size) => {
    const stat = await fsp.stat(target);
    if (stat.isDirectory()) return null;
    return cacheKey(target, { size, mtimeMs: stat.mtimeMs, fileSize: stat.size });
  };

  // The kept JPEGs held to a size: past CACHE_MAX the oldest go, down to
  // CACHE_TRIM, looked at once a run, after the window is up.
  setTimeout(() => { trimCache(dir).catch(() => {}); }, 30000).unref?.();

  /** The JPEG bytes for a file, from the cache or made now; null when there are none to be had. */
  const bytesFor = async (target, { size = THUMB_SIZE } = {}) => {
    let key;
    try {
      key = await keyFor(target, size);
    } catch {
      return null;
    }
    if (!key) return null;
    const file = path.join(dir, `${key}.jpg`);
    try {
      return await fsp.readFile(file);
    } catch {
      /* not made yet */
    }
    if (refused.has(key)) return null;
    if (inflight.has(key)) return inflight.get(key);
    const job = queue.push({ target, size, key }).finally(() => inflight.delete(key));
    inflight.set(key, job);
    return job;
  };

  return {
    dir,
    bytesFor,
    /** An HTTP answer for the protocol: the JPEG, or 404 with a header that says why. */
    respond: async (target, options) => {
      const bytes = await bytesFor(target, options);
      if (!bytes) return new Response('no thumbnail', { status: 404, headers: { 'x-rutba-thumb': 'none' } });
      return new Response(bytes, {
        status: 200,
        headers: { 'content-type': 'image/jpeg', 'content-length': String(bytes.length), 'cache-control': 'max-age=86400' },
      });
    },
    /** A thumbnail a window made itself, kept as if the system had made it. */
    put: async (target, bytes, { size = THUMB_SIZE } = {}) => {
      const key = await keyFor(target, size).catch(() => null);
      if (!key) return { stored: false };
      await fsp.writeFile(path.join(dir, `${key}.jpg`), Buffer.from(bytes));
      refused.delete(key);
      return { stored: true };
    },
    /** For the checks: how many are waiting and how many are being made. */
    get load() {
      return { pending: queue.pending, running: queue.running };
    },
  };
}
