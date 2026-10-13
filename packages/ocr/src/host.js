// Text recognition on a pool of threads of its own.
//
// Reading a page takes from a fraction of a second to several, and the
// suite's main process serves every window: recognition there would hold
// them all, as the password hashing of 1.44 did. So it never runs there,
// nor on a window's thread: each page goes to one of a few worker threads
// (worker.js), one page a thread at a time, the rest waiting in a queue in
// the order they were asked for, so a scan of three hundred pages takes
// its turn and holds nothing else up. A thread keeps its engine and its
// last languages between pages; a job is given, when it can be, to a free
// thread that already holds its languages.
//
// Unlike the password pool there is no fallback to the calling thread: a
// thread that cannot start, or dies, fails the jobs it held with an error
// the window can show.

import os from 'node:os';
import { Worker } from 'node:worker_threads';
import { checkLanguages } from './languages.js';

const WORKER = new URL('./worker.js', import.meta.url);

/** A thread with nothing to do for this long is stopped, its memory given back. */
const IDLE_MS = 60000;

/**
 * One thread a core at most, and never more than four: each holds an engine
 * of 40 to 200 MB (docs/OCR.md), and the machine keeps a core of its own.
 */
export const defaultThreads = () => Math.max(1, Math.min(4, Math.floor((os.availableParallelism?.() ?? os.cpus().length) / 2), (os.availableParallelism?.() ?? os.cpus().length) - 1));

export class OcrCancelledError extends Error {
  constructor() {
    super('Text recognition was cancelled.');
    this.name = 'OcrCancelledError';
    this.code = 'OCR_CANCELLED';
  }
}

function rebuilt(error) {
  return Object.assign(new Error(error?.message || 'Text recognition failed.'), { name: error?.name || 'Error', code: error?.code });
}

/** The pixels as a copy the thread may own, so the caller keeps its own. */
function portable(image) {
  if (image instanceof ArrayBuffer) return { image: image.slice(0), transfer: [] };
  if (ArrayBuffer.isView(image)) {
    const copy = new Uint8Array(image.byteLength);
    copy.set(new Uint8Array(image.buffer, image.byteOffset, image.byteLength));
    return { image: copy, transfer: [copy.buffer] };
  }
  const data = image?.data;
  if (!data) return { image, transfer: [] };
  const copy = new Uint8Array(data.length);
  copy.set(data);
  return { image: { width: image.width, height: image.height, channels: image.channels, data: copy }, transfer: [copy.buffer] };
}

/**
 * @param {{ threads?: number, vendor?: string, idleMs?: number }} [options]
 *   `vendor` the folder holding core/ and tessdata/ (the package's own by default)
 */
export function createOcrHost({ threads = defaultThreads(), vendor, idleMs = IDLE_MS } = {}) {
  let seq = 0;
  /** { thread, job, languages, timer } */
  const pool = [];
  /** jobs waiting: { id, args, transfer, resolve, reject, signal, onAbort } */
  const queue = [];
  const stats = { done: 0, failed: 0, cancelled: 0, started: 0 };

  const settle = (job) => {
    if (job.signal && job.onAbort) job.signal.removeEventListener('abort', job.onAbort);
  };

  const idleLater = (entry) => {
    clearTimeout(entry.timer);
    if (entry.job) return;
    entry.timer = setTimeout(() => {
      if (entry.job) return;
      drop(entry);
      entry.thread.terminate().catch(() => {});
    }, idleMs);
    entry.timer.unref?.();
  };

  function drop(entry) {
    const at = pool.indexOf(entry);
    if (at >= 0) pool.splice(at, 1);
    clearTimeout(entry.timer);
  }

  function startThread() {
    const entry = { thread: new Worker(WORKER, { workerData: { vendor } }), job: null, languages: '', timer: null };
    stats.started++;
    entry.thread.unref?.();
    entry.thread.on('message', ({ id, ok, value, error }) => {
      const job = entry.job;
      if (!job || job.id !== id) return;
      entry.job = null;
      settle(job);
      if (ok) {
        stats.done++;
        job.resolve(value);
      } else {
        stats.failed++;
        job.reject(rebuilt(error));
      }
      idleLater(entry);
      pump();
    });
    const lost = (err) => {
      drop(entry);
      const job = entry.job;
      entry.job = null;
      if (job && !job.cancelled) {
        settle(job);
        stats.failed++;
        job.reject(Object.assign(new Error('Text recognition stopped unexpectedly.'), { code: 'OCR_THREAD', cause: err }));
      }
      pump();
    };
    entry.thread.on('error', lost);
    entry.thread.on('exit', () => lost());
    pool.push(entry);
    return entry;
  }

  /** Hand waiting jobs to free threads, starting threads while the pool is short. */
  function pump() {
    while (queue.length) {
      const job = queue[0];
      const free = pool.filter((e) => !e.job);
      let entry = free.find((e) => e.languages === job.key) || free[0];
      if (!entry && pool.length < threads) {
        try {
          entry = startThread();
        } catch (err) {
          queue.shift();
          settle(job);
          stats.failed++;
          job.reject(Object.assign(new Error('Text recognition could not start.'), { code: 'OCR_THREAD', cause: err }));
          continue;
        }
      }
      if (!entry) return;
      queue.shift();
      clearTimeout(entry.timer);
      entry.job = job;
      entry.languages = job.key;
      try {
        entry.thread.postMessage({ id: job.id, args: job.args }, job.transfer);
      } catch (err) {
        entry.job = null;
        settle(job);
        stats.failed++;
        job.reject(err);
      }
    }
  }

  function cancel(job) {
    if (job.cancelled) return;
    job.cancelled = true;
    stats.cancelled++;
    settle(job);
    const waiting = queue.indexOf(job);
    if (waiting >= 0) queue.splice(waiting, 1);
    const running = pool.find((e) => e.job === job);
    if (running) {
      // A page part read cannot be stopped inside the engine: the thread goes, and a new one starts for the next job.
      running.job = null;
      drop(running);
      running.thread.terminate().catch(() => {});
    }
    job.reject(new OcrCancelledError());
    pump();
  }

  return {
    /**
     * Read the text in a picture.
     * @param image `{ width, height, data }` pixels (1 to 4 channels, 8 bits) or a PNG's bytes
     * @param {{ languages?: string[] | string, dpi?: number, mode?: string, signal?: AbortSignal }} [options]
     * @returns {Promise<{ text, lines, words, width, height, languages, angle, ms }>}
     */
    recognize(image, { languages, dpi, mode, signal, ...rest } = {}) {
      let langs;
      try {
        langs = checkLanguages(languages);
      } catch (err) {
        return Promise.reject(err);
      }
      if (signal?.aborted) return Promise.reject(new OcrCancelledError());
      const { image: own, transfer } = portable(image);
      return new Promise((resolve, reject) => {
        const job = { id: ++seq, key: langs.join('+'), args: { ...rest, image: own, languages: langs, dpi, mode }, transfer, resolve, reject, signal };
        if (signal) {
          job.onAbort = () => cancel(job);
          signal.addEventListener('abort', job.onAbort, { once: true });
        }
        queue.push(job);
        pump();
      });
    },

    /** Threads running, jobs waiting and reading, and counts so far; for the checks. */
    get stats() {
      return { ...stats, threads: pool.length, busy: pool.filter((e) => e.job).length, queued: queue.length, limit: threads };
    },

    /** Stop every thread and fail what waits; a later job starts them again. */
    async stop() {
      for (const job of queue.splice(0)) {
        settle(job);
        job.reject(new OcrCancelledError());
      }
      await Promise.all(pool.slice().map((entry) => {
        const job = entry.job;
        entry.job = null;
        if (job) {
          settle(job);
          job.reject(new OcrCancelledError());
        }
        drop(entry);
        return entry.thread.terminate().catch(() => {});
      }));
    },
  };
}
