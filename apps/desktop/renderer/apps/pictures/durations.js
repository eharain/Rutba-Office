// A clip's length, read once and kept for the life of the window.
//
// A tile's picture is a JPEG the platform already made (thumbs.js, in the
// main process, or a frame the window drew itself — frames.js); nothing
// about either says how long the clip runs. Only a <video> or <audio>
// element loading the file's metadata can say that, so this reads it once
// per clip, a few at a time as their tiles come into view — the platform's
// own thumbnailer makes its JPEGs the same way — and keeps what it found,
// so a tile drawn again never asks twice.

import { fileUrl } from './library.js';

const CONCURRENCY = 3;
const TIMEOUT = 8000;

const jobs = new Map(); // path -> Promise<number|null>
const found = new Map(); // path -> number|null, once that promise has settled
const waiting = [];
let running = 0;

function pump() {
  while (running < CONCURRENCY && waiting.length) {
    const job = waiting.shift();
    running += 1;
    job().finally(() => {
      running -= 1;
      pump();
    });
  }
}

function read(path, kind) {
  return new Promise((resolve) => {
    const el = document.createElement(kind === 'audio' ? 'audio' : 'video');
    el.muted = true;
    el.preload = 'metadata';
    let done = false;
    const finish = (value) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      el.removeAttribute('src');
      el.load();
      resolve(value);
    };
    const timer = setTimeout(() => finish(null), TIMEOUT);
    el.onerror = () => finish(null);
    el.onloadedmetadata = () => finish(Number.isFinite(el.duration) && el.duration >= 0 ? el.duration : null);
    el.src = fileUrl(path);
  });
}

/**
 * A clip's length in seconds, read once per window; null when it could not
 * be read. `kind` is `'video'` or `'audio'`, as `kindOf` (library.js) names
 * it.
 */
export function durationOf(path, kind) {
  if (jobs.has(path)) return jobs.get(path);
  const job = new Promise((resolve) => {
    waiting.push(() => read(path, kind).then((value) => {
      found.set(path, value);
      resolve(value);
    }));
    pump();
  });
  jobs.set(path, job);
  return job;
}

/**
 * The length already found for a path, without starting a read —
 * `undefined` when nothing has asked for it yet or it is still being read.
 */
export function peekLength(path) {
  return found.get(path);
}
