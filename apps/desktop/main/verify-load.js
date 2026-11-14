// How busy the machine was over a wait that gave up, said beside its cause in
// the log and the failure list, so a wait that ran out on a busy machine reads
// as one (a gate that co-exists, 13 Nov 2026: the load beside every timed
// result).
//
// The processors are read in a worker every two seconds, never on the main
// thread: on a busy Windows machine os.cpus() can hold its caller for over a
// second (review of 14 Nov: p90 615 ms, max 1,363 ms at 100% busy), and on
// Electron's main thread that would hold back the very replies a wait waits
// for. A wait notes when it began; at give-up its share is read between the
// worker's reading nearest before that moment and its latest. The share is the
// whole machine's, the app under test included.

import os from 'node:os';
import { Worker } from 'node:worker_threads';

const EVERY_MS = 2000;
// About four minutes of readings: longer than any wait (SIGNAL_WAIT is a minute).
const KEPT = 128;

const readings = [];
let worker = null;
let failed = false;

const SAMPLER = `
const { parentPort } = require('node:worker_threads');
const os = require('node:os');
// Stamped once read: on a busy machine os.cpus() can take over a second.
const say = () => { const times = os.cpus().map((c) => c.times); parentPort.postMessage({ at: Date.now(), times }); };
say();
setInterval(say, ${EVERY_MS});
`;

/** The worker started once, unref'd so it never keeps a run alive. A worker that cannot start leaves the load unsaid. */
function sampling() {
  if (worker || failed) return;
  try {
    worker = new Worker(SAMPLER, { eval: true });
    worker.on('message', (r) => {
      readings.push(r);
      if (readings.length > KEPT) readings.shift();
    });
    worker.on('error', () => { failed = true; worker = null; });
    // Last: a 'message' listener added after unref() refs the worker's port again, and the run would never end.
    worker.unref();
  } catch {
    failed = true;
  }
}

/** The processors' busy share between two readings of each core's times, 0 to 1, or null when nothing could be compared. */
export function busyBetween(from, to) {
  const sum = (t) => Object.values(t).reduce((a, b) => a + b, 0);
  let idle = 0;
  let all = 0;
  to.forEach((t, i) => {
    if (!from[i]) return;
    all += sum(t) - sum(from[i]);
    idle += t.idle - from[i].idle;
  });
  return all > 0 ? Math.min(1, Math.max(0, 1 - idle / all)) : null;
}

/**
 * The load over the time since `since` (ms), from `kept` readings
 * ({ at, times }, oldest first): the busy share between the last reading at
 * or before `since` (the oldest if none is) and the latest, as a whole
 * percentage, or null with fewer than two readings to compare; and the memory
 * free now, in GiB to one place.
 */
export function loadFrom(since, kept, { freemem = os.freemem } = {}) {
  const freeGb = Math.round((freemem() / 2 ** 30) * 10) / 10;
  if (kept.length < 2) return { busy: null, freeGb };
  let from = kept[0];
  for (const r of kept) if (r.at <= since) from = r;
  const to = kept[kept.length - 1];
  const share = from === to ? null : busyBetween(from.times, to.times);
  return { busy: share === null ? null : Math.round(share * 100), freeGb };
}

/** Called as a wait begins: starts the sampler if it has not, and answers the moment to measure from. */
export function loadMark() {
  sampling();
  return Date.now();
}

/** The load since a mark, from the sampler's readings so far. */
export const loadSince = (mark) => loadFrom(mark, readings);

/** The load as the log and the failure list say it ('' for none). */
export function saidLoad(load) {
  if (!load) return '';
  const free = `${load.freeGb} GB free`;
  return load.busy === null ? `the processors not read over the wait, ${free}` : `the processors ${load.busy}% busy over the wait, ${free}`;
}
