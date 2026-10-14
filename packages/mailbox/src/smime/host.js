// The S/MIME thread, with a time and a memory budget.
//
// Reading a signed or encrypted message, or opening a .p12, runs on one
// worker thread (worker.js), one job at a time, the rest waiting in turn. A
// job that runs past its time is ended with the thread, and a thread that
// runs out of its memory dies alone; either way the job fails with a code the
// window can say ("this message's signature or encryption could not be
// read"), a fresh thread takes the next job, and the rest of the application
// carries on.

import { Worker } from 'node:worker_threads';

const WORKER = new URL('./worker.js', import.meta.url);

/** How long one message may take to read. A real one takes milliseconds. */
export const READ_MS = 10_000;
/** How long a .p12 may take to open: its key is derived as many times as it says, within the cap. */
export const PKCS12_MS = 60_000;
/** The thread's heap. A message's structures are capped well below this (der.js). */
const HEAP_MB = 512;

export class SmimeTimeoutError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'SmimeTimeoutError';
    this.code = code;
  }
}

export function createSmimeHost({ readMs = READ_MS, pkcs12Ms = PKCS12_MS, heapMb = HEAP_MB } = {}) {
  let thread = null;
  let seq = 0;
  let current = null; // { seq, resolve, reject, timer }
  const queue = [];
  const stats = { started: 0, timeouts: 0, crashes: 0 };

  const start = () => {
    stats.started++;
    const t = new Worker(WORKER, { resourceLimits: { maxOldGenerationSizeMb: heapMb, maxYoungGenerationSizeMb: 64 } });
    t.unref?.();
    t.on('message', (msg) => {
      if (!current || msg.seq !== current.seq) return;
      const job = current;
      current = null;
      clearTimeout(job.timer);
      if (msg.ok) job.resolve(msg.value);
      else job.reject(Object.assign(new Error(msg.error?.message || 'The message could not be read.'), { name: msg.error?.name || 'Error', code: msg.error?.code }));
      pump();
    });
    const lost = () => {
      if (thread !== t) return;
      thread = null;
      if (current) {
        const job = current;
        current = null;
        clearTimeout(job.timer);
        stats.crashes++;
        job.reject(new SmimeTimeoutError('unreadable', 'The message used more memory than any real one does, so it was not read.'));
      }
      pump();
    };
    t.on('error', lost);
    t.on('exit', lost);
    return t;
  };

  function pump() {
    if (current || !queue.length) return;
    const job = queue.shift();
    if (!thread) thread = start();
    thread.ref?.();
    current = job;
    job.timer = setTimeout(() => {
      if (current !== job) return;
      current = null;
      stats.timeouts++;
      const t = thread;
      thread = null;
      t?.terminate().catch(() => {});
      job.reject(new SmimeTimeoutError('unreadable', 'The message took longer to read than any real one does, so it was not read.'));
      pump();
    }, job.ms);
    try {
      thread.postMessage({ seq: job.seq, op: job.op, args: job.args });
    } catch (err) {
      clearTimeout(job.timer);
      current = null;
      job.reject(err);
      pump();
    }
  }

  const ask = (op, args, ms) =>
    new Promise((resolve, reject) => {
      queue.push({ seq: ++seq, op, args, ms, resolve, reject });
      pump();
    }).finally(() => {
      if (!current && !queue.length) thread?.unref?.();
    });

  return {
    /** Read an S/MIME message: { message, signed, encrypted, error, certificates (DER) }. */
    read: (args) => ask('read', args, readMs),
    /** Open a .p12/.pfx: { key (PEM), cert (DER), chain (DER) }. */
    pkcs12: (args) => ask('pkcs12', args, pkcs12Ms),
    get stats() {
      return { ...stats };
    },
    async stop() {
      const t = thread;
      thread = null;
      await t?.terminate().catch(() => {});
    },
  };
}
