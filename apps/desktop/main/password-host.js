// Password hashing off the document service's thread.
//
// Every password Office writes is hashed 100,000 times: Excel's Protect
// Sheet and Protect Workbook, Word's Restrict Editing, and the key of a
// file encrypted with a password. A third of a second on a quiet machine,
// several seconds on a busy one, and the document service answers every
// window on one thread: while one window checked a password, every other
// window's typing waited behind it.
//
// The hashes are worked out here instead, on a small pool of threads
// (password-worker.js) running the very functions the engines run, so
// what comes back is what the engine would have made. The document
// service hands each answer to the engine before the operation that wants
// it runs (protection.js `rememberPasswordHash`, crypt.js's key cache), and
// its own thread is free the whole time.
//
// Where a thread cannot start, the work is done in place, as before.

import os from 'node:os';
import { Worker } from 'node:worker_threads';
import { hashPassword } from '@rutba/ooxml/protection';
import { spinPassword, decryptPackage, EncryptedFileError } from '@rutba/office-formats/crypt';

const WORKER = new URL('./password-worker.js', import.meta.url);

/** A thread with nothing to do for this long is stopped; the next job starts one again. */
const IDLE_MS = 30000;

/** Two threads at most: two windows can check at once, and the rest of the machine keeps its cores. */
const defaultThreads = () => Math.max(1, Math.min(2, (os.availableParallelism?.() ?? os.cpus().length) - 1));

/** The same jobs, worked out on the calling thread. */
const IN_PLACE = {
  hash: ({ password, salt, spinCount, algorithm }) => hashPassword(password, { salt, spinCount, algorithm }),
  spin: ({ password, salt, spinCount, alg }) => spinPassword(password, Buffer.from(salt), spinCount, alg),
  decrypt: ({ bytes, password }) => decryptPackage(Buffer.from(bytes), password),
};

/** An error from the thread as the engine would have thrown it. */
function rebuilt(error) {
  if (error?.name === 'EncryptedFileError') return new EncryptedFileError(error.code, error.message);
  return Object.assign(new Error(error?.message || 'The password could not be worked out.'), { name: error?.name || 'Error', code: error?.code });
}

const asBuffer = (bytes) => (Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength));

/**
 * @param {{ threads?: number, worker?: boolean }} [options] `worker: false`
 *   works everything out in place (a test's comparison, a platform that
 *   refuses threads)
 */
export function createPasswordHost({ threads = defaultThreads(), worker = true } = {}) {
  let failed = !worker;
  let seq = 0;
  /** { thread, jobs: Map(id -> job), timer } */
  const pool = [];
  const stats = { threaded: 0, inPlace: 0 };

  const retire = (entry) => {
    const at = pool.indexOf(entry);
    if (at >= 0) pool.splice(at, 1);
    clearTimeout(entry.timer);
    return entry.jobs;
  };

  const idleLater = (entry) => {
    clearTimeout(entry.timer);
    if (entry.jobs.size) return;
    entry.timer = setTimeout(() => {
      if (entry.jobs.size) return;
      retire(entry);
      entry.thread.terminate().catch(() => {});
    }, IDLE_MS);
    entry.timer.unref?.();
  };

  function startThread() {
    const entry = { thread: new Worker(WORKER), jobs: new Map(), timer: null };
    entry.thread.unref?.();
    entry.thread.on('message', ({ id, ok, value, error }) => {
      const job = entry.jobs.get(id);
      if (!job) return;
      entry.jobs.delete(id);
      idleLater(entry);
      if (ok) job.resolve(value);
      else job.reject(rebuilt(error));
    });
    const lost = () => {
      // Whatever it held is worked out here instead: an answer is still owed.
      const owed = [...retire(entry).values()];
      entry.jobs.clear();
      for (const job of owed) runInPlace(job.name, job.args).then(job.resolve, job.reject);
    };
    entry.thread.on('error', () => {
      failed = true;
      lost();
    });
    entry.thread.on('exit', lost);
    pool.push(entry);
    return entry;
  }

  /** The thread with the least to do, a new one while the pool is short of `threads`. */
  function pick() {
    const free = pool.find((e) => !e.jobs.size);
    if (free) return free;
    if (pool.length < threads) return startThread();
    return pool.reduce((a, b) => (b.jobs.size < a.jobs.size ? b : a));
  }

  function runInPlace(name, args) {
    stats.inPlace++;
    return new Promise((resolve, reject) => {
      // A turn of the loop first, so the caller's promise is handed out before the work holds the thread.
      setImmediate(() => {
        try {
          resolve(IN_PLACE[name](args));
        } catch (err) {
          reject(err);
        }
      });
    });
  }

  async function run(name, args, transfer = []) {
    if (!failed) {
      let entry = null;
      try {
        entry = pick();
      } catch {
        failed = true;
      }
      if (entry) {
        stats.threaded++;
        return new Promise((resolve, reject) => {
          const id = ++seq;
          clearTimeout(entry.timer);
          entry.jobs.set(id, { name, args, resolve, reject });
          try {
            entry.thread.postMessage({ id, job: name, args }, transfer);
          } catch (err) {
            entry.jobs.delete(id);
            reject(err);
          }
        });
      }
    }
    return runInPlace(name, args);
  }

  return {
    /** `hashPassword`'s answer for this spec: { password, salt, spinCount, algorithm }. */
    hash: (spec) => run('hash', { password: String(spec.password), salt: spec.salt, spinCount: spec.spinCount, algorithm: spec.algorithm }),

    /** `spinPassword`'s answer: the spun hash of a password, as a Buffer. */
    spin: async ({ password, salt, spinCount, alg = 'sha512' }) => asBuffer(await run('spin', { password: String(password), salt: new Uint8Array(salt), spinCount, alg })),

    /**
     * `decryptPackage`'s answer for an encrypted file: its package, or the
     * same EncryptedFileError the engine throws. The bytes are copied to
     * the thread, never handed over: the caller keeps its own.
     */
    decrypt: async (bytes, password) => {
      const out = await run('decrypt', { bytes: new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength), password: String(password) });
      return { ...out, bytes: asBuffer(out.bytes) };
    },

    /** How many jobs ran on a thread and how many in place; for the checks. */
    get stats() { return { ...stats, threads: pool.length }; },

    /** Stop every thread; a later job starts one again. */
    async stop() {
      await Promise.all(pool.slice().map((entry) => entry.thread.terminate().catch(() => {})));
    },
  };
}

/** The one pool the document service uses. */
let shared = null;
export function passwordHost() {
  shared ||= createPasswordHost();
  return shared;
}
