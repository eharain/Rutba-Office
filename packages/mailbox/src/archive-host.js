// Archives opened in place, each on a thread of its own.
//
// A ten-year .pst holds a hundred thousand messages, and indexing them takes
// half a minute of steady work. The suite's main process serves every
// window, so that work never runs there: each archive opened is given a
// worker thread (archive-worker.js) that opens the file, reads it for the
// window, indexes it and searches its index. The main process only passes
// messages to and fro - the longest it is held is the copying of one answer.
//
// One thread an archive, not a shared pool: an archive's index lives in its
// thread, so the thread that indexed it is the one that must search it.
// Closing an archive stops its thread and gives its memory back. A thread
// that dies fails what it was asked with an error the window can show, and
// the next question about that archive starts a new one.

import { Worker } from 'node:worker_threads';

const WORKER = new URL('./archive-worker.js', import.meta.url);

function rebuilt(error) {
  return Object.assign(new Error(error?.message || 'The archive could not be read.'), { name: error?.name || 'Error', code: error?.code });
}

/**
 * @param {{ onProgress?: (id: string, p: { state: string, done: number, total: number, error?: string|null }) => void, slice?: number, part?: number }} [options]
 *   `slice` ms of indexing between a thread's messages; `part` messages a
 *   searchable part of a folder being indexed
 */
export function createArchiveHost({ onProgress = () => {}, slice, part } = {}) {
  /** id -> { worker, file, name, seq, pending: Map, opened: Promise, busy: boolean, closing: boolean } */
  const threads = new Map();
  const stats = { started: 0, lost: 0 };

  // A thread keeps the process alive only while it is asked something or indexing.
  const hold = (entry) => {
    if (entry.pending.size || entry.busy) entry.worker.ref?.();
    else entry.worker.unref?.();
  };

  function send(entry, op, args) {
    return new Promise((resolve, reject) => {
      const seq = ++entry.seq;
      entry.pending.set(seq, { resolve, reject });
      hold(entry);
      try {
        entry.worker.postMessage({ seq, op, args });
      } catch (err) {
        entry.pending.delete(seq);
        hold(entry);
        reject(err);
      }
    });
  }

  function start(id, file, name, kind) {
    const worker = new Worker(WORKER, { workerData: { slice, part } });
    stats.started++;
    const entry = { worker, file, name, seq: 0, pending: new Map(), opened: null, busy: false, closing: false };
    worker.on('message', (msg) => {
      if (msg.event === 'progress') {
        const { event, ...p } = msg;
        entry.busy = p.state === 'indexing';
        hold(entry);
        onProgress(id, p);
        return;
      }
      const job = entry.pending.get(msg.seq);
      if (!job) return;
      entry.pending.delete(msg.seq);
      hold(entry);
      if (msg.ok) job.resolve(msg.value);
      else job.reject(rebuilt(msg.error));
    });
    let gone = false;
    const lost = (err) => {
      if (gone) return;
      gone = true;
      if (threads.get(id) === entry) threads.delete(id);
      const jobs = [...entry.pending.values()];
      entry.pending.clear();
      if (entry.closing) return;
      stats.lost++;
      const error = Object.assign(new Error('The archive stopped being read unexpectedly. Open it again to carry on.'), { code: 'ARCHIVE_THREAD', cause: err });
      for (const job of jobs) job.reject(error);
      onProgress(id, { state: 'failed', done: 0, total: 0, error: error.message });
    };
    worker.on('error', lost);
    worker.on('exit', () => lost());
    threads.set(id, entry);
    entry.opened = send(entry, 'open', { path: file, id, name, kind });
    // A file that will not open takes its thread with it.
    entry.opened.catch(() => {
      entry.closing = true;
      if (threads.get(id) === entry) threads.delete(id);
      worker.terminate().catch(() => {});
    });
    return entry;
  }

  /** The archive's thread, started (and the file opened) if it is not running. */
  function thread(id, file, name, kind) {
    let entry = threads.get(id);
    if (!entry) {
      if (!file) return Promise.reject(new Error('That archive is not open.'));
      entry = start(id, file, name, kind);
    }
    return entry.opened.then(() => entry);
  }

  const host = {
    /**
     * Open an archive on a thread of its own.
     * @returns {Promise<{ kind, label, bytes, folders }>}
     */
    async open(id, file, name) {
      const entry = threads.get(id);
      if (entry && entry.file !== file) await host.close(id);
      return (await thread(id, file, name)).opened;
    },

    /**
     * Ask an archive's thread: `folders`, `list`, `message`, `load`, `text`,
     * `search`, `index`, `pause`, `status`. `where` names the file, so a
     * thread that is not running (closed, or lost) can be started again;
     * `where.kind: 'accounts'` makes it the accounts' thread, `file` the
     * mail store's folder (`sync`, `search`, `drop`).
     */
    async call(id, op, args = {}, where = null) {
      const entry = await thread(id, where?.file, where?.name, where?.kind);
      return send(entry, op, args);
    },

    /** Is a thread running for this archive. */
    has: (id) => threads.has(id),

    /** Stop an archive's thread. Its file is not touched. */
    async close(id) {
      const entry = threads.get(id);
      if (!entry) return;
      threads.delete(id);
      entry.closing = true;
      const jobs = [...entry.pending.values()];
      entry.pending.clear();
      for (const job of jobs) job.reject(new Error('That archive was closed.'));
      await entry.worker.terminate().catch(() => {});
    },

    /** Threads running, for the checks. */
    get stats() {
      return { ...stats, threads: threads.size };
    },

    /** Stop every thread. */
    async stop() {
      await Promise.all([...threads.keys()].map((id) => host.close(id)));
    },
  };
  return host;
}
