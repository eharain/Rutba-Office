// The main process's end of the document service, which runs on a thread of
// its own (doc-worker.js).
//
// Two faces of the same service:
//
// - `doc`, for the windows, the print service and the main process: every
//   method answers with a promise, and the main process goes on answering
//   everything else while the thread works.
// - `direct`, for the checks, which call the service as a function and read
//   its answer on the next line: every method blocks until the thread has
//   answered (a SharedArrayBuffer the thread wakes), and a method that was
//   asynchronous already still answers with a promise. Nothing a check runs
//   asks the main process to wait on it in turn, so nothing deadlocks: what
//   the thread asks of the main process (an equation laid out by Chromium)
//   is asked only from an asynchronous method, after its promise is given.
//
// The thread mints its own ids for the pictures it hands windows, and sends
// each picture's bytes before the answer that names it; the main process
// holds them for the window whose request it was, as before. If the thread
// stops, it is started again; the documents it held are gone, their
// recovery copies on disk, and a window's next request is told so.

import { Worker, MessageChannel, receiveMessageOnPort } from 'node:worker_threads';
import { fileURLToPath } from 'node:url';

const WORKER = fileURLToPath(new URL('./doc-worker.js', import.meta.url));

/**
 * Buffers again, where a copy between threads made them plain Uint8Arrays:
 * the engines and the checks were written to Node's Buffers (the PDF writer
 * refuses anything else), and a Buffer over the same bytes costs nothing.
 * Changes what it is given, which is always a fresh copy.
 */
export function asBuffers(value, seen = new Set()) {
  if (value === null || typeof value !== 'object') return value;
  if (value.constructor === Uint8Array) return Buffer.from(value.buffer, value.byteOffset, value.byteLength);
  if (ArrayBuffer.isView(value) || value instanceof ArrayBuffer || seen.has(value)) return value;
  seen.add(value);
  if (value instanceof Map) {
    for (const [k, v] of value) if (v !== null && typeof v === 'object') value.set(k, asBuffers(v, seen));
    return value;
  }
  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) if (value[i] !== null && typeof value[i] === 'object') value[i] = asBuffers(value[i], seen);
    return value;
  }
  for (const key of Object.keys(value)) if (value[key] !== null && typeof value[key] === 'object') value[key] = asBuffers(value[key], seen);
  return value;
}

/** How long a check waits on one call before saying the thread has stopped answering. */
const SYNC_LIMIT_MS = 10 * 60 * 1000;

/** How long the thread may take to load before the main process takes the service on. */
const READY_LIMIT_MS = Number(process.env.RUTBA_DOC_READY_MS || 60000);

/**
 * @param {object} o
 * @param {(bytes, type, name, options) => object} o.holdBlob the shell's, with `id` and `owner` options
 * @param {(id: string) => void} o.releaseBlob
 * @param {(id: string) => string} o.blobUrl where a held blob is fetched
 * @param {number} [o.readyLimitMs] how long the thread may take to get ready
 */
export function createDocumentHost({ holdBlob, releaseBlob, blobUrl, readyLimitMs = READY_LIMIT_MS }) {
  // What the main process does for the thread, filled in once it can.
  const handlers = { measureMath: null, teach: null, setting: null };
  let init = null;
  let thread = null;
  let generation = 0;
  let seq = 0;
  let readyNames = null;
  // Ready once, by the first thread or, if no thread can start, by the
  // service made on the main process instead (`local`).
  let readyOf;
  const firstReady = new Promise((resolve) => { readyOf = resolve; });
  let local = null;
  let fallingBack = null;
  const waiting = new Map(); // seq -> { resolve, reject } for the promised answers
  const answered = new Map(); // seq -> message, for the one call a check is blocked on
  // [0] an answer is on its way to a blocked caller; [1] the thread has ended.
  const signal = new Int32Array(new SharedArrayBuffer(8));

  // Who is waiting for every answer to be in (`idle`), let go when the last one is.
  let quiet = [];
  const settled = () => {
    const now = quiet;
    quiet = [];
    for (const resolve of now) resolve(true);
  };

  const failed = (error) => Object.assign(new Error(error?.message || 'The document service could not answer.'), { name: error?.name || 'Error', code: error?.code });

  function start() {
    generation += 1;
    Atomics.store(signal, 1, 0);
    const { port1, port2 } = new MessageChannel();
    const mine = { port: port2, worker: null, alive: true, generation };
    mine.worker = new Worker(WORKER, {
      workerData: { port: port1, signal: signal.buffer, generation, blobPrefix: blobUrl('') },
      transferList: [port1],
    });
    port2.on('message', (m) => handle(mine, m));
    // A thread started again is told what the first was.
    if (init) port2.postMessage({ t: 'init', ...init });
    mine.worker.on('error', (err) => {
      console.error('the document service stopped:', err?.stack || err?.message || err);
    });
    mine.worker.on('exit', (code) => {
      mine.alive = false;
      port2.close();
      // Every answer still awaited from it is a document that is no more.
      const gone = new Error('The document engine stopped, and the documents it held were closed. Unsaved work is kept as a recovery copy.');
      for (const [, w] of waiting) w.reject(gone);
      waiting.clear();
      settled();
      if (thread === mine && !stopping && !mine.everReady) {
        // A thread that never got ready will not get ready if started again.
        fallBack(`it ended (${code}) before it was ready`);
      } else if (thread === mine && !stopping && !fallingBack) {
        console.error(`the document service ended (${code}); starting it again`);
        thread = start();
      }
    });
    return mine;
  }

  /**
   * The service on the main process, as it ran before it had a thread: for
   * a machine where the thread cannot start or never gets ready, so the
   * suite still opens documents, only without the thread's freedom.
   */
  function fallBack(why) {
    fallingBack ||= fallBackNow(why);
    return fallingBack;
  }

  async function fallBackNow(why) {
    console.error(`the document service could not start on a thread of its own (${why}); it runs on the main process`);
    try {
      await thread?.worker?.terminate();
    } catch {
      /* gone already */
    }
    const [{ createDocumentService }, { createProofing }] = await Promise.all([import('./documents.js'), import('./proofing.js')]);
    const mirror = { ...(init?.settings || {}) };
    const proofing = createProofing({
      stores: { settings: { get: (key, fallback) => (key in mirror ? mirror[key] : fallback), set: (key, value) => { mirror[key] = value; handlers.setting?.(key, value); } } },
      locale: () => init?.locale || 'en-GB',
    });
    local = createDocumentService({ holdBlob, releaseBlob, recoveryDir: init?.recoveryDir ?? null, measureMath: (list) => handlers.measureMath(list), proofing });
    readyNames = Object.keys(local).filter((k) => typeof local[k] === 'function');
    readyOf(readyNames);
  }

  function handle(from, m) {
    switch (m.t) {
      case 'ready':
        readyNames = m.names;
        from.everReady = true;
        readyOf(m.names);
        return;
      case 'reply':
      case 'settle': {
        if (m.pending) return void answered.set(m.seq, m);
        const w = waiting.get(m.seq);
        if (w) {
          waiting.delete(m.seq);
          if (m.ok) w.resolve(m.value);
          else w.reject(failed(m.error));
          if (!waiting.size) settled();
        } else {
          answered.set(m.seq, m);
        }
        return;
      }
      case 'hold':
        holdBlob(m.bytes, m.type, m.name, { ...m.options, id: m.id, owner: m.owner });
        return;
      case 'release':
        releaseBlob(m.id);
        return;
      case 'ask': {
        const fn = handlers[m.fn];
        Promise.resolve()
          .then(() => {
            if (typeof fn !== 'function') throw new Error(`The main process cannot ${m.fn} here.`);
            return fn(m.args);
          })
          .then(
            (value) => from.alive && from.port.postMessage({ t: 'answer', seq: m.seq, ok: true, value }),
            (err) => from.alive && from.port.postMessage({ t: 'answer', seq: m.seq, ok: false, error: { message: err?.message || String(err), name: err?.name } }),
          );
        return;
      }
      case 'teach':
        try {
          handlers.teach?.(m.win, m.word, m.add);
        } catch {
          /* a window on its way out */
        }
        return;
      case 'setting':
        handlers.setting?.(m.key, m.value);
        return;
      default:
    }
  }

  const winOf = (win) => (win && typeof win === 'object' ? win.id ?? null : win ?? null);

  // RUTBA_DOC_TRACE=1: every call, its operations, and how long its answer took.
  const trace = process.env.RUTBA_DOC_TRACE
    ? (name, args, win, how, t0, failure) => console.log(`     [doc] ${how} ${name}${Array.isArray(args?.ops) ? `(${args.ops.map((o) => o.op).join(', ')})` : ''} #${winOf(win) ?? '-'} ${Math.round(performance.now() - t0)} ms${failure ? ` FAILED ${failure}` : ''}`)
    : null;

  /** A call answered with a promise. */
  function call(name, args, win) {
    if (local) return Promise.resolve().then(() => local[name](args, win));
    const t = thread;
    if (!t?.alive) return Promise.reject(new Error('The document engine is starting again; try that once more.'));
    if (trace) {
      const t0 = performance.now();
      return callNow(t, name, args, win).then((v) => (trace(name, args, win, 'async', t0), v), (e) => { trace(name, args, win, 'async', t0, e.message); throw e; });
    }
    return callNow(t, name, args, win);
  }

  function callNow(t, name, args, win) {
    return new Promise((resolve, reject) => {
      const n = ++seq;
      waiting.set(n, { resolve, reject });
      try {
        t.port.postMessage({ t: 'call', seq: n, name, args, win: winOf(win), sync: false });
      } catch (err) {
        waiting.delete(n);
        reject(new Error(`That could not be sent to the document service: ${err?.message || err}`));
      }
    });
  }

  /** A call the caller blocks on, as the checks call the service. */
  function callSync(name, args, win) {
    if (local) return local[name](args, win);
    const t = thread;
    if (!t?.alive) throw new Error('The document engine is not running.');
    const n = ++seq;
    Atomics.store(signal, 0, 0);
    t.port.postMessage({ t: 'call', seq: n, name, args, win: winOf(win), sync: true });
    const until = Date.now() + SYNC_LIMIT_MS;
    for (;;) {
      // Everything the thread has sent so far, in order: the pictures an
      // answer names come before it.
      for (let got = receiveMessageOnPort(t.port); got; got = receiveMessageOnPort(t.port)) {
        handle(t, got.message);
        if (answered.has(n)) break;
      }
      if (answered.has(n)) break;
      if (!t.alive || Atomics.load(signal, 1) === 1) throw new Error('The document engine stopped while answering.');
      if (Date.now() > until) throw new Error(`The document service did not answer ${name} in ten minutes.`);
      Atomics.wait(signal, 0, 0, 50);
    }
    const m = answered.get(n);
    answered.delete(n);
    if (m.pending) {
      // Asynchronous in the service too: its answer follows by message.
      return new Promise((resolve, reject) => waiting.set(n, { resolve: (v) => resolve(asBuffers(v)), reject }));
    }
    if (!m.ok) throw failed(m.error);
    return asBuffers(m.value);
  }

  let stopping = false;
  const faces = { doc: null, direct: null };

  return {
    /** Start the thread: its modules load while the main process gets ready. */
    start() {
      thread ||= start();
      return this;
    },

    /** Where the profile is, the locale and the settings: the service is made with them. */
    init(options = {}) {
      init = options;
      if (thread) thread.port.postMessage({ t: 'init', ...init });
      else thread = start();
      // A thread that has not got ready in a minute never will: the main
      // process takes the service on rather than open no documents at all.
      const waitingFor = thread;
      const late = setTimeout(() => { if (!waitingFor.everReady && thread === waitingFor) fallBack(`it was not ready in ${Math.round(readyLimitMs / 1000)} s`); }, readyLimitMs);
      late.unref?.();
      return this;
    },

    /** Whether the service runs on its own thread (false once it has fallen back to the main process). */
    get threaded() { return !local; },

    /** What only the main process can do: `measureMath`, `teach`, `setting`. */
    provide(more) {
      Object.assign(handlers, more);
    },

    /** The service's method names, once the thread has loaded. */
    async ready() {
      await firstReady;
      if (!faces.doc) {
        faces.doc = Object.fromEntries(readyNames.map((name) => [name, (args, win) => call(name, args, win)]));
        faces.direct = Object.fromEntries(readyNames.map((name) => [name, (args, win) => callSync(name, args, win)]));
      }
      return faces;
    },

    get doc() { return faces.doc; },
    get direct() { return faces.direct; },

    /**
     * For the checks: settled once every window request the thread is working
     * on has its answer (or after `limit` ms, whichever is first), as every
     * request was done before a check could look when the documents were
     * worked on in the main process. True when it settled, false when it gave up.
     */
    idle(limit = 3000) {
      if (local || !waiting.size) return Promise.resolve(true);
      return new Promise((resolve) => {
        const timer = setTimeout(() => { quiet = quiet.filter((r) => r !== done); resolve(false); }, limit);
        const done = (v) => { clearTimeout(timer); resolve(v); };
        quiet.push(done);
      });
    },

    /** For the checks: the thread, to be stopped and seen to start again. */
    get threadId() { return thread?.worker?.threadId ?? null; },

    /** For the checks: the thread stopped as a fault stops it; it is started again. */
    async crashForCheck() {
      await thread?.worker?.terminate();
    },

    async stop() {
      stopping = true;
      await thread?.worker?.terminate();
    },
  };
}
