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
//
// Files opened from the Locker in memory (memory-files.js) are held here,
// on the main process, and the thread sees them through a mirror
// (memory-mirror.js): sent before any call when it has changed, so it is
// always there ahead of the call that names one of them, and sent to a
// thread started again before its first call. The thread asks for their
// bytes, saves and recovery copies; a lock asks the thread for a file's
// unsaved work, and waits only so long; a file let go is told to the thread.

import { Worker, MessageChannel, receiveMessageOnPort } from 'node:worker_threads';
import { fileURLToPath } from 'node:url';
import { memoryFiles, onTableSet, onDropped, addUnsavedSource, addBusySource, lockedCopyPaths } from './memory-files.js';

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

/** Each mirror's number, in the order they are made: the thread applies only a newer one than it has. */
let snapshots = 0;

/** The table of files opened in memory as the thread is sent it: names, sizes and marks, never bytes. */
export function memorySnapshot(table = memoryFiles()) {
  snapshots += 1;
  if (!table) return { seq: snapshots, root: null, entries: [], copies: [] };
  return {
    seq: snapshots,
    root: table.root,
    entries: table.entries().map((e) => ({
      token: e.token,
      path: e.path,
      name: e.name,
      size: e.size,
      modified: e.modified,
      unsaved: Boolean(e.info?.unsaved),
      saves: Number(e.info?.saves) || 0,
      recover: typeof e.recover === 'function',
      forget: typeof e.forget === 'function',
      recovered: typeof e.recovered === 'function',
      save: typeof e.save === 'function',
      saveAs: typeof e.saveAs === 'function' && Boolean(e.bytes),
      ranged: !e.bytes,
    })),
    copies: lockedCopyPaths(),
  };
}

const memoryError = (code, message) => Object.assign(new Error(message), { code });
const wipe = (b) => {
  if (b && typeof b.fill === 'function' && b.byteLength) b.fill(0);
};
/** Bytes of their own for the trip to the thread (a small Buffer is a view into a shared pool); the original cleared. */
const ownBytes = (b) => {
  if (!b) return b;
  const out = new Uint8Array(b.byteLength);
  out.set(b);
  wipe(b);
  return out;
};
/** A mirror's text without its number: two mirrors of the same table are the same. */
const mirrorKey = (snapshot) => JSON.stringify({ ...snapshot, seq: undefined });

/**
 * What the thread asks of the table: a file's bytes, a save, a recovery copy. Answers with the
 * mirror as it is after it. Bytes that came over are the table's to keep a copy of, and are cleared here.
 */
export async function memoryAsk({ op, p, bytes = null, name = null, how = {} }) {
  const table = memoryFiles();
  if (!table) throw memoryError('ENOENT', 'This file was opened from the Locker, and the Locker has locked since. Unlock it and open the file again.');
  const given = bytes ? asBuffers(bytes) : null;
  try {
    let value = null;
    if (op === 'read') value = ownBytes(await table.read(p));
    else if (op === 'write') await table.write(p, given, how);
    else if (op === 'saveAs') value = await table.saveAs(p, name, given, how);
    else if (op === 'recover') {
      const e = table.entry(p);
      const copy = Buffer.from(given || []);
      try {
        value = e?.recover ? Boolean(await e.recover(copy, how)) : false;
      } finally {
        wipe(copy);
      }
    } else if (op === 'forget') {
      const e = table.entry(p);
      if (e?.forget) await e.forget();
    } else if (op === 'recovered') {
      const e = table.entry(p);
      value = e?.recovered ? ownBytes(await e.recovered()) : null;
    } else throw new Error(`The table has no ${op}.`);
    return { value, snapshot: memorySnapshot(table) };
  } finally {
    wipe(given);
  }
}

/** The one request the thread waits for: a file's bytes for an open a check is blocked on. */
function memoryAskSync({ op, p }) {
  const table = memoryFiles();
  if (!table || op !== 'read') throw memoryError('ENOENT', 'This file was opened from the Locker, and the Locker has locked since. Unlock it and open the file again.');
  return { value: ownBytes(table.readSync(p)), snapshot: memorySnapshot(table) };
}

/**
 * @param {object} o
 * @param {(bytes, type, name, options) => object} o.holdBlob the shell's, with `id` and `owner` options
 * @param {(id: string) => void} o.releaseBlob
 * @param {(id: string) => string} o.blobUrl where a held blob is fetched
 * @param {number} [o.readyLimitMs] how long the thread may take to get ready
 */
export function createDocumentHost({ holdBlob, releaseBlob, blobUrl, readyLimitMs = READY_LIMIT_MS }) {
  // What the main process does for the thread, filled in once it can. The table of files opened in memory is here from the start.
  const handlers = { measureMath: null, teach: null, setting: null, mem: memoryAsk };
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
    // The channel the thread waits on, for the bytes of a file opened in memory that a check's open needs at once.
    const sync = new MessageChannel();
    const syncSignal = new Int32Array(new SharedArrayBuffer(4));
    const mine = { port: port2, syncPort: sync.port2, syncSignal, worker: null, alive: true, generation, mirrored: null, mirroredSeq: -1 };
    mine.worker = new Worker(WORKER, {
      workerData: { port: port1, signal: signal.buffer, generation, blobPrefix: blobUrl(''), syncPort: sync.port1, syncSignal: syncSignal.buffer },
      transferList: [port1, sync.port1],
    });
    port2.on('message', (m) => handle(mine, m));
    sync.port2.on('message', (m) => answerSync(mine, m));
    // A thread started again is told what the first was, and is sent the table's mirror before any call.
    if (init) port2.postMessage({ t: 'init', ...init });
    mirrorTo(mine);
    mine.worker.on('error', (err) => {
      console.error('the document service stopped:', err?.stack || err?.message || err);
    });
    mine.worker.on('exit', (code) => {
      mine.alive = false;
      port2.close();
      sync.port2.close();
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

  /** What the thread has as its mirror: the newest one sent it, by number. */
  function sent(t, snapshot) {
    if (!snapshot || !(snapshot.seq > t.mirroredSeq)) return;
    t.mirroredSeq = snapshot.seq;
    t.mirrored = mirrorKey(snapshot);
  }

  /** The table's mirror sent to a thread, when it differs from the newest one it was sent. */
  function mirrorTo(t) {
    if (!t?.alive) return;
    const snapshot = memorySnapshot();
    if (t.mirrored === mirrorKey(snapshot)) return;
    try {
      t.port.postMessage({ t: 'mem', snapshot });
      sent(t, snapshot);
    } catch {
      t.mirrored = null;
    }
  }

  /** A request the thread is waiting for, answered at once. */
  function answerSync(from, m) {
    if (!from.alive || !m) return;
    let reply;
    try {
      const out = memoryAskSync(m.args || {});
      reply = { seq: m.seq, ok: true, value: out };
    } catch (err) {
      reply = { seq: m.seq, ok: false, error: { message: err?.message || String(err), name: err?.name, code: err?.code } };
    }
    try {
      from.syncPort.postMessage(reply);
      sent(from, reply.value?.snapshot);
    } finally {
      wipe(reply.value?.value);
      Atomics.store(from.syncSignal, 0, 1);
      Atomics.notify(from.syncSignal, 0);
    }
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
            (value) => {
              try {
                if (from.alive) {
                  from.port.postMessage({ t: 'answer', seq: m.seq, ok: true, value });
                  if (m.fn === 'mem') sent(from, value?.snapshot);
                }
              } finally {
                // A file's bytes are copied over: the main process's copy is cleared once it is on its way.
                if (m.fn === 'mem') wipe(value?.value);
              }
            },
            (err) => from.alive && from.port.postMessage({ t: 'answer', seq: m.seq, ok: false, error: { message: err?.message || String(err), name: err?.name, code: err?.code } }),
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
    mirrorTo(t);
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
    mirrorTo(t);
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
      // The thread may be waiting on this process in turn, for the bytes of a file opened in memory.
      for (let got = receiveMessageOnPort(t.syncPort); got; got = receiveMessageOnPort(t.syncPort)) answerSync(t, got.message);
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

  // The table, followed: its mirror sent on every change, a file let go told to the thread (which closes
  // its session), and a lock's question about a file's unsaved work asked of the thread. In the service
  // on the main process (`local`) documents.js holds the table itself, and none of this is used.
  const unfollow = [];
  unfollow.push(onTableSet((table) => {
    unfollow.push(table.onChange(() => { if (!local) mirrorTo(thread); }));
  }));
  unfollow.push(onDropped((p) => {
    if (local || !thread?.alive) return;
    mirrorTo(thread);
    try {
      thread.port.postMessage({ t: 'memDropped', p });
    } catch {
      /* the thread is on its way out; its documents go with it */
    }
  }));
  unfollow.push(addUnsavedSource((p) => {
    if (local || !thread?.alive || !readyNames?.includes('unsavedBytes')) return null;
    return call('unsavedBytes', { path: p }).catch(() => null);
  }));
  unfollow.push(addBusySource((p) => {
    if (local || !thread?.alive || !readyNames?.includes('memoryBusy')) return false;
    return call('memoryBusy', { path: p }).catch(() => false);
  }));

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
      for (const off of unfollow.splice(0)) {
        try {
          off?.();
        } catch {
          /* gone already */
        }
      }
      await thread?.worker?.terminate();
    },
  };
}
