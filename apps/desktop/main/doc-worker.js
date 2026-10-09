// The document service on a thread of its own.
//
// Opening, editing, laying out and saving every document ran on the main
// process, the one thread that also answers every window, the menu and
// mail: a large workbook's recalculation or a long document's layout held
// all of them until it was done. They run here now (documents.js, as it
// was), and the main process only passes the windows' requests in and the
// answers out (doc-host.js).
//
// What only the main process can do is asked of it by message: an
// equation laid out by Chromium, a word taught to a window's spelling
// underline, a setting written. A picture handed to a window is held here
// under an id minted here, and its bytes sent ahead of the answer that
// names it, so the main process has it before any window can ask.

import { workerData } from 'node:worker_threads';
import { AsyncLocalStorage } from 'node:async_hooks';
import { createDocumentService } from './documents.js';
import { createProofing } from './proofing.js';
import { asBuffers } from './doc-host.js';

const { port, signal: shared, generation, blobPrefix } = workerData;
const signal = new Int32Array(shared);
const post = (message, transfer = []) => port.postMessage(message, transfer);

/** The window whose request is being answered, for the pictures it is handed. */
const owner = new AsyncLocalStorage();

let blobs = 0;
function holdBlob(bytes, type = 'application/octet-stream', name = '', options = {}) {
  const id = `w${generation}-${++blobs}`;
  // A copy of just these bytes, handed over whole: a view into a package's
  // buffer would otherwise take the whole buffer across with it.
  const view = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const copy = new Uint8Array(view.byteLength);
  copy.set(view);
  // Read before it goes: a buffer handed over is empty here afterwards.
  const size = copy.byteLength;
  post({ t: 'hold', id, bytes: copy, type, name, options: { group: options.group ?? null, generation: options.generation ?? null }, owner: owner.getStore() ?? null }, [copy.buffer]);
  return { id, url: `${blobPrefix}${id}`, size };
}

const releaseBlob = (id) => post({ t: 'release', id });

// Requests to the main process, answered by message.
let asked = 0;
const answers = new Map();
const ask = (fn, args) => new Promise((resolve, reject) => {
  const seq = ++asked;
  answers.set(seq, { resolve, reject });
  post({ t: 'ask', seq, fn, args });
});

/**
 * The service, made when the main process says where the profile is: the
 * modules above load the moment the thread starts, while the main process
 * is still getting ready.
 */
let service = null;
function setUp({ recoveryDir = null, locale = 'en-GB', settings = {} }) {
  // The settings proofing keeps, mirrored here; proofing is the only thing
  // that writes them, and each write goes to the main process's store too.
  const mirror = { ...settings };
  const stores = {
    settings: {
      get: (key, fallback) => (key in mirror ? mirror[key] : fallback),
      set: (key, value) => {
        mirror[key] = value;
        post({ t: 'setting', key, value });
      },
    },
  };
  const proofing = createProofing({
    stores,
    locale: () => locale || 'en-GB',
    teach: (win, word, add) => post({ t: 'teach', win: win?.id ?? null, word, add }),
  });
  service = createDocumentService({
    holdBlob,
    releaseBlob,
    recoveryDir,
    measureMath: (list) => ask('measureMath', list),
    proofing,
  });
  post({ t: 'ready', names: Object.keys(service).filter((k) => typeof service[k] === 'function') });
}

/**
 * An answer made fit to send: a view into a larger buffer (a picture's
 * bytes in its package) is copied out, so only its own bytes go, and the
 * copy is handed over rather than copied again. What holds such a view is
 * copied on the way, never changed: the service's own objects stay as
 * they are.
 */
function portable(value, transfer, seen = new Map()) {
  if (value === null || typeof value !== 'object') return value;
  if (ArrayBuffer.isView(value)) {
    if (value.byteLength === value.buffer.byteLength) return value;
    const copy = new Uint8Array(value.byteLength);
    copy.set(new Uint8Array(value.buffer, value.byteOffset, value.byteLength));
    transfer.push(copy.buffer);
    return copy;
  }
  if (value instanceof ArrayBuffer || value instanceof Date || value instanceof RegExp || value instanceof Map || value instanceof Set) return value;
  if (seen.has(value)) return seen.get(value);
  seen.set(value, value);
  let out = null;
  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) {
      const v = value[i];
      if (v === null || typeof v !== 'object') continue;
      const p = portable(v, transfer, seen);
      if (p !== v) (out ||= value.slice())[i] = p;
    }
    if (out) seen.set(value, out);
    return out || value;
  }
  for (const key of Object.keys(value)) {
    const v = value[key];
    if (v === null || typeof v !== 'object') continue;
    const p = portable(v, transfer, seen);
    if (p !== v) (out ||= { ...value })[key] = p;
  }
  if (out) seen.set(value, out);
  return out || value;
}

const failure = (err) => ({ message: err?.message || String(err), name: err?.name || 'Error', code: err?.code, stack: err?.stack });

/** An answer sent; a synchronous caller woken once it is on its way. */
function answer(t, seq, ok, value, { sync = false, pending = false } = {}) {
  const transfer = [];
  try {
    if (pending) post({ t, seq, ok: true, pending: true });
    else if (ok) post({ t, seq, ok: true, value: portable(value, transfer) }, transfer);
    else post({ t, seq, ok: false, error: failure(value) });
  } catch (err) {
    // An answer that cannot be sent (something in it that does not copy)
    // is a fault of its own, said in a sentence rather than lost.
    post({ t, seq, ok: false, error: failure(new Error(`The answer to that could not be sent: ${err?.message || err}`)) });
  }
  if (sync) {
    Atomics.store(signal, 0, 1);
    Atomics.notify(signal, 0);
  }
}

port.on('message', (m) => {
  if (m.t === 'answer') {
    const waiting = answers.get(m.seq);
    if (!waiting) return;
    answers.delete(m.seq);
    if (m.ok) waiting.resolve(asBuffers(m.value));
    else waiting.reject(Object.assign(new Error(m.error?.message || 'The main process could not answer.'), { name: m.error?.name || 'Error' }));
    return;
  }
  if (m.t === 'init') return void setUp(m);
  if (m.t !== 'call') return;
  const win = m.win != null ? { id: m.win } : null;
  owner.run(m.win ?? null, () => {
    let result;
    try {
      if (typeof service?.[m.name] !== 'function') throw new Error(`Rutba Office: the document service has no ${m.name}.`);
      // A check's arguments were Buffers where it made them; a window's never were.
      result = service[m.name](m.sync ? asBuffers(m.args) : m.args, win);
    } catch (err) {
      answer('reply', m.seq, false, err, { sync: m.sync });
      return;
    }
    if (result && typeof result.then === 'function') {
      // A synchronous caller is told now that the answer will follow, and
      // has a promise of it, as it would have had from the service itself.
      if (m.sync) answer('reply', m.seq, true, undefined, { sync: true, pending: true });
      const t = m.sync ? 'settle' : 'reply';
      result.then((value) => answer(t, m.seq, true, value), (err) => answer(t, m.seq, false, err));
    } else {
      answer('reply', m.seq, true, result, { sync: m.sync });
    }
  });
});

// A check blocked on an answer learns at once that none is coming.
process.on('exit', () => {
  Atomics.store(signal, 1, 1);
  Atomics.notify(signal, 0);
});
