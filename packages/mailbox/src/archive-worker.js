// One opened archive on a thread of its own (archive-host.js starts it).
//
// The thread opens the file, reads its folders and messages for the window,
// and builds the archive's part of the search index - and searches it - so
// none of that holds the main process, which answers every window. Indexing
// runs in slices of a few milliseconds with the thread's own messages
// answered between them: a search asked for while indexing goes on is
// answered from what has been indexed so far, and says how far that is.
// Indexing can be paused (the window's Stop) and picked up again where it
// stopped.
//
// Opened with `kind: 'accounts'`, a thread does the same for the accounts'
// mail instead: it reads the mail store's folders (read only - the main
// process alone writes them), keeps their part of the index up to date when
// asked, and searches it.

import { parentPort, workerData } from 'node:worker_threads';
import { openArchive, indexSteps, runSliced } from './archive.js';
import { MailStore } from './store.js';
import { SearchIndex } from './search-index.js';

/** How long a slice of indexing runs before the thread looks at its messages. */
const SLICE_MS = workerData?.slice || 20;
/** A folder is made searchable in parts of this many messages as it is indexed. */
const PART = workerData?.part || 2000;
/** Progress goes to the host at most this often. */
const TELL_MS = 250;

let source = null;
let info = null;
const index = new SearchIndex();
const hooks = {};
let steps = null;
let running = false;
let current = null;
const complete = new Set();
const progress = { state: 'waiting', done: 0, total: 0, error: null };
let told = 0;

const tell = (force = false) => {
  const now = Date.now();
  if (!force && now - told < TELL_MS) return;
  told = now;
  parentPort.postMessage({ event: 'progress', ...progress });
};

/** One slice of indexing, then the thread's messages, then the next slice. */
function pump() {
  if (!running || !steps) return;
  const until = Date.now() + SLICE_MS;
  try {
    while (Date.now() < until) {
      const r = steps.next();
      if (r.done) {
        for (const f of source.folders()) complete.add(f.path);
        steps = null;
        running = false;
        progress.state = 'ready';
        progress.done = progress.total;
        tell(true);
        return;
      }
      if (current !== r.value.folder) {
        if (current != null) complete.add(current);
        current = r.value.folder;
      }
      progress.done = r.value.done;
    }
  } catch (err) {
    steps = null;
    running = false;
    progress.state = 'failed';
    progress.error = err?.message || String(err);
    tell(true);
    return;
  }
  tell();
  setImmediate(pump);
}

/** The segments of one folder, in the order they were made. */
const keysOf = (folder) => [...index.segments.values()].filter((s) => s.folder.path === folder).map((s) => s.key);

const indexed = () => ({ state: progress.state, done: progress.done, total: progress.total });

/** Result rows as small as an export needs: where each message is, and when. */
const slimmed = (rows) => rows.map((r) => ({ accountId: r.accountId, folder: r.folder, id: r.id, date: r.date, t: r.t }));

/** The accounts' thread: the store read from its folder, and its index brought up to date one sync at a time. */
let store = null;
let syncing = Promise.resolve();

const ACCOUNT_OPS = {
  /** Every account's part of the index up to date (a folder changed since is indexed again). */
  sync({ accounts = [], keys = null }) {
    const run = async () => {
      // A profile with a password: the store's files open with the keys the main process holds now.
      if (keys) store.useKeys(keys);
      const live = new Set(accounts.map((a) => a.id));
      for (const a of accounts) await runSliced(store.indexSteps(a.id, { name: a.name || a.id }), { slice: SLICE_MS });
      // An account no longer listed takes its part of the index with it.
      for (const key of [...store.index.segments.keys()]) {
        const id = key.split('\u0000')[0];
        if (!live.has(id)) store.dropIndex(id);
      }
      return { accounts: accounts.length };
    };
    syncing = syncing.then(run, run);
    return syncing;
  },
  search({ query, limit = 200, offset = 0, sources = null, slim = false }) {
    const found = store.index.search(query, { limit, offset, sources });
    return { ...found, rows: slim ? slimmed(found.rows) : found.rows };
  },
  drop({ accountId }) {
    store.dropIndex(accountId);
    return { dropped: true };
  },
};

const OPS = {
  open({ path, id, name, kind }) {
    if (kind === 'accounts') {
      store = new MailStore(path);
      return { kind: 'accounts' };
    }
    source = openArchive(path);
    info = { id, name };
    const folders = source.folders();
    progress.total = folders.reduce((n, f) => n + (f.total || 0), 0);
    return { kind: source.kind, label: source.label, bytes: source.bytes, folders };
  },

  folders: () => source.folders(),

  /** Start indexing, or carry on where a pause left it. */
  index() {
    if (progress.state === 'ready') return indexed();
    steps ||= indexSteps(source, index, info, { part: PART, hooks });
    if (!running) {
      running = true;
      progress.state = 'indexing';
      tell(true);
      setImmediate(pump);
    }
    return indexed();
  },

  /** Pause: what has been read becomes searchable, and the rest waits. */
  pause() {
    if (progress.state !== 'indexing') return indexed();
    running = false;
    hooks.flush?.();
    progress.state = 'paused';
    tell(true);
    return indexed();
  },

  status: () => indexed(),

  list({ folder, offset = 0, limit = 100, query = '', unreadOnly = false, slim = false }) {
    if (query) {
      const keys = keysOf(folder);
      if (keys.length) {
        hooks.flush?.();
        const found = index.search(query, { keys: keysOf(folder), limit, offset });
        return { total: found.total, rows: slim ? slimmed(found.rows) : found.rows, indexed: indexed() };
      }
    }
    if (complete.has(folder)) {
      let docs = [];
      for (const key of keysOf(folder)) docs = docs.concat(index.segments.get(key).docs);
      if (unreadOnly) docs = docs.filter((d) => d.unread);
      docs.sort((a, b) => b.t - a.t);
      const rows = docs.slice(offset, offset + limit).map((d) => (slim ? { id: d.id, folder, t: d.t, date: d.t ? new Date(d.t).toISOString() : null } : { ...d, date: d.t ? new Date(d.t).toISOString() : null, folder }));
      return { total: docs.length, rows };
    }
    // A folder the archive does not have (the window still on the folder of
    // the mailbox it was showing a moment ago) is an empty list, not an error.
    if (!source.folders().some((f) => f.path === folder)) return { total: 0, rows: [] };
    // A large .olm, which cannot be listed without reading every message,
    // is listed from what has been indexed of the folder so far; the window
    // asks again as indexing goes on.
    if (!source.listsCheaply) {
      if (slim && source.ids) {
        const ids = source.ids(folder);
        return { total: ids.length, rows: ids.slice(offset, offset + limit).map((id) => ({ id, folder, t: 0, date: null })) };
      }
      hooks.flush?.();
      let docs = [];
      for (const key of keysOf(folder)) docs = docs.concat(index.segments.get(key).docs);
      docs.sort((a, b) => b.t - a.t);
      return { total: docs.length, rows: docs.slice(offset, offset + limit).map((d) => ({ ...d, date: d.t ? new Date(d.t).toISOString() : null, folder })), indexed: indexed() };
    }
    const page = source.list(folder, { offset, limit });
    return { total: page.total, rows: page.rows.map((r) => ({ ...r, folder })) };
  },

  message: ({ folder, id }) => source.message(folder, id),

  /** A message and its bytes as the archive holds them, for an export. */
  load({ folder, id }) {
    const message = source.message(folder, id);
    return message ? { message, raw: source.raw(folder, id) } : null;
  },

  text: ({ folder, id }) => source.text(folder, id),

  /** Search what has been indexed so far. */
  search({ query, limit = 200, offset = 0, slim = false }) {
    hooks.flush?.();
    const found = index.search(query, { limit, offset });
    return { ...found, rows: slim ? slimmed(found.rows) : found.rows, indexed: indexed() };
  },

  close() {
    running = false;
    try {
      source?.close();
    } catch {
      /* already closed */
    }
    return { closed: true };
  },
};

const failed = (seq, err) => parentPort.postMessage({ seq, ok: false, error: { message: err?.message || String(err), name: err?.name, code: err?.code } });

parentPort.on('message', ({ seq, op, args }) => {
  try {
    const fn = op === 'open' ? OPS.open : store ? ACCOUNT_OPS[op] : OPS[op];
    if (!fn) throw new Error(`No such archive operation: ${op}`);
    if (op !== 'open' && !source && !store) throw new Error('That archive is not open.');
    const value = fn(args || {});
    if (value && typeof value.then === 'function') value.then((v) => parentPort.postMessage({ seq, ok: true, value: v }), (err) => failed(seq, err));
    else parentPort.postMessage({ seq, ok: true, value });
  } catch (err) {
    failed(seq, err);
  }
});
