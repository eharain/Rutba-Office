// Files opened in memory: a locked file that Office's own apps open without a
// plain copy ever being written to the disk.
//
// The Locker hands an app a path, `<root>/<token>/<Name.ext>`, and puts the
// file here under it. Nothing is written at that path. The name and the
// extension are the file's own, so every app's handling of kinds, titles and
// window addresses works as it does for any file. Every place that reads or
// writes a document's bytes in the main process asks this table first:
// documents.js, pdf.js, the shell's fs calls and the rutba://file protocol.
//
// Failing closed: `<root>/<token>` is a 0-byte FILE, not a folder. A write
// that misses this table (a code path not hooked) has no folder to land in,
// so it fails on every system and never leaves a plain file behind. A path
// under `<root>` whose token is not here (a document left open after a lock)
// reads as gone, never from the disk.
//
// A file is held one of two ways:
//   bytes:  the whole plain file in a Buffer, for documents and PDFs. Zeroed
//           when it is dropped.
//   ranged: read by range from the locked file (`readAt`), for media, so a
//           film is never held whole.
//
// A save hands the new bytes to the entry's `save`, which seals them into the
// locked file. The table never writes a plain byte anywhere.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { removeTree } from '@rutba/office-shell/electron/remove-tree';

/** Path text compared the way the system compares it. */
const norm = (p) => {
  const r = path.resolve(String(p || ''));
  return process.platform === 'win32' || process.platform === 'darwin' ? r.toLowerCase() : r;
};

export class MemoryFileError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

/** The most of a file read by range that is ever read whole (a sound for its waveform, say). */
export const RANGED_WHOLE = 256 * 1024 * 1024;

const gone = () => new MemoryFileError('ENOENT', 'This file was opened from the Locker, and the Locker has locked since. Unlock it and open the file again.');

/**
 * @param {object} o
 * @param {string} o.root  `<profile>/locker/mem`: only marker files are ever made here
 */
export function createMemoryFiles({ root }) {
  const rootN = norm(root);
  /** token -> entry */
  const byToken = new Map();
  const listeners = new Set();

  const tokenOf = (p) => {
    const n = norm(p);
    if (!n.startsWith(rootN + path.sep)) return null;
    const rest = n.slice(rootN.length + 1);
    const token = rest.split(/[\\/]/)[0];
    return token || null;
  };

  /** Whether a path is under the table's folder at all: such a path is never read from or written to the disk. */
  const isUnder = (p) => tokenOf(p) !== null;

  /** The entry for a path, or null. A path under the folder with no entry is not one either. */
  const entryOf = (p) => {
    const t = tokenOf(p);
    if (!t) return null;
    const e = byToken.get(t);
    if (!e) return null;
    return norm(e.path) === norm(p) ? e : null;
  };

  const emit = (what, e) => {
    for (const fn of listeners) {
      try {
        fn(what, e);
      } catch {
        /* a listener's fault is its own */
      }
    }
  };

  function marker(token) {
    fs.mkdirSync(root, { recursive: true });
    fs.writeFileSync(path.join(root, token), Buffer.alloc(0), { flag: 'wx' });
  }

  /**
   * A file put in the table. Answers its path.
   *
   * @param {object} o
   * @param {string} o.name           the file's own name, with its extension
   * @param {Buffer} [o.bytes]        the whole plain file (documents, PDFs)
   * @param {(start: number, length: number) => Promise<Buffer>} [o.readAt]  ranged reads (media)
   * @param {number} [o.size]         for a ranged file
   * @param {number} [o.modified]
   * @param {(bytes: Buffer) => Promise<object|void>} [o.save]  seals new bytes into the locked file; none means read-only
   * @param {(name: string, bytes: Buffer) => Promise<{name: string}>} [o.saveAs]  seals the bytes as a new file beside it in its locked folder; answers the name it was given there
   * @param {() => Promise<void>} [o.close]  for a ranged file: closes its reader
   * @param {(bytes: Buffer, how: object) => Promise<void>} [o.recover]  seals a recovery copy of unsaved work; none means no copy is kept
   * @param {() => Promise<void>} [o.forget]  the window gave up its unsaved work (closed without saving, or reverted): its recovery copy goes
   * @param {() => Promise<Buffer|null>} [o.recovered]  the work sealed in its recovery copy and not yet saved, or null
   * @param {() => boolean} [o.setAside]  that sealed work left where it is, for the next unlock to offer, and the file given a copy of its own; true when nothing is marked unsaved after it
   * @param {object} [o.info]         what the Locker needs to find it again (folder id, rel)
   */
  function add({ name, bytes = null, readAt = null, size = null, modified = Date.now(), save = null, saveAs = null, close = null, recover = null, forget = null, recovered = null, setAside = null, info = {} }) {
    const base = path.basename(String(name || ''));
    if (!base || base === '.' || base === '..') throw new MemoryFileError('EINVAL', 'A file opened in memory needs its name.');
    if (!bytes && !readAt) throw new MemoryFileError('EINVAL', 'A file opened in memory needs its bytes or a reader.');
    const token = crypto.randomBytes(12).toString('hex');
    marker(token);
    const e = {
      token,
      name: base,
      path: path.join(root, token, base),
      bytes: bytes ? Buffer.from(bytes) : null,
      readAt,
      size: bytes ? bytes.length : Number(size) || 0,
      modified,
      save,
      saveAs,
      close,
      recover,
      forget,
      recovered,
      setAside,
      info,
      saving: null,
      dirty: false,
    };
    // The caller's buffer is its own to clear; the table keeps its own copy.
    byToken.set(token, e);
    emit('add', e);
    return e.path;
  }

  /** The file's details, as fs.stat would give them, or null. */
  function stat(p) {
    const e = entryOf(p);
    if (!e) return null;
    return { path: e.path, name: e.name, dir: false, size: e.size, mtime: e.modified, mtimeMs: e.modified, ctime: e.modified, ext: path.extname(e.name).toLowerCase(), readonly: !e.save, locked: true, ranged: !e.bytes };
  }

  // Read by a window, or by the document service for one: the file has been shown, so the Locker's check
  // counts it as closed once no window shows it, even when its window opened and closed between two checks.
  const seen = (e) => {
    if (e.info && typeof e.info === 'object') e.info.seen = true;
  };

  /** The whole file. A ranged one is read through, up to RANGED_WHOLE: a film is played in parts, never held whole. */
  async function read(p) {
    const e = entryOf(p);
    if (!e) throw gone();
    seen(e);
    if (e.bytes) return Buffer.from(e.bytes);
    if (e.size > RANGED_WHOLE) throw new MemoryFileError('EFBIG', `${e.name} is played from the Locker in parts, and is too large to be read whole.`);
    return Buffer.from(await e.readAt(0, e.size));
  }

  /** The same, at once, for the code paths that read synchronously. Only a file held whole. */
  function readSync(p) {
    const e = entryOf(p);
    if (!e) throw gone();
    seen(e);
    if (!e.bytes) throw new MemoryFileError('EINVAL', 'This file is read in parts.');
    return Buffer.from(e.bytes);
  }

  /** Bytes from `start`, `length` of them at most. */
  async function readRange(p, start, length) {
    const e = entryOf(p);
    if (!e) throw gone();
    seen(e);
    const from = Math.max(0, Math.floor(Number(start) || 0));
    const n = Math.max(0, Math.min(Math.floor(Number(length) || 0), e.size - from));
    if (n <= 0) return Buffer.alloc(0);
    if (e.bytes) return Buffer.from(e.bytes.subarray(from, from + n));
    return Buffer.from(await e.readAt(from, n));
  }

  /**
   * New bytes for the file: sealed into its locked file by its `save`, and
   * only then held as the file's bytes. Saves of one file run one after another.
   */
  async function write(p, bytes, how = {}) {
    const e = entryOf(p);
    if (!e) throw gone();
    if (!e.save) throw new MemoryFileError('EACCES', 'This file was opened to read only.');
    const next = Buffer.from(bytes);
    const run = async () => {
      let out;
      try {
        out = await e.save(next, how || {});
      } catch (err) {
        next.fill(0);
        throw err;
      }
      if (!byToken.has(e.token)) {
        next.fill(0);
        return out;
      }
      if (e.bytes) {
        e.bytes.fill(0);
        e.bytes = next;
        e.size = next.length;
      } else {
        // A file read by range: its reader was over the locked file as it was, so the save hands back a new one.
        next.fill(0);
        if (typeof out?.readAt === 'function') {
          await e.close?.().catch?.(() => {});
          e.readAt = out.readAt;
          e.close = out.close || null;
          e.size = Number(out.size) || e.size;
        }
      }
      e.modified = Number(out?.modified) || Date.now();
      emit('save', e);
      return out;
    };
    const prev = e.saving || Promise.resolve();
    const mine = prev.catch(() => {}).then(run);
    e.saving = mine;
    try {
      return await mine;
    } finally {
      if (e.saving === mine) e.saving = null;
    }
  }

  /**
   * The file sealed as a new file beside it in its locked folder (a Word 97-2003 document saved as a
   * .docx, owner, 13 Nov): `name` is the name asked for, and the Locker may give it a free one. The
   * entry then is that new file, at a path under the same token, and its old path reads as gone.
   * Answers the new path.
   */
  async function saveAs(p, name, bytes, how = {}) {
    const e = entryOf(p);
    if (!e) throw gone();
    if (!e.saveAs || !e.bytes) throw new MemoryFileError('EACCES', 'This file cannot be saved as a new file.');
    const next = Buffer.from(bytes);
    const run = async () => {
      let out;
      try {
        out = await e.saveAs(path.basename(String(name)), next, how || {});
      } catch (err) {
        next.fill(0);
        throw err;
      }
      if (!byToken.has(e.token)) {
        next.fill(0);
        return e.path;
      }
      e.bytes.fill(0);
      e.bytes = next;
      e.size = next.length;
      e.name = path.basename(String(out?.name || name));
      // The window's address still names the file it opened: the Locker knows the window by that.
      e.aliases = [...(e.aliases || []), e.path];
      e.path = path.join(root, e.token, e.name);
      e.modified = Number(out?.modified) || Date.now();
      emit('save', e);
      return e.path;
    };
    const prev = e.saving || Promise.resolve();
    const mine = prev.catch(() => {}).then(run);
    e.saving = mine;
    try {
      return await mine;
    } finally {
      if (e.saving === mine) e.saving = null;
    }
  }

  /** A file let go: its bytes zeroed, its reader closed, its marker removed. */
  async function drop(p) {
    const e = typeof p === 'string' ? entryOf(p) : p;
    if (!e || byToken.get(e.token) !== e) return false;
    byToken.delete(e.token);
    if (e.bytes) e.bytes.fill(0);
    e.bytes = null;
    try {
      await e.close?.();
    } catch {
      /* closed already */
    }
    try {
      fs.rmSync(path.join(root, e.token), { force: true });
    } catch {
      /* gone with the folder */
    }
    emit('drop', e);
    tellDropped(e.path);
    return true;
  }

  /** Every file let go: on lock, on quit. */
  async function clear() {
    for (const e of [...byToken.values()]) await drop(e);
  }

  /** Markers left by a run that did not end cleanly: nothing behind them was ever on the disk. */
  function sweep() {
    let names = [];
    try {
      names = fs.readdirSync(root);
    } catch {
      return 0;
    }
    let n = 0;
    for (const name of names) {
      if (byToken.has(name)) continue;
      try {
        removeTree(path.join(root, name));
        n += 1;
      } catch {
        /* left for the next start */
      }
    }
    return n;
  }

  return {
    root,
    add,
    has: (p) => entryOf(p) !== null,
    isUnder,
    // A file from the Locker, opened in memory or as one of the plain copies it still makes (lockedCopy).
    isLocked: (p) => isUnder(p) || Boolean(p && lockedCopy(String(p))),
    entry: entryOf,
    stat,
    read,
    readSync,
    readRange,
    write,
    saveAs,
    // Whether two paths are under the same token: a new file asked for beside one opened in memory.
    besides: (a, b) => {
      const t = tokenOf(a);
      return t !== null && t === tokenOf(b) && byToken.has(t);
    },
    drop,
    clear,
    sweep,
    entries: () => [...byToken.values()],
    onChange: (fn) => {
      listeners.add(fn);
      return () => listeners.delete(fn);
    },
  };
}

/** Whether a path is one of the Locker's plain copies (a kind opened in another program, or a file over the size limit). */
let lockedCopy = () => false;
let lockedCopyList = () => [];
export const setLockedCopyCheck = (fn, list = null) => {
  lockedCopy = typeof fn === 'function' ? fn : () => false;
  lockedCopyList = typeof list === 'function' ? list : () => [];
};
/** The Locker's plain copies open now: sent to the document thread, which has no Locker of its own. */
export const lockedCopyPaths = () => {
  try {
    return [...(lockedCopyList() || [])].map(String);
  } catch {
    return [];
  }
};

/**
 * Where a file's unsaved work is: each app that holds documents in the main
 * process (documents.js) answers for its own. The Locker asks when it seals
 * what is open before it locks.
 */
const unsavedSources = new Set();
export function addUnsavedSource(fn) {
  unsavedSources.add(fn);
  return () => unsavedSources.delete(fn);
}
/** How long a lock waits on one source for a file's unsaved work: the document thread may be busy, or hung. */
export const UNSAVED_WAIT_MS = 10000;

/** A promise's answer, or null after `ms`; bytes that come later are zeroed, as no one is left to clear them. */
function inTime(promise, ms) {
  if (!(ms < Infinity)) return promise;
  let late = false;
  let timer = null;
  return Promise.race([
    promise.then((v) => {
      if (late && v && typeof v.fill === 'function') v.fill(0);
      return v;
    }),
    new Promise((resolve) => {
      timer = setTimeout(() => {
        late = true;
        resolve(null);
      }, Math.max(0, ms));
    }),
  ]).finally(() => clearTimeout(timer));
}

/**
 * The unsaved work of a file, from whichever source holds it, or null: `{ bytes, version, key }`, where
 * `key` names the document holding it and `version` how far its edits had gone, so the Locker can tell
 * it from what a save has put in the file since. A source may answer with the bytes alone (no version
 * known), and may answer later (the document thread does): each is given `within` ms, and one that has
 * not answered by then counts as none.
 */
export async function unsavedOf(p, { within = UNSAVED_WAIT_MS } = {}) {
  for (const fn of unsavedSources) {
    let got = null;
    try {
      got = await inTime(Promise.resolve().then(() => fn(p)), within);
    } catch {
      got = null;
    }
    const b = got && (ArrayBuffer.isView(got) ? got : got.bytes);
    if (b && b.length) {
      const out = Buffer.from(b);
      if (typeof b.fill === 'function') b.fill(0);
      const known = !ArrayBuffer.isView(got) && got.key != null && Number.isFinite(Number(got.version));
      return { bytes: out, version: known ? Number(got.version) : null, key: known ? String(got.key) : null };
    }
  }
  return null;
}

/**
 * Whether the app holding a file has a save of it on its way (the document thread may be serialising
 * one), so the Locker does not let the file go under it. A source that does not answer in `within` ms
 * counts as busy: the Locker looks again at its next check.
 */
const busySources = new Set();
export function addBusySource(fn) {
  busySources.add(fn);
  return () => busySources.delete(fn);
}
export async function busyOf(p, { within = 1500, onSilent = null } = {}) {
  for (const fn of busySources) {
    let busy = false;
    try {
      // Its answer made true or false first: only no answer in time reads as null, and so as busy. That is
      // told apart (`onSilent`): an app that says it is busy is answering; one that says nothing is not.
      busy = await inTime(Promise.resolve().then(() => fn(p)).then(Boolean), within).then((v) => {
        if (v !== null) return v;
        onSilent?.();
        return true;
      });
    } catch {
      busy = false;
    }
    if (busy) return true;
  }
  return false;
}

/** Told when a file is let go, by any table: a document held on it is closed with it. */
const dropped = new Set();
export function onDropped(fn) {
  dropped.add(fn);
  return () => dropped.delete(fn);
}
/** A file let go, told to every listener: by the table, and on the document thread by the main process's notice. */
export function tellDropped(p) {
  for (const fn of dropped) {
    try {
      fn(p);
    } catch {
      /* a listener's fault is its own */
    }
  }
}

/** The table the app runs with: set once at start, read by every hook. Null until then, so a hook with no table does what it always did. */
let current = null;
const tableSet = new Set();
export const setMemoryFiles = (t) => {
  current = t;
  for (const fn of tableSet) {
    try {
      fn(t);
    } catch {
      /* a listener's fault is its own */
    }
  }
};
/** Told when the table is set (the document host follows it to keep the thread's mirror). Told at once if it is set already. */
export function onTableSet(fn) {
  tableSet.add(fn);
  if (current) fn(current);
  return () => tableSet.delete(fn);
}
export const memoryFiles = () => current;
