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
   * @param {() => Promise<void>} [o.close]  for a ranged file: closes its reader
   * @param {(bytes: Buffer, how: object) => Promise<void>} [o.recover]  seals a recovery copy of unsaved work; none means no copy is kept
   * @param {object} [o.info]         what the Locker needs to find it again (folder id, rel)
   */
  function add({ name, bytes = null, readAt = null, size = null, modified = Date.now(), save = null, close = null, recover = null, info = {} }) {
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
      close,
      recover,
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
    return { path: e.path, name: e.name, dir: false, size: e.size, mtime: e.modified, mtimeMs: e.modified, ctime: e.modified, ext: path.extname(e.name).toLowerCase(), readonly: !e.save };
  }

  /** The whole file. A ranged one is read through. */
  async function read(p) {
    const e = entryOf(p);
    if (!e) throw gone();
    if (e.bytes) return Buffer.from(e.bytes);
    return Buffer.from(await e.readAt(0, e.size));
  }

  /** The same, at once, for the code paths that read synchronously. Only a file held whole. */
  function readSync(p) {
    const e = entryOf(p);
    if (!e) throw gone();
    if (!e.bytes) throw new MemoryFileError('EINVAL', 'This file is read in parts.');
    return Buffer.from(e.bytes);
  }

  /** Bytes from `start`, `length` of them at most. */
  async function readRange(p, start, length) {
    const e = entryOf(p);
    if (!e) throw gone();
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
  async function write(p, bytes) {
    const e = entryOf(p);
    if (!e) throw gone();
    if (!e.save) throw new MemoryFileError('EACCES', 'This file was opened to read only.');
    const next = Buffer.from(bytes);
    const run = async () => {
      const out = await e.save(next);
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
        fs.rmSync(path.join(root, name), { recursive: true, force: true });
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
    entry: entryOf,
    stat,
    read,
    readSync,
    readRange,
    write,
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

/** The table the app runs with: set once at start, read by every hook. Null until then, so a hook with no table does what it always did. */
let current = null;
export const setMemoryFiles = (t) => {
  current = t;
};
export const memoryFiles = () => current;
