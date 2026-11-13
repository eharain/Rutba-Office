// The table of files opened in memory, as the document thread sees it.
//
// Word, Sheets and Slides run on a thread of their own (doc-worker.js), and
// the table (memory-files.js) and the Locker live on the main process. This
// stand-in is set as the thread's table, so every hook in documents.js asks
// it as it would ask the table itself:
//
// - What is answered here, at once, from a mirror the main process sends
//   before every call that has changed it: which paths are under the table,
//   which are locked (the table's and the Locker's plain copies), each file's
//   name, size and time, and whether its work is marked unsaved. The mirror
//   holds no bytes.
// - What is asked of the main process: the bytes (read), a save (write, Save
//   As beside it), a recovery copy sealed (recover) or given up (forget).
//   Each answer carries the mirror as it is after it, so a saved file's new
//   path is known here before the save is answered.
//
// A path under the table that the mirror does not hold reads as gone, as it
// does in the table itself, and is never read from or written to the disk.

import path from 'node:path';
import { MemoryFileError } from './memory-files.js';

/** Path text compared the way the system compares it. */
const norm = (p) => {
  const r = path.resolve(String(p || ''));
  return process.platform === 'win32' || process.platform === 'darwin' ? r.toLowerCase() : r;
};

const gone = () => new MemoryFileError('ENOENT', 'This file was opened from the Locker, and the Locker has locked since. Unlock it and open the file again.');

/**
 * @param {object} o
 * @param {(args: object) => Promise<{value, snapshot}>} o.ask      a request to the main process
 * @param {(args: object) => {value, snapshot}} o.askSync          the same, waited for (a check's synchronous open)
 */
export function createMemoryMirror({ ask, askSync }) {
  let root = null;
  let rootN = null;
  let byToken = new Map();
  let copies = new Set();
  // The newest mirror applied: one made earlier and arriving later (with an answer) is not.
  let seq = -1;

  const tokenOf = (p) => {
    if (!rootN || !p) return null;
    const n = norm(p);
    if (!n.startsWith(rootN + path.sep)) return null;
    return n.slice(rootN.length + 1).split(/[\\/]/)[0] || null;
  };
  const isUnder = (p) => tokenOf(p) !== null;
  const entryOf = (p) => {
    const t = tokenOf(p);
    const e = t ? byToken.get(t) : null;
    return e && norm(e.path) === norm(p) ? e : null;
  };

  /** The mirror as the main process sends it, when it is newer than the one here. */
  function apply(snapshot) {
    if (!snapshot) return;
    const at = Number(snapshot.seq);
    if (Number.isFinite(at)) {
      if (at <= seq) return;
      seq = at;
    }
    root = snapshot.root || null;
    rootN = root ? norm(root) : null;
    byToken = new Map((snapshot.entries || []).map((e) => [e.token, e]));
    copies = new Set((snapshot.copies || []).map(norm));
  }

  const toMain = async (args) => {
    // Bytes of their own for the trip: a small Buffer is a view into a pool other buffers share, and the
    // whole pool would be copied with it. Cleared once the request is sent.
    let fresh = null;
    if (args.bytes) {
      fresh = new Uint8Array(args.bytes.byteLength);
      fresh.set(args.bytes);
      args = { ...args, bytes: fresh };
    }
    const asked = ask(args);
    fresh?.fill(0);
    const got = await asked;
    apply(got?.snapshot);
    return got?.value ?? null;
  };

  /** An entry as documents.js reads one: its unsaved mark, and its hooks, each a request to the main process. */
  const view = (e) => (e ? {
    path: e.path,
    name: e.name,
    size: e.size,
    modified: e.modified,
    info: { unsaved: Boolean(e.unsaved), saves: Number(e.saves) || 0 },
    recover: e.recover ? (bytes, how = {}) => toMain({ op: 'recover', p: e.path, bytes, how }) : null,
    forget: e.forget ? () => toMain({ op: 'forget', p: e.path }) : null,
    recovered: e.recovered ? () => toMain({ op: 'recovered', p: e.path }) : null,
  } : null);

  return {
    apply,
    get root() { return root; },
    isUnder,
    has: (p) => entryOf(p) !== null,
    isLocked: (p) => isUnder(p) || Boolean(p && copies.has(norm(p))),
    entry: (p) => view(entryOf(p)),
    entries: () => [...byToken.values()].map(view),
    stat(p) {
      const e = entryOf(p);
      if (!e) return null;
      return { path: e.path, name: e.name, dir: false, size: e.size, mtime: e.modified, mtimeMs: e.modified, ctime: e.modified, ext: path.extname(e.name).toLowerCase(), readonly: !e.save, locked: true, ranged: Boolean(e.ranged) };
    },
    besides: (a, b) => {
      const t = tokenOf(a);
      return t !== null && t === tokenOf(b) && byToken.has(t);
    },
    async read(p) {
      if (!entryOf(p)) throw gone();
      const v = await toMain({ op: 'read', p });
      return Buffer.isBuffer(v) ? v : Buffer.from(v);
    },
    readSync(p) {
      if (!entryOf(p)) throw gone();
      const got = askSync({ op: 'read', p });
      apply(got?.snapshot);
      return Buffer.isBuffer(got.value) ? got.value : Buffer.from(got.value);
    },
    async write(p, bytes, how = {}) {
      if (!entryOf(p)) throw gone();
      await toMain({ op: 'write', p, bytes, how });
    },
    async saveAs(p, name, bytes, how = {}) {
      if (!entryOf(p)) throw gone();
      return toMain({ op: 'saveAs', p, name, bytes, how });
    },
  };
}
