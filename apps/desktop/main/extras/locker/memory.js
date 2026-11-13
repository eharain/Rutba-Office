// Locked files opened in memory (2.26.0, D-2026-11-13-20): the Locker's half
// of memory-files.js. Opening one of Office's own kinds reads it out of its
// locked file into the table, and a save from the app seals the new bytes
// back as the file's next version. No plain copy is written.
//
// What still opens from a plain copy in the profile, as in 2.25.0: a kind
// Office does not open (it goes to another program), the kinds whose app
// hands a path to a decoder or an encoder (HEIC, TIFF, PSD and metafile
// pictures; MOV and MKV films), Mail, Calendar, Contacts, Base and Recorder
// files, whose apps keep their own stores, and a document over the size
// limit.
//
// Locking with a file open follows 2.25.0: a file an Office window still
// shows stays open, and a save made while the Locker is locked is held here
// and sealed at the next unlock. A file no window shows any more is let go,
// its bytes zeroed.

import path from 'node:path';

/** Kinds opened whole into memory: their apps read and write them through documents.js, pdf.js or the shell. */
const WHOLE = new Set([
  // Word
  'doc', 'docm', 'docx', 'dotx', 'html', 'htm', 'md', 'odt', 'rtf', 'txt',
  // Sheets
  'csv', 'ods', 'tsv', 'xls', 'xlsm', 'xlsx', 'xltx',
  // Slides
  'odp', 'potx', 'ppsx', 'ppt', 'pptm', 'pptx',
  // PDF
  'pdf',
  // Pictures the window draws itself, and Image's own files
  'avif', 'bmp', 'gif', 'ico', 'jpeg', 'jpg', 'png', 'svg', 'webp', 'rimg',
]);

/** Kinds played by range, straight from the locked file: a film is never held whole. */
const RANGED = new Set(['flac', 'm4a', 'mp3', 'mp4', 'ogg', 'ogv', 'wav', 'webm']);

/** Documents and PDFs up to this size open in memory; a bigger one opens from a plain copy (owner's question 6). */
export const MEMORY_LIMIT = 1024 * 1024 * 1024;

const extOf = (rel) => path.extname(String(rel || '')).slice(1).toLowerCase();

/** How a file opens: 'whole', 'ranged', or null for a plain copy. */
export function memoryWay(rel) {
  const ext = extOf(rel);
  if (WHOLE.has(ext)) return 'whole';
  if (RANGED.has(ext)) return 'ranged';
  return null;
}

/**
 * @param {object} o
 * @param {object} o.locker  createLocker's: readBytes, openRange, saveBytes, isOpen
 * @param {object} o.table   createMemoryFiles's
 * @param {number} [o.limit]
 */
export function createMemoryOpener({ locker, table, limit = MEMORY_LIMIT }) {
  /** Saves made while the Locker was locked, sealed at the next unlock: path -> { id, rel, bytes }. */
  const held = new Map();

  const sealOrHold = (p, id, rel) => async (bytes) => {
    if (!locker.isOpen()) {
      held.get(p)?.bytes.fill(0);
      held.set(p, { id, rel, bytes: Buffer.from(bytes) });
      return { held: true, modified: Date.now() };
    }
    const out = await locker.saveBytes(id, rel, bytes);
    return { modified: Number(out?.modified) || Date.now() };
  };

  /**
   * A locked file put in the table. Answers `{ path, way }`, or null when it
   * should open from a plain copy instead (`why` says which reason).
   */
  async function open(id, rel) {
    const way = memoryWay(rel);
    if (!way) return null;
    const name = path.posix.basename(String(rel).replace(/\\/g, '/'));
    if (way === 'ranged') {
      const r = await locker.openRange(id, rel);
      let p = null;
      try {
        p = table.add({
          name,
          readAt: r.readAt,
          size: Number(r.meta.size),
          modified: Number(r.meta.modified) || Date.now(),
          close: r.close,
          info: { id, rel },
          save: async (bytes) => {
            const out = await sealOrHold(p, id, rel)(bytes);
            if (out.held) return out;
            // The reader was over the locked file as it was: one over the new one.
            const n = await locker.openRange(id, rel);
            return { ...out, readAt: n.readAt, close: n.close, size: Number(n.meta.size) };
          },
        });
      } catch (err) {
        await r.close().catch(() => {});
        throw err;
      }
      return { path: p, way };
    }
    let got;
    try {
      got = await locker.readBytes(id, rel, { limit });
    } catch (err) {
      if (err?.code === 'too-large') return { path: null, way: null, why: 'too-large' };
      throw err;
    }
    let p = null;
    try {
      p = table.add({
        name,
        bytes: got.bytes,
        modified: Number(got.meta.modified) || Date.now(),
        info: { id, rel },
        save: (bytes) => sealOrHold(p, id, rel)(bytes),
      });
    } finally {
      got.bytes.fill(0);
    }
    return { path: p, way };
  }

  /** The paths in the table: what the idle lock and closing the Locker's window wait for while a window shows one. */
  const paths = () => table.entries().map((e) => e.path);

  /**
   * After a lock: every file no window shows is let go, its bytes zeroed. A
   * shown one stays, and its saves are held until the next unlock.
   */
  async function afterLock({ shown = () => false } = {}) {
    for (const e of table.entries()) if (!shown(e.path)) await table.drop(e);
  }

  /** Files whose window has closed, let go: the reaper's turn. A held save keeps its file until it is sealed. */
  async function reap({ shown = () => false } = {}) {
    const gone = [];
    for (const e of table.entries()) {
      if (shown(e.path) || held.has(e.path)) continue;
      await table.drop(e);
      gone.push(e.info);
    }
    return gone;
  }

  /** At unlock: the saves held while locked, sealed in the order they were made. Answers what would not go. */
  async function flushHeld() {
    const failed = [];
    for (const [p, h] of [...held]) {
      try {
        await locker.saveBytes(h.id, h.rel, h.bytes);
        h.bytes.fill(0);
        held.delete(p);
      } catch (err) {
        failed.push({ id: h.id, rel: h.rel, code: err?.code || 'error', message: err?.message || String(err) });
      }
    }
    return failed;
  }

  /** On quit: everything let go. A held save that could not be sealed is lost with it; the caller asks first. */
  async function closeAll() {
    for (const h of held.values()) h.bytes.fill(0);
    held.clear();
    await table.clear();
  }

  return {
    open,
    paths,
    afterLock,
    reap,
    flushHeld,
    closeAll,
    heldCount: () => held.size,
    hasOpen: () => table.entries().length > 0,
  };
}
