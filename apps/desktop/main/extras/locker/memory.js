// Locked files opened in memory (2.26.0, D-2026-11-13-20): the Locker's half
// of memory-files.js. Opening one of Office's own kinds reads it out of its
// locked file into the table, and a save from the app seals the new bytes
// back as the file's next version. No plain copy is written.
//
// What still opens from a plain copy in the profile, as in 2.25.0: a kind
// Office does not open (it goes to another program), the kinds whose app
// hands a path to a decoder (HEIC, TIFF, PSD and metafile pictures; MOV and
// MKV films, which play but are not edited), Mail, Calendar, Contacts, Base
// and Recorder files, whose apps keep their own stores, and a document over
// the size limit.
//
// Locking with a file open (owner, 13 Nov, D-2026-11-13-22): while the keys
// are still there, each file's unsaved work is sealed as a recovery copy
// under its folder's key, its window closes, and its bytes are zeroed. The
// next unlock offers "Reopen the files that were open". The same copy, kept
// up to date as the app autosaves, is what a crash leaves; it is offered the
// same way. A copy goes once its file is saved, or its window is closed by
// the person.
//
// The recovery copies are `<profile>/locker/recovery/<random>.rlock`: a
// locked file of the folder's own, whose header says which folder (and
// nothing else), and whose sealed details name the file it belongs to.

import crypto from 'node:crypto';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { readHeader, LockerError } from './format.js';

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

/** Kinds Office opens, but through a decoder that reads a path: these still take a plain copy, and the window says why. */
const BY_PATH = new Set(['heic', 'heif', 'tif', 'tiff', 'psd', 'emf', 'wmf', 'mov', 'mkv']);

/**
 * Older formats Office reads but does not write in their own format (a Word 97-2003 document, a web
 * page saved as .htm, an Excel 97-2003 workbook, a PowerPoint 97-2003 deck). They open in memory like
 * the rest, but a save would need a new file, and a locked file is saved only back into itself
 * (answer 2), so the window is told at open that changes to one cannot be saved while it is locked.
 * Whether Save As may make a new file in the same locked folder is the owner's to say.
 */
const OLD_KINDS = new Set(['doc', 'htm', 'xls', 'ppt']);

/** A note for the window about a file opened in memory: 'old-kind' when its changes cannot be saved while it is locked. */
export const memoryNote = (rel) => (OLD_KINDS.has(extOf(rel)) ? 'old-kind' : null);

/** Documents and PDFs up to this size open in memory; a bigger one opens from a plain copy (owner's answer 6). */
export const MEMORY_LIMIT = 1024 * 1024 * 1024;

const extOf = (rel) => path.extname(String(rel || '')).slice(1).toLowerCase();
const SLOT = /^[0-9a-f]{24}\.rlock$/;

/** How a file opens: 'whole', 'ranged', or null for a plain copy. */
export function memoryWay(rel) {
  const ext = extOf(rel);
  if (WHOLE.has(ext)) return 'whole';
  if (RANGED.has(ext)) return 'ranged';
  return null;
}

/** Why a file opens from a plain copy: 'by-path' (its decoder reads a path), or 'kind' (another program opens it). */
export const plainWhy = (rel) => (BY_PATH.has(extOf(rel)) ? 'by-path' : 'kind');

const lockedNow = () => new LockerError('locked', 'The Locker has locked, so this file cannot be saved. Unlock it and reopen the file.');

/**
 * @param {object} o
 * @param {object} o.locker       createLocker's: readBytes, openRange, saveBytes, sealRecovery, openRecovery, folders, isOpen
 * @param {object} o.table        createMemoryFiles's
 * @param {string} o.recoveryDir  `<profile>/locker/recovery`
 * @param {(p: string) => Buffer|null} [o.unsavedOf]  a file's unsaved work, from the app that holds it
 * @param {number} [o.limit]
 */
export function createMemoryOpener({ locker, table, recoveryDir, unsavedOf = () => null, limit = MEMORY_LIMIT }) {
  const slotPath = (slot) => path.join(recoveryDir, slot);
  const newSlot = () => `${crypto.randomBytes(12).toString('hex')}.rlock`;
  const forgetSlot = (slot) => (slot ? fsp.rm(slotPath(slot), { force: true }).catch(() => {}) : Promise.resolve());
  // The Locker's own files in the table; anything else there (a print job, print.js) is not one of them.
  const ours = () => table.entries().filter((e) => e.info?.slot);
  // Unsaved work an app sealed itself (PDF and Image, through doc.keepCopy): marks the file, so a lock
  // that has no newer bytes for it never seals nothing over that copy.
  const recoverInto = (getEntry) => async (bytes) => {
    const e = getEntry();
    if (!e) return false;
    const done = await sealSlot(e, bytes);
    if (done) e.info.unsaved = true;
    return done;
  };

  /** The work of a file sealed into its recovery copy: what was saved last, or what the window had unsaved. */
  async function sealSlot(e, bytes, { unsaved = true } = {}) {
    if (!locker.isOpen()) return false;
    const { id, rel, slot } = e.info;
    await locker.sealRecovery(id, slotPath(slot), bytes || Buffer.alloc(0), { name: rel, unsaved: Boolean(unsaved), at: Date.now() });
    return true;
  }

  const saver = (getEntry) => async (bytes) => {
    if (!locker.isOpen()) throw lockedNow();
    const e = getEntry();
    const out = await locker.saveBytes(e.info.id, e.info.rel, bytes);
    // What is in the locked file now is what was on screen: the recovery copy is not wanted.
    await forgetSlot(e.info.slot);
    e.info.unsaved = false;
    return { modified: Number(out?.modified) || Date.now() };
  };

  /**
   * A locked file put in the table. Answers `{ path, way }`, or `{ path: null,
   * why }` when it opens from a plain copy instead. `recovered` puts the
   * bytes of a recovery copy in place of the locked file's, marked unsaved.
   */
  async function open(id, rel, { recovered = null, slot = null } = {}) {
    const way = recovered ? 'whole' : memoryWay(rel);
    if (!way) return { path: null, way: null, why: plainWhy(rel) };
    const name = path.posix.basename(String(rel).replace(/\\/g, '/'));
    const info = { id, rel, slot: slot || newSlot(), unsaved: Boolean(recovered) };
    let p = null;
    const entry = () => table.entry(p);
    if (way === 'ranged') {
      const r = await locker.openRange(id, rel);
      try {
        p = table.add({
          name,
          readAt: r.readAt,
          size: Number(r.meta.size),
          modified: Number(r.meta.modified) || Date.now(),
          close: r.close,
          info,
          recover: recoverInto(entry),
          save: async (bytes) => {
            const out = await saver(entry)(bytes);
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
    let bytes = recovered;
    let modified = Date.now();
    if (!bytes) {
      try {
        const got = await locker.readBytes(id, rel, { limit });
        bytes = got.bytes;
        modified = Number(got.meta.modified) || modified;
      } catch (err) {
        if (err?.code === 'too-large') return { path: null, way: null, why: 'too-large' };
        throw err;
      }
    }
    try {
      p = table.add({ name, bytes, modified, info, recover: recoverInto(entry), save: saver(entry) });
    } finally {
      bytes.fill(0);
    }
    return { path: p, way, note: memoryNote(rel) };
  }

  /** The paths in the table: what the idle lock and closing the Locker's window wait for while a window shows one. */
  const paths = () => ours().map((e) => e.path);

  /**
   * Before a lock, while the keys are still there: every open file's unsaved
   * work sealed as its recovery copy, marked to be offered again, by
   * `deadline`. Answers the paths, whose windows the caller closes.
   */
  async function sealOpen({ deadline = Infinity } = {}) {
    const out = [];
    for (const e of ours()) {
      out.push(e.path);
      if (Date.now() > deadline) continue;
      const bytes = unsavedOf(e.path);
      try {
        // Work the app sealed already and nothing newer to hand: that copy is the one to offer, kept as it is.
        if (!bytes && e.info.unsaved) continue;
        await sealSlot(e, bytes, { unsaved: Boolean(bytes) });
      } catch {
        /* sealed as far as it went; the window still closes */
      } finally {
        bytes?.fill(0);
      }
    }
    return out;
  }

  /** After a lock: every file let go, its bytes zeroed. Its recovery copy stays for the next unlock. */
  const closeAll = () => table.clear();

  /**
   * Files whose window the person closed: let go, and their recovery copies with them. Only a file a
   * window has shown counts as closed, as with the plain copies: one whose window is still on its way
   * (Open, or Reopen opening several one after another) is kept, until a window shows it or the Locker locks.
   */
  async function reap({ shown = () => false } = {}) {
    const gone = [];
    for (const e of ours()) {
      if (shown(e.path)) {
        e.info.seen = true;
        continue;
      }
      if (!e.info.seen) continue;
      await table.drop(e);
      await forgetSlot(e.info.slot);
      gone.push({ id: e.info.id, rel: e.info.rel });
    }
    return gone;
  }

  /**
   * At unlock: the recovery copies there are, the ones a lock sealed and the
   * ones a crash left, read for what file each belongs to. A copy whose
   * folder is not in the Locker any more, or will not open, is left alone.
   */
  async function offers() {
    let names = [];
    try {
      names = (await fsp.readdir(recoveryDir)).filter((n) => SLOT.test(n));
    } catch {
      return [];
    }
    const open = new Set(ours().map((e) => e.info.slot));
    const folders = new Set((await locker.folders().catch(() => [])).map((f) => f.id));
    const out = [];
    for (const slot of names) {
      if (open.has(slot)) continue;
      try {
        const id = (await readHeader(slotPath(slot))).folderId.toString('hex');
        if (!folders.has(id)) continue;
        const { bytes, meta } = await locker.openRecovery(id, slotPath(slot));
        bytes.fill(0);
        out.push({ key: slot, id, rel: meta.name, name: path.posix.basename(String(meta.name)), unsaved: meta.unsaved === true, at: Number(meta.at) || 0 });
      } catch {
        /* not one that opens now */
      }
    }
    return out.sort((a, b) => a.at - b.at);
  }

  /** One offered file opened again: with its unsaved work when it had some, else as it is in its locked file. */
  async function reopen(key) {
    if (!SLOT.test(String(key))) throw new LockerError('state', 'That file is no longer offered.');
    const id = (await readHeader(slotPath(key))).folderId.toString('hex');
    const { bytes, meta } = await locker.openRecovery(id, slotPath(key));
    if (meta.unsaved === true) return { ...(await open(id, meta.name, { recovered: bytes, slot: key })), unsaved: true };
    bytes.fill(0);
    await forgetSlot(key);
    return { ...(await open(id, meta.name)), unsaved: false };
  }

  /** An offer turned down: its recovery copy goes. */
  const dismiss = (key) => (SLOT.test(String(key)) ? forgetSlot(key) : Promise.resolve());

  return {
    open,
    paths,
    sealOpen,
    closeAll,
    reap,
    offers,
    reopen,
    dismiss,
    hasOpen: () => ours().length > 0,
  };
}
