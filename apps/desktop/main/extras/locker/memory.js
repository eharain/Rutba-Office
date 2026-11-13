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
import { busyOf as busyOfApps } from '../../memory-files.js';

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
 * the rest, and a save makes a new file in the current format beside it, in the same locked folder
 * (owner, 13 Nov, D-2026-11-13-26: Budget.xls is saved as Budget.xlsx). The window is told so at open.
 */
const OLD_KINDS = new Set(['doc', 'htm', 'xls', 'ppt']);

/** A note for the window about a file opened in memory: 'old-kind' when a save makes a new file beside it. */
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

/** `name` in `dir`, or the first "Name (n).ext" not in `taken` (lower-case rels). */
export function freeRel(dir, name, taken) {
  const at = (n) => (dir && dir !== '.' ? `${dir}/${n}` : n);
  if (!taken.has(at(name).toLowerCase())) return at(name);
  const ext = path.posix.extname(name);
  const stem = name.slice(0, name.length - ext.length);
  for (let i = 2; ; i += 1) {
    const rel = at(`${stem} (${i})${ext}`);
    if (!taken.has(rel.toLowerCase())) return rel;
  }
}

/** How long a lock waits on the app holding a file for its unsaved work, at most. */
export const UNSAVED_WAIT_MS = 10000;

/**
 * The promise's answer, or null after `ms`, whichever is first; never rejects. Bytes that come after the
 * time is up are zeroed, as no one is left to clear them.
 */
function inTime(promise, ms, onLate = null) {
  if (!promise) return Promise.resolve(null);
  let late = false;
  let timer = null;
  return Promise.race([
    Promise.resolve(promise).then(
      (v) => {
        if (late && v && typeof v.fill === 'function') v.fill(0);
        return v;
      },
      () => null,
    ),
    new Promise((resolve) => {
      timer = setTimeout(() => {
        late = true;
        onLate?.();
        resolve(null);
      }, Math.max(0, ms));
    }),
  ]).finally(() => clearTimeout(timer));
}

/**
 * How long a request to the app holding the files must have had to answer before no answer counts as
 * silence. One asked in the last moments of a file's time (an answer of "busy" that is a little slow) only
 * ends that file's wait. A file's whole time set shorter than this is the floor instead.
 */
const SILENT_FLOOR_MS = 250;

/**
 * Until the app holding a file has no save or copy of it on its way, or `ms` is up. An app that answers
 * "busy" is asked again until then; one that does not answer a request it had `floor` or more for
 * is told (`onSilent`) and not waited for.
 */
async function untilIdleOf(busyOf, p, ms, onSilent = null, floor = SILENT_FLOOR_MS) {
  const until = Date.now() + Math.max(0, ms);
  let quiet = false;
  for (;;) {
    const left = until - Date.now();
    if (left <= 0) return false;
    // This request's window: no answer in it is silence only when it was long enough to answer in. A few ms
    // under the floor still counts: the clock can tick between setting `until` and the first window.
    const said = () => {
      if (!quiet && left >= floor - 5) onSilent?.();
      quiet = true;
    };
    const busy = await inTime(Promise.resolve().then(() => busyOf(p, { within: left, onSilent: said })), left, said);
    if (quiet) return false;
    if (busy === false) return true;
    await new Promise((resolve) => setTimeout(resolve, Math.min(20, Math.max(0, until - Date.now()))));
  }
}

const lockedNow = () => new LockerError('locked', 'The Locker has locked, so this file cannot be saved. Unlock it and reopen the file.');

/**
 * @param {object} o
 * @param {object} o.locker       createLocker's: readBytes, openRange, saveBytes, sealRecovery, openRecovery, folders, isOpen
 * @param {object} o.table        createMemoryFiles's
 * @param {string} o.recoveryDir  `<profile>/locker/recovery`
 * @param {(p: string, how: object) => Promise<{bytes, version, key}|Buffer|null>} [o.unsavedOf]  a file's unsaved work, from the app that holds it
 * @param {(p: string) => Promise<boolean>} [o.busyOf]  whether the app holding a file has a save of it on its way
 * @param {number} [o.limit]
 */
export function createMemoryOpener({ locker, table, recoveryDir, unsavedOf = () => null, busyOf = busyOfApps, limit = MEMORY_LIMIT, unsavedWaitMs = UNSAVED_WAIT_MS }) {
  const slotPath = (slot) => path.join(recoveryDir, slot);
  // A file's time set below the floor: a request that had all of it counts as unanswered when it times out.
  const silentFloor = Math.min(SILENT_FLOOR_MS, Math.max(1, unsavedWaitMs));
  const untilIdle = (p, ms, onSilent) => untilIdleOf(busyOf, p, ms, onSilent, silentFloor);
  // A locked file, as one key: its folder and its path in it, compared as the folder compares them.
  const relKey = (id, rel) => `${id}/${String(rel || '').replace(/\\/g, '/').toLowerCase()}`;
  // Names picked for a new file beside another and not yet in the folder or the table (a save beside on
  // its way, a reopen opening): taken for every other pick until then.
  const reserved = new Set();
  const newSlot = () => `${crypto.randomBytes(12).toString('hex')}.rlock`;
  const forgetSlot = (slot) => (slot ? fsp.rm(slotPath(slot), { force: true }).catch(() => {}) : Promise.resolve());
  // The Locker's own files in the table; anything else there (a print job, print.js) is not one of them.
  const ours = () => table.entries().filter((e) => e.info?.slot);
  // Every path a file has had: a window's address keeps the one it opened, after a Save As beside it.
  const pathsOf = (e) => [e.path, ...(e.aliases || [])];
  // Which work is newer. Word, Sheets and Slides say which document a save or a copy comes from (`key`)
  // and how far its edits had gone (`version`): a copy is kept only while it is newer than what a save
  // has put in the file, and a save takes the copy away only when it is at least as new as the copy.
  // A document opened again (another key) is newer than all an earlier one did. PDF and Image say
  // neither, and are judged by the count of saves when their bytes were taken (`how.saves`).
  const versionOf = (how) => (how && how.key != null && Number.isFinite(Number(how.version)) ? { key: String(how.key), version: Number(how.version) } : null);
  const newerThanSaved = (e, v) => v.key !== e.info.savedKey || v.version > e.info.savedVersion;
  const newerThanCopy = (e, v) => !e.info.slotKey || v.key !== e.info.slotKey || v.version > e.info.slotVersion;
  const copyIs = (e, v) => {
    e.info.slotKey = v ? v.key : null;
    e.info.slotVersion = v ? v.version : -1;
  };

  /** A seal of this file's recovery copy, after every seal of it before: one file's seals never cross. */
  const inChain = (e, fn) => {
    e.info.pending = (e.info.pending || 0) + 1;
    const mine = (e.info.sealing || Promise.resolve()).catch(() => {}).then(fn).finally(() => {
      e.info.pending -= 1;
    });
    e.info.sealing = mine;
    return mine;
  };

  // Unsaved work an app sealed itself (Word, Sheets and Slides by their autosave, PDF and Image through
  // doc.keepCopy): marks the file, so a lock that has no newer bytes for it keeps that copy. Once a lock
  // has reached the file (`final`) no copy from Word, Sheets or Slides is taken: the lock asked for their
  // work itself, and what it seals is the last word. PDF and Image hand the lock nothing, so a copy of
  // theirs that was on its way is still sealed, while the Locker is open (sealSlot).
  const recoverInto = (getEntry) => (bytes, how = {}) => {
    const e = getEntry();
    const v = versionOf(how);
    if (!e || (e.info.final && v)) return Promise.resolve(false);
    const saves = Number.isFinite(how?.saves) ? how.saves : e.info.saves || 0;
    const stale = () => (v ? !newerThanSaved(e, v) : (e.info.saves || 0) !== saves);
    return inChain(e, async () => {
      if (stale()) return false;
      // A copy as new is sealed already: this one is in it.
      if (v && !newerThanCopy(e, v)) return true;
      const done = await sealSlot(e, bytes);
      // A save that landed while it was being sealed: what is in the locked file is as new, or newer.
      if (stale()) {
        await forgetSlot(e.info.slot);
        copyIs(e, null);
        e.info.unsaved = false;
        return false;
      }
      if (done) {
        e.info.unsaved = true;
        copyIs(e, v);
      }
      return done;
    });
  };

  /** The work sealed in a file's recovery copy and not yet saved, for its document opened again, or null. */
  const recoveredOf = (getEntry) => async () => {
    const e = getEntry();
    if (!e || !e.info.unsaved || !locker.isOpen()) return null;
    await e.info.sealing?.catch(() => {});
    try {
      const { bytes, meta } = await locker.openRecovery(e.info.id, slotPath(e.info.slot));
      if (meta?.unsaved !== true || !bytes.length) {
        bytes.fill(0);
        return null;
      }
      return bytes;
    } catch {
      return null;
    }
  };

  /**
   * Sealed work a document could not be opened with (its copy would not open): left in its recovery copy,
   * and the file given a new copy of its own, with nothing marked unsaved. The document then opens as the
   * file was saved, and nothing it does seals over or forgets that work. In the file's chain, after any
   * seal of it already under way, so none lands in the old copy and marks the new one unsaved. Answers
   * 'offered' when the old copy opens (it is offered once the file is closed, or at the next unlock),
   * 'unread' when it does not (it is kept, and offered only if it opens at a later unlock), 'none' when
   * nothing was marked unsaved by then, and false when the file is gone.
   */
  const setAsideOf = (getEntry) => () => {
    const e = getEntry();
    if (!e) return Promise.resolve(false);
    return inChain(e, async () => {
      if (!e.info.unsaved) return 'none';
      const old = e.info.slot;
      e.info.slot = newSlot();
      e.info.unsaved = false;
      copyIs(e, null);
      try {
        const { bytes } = await locker.openRecovery(e.info.id, slotPath(old));
        bytes.fill(0);
        return 'offered';
      } catch {
        return 'unread';
      }
    });
  };

  /** The work of a file sealed into its recovery copy: what was saved last, or what the window had unsaved. */
  async function sealSlot(e, bytes, { unsaved = true } = {}) {
    if (!locker.isOpen()) return false;
    const { id, rel, slot } = e.info;
    await locker.sealRecovery(id, slotPath(slot), bytes || Buffer.alloc(0), { name: rel, unsaved: Boolean(unsaved), at: Date.now() });
    return true;
  }

  /** The window gave its work up (closed without saving, or reverted): its copy goes, and the file is not unsaved. */
  const forgetter = (getEntry) => async () => {
    const e = getEntry();
    if (!e) return;
    // A copy still being sealed is stale from now on.
    e.info.saves = (e.info.saves || 0) + 1;
    await inChain(e, async () => {
      e.info.unsaved = false;
      copyIs(e, null);
      await forgetSlot(e.info.slot);
    });
  };

  /**
   * A save is in the locked file: noted at once (a copy being sealed sees it), and the recovery copy goes
   * unless it holds newer work, decided in the file's chain, after every seal of it before: a copy that
   * was being sealed when the save landed is judged as it is once sealed, never forgotten under it.
   */
  async function saved(e, how) {
    e.info.saves = (e.info.saves || 0) + 1;
    const v = versionOf(how);
    if (v && (v.key !== e.info.savedKey || v.version > e.info.savedVersion)) {
      e.info.savedKey = v.key;
      e.info.savedVersion = v.version;
    }
    await inChain(e, async () => {
      // Typed after this save's bytes were taken, and sealed since (a lock, an autosave): kept.
      if (v && e.info.unsaved && e.info.slotKey === v.key && e.info.slotVersion > v.version) return;
      copyIs(e, null);
      await forgetSlot(e.info.slot);
      e.info.unsaved = false;
    });
  }

  const saver = (getEntry) => async (bytes, how = {}) => {
    if (!locker.isOpen()) throw lockedNow();
    const e = getEntry();
    if (!e) throw new LockerError('state', 'That file is no longer open.');
    const out = await locker.saveBytes(e.info.id, e.info.rel, bytes);
    // What is in the locked file now is what was on screen: the recovery copy is not wanted, unless it is newer.
    await saved(e, how);
    return { modified: Number(out?.modified) || Date.now() };
  };

  /**
   * Save As from a file opened in memory (owner, 13 Nov, D-2026-11-13-26): the bytes sealed as a new
   * file beside it in the same locked folder, under the name asked for or, when that is taken, the
   * first free "Name (2).ext". The new file is what the window has open from then on.
   */
  const besideAs = (getEntry) => async (name, bytes, how = {}) => {
    if (!locker.isOpen()) throw lockedNow();
    const e = getEntry();
    if (!e) throw new LockerError('state', 'That file is no longer open.');
    const dir = path.posix.dirname(String(e.info.rel).replace(/\\/g, '/'));
    // Taken: the folder's files, and the names open now that are only in memory (a reopened copy beside a file).
    const rel = await pickFree(e.info.id, dir, name, e);
    const hold = relKey(e.info.id, rel);
    let out;
    try {
      out = await locker.saveBytes(e.info.id, rel, bytes);
      e.info.rel = rel;
    } finally {
      reserved.delete(hold);
    }
    await saved(e, how);
    return { name: path.posix.basename(rel), modified: Number(out?.modified) || Date.now() };
  };

  /**
   * A locked file put in the table. Answers `{ path, way }`, or `{ path: null,
   * why }` when it opens from a plain copy instead. `recovered` puts the
   * bytes of a recovery copy in place of the locked file's, marked unsaved.
   */
  // Opens and reopens one at a time: each sees what the one before it put in the table.
  let opening = Promise.resolve();
  const inTurn = (fn) => {
    const run = opening.then(fn);
    opening = run.catch(() => {});
    return run;
  };
  const open = (id, rel, how = {}) => inTurn(() => openNow(id, rel, how));

  async function openNow(id, rel, { recovered = null, slot = null } = {}) {
    const way = recovered ? 'whole' : memoryWay(rel);
    if (!way) return { path: null, way: null, why: plainWhy(rel) };
    // Open already (a reopen just put it there, or another window has it): that one, never a second entry
    // for the same locked file, whose saves would land over each other.
    const held0 = recovered ? null : ours().find((e) => relKey(e.info.id, e.info.rel) === relKey(id, rel));
    if (held0) return { path: held0.path, way, note: memoryNote(rel), already: true };
    const name = path.posix.basename(String(rel).replace(/\\/g, '/'));
    const info = { id, rel, slot: slot || newSlot(), unsaved: Boolean(recovered), saves: 0, savedKey: null, savedVersion: -1, slotKey: null, slotVersion: -1, pending: 0, final: false };
    let p = null;
    // The entry itself, while it is in the table: Save As beside it moves it to a new path (memory-files saveAs).
    let held = null;
    const entry = () => (held && table.entries().includes(held) ? held : null);
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
          forget: forgetter(entry),
          setAside: setAsideOf(entry),
          save: async (bytes, how) => {
            const out = await saver(entry)(bytes, how);
            // The reader was over the locked file as it was: one over the new one.
            const n = await locker.openRange(id, rel);
            return { ...out, readAt: n.readAt, close: n.close, size: Number(n.meta.size) };
          },
        });
        held = table.entry(p);
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
      p = table.add({ name, bytes, modified, info, recover: recoverInto(entry), forget: forgetter(entry), recovered: recoveredOf(entry), setAside: setAsideOf(entry), save: saver(entry), saveAs: besideAs(entry) });
      held = table.entry(p);
    } finally {
      bytes.fill(0);
    }
    return { path: p, way, note: memoryNote(rel) };
  }

  /** The paths in the table: what the idle lock and closing the Locker's window wait for while a window shows one. */
  const paths = () => ours().flatMap(pathsOf);

  /**
   * Before a lock, while the keys are still there: every open file's unsaved
   * work sealed as its recovery copy, marked to be offered again, by
   * `deadline`. Answers the paths, whose windows the caller closes.
   */
  async function sealOpen({ deadline = Infinity } = {}) {
    const out = [];
    // Once a request to the app holding the files gets no answer in time (its unsaved work, or whether a
    // save is on its way), no later file waits on it: each keeps the copy sealed before, and the lock
    // reaches every file by its deadline. An app that answers "busy", or a long save of the Locker's own,
    // uses up only that file's time.
    let silent = false;
    const hushed = () => {
      silent = true;
    };
    for (const e of ours()) {
      out.push(...pathsOf(e));
      if (Date.now() > deadline) continue;
      // No copy is taken from here on: what this seals is the last word on the file.
      e.info.final = true;
      // How long this file may be waited on: the app holding it (the document thread) may be busy or hung,
      // and the lock still finishes by its deadline, keeping the copy already sealed.
      const within = () => Math.max(0, Math.min(unsavedWaitMs, deadline - Date.now()));
      const fromApp = () => (silent ? 0 : within());
      // A save on its way and the copies already asked for finish first: what they leave is what is judged.
      // That includes what the app holding it has not handed over yet (a PDF's copy behind a big autosave).
      await inTime(e.saving, within());
      if (!silent) await untilIdle(e.path, fromApp(), hushed);
      await inTime(e.info.sealing, within());
      const saves = e.info.saves || 0;
      const wait = fromApp();
      // No answer is silence only when the app had long enough to give one (the lock's deadline may leave less).
      const raw = silent ? null : await inTime(Promise.resolve().then(() => unsavedOf(e.path, { within: wait })), wait, wait >= silentFloor ? hushed : null);
      const got = raw && (ArrayBuffer.isView(raw) ? { bytes: Buffer.from(raw), version: null, key: null } : raw);
      if (raw && ArrayBuffer.isView(raw)) raw.fill(0);
      // And again: a save or a copy that set off while the work was asked for.
      await inTime(e.saving, within());
      await inTime(e.info.sealing, within());
      let bytes = got?.bytes?.length ? got.bytes : null;
      const v = versionOf(got);
      try {
        // Work the app holds that a save, or a copy, already has: nothing newer to seal.
        if (bytes && v && (!newerThanSaved(e, v) || !newerThanCopy(e, v))) {
          bytes.fill(0);
          bytes = null;
        }
        // Work the app sealed already and nothing newer to hand: that copy is the one to offer, kept as it is.
        if (!bytes && e.info.unsaved) continue;
        const sealing = bytes;
        await inChain(e, async () => {
          // And a copy sealed while this waited its turn: kept as it is too.
          if (!sealing && e.info.unsaved) return;
          await sealSlot(e, sealing, { unsaved: Boolean(sealing) });
          e.info.unsaved = Boolean(sealing);
          copyIs(e, sealing ? v : null);
          // Bytes of no known version, and a save that landed while they were asked for: the file is newer.
          if (sealing && !v && (e.info.saves || 0) !== saves) {
            await forgetSlot(e.info.slot);
            e.info.unsaved = false;
          }
        });
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
      // A file a lock has reached: the lock closed its window, and it lets the file go itself, keeping the
      // copy it sealed for the next unlock. A check that lands while the lock is under way leaves it alone.
      if (e.info.final) continue;
      if (pathsOf(e).some((p) => shown(p))) {
        e.info.seen = true;
        continue;
      }
      if (!e.info.seen) continue;
      // A save or a copy of it on its way (a window that closed while its document was still being
      // serialised for a save): looked at again at the next check, so neither is lost with the file.
      if (e.saving || e.info.pending > 0) continue;
      if (await busyOf(e.path).catch(() => true)) continue;
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
    // A copy of a file that is open now (work set aside when it would not open with it) waits until the
    // file is closed: reopened beside it, the two would save over each other.
    const openRels = new Set(ours().map((e) => relKey(e.info.id, e.info.rel)));
    const folders = new Set((await locker.folders().catch(() => [])).map((f) => f.id));
    const out = [];
    for (const slot of names) {
      if (open.has(slot)) continue;
      try {
        const id = (await readHeader(slotPath(slot))).folderId.toString('hex');
        if (!folders.has(id)) continue;
        const { bytes, meta } = await locker.openRecovery(id, slotPath(slot));
        bytes.fill(0);
        if (openRels.has(relKey(id, meta.name))) continue;
        out.push({ key: slot, id, rel: meta.name, name: path.posix.basename(String(meta.name)), unsaved: meta.unsaved === true, at: Number(meta.at) || 0 });
      } catch {
        /* not one that opens now */
      }
    }
    return out.sort((a, b) => a.at - b.at);
  }

  /**
   * One offered file opened again: with its unsaved work when it had some, else as it is in its locked file.
   * Two copies of one file's work can be offered together (work set aside, then more typed and sealed into
   * the file's new copy). The file is opened once: unsaved work for a file open already opens as a new file
   * beside it in the same locked folder, "Name (2).ext" (`note: 'beside'`), and a copy with no unsaved work
   * has nothing to add, so it goes (`path: null`). Never two entries for one locked file, whose saves would
   * land over each other.
   */
  // One reopen at a time, and in turn with opens: two at once (two Locker windows, a second press, an Open
  // of the same file) would each see the file closed.
  const reopen = (key) => inTurn(() => reopenNow(key));

  async function reopenNow(key) {
    if (!SLOT.test(String(key))) throw new LockerError('state', 'That file is no longer offered.');
    if (ours().some((e) => e.info.slot === key)) throw new LockerError('state', 'That file is open already.');
    let id;
    try {
      id = (await readHeader(slotPath(key))).folderId.toString('hex');
    } catch (err) {
      // Its copy gone already (a reopen with nothing unsaved lets it go, or it was turned down).
      if (err?.code === 'ENOENT') throw new LockerError('state', 'That file is no longer offered.');
      throw err;
    }
    const { bytes, meta } = await locker.openRecovery(id, slotPath(key));
    const already = ours().find((e) => relKey(e.info.id, e.info.rel) === relKey(id, meta.name));
    if (already) {
      if (meta.unsaved !== true) {
        bytes.fill(0);
        await forgetSlot(key);
        return { path: null, way: null, unsaved: false, already: already.path };
      }
      const rel = await besideRel(id, meta.name);
      let opened;
      try {
        opened = await openNow(id, rel, { recovered: bytes, slot: key });
      } finally {
        reserved.delete(relKey(id, rel));
      }
      // Its copy sealed again at once under the new name, while the keys are here: a crash or a lock before
      // the window has it would otherwise offer it back as the file it is beside, and a save there would
      // write this work over that file.
      const e = table.entry(opened.path);
      if (e) {
        const copy = Buffer.from(e.bytes || []);
        try {
          await inChain(e, () => sealSlot(e, copy, { unsaved: true }));
        } catch (err) {
          await table.drop(e).catch(() => {});
          throw err;
        } finally {
          copy.fill(0);
        }
      }
      return { ...opened, unsaved: true, note: 'beside', beside: path.posix.basename(rel), oldKind: opened.note === 'old-kind' };
    }
    if (meta.unsaved === true) return { ...(await openNow(id, meta.name, { recovered: bytes, slot: key })), unsaved: true };
    bytes.fill(0);
    await forgetSlot(key);
    return { ...(await openNow(id, meta.name)), unsaved: false };
  }

  /**
   * The offered files opened again, the newest copy of each file first: when two copies of one file's work
   * are offered, the newer opens as the file itself and the older beside it, so a save of the file never
   * puts the older work over the newer. Answers each key's reopen, or `{ key, failed }`.
   */
  async function reopenAll(keys = []) {
    const at = new Map((await offers()).map((o) => [o.key, o.at]));
    const order = [...keys].sort((a, b) => (at.get(b) ?? -Infinity) - (at.get(a) ?? -Infinity));
    const out = [];
    for (const key of order) {
      try {
        out.push({ key, ...(await reopen(key)) });
      } catch (err) {
        out.push({ key, failed: err?.message || String(err) });
      }
    }
    return out;
  }

  /**
   * A free name for a new file in a locked folder: `name` in `dir`, or the first free "Name (n).ext". Taken
   * are the folder's files, those open now (some only in memory) and the names reserved. Once the folder is
   * scanned, the set is built, the name picked and reserved in one step, so no other pick comes between:
   * the caller lets the reservation go once the file is in the folder or the table.
   */
  async function pickFree(id, dir, name, except = null, also = null) {
    const scanned = await locker.scan(id);
    const lower = (r) => String(r).replace(/\\/g, '/').toLowerCase();
    const taken = new Set(scanned.map((x) => lower(x.rel)));
    for (const e of ours()) if (e !== except && e.info.id === id) taken.add(lower(e.info.rel));
    for (const k of reserved) if (k.startsWith(`${id}/`)) taken.add(k.slice(id.length + 1));
    if (also) taken.add(lower(also));
    const rel = freeRel(dir, name, taken);
    reserved.add(relKey(id, rel));
    return rel;
  }

  /** The first free "Name (n).ext" beside a locked file; held until the reopen has it in the table (reopenNow lets go). */
  function besideRel(id, rel) {
    const norm = String(rel).replace(/\\/g, '/');
    return pickFree(id, path.posix.dirname(norm), path.posix.basename(norm), null, norm);
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
    reopenAll,
    dismiss,
    hasOpen: () => ours().length > 0,
  };
}
