// The safety copy: the plain files of a folder, kept in one sealed archive in
// Office's profile from before the Locker removes any of them, until the
// person says in the Locker that the folder unlocks as they expect.
//
// What it is, plainly. A ZIP archive (stored, not compressed, ZIP64) holding
// a list of the files first and then the files themselves, sealed as one
// ordinary locked file of version 1 (format.js) under a key that this Office
// profile makes once and keeps beside it, in `safety/key.json`, sealed by the
// system's own safe store for the account signed in (Electron's safeStorage:
// DPAPI on Windows, the Keychain on macOS, the secret service on Linux; given
// in by index.js, so nothing here is tied to one system). Where the Office
// profile has a password of its own, the key file is sealed with that too (the
// profile's vault), so another Office profile on the same user account
// cannot open it. That key is not the Locker password. The copy opens without
// it, for this user account (and this Office profile), and the window says
// so: it is a net while a person gets to know the Locker, never a second lock.
//
// The rules of this module, and the tests that hold it to them
// (tests/locker-safety.test.js):
//   - a copy is made whole, read back and checked before the Locker may remove
//     any plain file it holds; a file that changes while it is copied, or a
//     drive without the room, stops it and nothing is locked;
//   - the space it needs is worked out and checked before anything is written;
//   - a restore writes into a new folder of its own and never overwrites;
//   - a copy is wiped only when the caller has checked the locked folder
//     against it (locker.js), or the folder is no longer in the locker;
//   - what a stop midway leaves (a part written, or a copy written but not
//     yet listed) is cleared before the next copy is made.

import crypto from 'node:crypto';
import fsp from 'node:fs/promises';
import path from 'node:path';
import zlib from 'node:zlib';
import { LockerError, newHeader, parseHeader, fileKey, sealFrame, openLocked, verifyLocked, checkVersion, HEADER_SIZE, META_PLAIN, META_FRAME, CHUNK, MODE_FOLDER } from './format.js';
import { writeJsonAtomic, secretKey } from './keys.js';
import { wipeFile } from './folder.js';

export const SAFETY_DIR = 'safety';
const KEY_FILE = 'key.json';
const LIST_FILE = 'copies.json';

/** No safe store to seal the key with (Linux without a keyring, say). The key is then never written at all, plain or otherwise. */
export const noSafeStore = () => new LockerError('protect', 'Safety copies cannot be made on this computer, because it has no keyring to keep their key in, such as GNOME Keyring or KWallet. The folder still locks as usual.');
const notSealed = () => new LockerError('protect', 'This computer\'s safe store would not keep the safety copy\'s key, so no safety copy was made.');
const notOpened = () => new LockerError('protect', 'This computer\'s safe store would not open the safety copy\'s key. It opens only for the user account that made it, on this computer.');
const EXT = '.rsafe';
/** The archive's first entry: the list of the files, in the order they follow. */
export const MANIFEST = '.rutba-safety-copy.json';
/** Room left on the drive after the copy, so a full profile drive does not stop Windows or Office. */
export const RESERVE = 512 * 1024 * 1024;
/**
 * A leftover in the safety folder is swept only once it is this old: a copy
 * another program (the Locker's command line, say) is writing on the same
 * profile at this moment is newer than that, and is left alone.
 */
export const STALE_MS = 15 * 60 * 1000;
/** How a file sealed by an Office profile's vault begins (profiles.js isSealed). */
const VAULT_SEALED = '{"rutbaSealed":';

const FRAME_EXTRA = 7 + 16;
const LOCAL = 0x04034b50;
const CENTRAL = 0x02014b50;
const ZIP64_END = 0x06064b50;
const ZIP64_LOCATOR = 0x07064b50;
const END = 0x06054b50;
const MAX32 = 0xFFFFFFFF;
const UTF8_FLAG = 0x0800;

/** A size for a message from the engine, which the window shows as it is. */
export function sizeWords(n) {
  const units = ['bytes', 'KB', 'MB', 'GB', 'TB'];
  let v = Number(n) || 0;
  let i = 0;
  while (v >= 1024 && i < units.length - 1) {
    v /= 1024;
    i += 1;
  }
  return i === 0 ? `${v} bytes` : `${v < 10 ? v.toFixed(1) : Math.round(v)} ${units[i]}`;
}

// CRC-32 as ZIP keeps it: node's own where it has one.
let table = null;
function crc32(buf, prev = 0) {
  if (typeof zlib.crc32 === 'function') return zlib.crc32(buf, prev);
  if (!table) {
    table = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
      let c = n;
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
  }
  let c = (prev ^ MAX32) >>> 0;
  for (const b of buf) c = table[(c ^ b) & 0xFF] ^ (c >>> 8);
  return (c ^ MAX32) >>> 0;
}

function dosStamp(ms) {
  const d = new Date(Number.isFinite(ms) ? ms : Date.now());
  if (d.getFullYear() < 1980) return { time: 0, date: (1 << 5) | 1 };
  return {
    time: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
    date: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
  };
}

/** A ZIP local header, every entry with its sizes in a ZIP64 field so one rule fits any size. */
function localHeader(name, size, crc, modified) {
  const n = Buffer.from(name, 'utf8');
  const h = Buffer.alloc(30 + n.length + 20);
  const { time, date } = dosStamp(modified);
  h.writeUInt32LE(LOCAL, 0);
  h.writeUInt16LE(45, 4);
  h.writeUInt16LE(UTF8_FLAG, 6);
  h.writeUInt16LE(0, 8);
  h.writeUInt16LE(time, 10);
  h.writeUInt16LE(date, 12);
  h.writeUInt32LE(crc, 14);
  h.writeUInt32LE(MAX32, 18);
  h.writeUInt32LE(MAX32, 22);
  h.writeUInt16LE(n.length, 26);
  h.writeUInt16LE(20, 28);
  n.copy(h, 30);
  const x = 30 + n.length;
  h.writeUInt16LE(0x0001, x);
  h.writeUInt16LE(16, x + 2);
  h.writeBigUInt64LE(BigInt(size), x + 4);
  h.writeBigUInt64LE(BigInt(size), x + 12);
  return h;
}

function centralEntry(name, size, crc, modified, offset) {
  const n = Buffer.from(name, 'utf8');
  const h = Buffer.alloc(46 + n.length + 28);
  const { time, date } = dosStamp(modified);
  h.writeUInt32LE(CENTRAL, 0);
  h.writeUInt16LE(45, 4);
  h.writeUInt16LE(45, 6);
  h.writeUInt16LE(UTF8_FLAG, 8);
  h.writeUInt16LE(0, 10);
  h.writeUInt16LE(time, 12);
  h.writeUInt16LE(date, 14);
  h.writeUInt32LE(crc, 16);
  h.writeUInt32LE(MAX32, 20);
  h.writeUInt32LE(MAX32, 24);
  h.writeUInt16LE(n.length, 28);
  h.writeUInt16LE(28, 30);
  h.writeUInt32LE(MAX32, 42);
  n.copy(h, 46);
  const x = 46 + n.length;
  h.writeUInt16LE(0x0001, x);
  h.writeUInt16LE(24, x + 2);
  h.writeBigUInt64LE(BigInt(size), x + 4);
  h.writeBigUInt64LE(BigInt(size), x + 12);
  h.writeBigUInt64LE(BigInt(offset), x + 20);
  return h;
}

function archiveEnd(count, cdOffset, cdSize) {
  const z = Buffer.alloc(56 + 20 + 22);
  z.writeUInt32LE(ZIP64_END, 0);
  z.writeBigUInt64LE(44n, 4);
  z.writeUInt16LE(45, 12);
  z.writeUInt16LE(45, 14);
  z.writeBigUInt64LE(BigInt(count), 24);
  z.writeBigUInt64LE(BigInt(count), 32);
  z.writeBigUInt64LE(BigInt(cdSize), 40);
  z.writeBigUInt64LE(BigInt(cdOffset), 48);
  z.writeUInt32LE(ZIP64_LOCATOR, 56);
  z.writeBigUInt64LE(BigInt(cdOffset + cdSize), 64);
  z.writeUInt32LE(1, 72);
  z.writeUInt32LE(END, 76);
  z.writeUInt16LE(0xFFFF, 84);
  z.writeUInt16LE(0xFFFF, 86);
  z.writeUInt32LE(MAX32, 88);
  z.writeUInt32LE(MAX32, 92);
  return z;
}

const nameBytes = (rel) => Buffer.byteLength(rel, 'utf8');

/** The archive's size for these files and this list, to the byte. */
function archiveSize(files, manifestLength) {
  let n = 0;
  for (const [rel, size] of [[MANIFEST, manifestLength], ...files.map((f) => [f.rel, f.size])]) {
    n += 30 + nameBytes(rel) + 20 + size + 46 + nameBytes(rel) + 28;
  }
  return n + 56 + 20 + 22;
}

/** What the sealed file comes to: the header and details frame, then the archive in 64 KB frames. */
function sealedSize(plainBytes) {
  return HEADER_SIZE + META_FRAME + Math.max(1, Math.ceil(plainBytes / CHUNK)) * FRAME_EXTRA + plainBytes;
}

/** The room a copy of these files needs, worked out before anything is read: the list's length allowed for generously. */
export function spaceNeeded(files) {
  const listGuess = 200 + files.reduce((n, f) => n + nameBytes(f.rel) * 2 + 140, 0);
  return sealedSize(archiveSize(files, listGuess));
}

/** One file's CRC-32, SHA-256 and size, read as it is now. */
async function measure(file) {
  const fh = await fsp.open(file, 'r');
  try {
    const sha = crypto.createHash('sha256');
    let crc = 0;
    let size = 0;
    const buf = Buffer.allocUnsafe(1024 * 1024);
    for (;;) {
      const { bytesRead } = await fh.read(buf, 0, buf.length, size);
      if (!bytesRead) break;
      const part = buf.subarray(0, bytesRead);
      sha.update(part);
      crc = crc32(part, crc);
      size += bytesRead;
    }
    return { crc, sha256: sha.digest('hex'), size };
  } finally {
    await fh.close().catch(() => {});
  }
}

/** A file that would not hold still: named, so the person can leave that one plain and lock the rest. */
const changedWhileCopied = (rel) => Object.assign(new LockerError('busy', `${rel} changed while the safety copy was being made, so nothing was locked. Try again once it has been saved, or leave it plain and lock the rest.`), { rel });
const heldOpen = (rel) => Object.assign(new LockerError('busy', `${rel} is in use by another program, or could not be read, so the safety copy could not be made and nothing was locked. Close it and try again, or leave it plain and lock the rest.`), { rel });
const HELD = new Set(['EBUSY', 'EPERM', 'EACCES']);

/** The archive as bytes, in order: the list, every file read again and matched with what was measured, then the directory. */
async function* archive(manifestBuf, files, measured) {
  const central = [];
  let at = 0;
  const entry = (name, size, crc, modified) => {
    central.push(centralEntry(name, size, crc, modified, at));
    const h = localHeader(name, size, crc, modified);
    at += h.length + size;
    return h;
  };
  yield entry(MANIFEST, manifestBuf.length, crc32(manifestBuf), Date.now());
  yield manifestBuf;
  for (let i = 0; i < files.length; i += 1) {
    const m = measured[i];
    yield entry(files[i].rel, m.size, m.crc, files[i].modified);
    const fh = await fsp.open(files[i].path, 'r').catch((err) => { throw HELD.has(err?.code) ? heldOpen(files[i].rel) : err; });
    try {
      const sha = crypto.createHash('sha256');
      let size = 0;
      for (;;) {
        const buf = Buffer.allocUnsafe(CHUNK);
        const { bytesRead } = await fh.read(buf, 0, Math.min(CHUNK, m.size - size + 1), size);
        if (!bytesRead) break;
        size += bytesRead;
        if (size > m.size) throw changedWhileCopied(files[i].rel);
        const part = buf.subarray(0, bytesRead);
        sha.update(part);
        yield part;
      }
      if (size !== m.size || sha.digest('hex') !== m.sha256) throw changedWhileCopied(files[i].rel);
    } finally {
      await fh.close().catch(() => {});
    }
  }
  const cd = Buffer.concat(central);
  yield cd;
  yield archiveEnd(central.length, at, cd.length);
}

/**
 * Bytes sealed into a locked file at `stagePath`, frame by frame as they come:
 * the same layout format.js writes and opens, so `openLocked` reads it back.
 */
async function sealStream(chunks, stagePath, rootKey, headerBuf, meta) {
  const header = parseHeader(headerBuf);
  const key = fileKey(rootKey, header);
  const hash = crypto.createHash('sha256').update(header.raw).digest();
  const out = await fsp.open(stagePath, 'wx', 0o600);
  try {
    await out.write(headerBuf, 0, HEADER_SIZE, 0);
    await out.write(Buffer.alloc(META_FRAME), 0, META_FRAME, HEADER_SIZE);
    const sha = crypto.createHash('sha256');
    let size = 0;
    let index = 0;
    let pos = HEADER_SIZE + META_FRAME;
    let held = null;
    const emit = async (plain, last) => {
      index += 1;
      const frame = sealFrame(key, hash, index, last, plain);
      await out.write(frame, 0, frame.length, pos);
      pos += frame.length;
    };
    let fill = Buffer.allocUnsafe(CHUNK);
    let used = 0;
    for await (const chunk of chunks) {
      sha.update(chunk);
      size += chunk.length;
      let from = 0;
      while (from < chunk.length) {
        const n = Math.min(CHUNK - used, chunk.length - from);
        chunk.copy(fill, used, from, from + n);
        used += n;
        from += n;
        if (used === CHUNK) {
          // Sealed one step late, so the last frame is known to be last.
          if (held) await emit(held, false);
          held = fill;
          fill = Buffer.allocUnsafe(CHUNK);
          used = 0;
        }
      }
    }
    if (used) {
      if (held) await emit(held, false);
      held = fill.subarray(0, used);
    }
    await emit(held || Buffer.alloc(0), true);
    const digest = sha.digest('hex');
    const json = Buffer.from(JSON.stringify({ v: 1, ...meta, size, sha256: digest, frames: index, modified: Date.now() }), 'utf8');
    const details = Buffer.alloc(META_PLAIN);
    json.copy(details, 0);
    const frame0 = sealFrame(key, hash, 0, false, details);
    await out.write(frame0, 0, frame0.length, HEADER_SIZE);
    await out.sync();
    return { size, sha256: digest, sealed: pos };
  } finally {
    key.fill(0);
    await out.close().catch(() => {});
  }
}

/** Exactly the bytes asked for, from a stream of frames of any length. */
function byteReader(frames) {
  const it = frames[Symbol.asyncIterator]();
  let buf = Buffer.alloc(0);
  const more = async () => {
    const { value, done } = await it.next();
    if (done) return false;
    buf = buf.length ? Buffer.concat([buf, value]) : value;
    return true;
  };
  return {
    async take(n) {
      while (buf.length < n) if (!(await more())) throw new LockerError('tampered', 'The safety copy is cut short.');
      const out = buf.subarray(0, n);
      buf = buf.subarray(n);
      return out;
    },
    async *pipe(n) {
      let left = n;
      while (left > 0) {
        if (!buf.length && !(await more())) throw new LockerError('tampered', 'The safety copy is cut short.');
        const part = buf.subarray(0, Math.min(left, buf.length));
        buf = buf.subarray(part.length);
        left -= part.length;
        yield part;
      }
    },
    /** The rest read through, so the locked file's own check of the whole runs. */
    async drain() {
      buf = Buffer.alloc(0);
      while (await more()) buf = Buffer.alloc(0);
    },
    stop: () => it.return?.(),
  };
}

/** The next entry's header: its name, size and CRC, or null where the directory starts. */
async function nextEntry(r) {
  const sig = (await r.take(4)).readUInt32LE(0);
  if (sig === CENTRAL) return null;
  if (sig !== LOCAL) throw new LockerError('tampered', 'The safety copy could not be read.');
  const h = await r.take(26);
  const crc = h.readUInt32LE(10);
  const nameLen = h.readUInt16LE(22);
  const extraLen = h.readUInt16LE(24);
  const name = (await r.take(nameLen)).toString('utf8');
  const extra = await r.take(extraLen);
  let size = h.readUInt32LE(18);
  for (let at = 0; at + 4 <= extra.length;) {
    const id = extra.readUInt16LE(at);
    const len = extra.readUInt16LE(at + 2);
    if (id === 0x0001 && len >= 8) size = Number(extra.readBigUInt64LE(at + 4));
    at += 4 + len;
  }
  return { name, size, crc };
}

const BAD_PART = /[<>:"\\|?*\x00-\x1f]|[. ]$/;
const RESERVED_PART = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;

/** A recorded path, if it is one inside the folder: never a way out of it. */
function safeRel(rel) {
  const parts = String(rel || '').split('/');
  if (!rel || parts.some((p) => !p || p === '.' || p === '..' || BAD_PART.test(p) || RESERVED_PART.test(p))) {
    throw new LockerError('tampered', 'The safety copy records a name that cannot be used.');
  }
  return parts;
}

/** A folder or file name made usable on this system: never empty, never a way out, never a reserved name. */
function usableName(name, fallback) {
  let n = String(name || '').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').replace(/[. ]+$/, '').trim();
  if (!n || n === '.' || n === '..' || RESERVED_PART.test(n)) n = fallback;
  return n.slice(0, 120);
}

function isInside(p, root) {
  const rel = path.relative(root, p);
  return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
}

/**
 * The safety copies of one locker profile.
 *
 * @param {object} o
 * @param {string} o.dir     the locker's folder in Office's profile
 * @param {object} [o.hooks] for the tests: `freeSpace(dir)`, `at(step)`, `protector`
 * @param {object} [o.protector] what seals the key for this account: `{ name, available(), seal(bytes), open(bytes) }`, the system's safe store (index.js). Without one, no copies are made: the key is never kept plain
 * @param {object} [o.vault]     the Office profile's vault (profiles.js): where the profile has a password, the key file is sealed with it as well
 */
export function createSafety({ dir, hooks = null, protector = undefined, vault = null }) {
  const home = path.join(dir, SAFETY_DIR);
  const keyFile = path.join(home, KEY_FILE);
  const listFile = path.join(home, LIST_FILE);
  // A test's own stand-in, or none at all (null), wins; then the caller's. There is no default of this module's own.
  const chosen = hooks && 'protector' in hooks ? hooks.protector : protector;
  const guard = chosen || null;
  // Sealed again by the vault itself whenever the profile's password is set, changed or taken away.
  vault?.cover?.(keyFile);

  async function freeSpace() {
    if (hooks?.freeSpace) return hooks.freeSpace(home);
    const st = await fsp.statfs(dir);
    return Number(st.bavail) * Number(st.bsize);
  }

  /**
   * The folder made before anything goes in it. It is in Office's profile,
   * which is already this account's own; on macOS and Linux it is also made
   * readable by this account only. No system tool is run for it.
   */
  async function ready() {
    await fsp.mkdir(home, { recursive: true, mode: 0o700 });
    if (process.platform !== 'win32') await fsp.chmod(home, 0o700);
  }

  /**
   * The key file as it is: its record, null when there is none, 'unreadable'
   * when it is cut short or not a key, or 'held' when it is sealed with an
   * Office profile password this run does not have (never replaced).
   */
  async function readKeyRecord() {
    let text;
    try {
      text = await fsp.readFile(keyFile, 'utf8');
    } catch (err) {
      if (err.code === 'ENOENT') return null;
      throw new LockerError('format', 'The safety copy\'s key could not be read.');
    }
    const wasSealed = text.startsWith(VAULT_SEALED);
    if (vault && text) {
      text = vault.readText(keyFile);
      if (text === null) return 'held';
    }
    let rec;
    try {
      rec = JSON.parse(text);
    } catch {
      return 'unreadable';
    }
    // Sealed by an Office profile's vault, read here without it: another profile's, or this one's before its password.
    if (rec && typeof rec === 'object' && 'rutbaSealed' in rec) return 'held';
    if (!(rec && typeof rec.sealed === 'string' && rec.id && rec.salt)) return 'unreadable';
    // Written plain by something without the vault (the command line) in a profile that has a password: sealed now.
    if (vault?.sealed?.() && !wasSealed) await sealWithVault(rec).catch(() => {});
    return rec;
  }

  /** The key file written: made new with 'wx', so a second writer finds it there and uses the first one's. */
  async function writeKeyRecord(rec) {
    let fh;
    try {
      fh = await fsp.open(keyFile, 'wx', 0o600);
    } catch (err) {
      if (err.code === 'EEXIST') return false;
      throw err;
    }
    try {
      if (!vault) {
        await fh.writeFile(JSON.stringify(rec, null, 2));
        await fh.sync();
      }
    } finally {
      await fh.close();
    }
    // Through the vault, it replaces the empty file just claimed, sealed with the profile's password where it has one.
    if (vault) await sealWithVault(rec);
    return true;
  }

  /** The key file written through the vault, and made to stay on the disk: a copy listed must never outlive its key. */
  async function sealWithVault(rec) {
    vault.writeJson(keyFile, rec);
    const fh = await fsp.open(keyFile, 'r+');
    try {
      await fh.sync();
    } finally {
      await fh.close();
    }
  }

  /**
   * The key made: created only where there is none, never over one, so two
   * copies made at once share it. A key file cut short before any copy used
   * it is made again; one that copies need is never replaced.
   */
  let making = null;
  function makeKey() {
    making ||= (async () => {
      let rec = await readKeyRecord();
      if (rec === 'unreadable' && !(await readList()).length) {
        await fsp.rm(keyFile, { force: true });
        rec = null;
      }
      if (rec) return rec;
      const secret = crypto.randomBytes(32);
      try {
        const sealed = await guard.seal(secret).catch((err) => { throw err instanceof LockerError ? err : notSealed(); });
        rec = { v: 1, id: crypto.randomBytes(16).toString('hex'), salt: crypto.randomBytes(32).toString('hex'), by: guard.name, sealed: Buffer.from(sealed).toString('base64'), made: new Date().toISOString() };
      } finally {
        secret.fill(0);
      }
      return (await writeKeyRecord(rec)) ? rec : readKeyRecord();
    })().finally(() => { making = null; });
    return making;
  }

  const heldKey = () => new LockerError('protect', 'The safety copy\'s key is sealed with this Office profile\'s password. Open the profile with it first.');

  /** The Office profile's own key: made once, kept here sealed by the system's safe store, never the password. */
  async function profileKey({ make = false } = {}) {
    if (!guard) throw noSafeStore();
    let rec = await readKeyRecord();
    if (rec === 'held') throw heldKey();
    if (!rec && !make) throw new LockerError('format', 'The safety copy\'s key is missing from this profile.');
    if (make && (!rec || rec === 'unreadable')) rec = await makeKey();
    if (rec === 'held') throw heldKey();
    if (!rec || rec === 'unreadable') throw new LockerError('format', 'The safety copy\'s key could not be read.');
    checkVersion(rec.v);
    if (rec.by !== guard.name) throw new LockerError('protect', 'The safety copy\'s key was sealed on another kind of system, and cannot be opened here.');
    const secret = Buffer.from(await guard.open(Buffer.from(rec.sealed, 'base64')).catch((err) => { throw err instanceof LockerError ? err : notOpened(); }));
    try {
      if (secret.length !== 32) throw new LockerError('format', 'The safety copy\'s key could not be read.');
      return { id: Buffer.from(rec.id, 'hex'), key: secretKey(secret, rec.salt, 'safety') };
    } finally {
      secret.fill(0);
    }
  }

  async function readList() {
    try {
      const rec = JSON.parse(await fsp.readFile(listFile, 'utf8'));
      checkVersion(rec.v);
      return rec.copies || [];
    } catch (err) {
      if (err.code === 'ENOENT') return [];
      if (err instanceof LockerError) throw err;
      throw new LockerError('format', 'The list of safety copies could not be read.');
    }
  }
  // Read and written as one step, one at a time: two folders' copies can finish together, and the list is one file.
  let saving = Promise.resolve();
  const updateList = (change) => {
    saving = saving.catch(() => {}).then(async () => writeJsonAtomic(listFile, { v: 1, copies: change(await readList()) }));
    return saving;
  };

  const fileOf = (copy) => path.join(home, path.basename(copy.file));
  /** Copies being restored now: never wiped under a restore. */
  const restoring = new Set();

  /**
   * What a stop midway left, cleared once before the first copy of this run:
   * a part of a copy never finished, and a copy written whole but stopped
   * before it was listed (its folder still asks for one, and nothing of the
   * folder was locked, so a new one is made). Once only, and before any copy
   * of this run is begun, so nothing being written now is ever taken; and only
   * what is STALE_MS old, so nothing another program is writing is taken.
   */
  let tidied = null;
  const tidy = () => {
    tidied ||= (async () => {
      const out = { parts: 0, unlisted: 0 };
      const names = await fsp.readdir(home).catch(() => []);
      if (!names.length) return out;
      // A list that cannot be read keeps every copy.
      const list = await readList().catch(() => null);
      const listed = new Set(list ? list.map((c) => path.basename(c.file)) : names);
      const now = Date.now();
      const old = async (p) => {
        const st = await fsp.stat(p).catch(() => null);
        return Boolean(st) && now - st.mtimeMs >= (hooks?.staleMs ?? STALE_MS);
      };
      for (const name of names) {
        const p = path.join(home, name);
        const part = /^\.rl-[0-9a-f]{24}\.part$/.test(name);
        const unlisted = name.endsWith(EXT) && !listed.has(name);
        if (!(part || unlisted) || !(await old(p))) continue;
        await wipeFile(p).catch(() => {});
        out[part ? 'parts' : 'unlisted'] += 1;
      }
      return out;
    })();
    return tidied;
  };

  /** A copy's archive opened: its list, and a reader positioned on the first file. */
  async function openCopy(copy) {
    const { key } = await profileKey();
    const o = await openLocked(fileOf(copy), key).finally(() => key.fill(0));
    const r = byteReader(o.frames());
    try {
      const first = await nextEntry(r);
      if (!first || first.name !== MANIFEST || first.size > 64 * 1024 * 1024) throw new LockerError('tampered', 'The safety copy could not be read.');
      const manifest = JSON.parse(Buffer.concat(await collect(r.pipe(first.size))).toString('utf8'));
      checkVersion(manifest.v);
      if (!Array.isArray(manifest.files)) throw new LockerError('tampered', 'The safety copy could not be read.');
      return { manifest, r, close: async () => { await r.stop()?.catch?.(() => {}); await o.close(); } };
    } catch (err) {
      await o.close();
      throw err instanceof SyntaxError ? new LockerError('tampered', 'The safety copy could not be read.') : err;
    }
  }

  const api = {
    /** Every copy kept, with what can be said of it without a key. */
    list: readList,

    /** Whether copies can be made here: something to seal their key with. */
    available: async () => Boolean(guard && (await guard.available())),

    /**
     * Who besides the Locker password opens the copies: 'profile' where the key
     * is sealed with the Office profile's password too, 'account' where
     * anyone at this Windows account can.
     */
    async scope() {
      const raw = await fsp.readFile(keyFile, 'utf8').catch(() => null);
      if (raw) return raw.startsWith(VAULT_SEALED) ? 'profile' : 'account';
      return vault?.sealed?.() ? 'profile' : 'account';
    },

    /** Whether a copy opens here: its file there, its key opened, its list read. */
    async opens(copy) {
      try {
        await api.manifest(copy);
        return true;
      } catch {
        return false;
      }
    },

    tidy,

    async forFolder(folderId) {
      return (await readList()).find((c) => c.folderId === folderId) || null;
    },

    async byId(id) {
      const c = (await readList()).find((x) => x.id === id);
      if (!c) throw new LockerError('state', 'That safety copy is not here any more.');
      return c;
    },

    /** What a copy of these files needs, and what the drive has. */
    async estimate(files) {
      const needed = spaceNeeded(files);
      const free = await freeSpace();
      return { files: files.length, bytes: files.reduce((n, f) => n + f.size, 0), needed, free, enough: free >= needed + RESERVE, available: await api.available(), scope: await api.scope() };
    },

    /**
     * A copy made of a folder's waiting files: measured, the room checked,
     * written, read back whole and matched, and only then listed. Anything
     * short of that leaves no copy and throws, so the caller locks nothing.
     *
     * `replace`: the folder's copy already listed is taken over by this one,
     * and wiped only once this one is listed.
     *
     * @param {{ folderId: string, root: string, files: Array<{ path, rel, size, modified }>, replace?: boolean }} o
     */
    async make({ folderId, root, files, replace = false }) {
      if (!replace && await api.forFolder(folderId)) throw new LockerError('exists', 'This folder already has a safety copy.');
      if (!(await api.available())) throw noSafeStore();
      await ready();
      await tidy();
      const first = await api.estimate(files);
      if (!first.enough) {
        const err = new LockerError('space', `The safety copy needs about ${sizeWords(first.needed)}, and the drive with Office's profile has ${sizeWords(first.free)} free. Nothing was locked. Make room, or lock without a safety copy.`);
        Object.assign(err, { needed: first.needed, free: first.free });
        throw err;
      }
      const measured = [];
      for (const f of files) {
        const m = await measure(f.path).catch((err) => { throw HELD.has(err?.code) ? heldOpen(f.rel) : err; });
        if (m.size !== f.size) throw changedWhileCopied(f.rel);
        measured.push(m);
      }
      await hooks?.at?.('safety-measured');
      const made = new Date().toISOString();
      const name = path.basename(root);
      const manifestBuf = Buffer.from(JSON.stringify({ v: 1, folder: name, made, files: files.map((f, i) => ({ rel: f.rel, size: measured[i].size, sha256: measured[i].sha256, modified: f.modified })) }), 'utf8');
      const plainBytes = archiveSize(files, manifestBuf.length);
      const free = await freeSpace();
      if (free < sealedSize(plainBytes) + RESERVE) {
        const err = new LockerError('space', `The safety copy needs about ${sizeWords(sealedSize(plainBytes))}, and the drive with Office's profile has ${sizeWords(free)} free. Nothing was locked. Make room, or lock without a safety copy.`);
        Object.assign(err, { needed: sealedSize(plainBytes), free });
        throw err;
      }
      const id = crypto.randomBytes(16).toString('hex');
      const file = `${id}${EXT}`;
      const stage = path.join(home, `.rl-${crypto.randomBytes(12).toString('hex')}.part`);
      const { id: keyId, key } = await profileKey({ make: true });
      const final = path.join(home, file);
      let renamed = false;
      let listed = false;
      try {
        const sealed = await sealStream(archive(manifestBuf, files, measured), stage, key, newHeader({ mode: MODE_FOLDER, folderId: keyId }), { name: `Safety copy of ${name}`, kind: 'safety', made });
        await hooks?.at?.('safety-written');
        const back = await verifyLocked(stage, key);
        if (back.sha256 !== sealed.sha256 || back.size !== sealed.size) throw new LockerError('failed', 'The safety copy did not read back as it was written, so nothing was locked.');
        await fsp.rename(stage, final);
        renamed = true;
        await hooks?.at?.('safety-renamed');
        const copy = { id, folderId, root, name, file, made, files: files.length, bytes: files.reduce((n, f) => n + f.size, 0), size: sealed.sealed };
        let before = [];
        await updateList((copies) => {
          before = copies.filter((c) => c.folderId === folderId);
          return [...copies.filter((c) => c.folderId !== folderId), copy];
        });
        listed = true;
        // Still there once listed: nothing else took it meanwhile.
        if (!(await fsp.access(final).then(() => true, () => false))) {
          await updateList((copies) => copies.filter((c) => c.id !== id)).catch(() => {});
          throw new LockerError('failed', 'The safety copy was taken away while it was being made, so nothing was locked.');
        }
        for (const old of before) if (old.file !== file) await wipeFile(fileOf(old)).catch(() => {});
        return copy;
      } catch (err) {
        await wipeFile(stage).catch(() => {});
        // Written but never listed: not left for the next run to find.
        if (renamed && !listed) await wipeFile(final).catch(() => {});
        if (err?.code === 'ENOSPC') {
          throw Object.assign(new LockerError('space', 'The drive with Office\'s profile ran out of room while the safety copy was being made. Nothing was locked. Make room, or lock without a safety copy.'), { needed: sealedSize(plainBytes), free: await freeSpace().catch(() => 0) });
        }
        throw err;
      } finally {
        key.fill(0);
      }
    },

    /** The list of files a copy holds, read from the front of its archive. */
    async manifest(copy) {
      const o = await openCopy(copy);
      await o.close();
      return o.manifest;
    },

    /**
     * Every file of a copy written, unlocked, into a new folder inside
     * `parent`. Nothing that is there already is touched: the folder is new,
     * and each file in it is made new. A file that does not match its record
     * is taken away again and named in `failed`.
     *
     * @returns {Promise<{ dest: string, restored: number, failed: Array<{ rel, message }> }>}
     */
    async restore(copy, parent) {
      const base = path.resolve(parent);
      if (!(await fsp.stat(base).then((s) => s.isDirectory(), () => false))) throw new LockerError('state', 'That folder could not be found.');
      if (isInside(base, path.resolve(dir))) throw new LockerError('choice', 'Choose a folder outside the Locker\'s own folder in Office\'s profile.');
      if (restoring.has(copy.id)) throw new LockerError('busy', 'This safety copy is being restored already.');
      restoring.add(copy.id);
      let o;
      try {
        o = await openCopy(copy);
      } catch (err) {
        restoring.delete(copy.id);
        throw err;
      }
      let dest = null;
      const failed = [];
      let restored = 0;
      // The name from the sealed list, which nothing outside the copy can change.
      const folderName = usableName(o.manifest.folder || copy.name, 'Folder');
      try {
        for (let i = 1; !dest; i += 1) {
          const tryDest = path.join(base, `${folderName} (safety copy${i > 1 ? ` ${i}` : ''})`);
          try {
            await fsp.mkdir(tryDest);
            dest = tryDest;
          } catch (err) {
            if (err.code !== 'EEXIST') throw err;
            if (i > 999) throw err;
          }
        }
        for (const want of o.manifest.files) {
          const head = await nextEntry(o.r);
          if (!head || head.name !== want.rel || head.size !== want.size) throw new LockerError('tampered', 'The safety copy does not match its own list.');
          // A name this system cannot use, or one already written (names that differ only in case): that file is passed over, the rest go on.
          const skip = async (message) => {
            for await (const part of o.r.pipe(head.size)) void part;
            failed.push({ rel: want.rel, message });
          };
          let target;
          try {
            target = path.join(dest, ...safeRel(want.rel));
          } catch {
            await skip('This file\'s name cannot be used on this computer, so it was not restored.');
            continue;
          }
          await fsp.mkdir(path.dirname(target), { recursive: true });
          const sha = crypto.createHash('sha256');
          let fh;
          try {
            fh = await fsp.open(target, 'wx');
          } catch (err) {
            if (err.code !== 'EEXIST' && err.code !== 'EISDIR' && err.code !== 'ENOTDIR') throw err;
            await skip('Another file of the same name was restored already, so this one was not.');
            continue;
          }
          let ok = false;
          try {
            for await (const part of o.r.pipe(head.size)) {
              sha.update(part);
              await fh.write(part);
            }
            await fh.sync();
            ok = sha.digest('hex') === want.sha256;
          } finally {
            await fh.close().catch(() => {});
            if (!ok) await fsp.rm(target, { force: true }).catch(() => {});
          }
          if (!ok) {
            failed.push({ rel: want.rel, message: 'This file did not match the safety copy\'s record, so it was not restored.' });
            continue;
          }
          if (Number.isFinite(want.modified)) await fsp.utimes(target, new Date(want.modified), new Date(want.modified)).catch(() => {});
          restored += 1;
        }
        // The directory and the locked file's whole-file check: every file has matched its own record already.
        await o.r.drain().catch(() => {});
        return { dest, restored, failed };
      } catch (err) {
        if (dest) {
          err.dest = dest;
          err.restored = restored;
        }
        throw err;
      } finally {
        restoring.delete(copy.id);
        await o.close();
      }
    },

    /** A copy overwritten and removed, and taken off the list. Never while it is being restored. */
    async remove(copy) {
      if (restoring.has(copy.id)) throw new LockerError('busy', 'This safety copy is being restored. Remove it once that is done.');
      const file = fileOf(copy);
      await wipeFile(file);
      if (await fsp.access(file).then(() => true, () => false)) throw new LockerError('busy', 'The safety copy is in use by another program, so it could not be removed yet.');
      await updateList((copies) => copies.filter((c) => c.id !== copy.id));
    },

    get dir() { return home; },
  };
  return api;
}

async function collect(gen) {
  const out = [];
  for await (const part of gen) out.push(Buffer.from(part));
  return out;
}
