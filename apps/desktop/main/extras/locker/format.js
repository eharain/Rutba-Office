// A locked file on disk, and how it is sealed and opened.
//
// One file is one locked file. It starts with a fixed 128-byte header, then a
// metadata frame, then the file's bytes in 64 KB frames, every frame sealed
// with AES-256-GCM under a key that belongs to this file alone:
//
//   header (128 bytes)
//     0   magic "RLOCKV1\n"
//     8   chunk size, u32 BE (65536, anything else is refused)
//     12  mode: 0 the folder's key, 1 a password of its own (a shared file)
//     16  the file's id, 16 random bytes
//     32  the file's salt, 32 random bytes (HKDF)
//     64  the folder's id, 16 bytes (zero for a shared file)
//     80  scrypt salt, 32 bytes (a shared file only)
//     112 scrypt N, r, p, u32 BE each (a shared file only)
//     124 reserved, zero
//   frame 0: the metadata, 4096 bytes padded with zero bytes, so a name's
//            length does not show
//   frames 1..n: the file, 64 KB each, the last one shorter; an empty file
//            is one empty frame
//
// A frame on disk is `epoch(7) || ciphertext || tag(16)`. Its nonce is
// `epoch(7) || index(u32 BE) || last(1)`, and its associated data is the
// context string, the SHA-256 of the header, the index and the last flag, so a
// frame cannot be moved to another place in the file, another file, or past
// the end, and a file cut short does not end on a frame marked last. The
// epoch is new each time a frame is sealed, so a frame that changes is never
// sealed twice under one nonce; a frame that did not change is copied as it
// stands, so a sync drive uploads only what changed. Which version a frame
// came from is settled by the SHA-256 of the whole file kept in the metadata,
// checked after every full read.
//
// The file's key is HKDF-SHA256 of the folder's key (or, for a shared file, of
// the key scrypt makes from its password), the file's salt, and its id. The
// algorithms are node:crypto's own: AES-256-GCM, HKDF-SHA256, scrypt, SHA-256.
// Nothing here makes up a primitive of its own.
//
// Every format has its version written in it (FORMAT.md lists them all). A
// released format never changes: a change is a new number, and the code that
// reads the old number stays, so a file locked today opens in every later
// Rutba Office. A number this copy does not know is refused as 'newer', which
// is never treated as damage and never written over.
//
// This module only ever reads and writes the files it is pointed at. It does
// not decide where they go: place.js is the only code that swaps a file in.

import crypto from 'node:crypto';
import fsp from 'node:fs/promises';

/** The errors a person's action can meet, each with a code the callers decide by. */
export class LockerError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'LockerError';
    this.code = code;
  }
}

/** The format this Rutba Office writes, and the only one so far. */
export const FORMAT_VERSION = 1;
export const MAGIC = Buffer.from('RLOCKV1\n', 'latin1');
const NEWER = 'This was locked by a newer Rutba Office. Update Rutba Office to open it.';

/** The version in the first eight bytes of a binary format ("RLOCKV1\n", "RPIECEV1"), or null when they are not that format's. */
export function versionIn(buf, stem) {
  if (buf.length < 8 || buf.subarray(0, stem.length).toString('latin1') !== stem) return null;
  const digits = buf.subarray(stem.length, 8).toString('latin1').replace(/\n$/, '');
  return /^[1-9][0-9]*$/.test(digits) ? Number(digits) : null;
}

/** A version number read from a key file, a list or a file's details: this one is read, a higher one is 'newer'. */
export function checkVersion(v) {
  if (Number.isInteger(v) && v > FORMAT_VERSION) throw new LockerError('newer', NEWER);
  if (v !== FORMAT_VERSION) throw new LockerError('format', 'This file does not say which version of the Locker made it.');
}

export const newerError = () => new LockerError('newer', NEWER);
export const HEADER_SIZE = 128;
export const CHUNK = 65536;
export const META_PLAIN = 4096;
const EPOCH = 7;
const TAG = 16;
export const META_FRAME = META_PLAIN + EPOCH + TAG;
export const DATA_FRAME = CHUNK + EPOCH + TAG;
const FIRST_DATA = HEADER_SIZE + META_FRAME;
const CONTEXT = Buffer.from('rutba-locker:v1:chunk', 'utf8');
const FILE_INFO = 'rutba-locker:v1:file:';

export const MODE_FOLDER = 0;
export const MODE_SHARED = 1;

/** A new header for a file about to be locked for the first time. */
export function newHeader({ mode = MODE_FOLDER, folderId = Buffer.alloc(16), scrypt = null } = {}) {
  const h = Buffer.alloc(HEADER_SIZE);
  MAGIC.copy(h, 0);
  h.writeUInt32BE(CHUNK, 8);
  h.writeUInt8(mode, 12);
  crypto.randomBytes(16).copy(h, 16);
  crypto.randomBytes(32).copy(h, 32);
  Buffer.from(folderId).copy(h, 64);
  if (mode === MODE_SHARED) {
    if (!scrypt) throw new LockerError('format', 'A shared file needs its scrypt settings.');
    Buffer.from(scrypt.salt).copy(h, 80);
    h.writeUInt32BE(scrypt.N, 112);
    h.writeUInt32BE(scrypt.r, 116);
    h.writeUInt32BE(scrypt.p, 120);
  }
  return h;
}

/** A header read back, checked, and split into its fields. */
export function parseHeader(buf) {
  const version = buf.length < HEADER_SIZE ? null : versionIn(buf, 'RLOCKV');
  if (version === null) throw new LockerError('format', 'This is not a locked file.');
  if (version !== FORMAT_VERSION) throw newerError();
  const chunk = buf.readUInt32BE(8);
  if (chunk !== CHUNK) throw newerError();
  const mode = buf.readUInt8(12);
  if (mode !== MODE_FOLDER && mode !== MODE_SHARED) throw newerError();
  return {
    version,
    raw: Buffer.from(buf.subarray(0, HEADER_SIZE)),
    mode,
    fileId: Buffer.from(buf.subarray(16, 32)),
    salt: Buffer.from(buf.subarray(32, 64)),
    folderId: Buffer.from(buf.subarray(64, 80)),
    scrypt: mode === MODE_SHARED
      ? { salt: Buffer.from(buf.subarray(80, 112)), N: buf.readUInt32BE(112), r: buf.readUInt32BE(116), p: buf.readUInt32BE(120) }
      : null,
  };
}

/** The key that belongs to this file: from the folder's key (or the shared password's), its salt and its id. */
export function fileKey(rootKey, header) {
  if (!Buffer.isBuffer(rootKey) || rootKey.length !== 32) throw new LockerError('format', 'A locked file needs a 32-byte key.');
  const info = Buffer.from(`${FILE_INFO}${header.fileId.toString('hex')}`, 'utf8');
  return Buffer.from(crypto.hkdfSync('sha256', rootKey, header.salt, info, 32));
}

const headerHash = (header) => crypto.createHash('sha256').update(header.raw).digest();

function aad(hash, index, last) {
  const tail = Buffer.alloc(9);
  tail.writeBigUInt64BE(BigInt(index), 0);
  tail.writeUInt8(last ? 1 : 0, 8);
  return Buffer.concat([CONTEXT, hash, tail]);
}

function nonce(epoch, index, last) {
  if (index > 0xFFFFFFFF) throw new LockerError('format', 'This file is too long to lock in one piece.');
  const n = Buffer.alloc(12);
  epoch.copy(n, 0);
  n.writeUInt32BE(index, EPOCH);
  n.writeUInt8(last ? 1 : 0, 11);
  return n;
}

/** One frame sealed: its bytes on disk. */
export function sealFrame(key, hash, index, last, plain) {
  const epoch = crypto.randomBytes(EPOCH);
  const c = crypto.createCipheriv('aes-256-gcm', key, nonce(epoch, index, last));
  c.setAAD(aad(hash, index, last));
  const body = Buffer.concat([c.update(plain), c.final()]);
  return Buffer.concat([epoch, body, c.getAuthTag()]);
}

/** One frame opened, or null when its tag does not verify: wrong key, wrong place, or altered. */
export function openFrame(key, hash, index, last, frame) {
  if (frame.length < EPOCH + TAG) return null;
  try {
    const d = crypto.createDecipheriv('aes-256-gcm', key, nonce(frame.subarray(0, EPOCH), index, last));
    d.setAAD(aad(hash, index, last));
    d.setAuthTag(frame.subarray(frame.length - TAG));
    return Buffer.concat([d.update(frame.subarray(EPOCH, frame.length - TAG)), d.final()]);
  } catch {
    return null;
  }
}

function packMeta(meta) {
  const json = Buffer.from(JSON.stringify(meta), 'utf8');
  if (json.length > META_PLAIN) throw new LockerError('format', 'This file name is too long to lock.');
  const out = Buffer.alloc(META_PLAIN);
  json.copy(out, 0);
  return out;
}

function unpackMeta(buf) {
  let meta;
  try {
    let end = buf.indexOf(0);
    if (end < 0) end = buf.length;
    meta = JSON.parse(buf.subarray(0, end).toString('utf8'));
  } catch {
    throw new LockerError('tampered', 'This locked file\'s details could not be read.');
  }
  checkVersion(meta.v);
  return meta;
}

async function readFull(fh, buf, offset, length, position) {
  let done = 0;
  while (done < length) {
    const { bytesRead } = await fh.read(buf, offset + done, length - done, position + done);
    if (bytesRead === 0) break;
    done += bytesRead;
  }
  return done;
}

/** Whole-frame reader over a source file, with a stable 64 KB buffer each time. */
async function* chunksOf(fh) {
  let pos = 0;
  for (;;) {
    const buf = Buffer.allocUnsafe(CHUNK);
    const got = await readFull(fh, buf, 0, CHUNK, pos);
    if (got === 0) return;
    yield buf.subarray(0, got);
    if (got < CHUNK) return;
    pos += got;
  }
}

/**
 * Seals a source file into a locked file at `stagePath`, which must not yet
 * hold anything that matters. Nothing in this function touches the source or
 * any other file; the caller decides what the staged file becomes.
 *
 * Frames that are the same as those of `previous` (an earlier locked copy of
 * the same file, with the same header) are copied as they stand, so an upload
 * sends only what changed.
 *
 * @param {object} o
 * @param {string} o.source      the plain file
 * @param {string} o.stagePath   where the sealed bytes go (created, flushed, closed)
 * @param {Buffer} o.rootKey     the folder's key
 * @param {Buffer} o.headerBuf   the file's header (newHeader, or the old file's for a save)
 * @param {object} o.meta        what the metadata frame says besides the hash and size (name, modified time)
 * @param {string} [o.previous]  the file's present locked copy, for unchanged frames
 * @returns {{ size: number, sha256: string, reused: number, frames: number, modified: number }}
 */
export async function sealToStage({ source, stagePath, rootKey, headerBuf, meta, previous = null, hooks = null }) {
  const header = parseHeader(headerBuf);
  const key = fileKey(rootKey, header);
  const hash = headerHash(header);
  const src = await fsp.open(source, 'r');
  const out = await fsp.open(stagePath, 'wx');
  let prev = null;
  try {
    if (previous) prev = await openPrevious(previous, header, key, hash);
    await out.write(headerBuf, 0, HEADER_SIZE, 0);
    await out.write(Buffer.alloc(META_FRAME), 0, META_FRAME, HEADER_SIZE);

    const sha = crypto.createHash('sha256');
    let size = 0;
    let index = 0;
    let reused = 0;
    let pos = FIRST_DATA;
    let held = null;
    // A frame is sealed one step late, so the last one is known to be last.
    const emit = async (plain, last) => {
      index += 1;
      let frame = null;
      if (prev && index <= prev.frames) {
        const old = await prev.frame(index);
        const same = old && old.plain.equals(plain) && (index === prev.frames) === last;
        if (same) {
          frame = old.bytes;
          reused += 1;
        }
      }
      if (!frame) frame = sealFrame(key, hash, index, last, plain);
      await out.write(frame, 0, frame.length, pos);
      pos += frame.length;
    };
    for await (const chunk of chunksOf(src)) {
      sha.update(chunk);
      size += chunk.length;
      if (held) await emit(held, false);
      held = Buffer.from(chunk);
    }
    await emit(held || Buffer.alloc(0), true);
    const digest = sha.digest('hex');

    const stat = await src.stat();
    if (stat.size !== size) throw new LockerError('busy', 'This file changed while it was being locked.');
    const m = packMeta({ v: 1, ...meta, size, sha256: digest, frames: index, modified: stat.mtimeMs });
    const frame0 = sealFrame(key, hash, 0, false, m);
    await out.write(frame0, 0, frame0.length, HEADER_SIZE);
    await out.sync();
    return { size, sha256: digest, reused, frames: index, modified: stat.mtimeMs };
  } finally {
    await out.close().catch(() => {});
    await src.close().catch(() => {});
    if (prev) await prev.close().catch(() => {});
    key.fill(0);
  }
}

/** An earlier locked copy, opened for its frames, when it is the same file under the same key. */
async function openPrevious(path, header, key, hash) {
  let fh = null;
  try {
    fh = await fsp.open(path, 'r');
    const headBuf = Buffer.alloc(HEADER_SIZE);
    if ((await readFull(fh, headBuf, 0, HEADER_SIZE, 0)) < HEADER_SIZE || !headBuf.equals(header.raw)) {
      await fh.close();
      return null;
    }
    const metaBuf = Buffer.alloc(META_FRAME);
    if ((await readFull(fh, metaBuf, 0, META_FRAME, HEADER_SIZE)) < META_FRAME) {
      await fh.close();
      return null;
    }
    const metaPlain = openFrame(key, hash, 0, false, metaBuf);
    if (!metaPlain) {
      await fh.close();
      return null;
    }
    const meta = unpackMeta(metaPlain);
    const frames = meta.frames;
    return {
      frames,
      close: () => fh.close(),
      frame: async (index) => {
        const buf = Buffer.alloc(DATA_FRAME);
        const got = await readFull(fh, buf, 0, DATA_FRAME, FIRST_DATA + (index - 1) * DATA_FRAME);
        const bytes = buf.subarray(0, got);
        const plain = openFrame(key, hash, index, index === frames, bytes);
        return plain ? { plain, bytes: Buffer.from(bytes) } : null;
      },
    };
  } catch {
    await fh?.close().catch(() => {});
    return null;
  }
}

/**
 * A locked file opened: its details, and its frames in order. Every frame is
 * authenticated before it is handed out, and the whole file's hash is checked
 * once the last one has been read.
 *
 * @param {string} path
 * @param {Buffer} rootKey
 * @returns {Promise<{ header, meta, frames: () => AsyncGenerator<Buffer>, close: () => Promise<void> }>}
 * @throws {LockerError} 'format', 'bad-key' (wrong key or a damaged start), 'truncated', 'tampered'
 */
export async function openLocked(path, rootKey) {
  const fh = await fsp.open(path, 'r');
  let key = null;
  try {
    const headBuf = Buffer.alloc(HEADER_SIZE);
    const gotHead = await readFull(fh, headBuf, 0, HEADER_SIZE, 0);
    if (gotHead < HEADER_SIZE) throw new LockerError('format', 'This is not a locked file.');
    const header = parseHeader(headBuf);
    key = fileKey(rootKey, header);
    const hash = headerHash(header);
    const metaBuf = Buffer.alloc(META_FRAME);
    const gotMeta = await readFull(fh, metaBuf, 0, META_FRAME, HEADER_SIZE);
    if (gotMeta < META_FRAME) throw new LockerError('truncated', 'This locked file is cut short.');
    const metaPlain = openFrame(key, hash, 0, false, metaBuf);
    if (!metaPlain) throw new LockerError('bad-key', 'This key does not open the file, or the file has been altered.');
    const meta = unpackMeta(metaPlain);
    const expected = Number(meta.frames);
    if (!Number.isInteger(expected) || expected < 1) throw new LockerError('tampered', 'This locked file\'s details could not be read.');
    const k = key;
    key = null; // the frames generator owns it now
    return {
      header,
      meta,
      close: async () => {
        k.fill(0);
        await fh.close().catch(() => {});
      },
      frames: async function* frames() {
        const digest = crypto.createHash('sha256');
        let size = 0;
        let index = 0;
        let pos = FIRST_DATA;
        for (;;) {
          index += 1;
          const last = index === expected;
          const buf = Buffer.alloc(DATA_FRAME);
          const got = await readFull(fh, buf, 0, DATA_FRAME, pos);
          if (got < EPOCH + TAG) throw new LockerError('truncated', 'This locked file is cut short.');
          const frame = buf.subarray(0, got);
          const plain = openFrame(k, hash, index, last, frame);
          if (!plain) {
            // A short tail where a whole frame was expected reads as cut short, not altered.
            if (got < DATA_FRAME && !last) throw new LockerError('truncated', 'This locked file is cut short.');
            throw new LockerError('tampered', 'A part of this locked file has been altered.');
          }
          digest.update(plain);
          size += plain.length;
          pos += got;
          yield plain;
          if (last) break;
        }
        // Anything after the last frame, and a whole-file check that catches frames from different saves.
        const probe = Buffer.alloc(1);
        const extra = (await fh.read(probe, 0, 1, pos)).bytesRead;
        if (extra) throw new LockerError('tampered', 'This locked file has bytes after its end.');
        if (size !== meta.size || digest.digest('hex') !== meta.sha256) throw new LockerError('tampered', 'This locked file does not match its own record.');
      },
    };
  } catch (err) {
    key?.fill(0);
    await fh.close().catch(() => {});
    throw err;
  }
}

/** The header of a locked file: no key needed, used by the folder scan and the rules. */
export async function readHeader(path) {
  const fh = await fsp.open(path, 'r');
  try {
    const buf = Buffer.alloc(HEADER_SIZE);
    const got = await readFull(fh, buf, 0, HEADER_SIZE, 0);
    if (got < HEADER_SIZE) throw new LockerError('format', 'This is not a locked file.');
    return parseHeader(buf);
  } finally {
    await fh.close().catch(() => {});
  }
}

/** A locked file read through, frame by frame: every byte checked. Answers its details. */
export async function verifyLocked(path, rootKey) {
  const o = await openLocked(path, rootKey);
  try {
    let bytes = 0;
    for await (const plain of o.frames()) bytes += plain.length;
    return { size: bytes, sha256: o.meta.sha256, meta: o.meta, header: o.header };
  } finally {
    await o.close();
  }
}

/** The SHA-256 of a file on disk as it is now. */
export async function fileHash(path) {
  const fh = await fsp.open(path, 'r');
  try {
    const sha = crypto.createHash('sha256');
    let size = 0;
    for await (const c of chunksOf(fh)) {
      sha.update(c);
      size += c.length;
    }
    return { size, sha256: sha.digest('hex') };
  } finally {
    await fh.close().catch(() => {});
  }
}
