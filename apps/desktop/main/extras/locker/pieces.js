// Large files: a small folder (`<name>.rlock.d`) of pieces that are named by
// their content and never change once written, plus a list per version. A
// sync drive that cannot upload part of a file uploads only new pieces.
//
//   item.json                  the file's id and salt (no secret)
//   p/<name>.rpiece            one piece: "RPIECEV1" nonce(12) ciphertext tag(16)
//   list-<seq>-<rand>.rlist    which pieces, in order, for one version, sealed
//                              like any other locked file (format.js)
//
// A piece's name is an HMAC of its bytes under a key from the folder's key, so
// its name says nothing and the same bytes need not be stored twice. Its own
// key comes from HKDF of the folder's key, the file's salt and its name; the
// nonce is random. The list is written last and is the only commit: until it
// is in place, the pieces around it mean nothing, and the version before stands.

import crypto from 'node:crypto';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { newHeader, openLocked, LockerError, CHUNK, FORMAT_VERSION, versionIn, checkVersion, newerError } from './format.js';
import { sealFile, placeStaged } from './place.js';
import { gearTable, cutPoint, MAX } from './fastcdc.js';
import { writeJsonAtomic } from './keys.js';

const MAGIC = Buffer.from('RPIECEV1', 'latin1');
const INFO = 'rutba-locker:v1:piece:';

function nameKey(folderKey, item) {
  return Buffer.from(crypto.hkdfSync('sha256', folderKey, Buffer.from(item.salt, 'hex'), Buffer.from(`${INFO}name:${item.fileId}`, 'utf8'), 32));
}

function pieceKey(folderKey, item, name) {
  return Buffer.from(crypto.hkdfSync('sha256', folderKey, Buffer.from(item.salt, 'hex'), Buffer.from(`${INFO}${item.fileId}:${name}`, 'utf8'), 32));
}

const aadOf = (item, name) => Buffer.from(`${INFO}${item.fileId}:${name}`, 'utf8');

function sealPiece(folderKey, item, name, plain) {
  const iv = crypto.randomBytes(12);
  const key = pieceKey(folderKey, item, name);
  const c = crypto.createCipheriv('aes-256-gcm', key, iv);
  c.setAAD(aadOf(item, name));
  const body = Buffer.concat([c.update(plain), c.final()]);
  const out = Buffer.concat([MAGIC, iv, body, c.getAuthTag()]);
  key.fill(0);
  return out;
}

function openPiece(folderKey, item, name, sealed) {
  const version = sealed.length < MAGIC.length + 28 ? null : versionIn(sealed, 'RPIECEV');
  if (version === null) throw new LockerError('tampered', 'A piece of this file is damaged.');
  if (version !== FORMAT_VERSION) throw newerError();
  const key = pieceKey(folderKey, item, name);
  try {
    const d = crypto.createDecipheriv('aes-256-gcm', key, sealed.subarray(8, 20));
    d.setAAD(aadOf(item, name));
    d.setAuthTag(sealed.subarray(sealed.length - 16));
    return Buffer.concat([d.update(sealed.subarray(20, sealed.length - 16)), d.final()]);
  } catch {
    throw new LockerError('tampered', 'A piece of this file is damaged.');
  } finally {
    key.fill(0);
  }
}

const piecePath = (dir, name) => path.join(dir, 'p', `${name}.rpiece`);
const listName = (seq, rand) => `list-${String(seq).padStart(8, '0')}-${rand}.rlist`;
const parseListName = (n) => {
  const m = /^list-(\d{8})-([0-9a-f]{6})\.rlist$/.exec(n);
  return m ? { seq: Number(m[1]), rand: m[2], file: n } : null;
};

/** The file's identity, written once: its id and salt. */
export async function readItem(dir) {
  const item = JSON.parse(await fsp.readFile(path.join(dir, 'item.json'), 'utf8'));
  checkVersion(item.v);
  return item;
}

async function ensureItem(dir, header) {
  try {
    return await readItem(dir);
  } catch (err) {
    if (err.code !== 'ENOENT') throw err;
  }
  const item = { v: 1, fileId: header.fileId, salt: header.salt };
  await writeJsonAtomic(path.join(dir, 'item.json'), item);
  return item;
}

/** The lists in a large file's folder, newest first. */
export async function listsIn(dir) {
  const names = (await fsp.readdir(dir)).map(parseListName).filter(Boolean);
  return names.sort((a, b) => b.seq - a.seq);
}

/** One list, opened and read: the piece names, sizes, and the file's details. */
async function readList(dir, entry, folderKey) {
  const o = await openLocked(path.join(dir, entry.file), folderKey);
  const chunks = [];
  try {
    for await (const c of o.frames()) chunks.push(c);
  } finally {
    await o.close();
  }
  const list = JSON.parse(Buffer.concat(chunks).toString('utf8'));
  checkVersion(list.v);
  return list;
}

/**
 * A large file sealed in pieces: every piece written and verified, then the
 * list swapped in. Answers the piece names so the caller can prune.
 *
 * @param {object} o
 * @param {string} o.source      the plain file
 * @param {string} o.dir         `<name>.rlock.d`
 * @param {Buffer} o.folderKey
 * @param {Buffer} o.folderId
 * @param {string} o.stageDir
 * @param {object} o.meta        name
 * @param {number} [o.keep]      lists to keep, the new one included (at least 1)
 */
export async function sealPieces({ source, dir, folderKey, folderId, stageDir, meta, keep = 3, hooks = null }) {
  const at = async (s) => { await hooks?.at?.(s); };
  await fsp.mkdir(path.join(dir, 'p'), { recursive: true });
  const seed = newHeader({ folderId });
  const item = await ensureItem(dir, { fileId: seed.subarray(16, 32).toString('hex'), salt: seed.subarray(32, 64).toString('hex') });
  const nkey = nameKey(folderKey, item);
  const table = gearTable(folderKey);
  const fh = await fsp.open(source, 'r');
  const pieces = [];
  const sha = crypto.createHash('sha256');
  let size = 0;
  let written = 0;
  try {
    const buf = Buffer.allocUnsafe(MAX);
    let fill = 0;
    let done = false;
    while (!done || fill > 0) {
      while (!done && fill < MAX) {
        const { bytesRead } = await fh.read(buf, fill, MAX - fill, null);
        if (bytesRead === 0) done = true;
        else fill += bytesRead;
      }
      if (fill === 0) break;
      const len = cutPoint(buf, fill, table);
      const plain = Buffer.from(buf.subarray(0, len));
      buf.copy(buf, 0, len, fill);
      fill -= len;
      sha.update(plain);
      size += plain.length;
      const name = crypto.createHmac('sha256', nkey).update(plain).digest('hex').slice(0, 40);
      pieces.push({ name, size: plain.length });
      written += await writePiece({ dir, folderKey, item, name, plain, stageDir, hooks });
    }
    const stat = await fh.stat();
    if (stat.size !== size) throw new LockerError('busy', 'This file changed while it was being locked.');
  } finally {
    await fh.close().catch(() => {});
    nkey.fill(0);
  }
  await at('pieces');

  // The list: written and swapped in last. The version before stands until this rename.
  const lists = await listsIn(dir);
  const seq = (lists[0]?.seq || 0) + 1;
  const name = listName(seq, crypto.randomBytes(3).toString('hex'));
  const body = Buffer.from(JSON.stringify({ v: 1, ...meta, size, sha256: sha.digest('hex'), pieces, modified: (await fsp.stat(source)).mtimeMs }), 'utf8');
  await fsp.mkdir(stageDir, { recursive: true });
  const plainStage = path.join(stageDir, `.rl-${crypto.randomBytes(12).toString('hex')}.part`);
  await fsp.writeFile(plainStage, body);
  try {
    await sealFile({ source: plainStage, target: path.join(dir, name), rootKey: folderKey, headerBuf: newHeader({ folderId }), meta: { name: 'list' }, stageDir, hooks });
  } finally {
    await fsp.rm(plainStage, { force: true }).catch(() => {});
  }
  await at('listed');
  await pruneLists(dir, folderKey, keep);
  return { pieces: pieces.length, written, size, sha256: JSON.parse(body.toString()).sha256 };
}

async function writePiece({ dir, folderKey, item, name, plain, stageDir, hooks }) {
  const target = piecePath(dir, name);
  if (fs.existsSync(target)) {
    // A piece already there is checked; a good one is left alone, a damaged one replaced below.
    try {
      const got = await fsp.readFile(target);
      if (openPiece(folderKey, item, name, got).equals(plain)) return 0;
    } catch {
      // Replaced by the good copy below.
    }
  }
  const sealed = sealPiece(folderKey, item, name, plain);
  await fsp.mkdir(stageDir, { recursive: true });
  const stage = path.join(stageDir, `.rl-${crypto.randomBytes(12).toString('hex')}.part`);
  const fh = await fsp.open(stage, 'wx');
  try {
    await fh.writeFile(sealed);
    await fh.sync();
  } finally {
    await fh.close();
  }
  try {
    if (!openPiece(folderKey, item, name, sealed).equals(plain)) throw new LockerError('failed', 'A piece did not read back.');
    await placeStaged(stage, target, async (near) => {
      if (!openPiece(folderKey, item, name, await fsp.readFile(near)).equals(plain)) throw new LockerError('failed', 'A piece did not read back.');
    }, { rename: hooks?.rename, retries: hooks?.retries });
  } catch (err) {
    await fsp.rm(stage, { force: true }).catch(() => {});
    throw err;
  }
  return 1;
}

/** A large file opened for reading: the newest list by default, every piece authenticated and the whole checked. */
export async function openPieces({ dir, folderKey, version = 0 }) {
  const lists = await listsIn(dir);
  const chosen = lists[version];
  if (!chosen) throw new LockerError('truncated', 'This locked file has no list of its pieces.');
  const list = await readList(dir, chosen, folderKey);
  const item = await readItem(dir);
  return {
    meta: { name: list.name, size: list.size, sha256: list.sha256, modified: list.modified, pieces: list.pieces.length },
    lists: lists.length,
    close: async () => {},
    frames: async function* frames() {
      const sha = crypto.createHash('sha256');
      let size = 0;
      for (const p of list.pieces) {
        const target = piecePath(dir, p.name);
        let sealed;
        try {
          sealed = await fsp.readFile(target);
        } catch {
          throw new LockerError('truncated', 'A piece of this file is missing. Your cloud may not have finished downloading it.');
        }
        const plain = openPiece(folderKey, item, p.name, sealed);
        if (plain.length !== p.size) throw new LockerError('tampered', 'A piece of this file is damaged.');
        sha.update(plain);
        size += plain.length;
        for (let i = 0; i < plain.length; i += CHUNK) yield plain.subarray(i, Math.min(i + CHUNK, plain.length));
      }
      if (size !== list.size || sha.digest('hex') !== list.sha256) throw new LockerError('tampered', 'This locked file does not match its own record.');
    },
  };
}

/** The lists beyond the kept number, removed, and the pieces only they used. Pieces no list has ever named stay a day, in case another computer's list has not arrived yet. */
async function pruneLists(dir, folderKey, keep) {
  const lists = await listsIn(dir);
  const kept = lists.slice(0, Math.max(1, keep));
  const dropped = lists.slice(Math.max(1, keep));
  const usedByKept = new Set();
  for (const l of kept) {
    try {
      for (const p of (await readList(dir, l, folderKey)).pieces) usedByKept.add(p.name);
    } catch {
      // A list that will not read is not removed from under anything: its pieces stay.
      return;
    }
  }
  const usedByDropped = new Set();
  for (const l of dropped) {
    for (const p of (await readList(dir, l, folderKey).catch(() => ({ pieces: [] }))).pieces) usedByDropped.add(p.name);
    await fsp.rm(path.join(dir, l.file), { force: true }).catch(() => {});
  }
  const names = await fsp.readdir(path.join(dir, 'p')).catch(() => []);
  const day = Date.now() - 24 * 3600 * 1000;
  for (const f of names) {
    if (!f.endsWith('.rpiece')) continue;
    const name = f.slice(0, -7);
    if (usedByKept.has(name)) continue;
    const full = path.join(dir, 'p', f);
    const old = usedByDropped.has(name) || (await fsp.stat(full).then((s) => s.mtimeMs < day, () => false));
    if (old) await fsp.rm(full, { force: true }).catch(() => {});
  }
}
