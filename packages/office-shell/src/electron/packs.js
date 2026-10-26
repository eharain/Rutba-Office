// The core's side of app packs (docs/APP-PACKS.md): which apps are on this
// computer, and whether each is what Rutba signed.
//
// An app not chosen is absent, not turned off: the core finds the packs in
// one folder at start, checks each, and loads only those that pass. A pack
// is three things side by side:
//
//   <id>.asar            the app's files, an Electron archive
//   <id>.asar.unpacked/  the few the operating system reads from disk (tiles)
//   <id>.pack.json       the signed manifest: what the pack is and every file's SHA-512
//
// A pack is refused, and the reason said in plain words, when its manifest
// is not signed by a key the core holds, when it was made for another
// version of the core, when a file in it is missing, changed or not in the
// manifest, or when a pack it needs is not installed (or was refused
// itself). Nothing here loads a pack: it says what to load, and main.js
// registers each pack's IPC from the entry module named here.
//
// Not yet wired into main.js; the move of the app code into packs does that.
// No Electron import, so the checks run under plain Node; under Electron the
// archive is read with original-fs, which sees an .asar as the file it is.

import crypto from 'node:crypto';
import path from 'node:path';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');

/** What a pack's manifest file says it is. */
export const PACK_FORMAT = 'rutba-office-pack';
export const PACK_FORMAT_VERSION = 1;
const ID = /^[a-z][a-z0-9-]{0,31}$/;

/**
 * The bytes a signature covers: the manifest as JSON with its keys in order
 * and no spaces, so the signer and the core read one string, whatever
 * order a tool wrote the keys in.
 */
export function canonical(value) {
  if (Array.isArray(value)) return `[${value.map((v) => (v === undefined ? 'null' : canonical(v))).join(',')}]`;
  if (value && typeof value === 'object') {
    return `{${Object.keys(value)
      .filter((k) => value[k] !== undefined)
      .sort()
      .map((k) => `${JSON.stringify(k)}:${canonical(value[k])}`)
      .join(',')}}`;
  }
  return JSON.stringify(value);
}

/** A public key from a KeyObject, PEM text or DER bytes. */
export function publicKeyOf(key) {
  if (key instanceof crypto.KeyObject) return key.type === 'public' ? key : crypto.createPublicKey(key);
  if (Buffer.isBuffer(key) || key instanceof Uint8Array) return crypto.createPublicKey({ key: Buffer.from(key), format: 'der', type: 'spki' });
  return crypto.createPublicKey(String(key));
}

/** A key's short name: the first 16 hex digits of the SHA-256 of its public half. */
export function keyIdOf(key) {
  const der = publicKeyOf(key).export({ format: 'der', type: 'spki' });
  return crypto.createHash('sha256').update(der).digest('hex').slice(0, 16);
}

/* ── the archive ───────────────────────────────────────────────────── */

/**
 * An Electron archive's table of files, read from its header: two Chromium
 * pickles, the first holding the size of the second, the second holding
 * the header as a JSON string. A file's bytes start at 8 + that size + its
 * offset; an unpacked one lives beside the archive in <archive>.unpacked/.
 * @returns {Map<string, { size: number, offset: number, unpacked: boolean, link?: string }>}
 */
export function readArchive(file) {
  const fd = fs.openSync(file, 'r');
  try {
    const head = Buffer.alloc(8);
    if (fs.readSync(fd, head, 0, 8, 0) !== 8) throw new Error('it is too short to be an archive');
    const size = head.readUInt32LE(4);
    if (head.readUInt32LE(0) !== 4 || size < 8 || size > 64 * 1024 * 1024) throw new Error('its header is not an archive header');
    const buf = Buffer.alloc(size);
    if (fs.readSync(fd, buf, 0, size, 8) !== size) throw new Error('its header is cut short');
    const length = buf.readUInt32LE(4);
    if (8 + length > size) throw new Error('its header is cut short');
    const header = JSON.parse(buf.toString('utf8', 8, 8 + length));
    const base = 8 + size;
    const out = new Map();
    const walk = (node, prefix) => {
      for (const [name, child] of Object.entries(node.files || {})) {
        const p = prefix ? `${prefix}/${name}` : name;
        if (child.files) walk(child, p);
        else if (child.link) out.set(p, { link: child.link, size: 0, offset: 0, unpacked: false });
        else out.set(p, { size: Number(child.size), offset: base + Number(child.offset || 0), unpacked: Boolean(child.unpacked) });
      }
    };
    walk(header, '');
    return out;
  } finally {
    fs.closeSync(fd);
  }
}

const chunk = Buffer.alloc(1024 * 1024);

/** The SHA-512 of `size` bytes at `at` in an open file, read in pieces; null when it is cut short. */
function hashAt(fd, at, size) {
  const hash = crypto.createHash('sha512');
  let left = size;
  while (left > 0) {
    const n = fs.readSync(fd, chunk, 0, Math.min(chunk.length, left), at);
    if (n <= 0) return null;
    hash.update(chunk.subarray(0, n));
    left -= n;
    at += n;
  }
  return hash.digest('base64');
}

/**
 * The SHA-512 of each file in a pack. The archive is opened once (opening a
 * file is what costs on a computer whose virus checker looks at every
 * open); a file beside it, unpacked, is opened on its own and must be
 * exactly its listed size.
 */
function hashesOf(archive, table) {
  const out = new Map();
  const fd = fs.openSync(archive, 'r');
  try {
    for (const [p, entry] of table) {
      if (entry.link) continue;
      if (!entry.unpacked) {
        out.set(p, hashAt(fd, entry.offset, entry.size));
        continue;
      }
      const file = path.join(`${archive}.unpacked`, ...p.split('/'));
      if (!fs.existsSync(file)) continue;
      const ufd = fs.openSync(file, 'r');
      try {
        out.set(p, fs.fstatSync(ufd).size === entry.size ? hashAt(ufd, 0, entry.size) : null);
      } finally {
        fs.closeSync(ufd);
      }
    }
  } finally {
    fs.closeSync(fd);
  }
  return out;
}

/** Every file in a folder, as paths with forward slashes relative to it. */
function filesIn(dir, prefix = '', out = []) {
  if (!fs.existsSync(dir)) return out;
  for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = prefix ? `${prefix}/${ent.name}` : ent.name;
    if (ent.isDirectory()) filesIn(path.join(dir, ent.name), p, out);
    else out.push(p);
  }
  return out;
}

/** A path a manifest may name: relative, forward slashes, no step out. */
const safePath = (p) => typeof p === 'string' && p.length > 0 && !p.startsWith('/') && !p.includes('\\') && !/^[a-zA-Z]:/.test(p) && !p.split('/').some((s) => s === '..' || s === '.' || s === '');

/* ── checking one pack ─────────────────────────────────────────────── */

class Refusal extends Error {
  constructor(code, reason) {
    super(reason);
    this.code = code;
  }
}

/**
 * Checks one pack against the keys and the core's version; answers what to
 * load, or throws a Refusal saying why not.
 */
function checkPack({ dir, file, keys, version }) {
  const stem = file.slice(0, -'.pack.json'.length);
  let signed;
  try {
    signed = JSON.parse(fs.readFileSync(path.join(dir, file), 'utf8'));
  } catch (err) {
    throw new Refusal('unreadable', `The ${stem} pack's manifest cannot be read (${err.message}).`);
  }
  const manifest = signed?.manifest;
  if (signed?.format !== PACK_FORMAT || !manifest || typeof manifest !== 'object' || typeof signed.signature !== 'string') {
    throw new Refusal('format', `${file} is not a Rutba Office pack manifest.`);
  }

  // The signature first: nothing else in the manifest is believed until it holds.
  const bytes = Buffer.from(canonical(manifest), 'utf8');
  const signature = Buffer.from(signed.signature, 'base64');
  const good = keys.some((k) => {
    try {
      return crypto.verify(null, bytes, k, signature);
    } catch {
      return false;
    }
  });
  if (!good) throw new Refusal('signature', `The ${stem} pack is not signed by Rutba, or was changed after it was signed. It was not loaded.`);

  const name = typeof manifest.name === 'string' && manifest.name ? manifest.name : stem;
  if (manifest.format !== PACK_FORMAT_VERSION) throw new Refusal('format', `The ${name} pack is in a format this Rutba Office does not read (${manifest.format}).`);
  if (!ID.test(String(manifest.id)) || manifest.id !== stem) throw new Refusal('id', `The ${name} pack's manifest names it "${manifest.id}" but its file is ${file}.`);
  if (manifest.version !== version) {
    throw new Refusal('version', `The ${name} pack is version ${manifest.version} and this Rutba Office is ${version}. A pack must match the Rutba Office it runs in: update it with Add or remove apps.`);
  }
  const files = manifest.files;
  if (!files || typeof files !== 'object' || !Object.keys(files).length) throw new Refusal('format', `The ${name} pack lists no files.`);
  for (const [p, f] of Object.entries(files)) {
    if (!safePath(p) || !Number.isSafeInteger(f?.size) || typeof f?.sha512 !== 'string') throw new Refusal('format', `The ${name} pack lists a file it may not: ${p}.`);
  }
  for (const key of ['main', 'windows']) {
    if (manifest[key] != null && (!safePath(manifest[key]) || !files[manifest[key]])) throw new Refusal('format', `The ${name} pack's ${key} module (${manifest[key]}) is not one of its files.`);
  }
  if (!manifest.main) throw new Refusal('format', `The ${name} pack names no main module.`);
  const needs = manifest.needs ?? [];
  if (!Array.isArray(needs) || !needs.every((n) => ID.test(String(n)))) throw new Refusal('format', `The ${name} pack's list of the packs it needs cannot be read.`);

  // The archive: exactly the files the manifest lists, each as it was signed.
  const archive = path.join(dir, manifest.archive || `${stem}.asar`);
  if (!safePath(manifest.archive || `${stem}.asar`) || path.dirname(archive) !== path.resolve(dir)) throw new Refusal('format', `The ${name} pack names an archive outside its folder.`);
  if (!fs.existsSync(archive)) throw new Refusal('missing', `The ${name} pack's archive (${path.basename(archive)}) is missing.`);
  let table;
  try {
    table = readArchive(archive);
  } catch (err) {
    throw new Refusal('unreadable', `The ${name} pack's archive cannot be read: ${err.message}.`);
  }
  for (const [p, entry] of table) {
    if (entry.link) throw new Refusal('unknown', `The ${name} pack holds a link, which a pack may not: ${p}.`);
    if (!files[p]) throw new Refusal('unknown', `The ${name} pack holds a file its manifest does not list: ${p}. It may have been tampered with.`);
  }
  // Beside the archive, only the unpacked files it lists.
  const listedUnpacked = new Set([...table].filter(([, e]) => e.unpacked).map(([p]) => p));
  for (const p of filesIn(`${archive}.unpacked`)) {
    if (!listedUnpacked.has(p)) throw new Refusal('unknown', `The ${name} pack's folder holds a file its manifest does not list: ${p}. It may have been tampered with.`);
  }
  for (const [p, f] of Object.entries(files)) {
    const entry = table.get(p);
    if (!entry) throw new Refusal('missing', `The ${name} pack is missing its file ${p}.`);
    if (Boolean(f.unpacked) !== entry.unpacked) throw new Refusal('changed', `The ${name} pack's file ${p} is not where it was signed.`);
    if (entry.unpacked && !fs.existsSync(path.join(`${archive}.unpacked`, ...p.split('/')))) throw new Refusal('missing', `The ${name} pack is missing its file ${p}.`);
  }
  const hashes = hashesOf(archive, table);
  for (const [p, f] of Object.entries(files)) {
    const entry = table.get(p);
    if (entry.size !== f.size || hashes.get(p) !== f.sha512) {
      throw new Refusal('changed', `The ${name} pack's file ${p} has been changed since it was signed. It may have been tampered with; install the pack again with Add or remove apps.`);
    }
  }

  const inArchive = (p) => (table.get(p)?.unpacked ? path.join(`${archive}.unpacked`, ...p.split('/')) : path.join(archive, ...p.split('/')));
  return {
    id: manifest.id,
    name,
    version: manifest.version,
    needs,
    uses: Array.isArray(manifest.uses) ? manifest.uses.filter((u) => ID.test(String(u))) : [],
    manifest: path.join(dir, file),
    archive,
    // Electron reads into an archive by path, so these are paths a require or import can take.
    main: inArchive(manifest.main),
    windows: manifest.windows ? inArchive(manifest.windows) : null,
    ipc: manifest.ipc && typeof manifest.ipc === 'object' ? manifest.ipc : {},
    events: Array.isArray(manifest.events) ? manifest.events : [],
    tiles: (Array.isArray(manifest.tiles) ? manifest.tiles : []).filter((t) => safePath(t?.icon) && files[t.icon]).map((t) => ({ ...t, icon: inArchive(t.icon) })),
    fileTypes: Array.isArray(manifest.fileTypes) ? manifest.fileTypes : [],
    tools: Array.isArray(manifest.tools) ? manifest.tools : [],
  };
}

/* ── finding the packs ─────────────────────────────────────────────── */

/**
 * The packs in a folder that may be loaded, and those that may not.
 *
 * @param {object} o
 * @param {string} o.dir             the folder the installer copies packs into
 * @param {*} o.publicKey            the key (or keys) packs must be signed with: KeyObject, PEM or DER
 * @param {string} o.version         the core's own version; a pack must carry the same
 * @returns {{ packs: object[], refused: { id: string, file: string, code: string, reason: string }[] }}
 *   `packs` in an order that loads what a pack needs before it.
 */
export function findPacks({ dir, publicKey, version }) {
  if (!version) throw new Error('findPacks needs the core\'s version.');
  dir = path.resolve(dir);
  const keys = (Array.isArray(publicKey) ? publicKey : [publicKey]).filter(Boolean).map(publicKeyOf);
  if (!keys.length) throw new Error('findPacks needs the public key packs are signed with.');
  const refused = [];
  const found = new Map();
  const names = fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => f.endsWith('.pack.json')).sort() : [];
  for (const file of names) {
    try {
      const pack = checkPack({ dir, file, keys, version });
      found.set(pack.id, pack);
    } catch (err) {
      if (!(err instanceof Refusal)) throw err;
      refused.push({ id: file.slice(0, -'.pack.json'.length), file, code: err.code, reason: err.message });
    }
  }

  // A pack whose needs are not all here is refused too, and so, in turn, is one that needs it.
  for (let changed = true; changed; ) {
    changed = false;
    for (const pack of [...found.values()]) {
      const missing = pack.needs.filter((n) => !found.has(n));
      if (!missing.length) continue;
      found.delete(pack.id);
      const bad = missing.filter((n) => refused.some((r) => r.id === n));
      const absent = missing.filter((n) => !bad.includes(n));
      const words = [`The ${pack.name} pack was not loaded.`];
      if (absent.length) words.push(`It needs ${absent.join(', ')}, which ${absent.length > 1 ? 'are' : 'is'} not installed: add ${absent.length > 1 ? 'them' : 'it'} with Add or remove apps.`);
      if (bad.length) words.push(`It needs ${bad.join(', ')}, which could not be loaded either.`);
      refused.push({ id: pack.id, file: path.basename(pack.manifest), code: 'needs', reason: words.join(' ') });
      changed = true;
    }
  }

  // What a pack needs is loaded before it; otherwise in the order of their names.
  const order = [];
  const placed = new Set();
  const visiting = new Set();
  const place = (pack) => {
    if (placed.has(pack.id) || visiting.has(pack.id)) return;
    visiting.add(pack.id);
    for (const n of [...pack.needs].sort()) place(found.get(n));
    visiting.delete(pack.id);
    placed.add(pack.id);
    order.push(pack);
  };
  for (const id of [...found.keys()].sort()) place(found.get(id));
  return { packs: order, refused };
}
