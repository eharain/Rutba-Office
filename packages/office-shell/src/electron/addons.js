// The add-ons on this computer: which may be used, and which may not and why
// (private design: office-pack-compatibility-design.md 1.2).
//
// An add-on is a plain folder beside its signed envelope, in the add-ons
// folder (apps/desktop/main/addons-folder.js says where):
//
//   lang-fr/              the pack's files
//   lang-fr.pack.json     { format, keyId, manifest, signature }
//   lang-fr.previous/     the version kept for rollback: lang-fr/ and lang-fr.pack.json
//   .staging/             downloads and staged updates, never loaded
//
// Each add-on is checked on its own, stopping at the first fault: what the
// envelope and the contract decide, by addons-compat.js (envelope, kind,
// signature for its kind, format, the engine's version and APIs, the speech
// data an engine reads), then the files on disk: every listed file there
// with its size and SHA-512, no file outside the list, no link or junction.
// A model is checked by size here; the engine hashes its bytes as it reads
// them. Per id the highest (version, revision) that passes is used, the
// other kept as the previous. Then `requires` is resolved among what passed,
// by id or by a name a pack provides.
//
// Data kinds are never imported or run. An engine's main is loaded like an
// app pack's (app-packs.js), and only the engine.
//
// No Electron import: the checks run under plain Node; under Electron files
// are read with original-fs, so nothing is taken for an archive.

import crypto from 'node:crypto';
import path from 'node:path';
import { createRequire } from 'node:module';
import { usable, compareVersions, safePath } from './addons-compat.js';
import { isListed } from './packs.js';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');

const ID = /^[a-z][a-z0-9-]{0,31}$/;
/** Kinds checked by size only at start; their bytes are checked as they are read. */
const SIZE_ONLY = ['model'];
const ENGINE = 'engine';

/** The SHA-512 of a file, read in pieces. */
export function sha512Of(file) {
  const hash = crypto.createHash('sha512');
  const fd = fs.openSync(file, 'r');
  try {
    const chunk = Buffer.alloc(1024 * 1024);
    for (let at = 0; ; ) {
      const n = fs.readSync(fd, chunk, 0, chunk.length, at);
      if (n <= 0) break;
      hash.update(chunk.subarray(0, n));
      at += n;
    }
  } finally {
    fs.closeSync(fd);
  }
  return hash.digest('base64');
}

/**
 * The files of an add-on's folder against its manifest: the refusal, or null.
 * `hashAll` hashes a model's files too, as moving a staged one in does.
 * @returns {{ code: 'files'|'link', reason: string } | null}
 */
export function checkFolder(dir, manifest, { hashAll = false } = {}) {
  const name = manifest.name || manifest.id;
  if (!fs.existsSync(dir) || !fs.lstatSync(dir).isDirectory()) return { code: 'files', reason: `${name} is missing its files.` };
  if (fs.lstatSync(dir).isSymbolicLink()) return { code: 'link', reason: `${name} is a link, which an add-on may not be.` };
  const files = manifest.files;
  if (!files || typeof files !== 'object' || !Object.keys(files).length) return { code: 'files', reason: `${name} lists no files.` };
  for (const [p, f] of Object.entries(files)) {
    if (!safePath(p) || !Number.isSafeInteger(f?.size) || typeof f?.sha512 !== 'string') return { code: 'files', reason: `${name} lists a file it may not: ${p}.` };
  }
  // Nothing on disk that the manifest does not list, and no link or junction anywhere.
  const walk = (rel) => {
    for (const e of fs.readdirSync(path.join(dir, rel), { withFileTypes: true })) {
      const p = rel ? `${rel}/${e.name}` : e.name;
      const full = path.join(dir, ...p.split('/'));
      if (e.isSymbolicLink() || fs.lstatSync(full).isSymbolicLink()) return { code: 'link', reason: `${name} holds a link, which an add-on may not: ${p}.` };
      if (e.isDirectory()) { const r = walk(p); if (r) return r; continue; }
      if (!isListed(files, p)) return { code: 'files', reason: `${name} holds a file it does not list: ${p}. It may have been tampered with; add it again.` };
    }
    return null;
  };
  const extra = walk('');
  if (extra) return extra;
  for (const [p, f] of Object.entries(files)) {
    const full = path.join(dir, ...p.split('/'));
    if (!fs.existsSync(full)) return { code: 'files', reason: `${name} is missing its file ${p}. Add it again.` };
    if (fs.statSync(full).size !== f.size) return { code: 'files', reason: `${name}'s file ${p} has been changed since it was signed. Add it again.` };
    if (SIZE_ONLY.includes(manifest.kind) && !hashAll) continue;
    if (sha512Of(full) !== f.sha512) return { code: 'files', reason: `${name}'s file ${p} has been changed since it was signed. Add it again.` };
  }
  return null;
}

function readEnvelope(file) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return null;
  }
}

/** The candidates for one id: the current and, when kept, the previous. */
function candidatesIn(root) {
  const out = new Map();
  if (!fs.existsSync(root)) return out;
  const add = (id, envelopeFile, folder, slot) => {
    if (!out.has(id)) out.set(id, []);
    out.get(id).push({ id, envelopeFile, folder, slot });
  };
  for (const name of fs.readdirSync(root).sort()) {
    if (name.startsWith('.')) continue;
    if (name.endsWith('.pack.json')) {
      const id = name.slice(0, -'.pack.json'.length);
      if (ID.test(id)) add(id, path.join(root, name), path.join(root, id), 'current');
    } else if (name.endsWith('.previous')) {
      const id = name.slice(0, -'.previous'.length);
      const env = path.join(root, name, `${id}.pack.json`);
      if (ID.test(id) && fs.existsSync(env)) add(id, env, path.join(root, name, id), 'previous');
    }
  }
  return out;
}

/**
 * The add-ons in a folder that may be used, and those that may not.
 *
 * @param {object} o
 * @param {string|null} o.dir       the add-ons folder, or null: none
 * @param {object} o.contract       the contract this Office decides with (addons-compat contractFor)
 * @param {string} o.office         this Office's version
 * @param {string} [o.platform]     this computer's, as an engine names it: win32-x64
 * @param {string[]} [o.skip]       ids not to load (removals waiting for the next start)
 * @param {number} [o.deadline]     a time (Date.now()) after which no more add-ons are checked this start: those left are `deferred`
 * @param {() => number} [o.now]
 * @returns {Registry}
 */
export function findAddons({ dir, contract, office, platform = `${process.platform}-${process.arch}`, skip = [], deadline = Infinity, now = Date.now }) {
  const refused = [];
  const deferred = [];
  const passed = new Map(); // id -> [{ ...candidate, manifest, envelope }]
  const groups = dir ? candidatesIn(path.resolve(dir)) : new Map();
  const keys = contract.keys || {};

  const check = (c, engine) => {
    try {
      return checkOne(c, engine);
    } catch (err) {
      // A fault reading one add-on (a file held open, a folder it may not read) is that add-on's alone.
      return { usable: false, id: c.id, code: 'files', reason: `${c.id} could not be read (${err?.code || err?.message || err}). Add it again.` };
    }
  };
  const checkOne = (c, engine) => {
    const envelope = readEnvelope(c.envelopeFile);
    const u = usable(envelope, { contract, office, keys, engine });
    if (!u.usable) return { ...u, id: c.id };
    const m = envelope.manifest;
    if (m.id !== c.id) return { usable: false, id: c.id, code: 'envelope', reason: `${c.id}.pack.json is for ${m.id}.` };
    if (m.kind === ENGINE && m.platform !== platform) return { usable: false, id: c.id, code: 'version', reason: `${m.name || c.id} is for ${m.platform}, and this computer is ${platform}.` };
    const disk = checkFolder(c.folder, m);
    if (disk) return { usable: false, id: c.id, ...disk };
    return { usable: true, manifest: m, envelope };
  };

  // Engines first: which speech data may be used depends on the engine.
  const kindOf = (c) => readEnvelope(c.envelopeFile)?.manifest?.kind;
  const ids = [...groups.keys()].filter((id) => !skip.includes(id));
  const engines = ids.filter((id) => groups.get(id).some((c) => kindOf(c) === ENGINE));
  const pick = (id, engine) => {
    if (now() > deadline) { deferred.push(id); return; }
    const ok = [];
    const bad = [];
    for (const c of groups.get(id)) {
      const r = check(c, engine);
      if (r.usable) ok.push({ ...c, manifest: r.manifest, envelope: r.envelope });
      else bad.push({ ...r, slot: c.slot });
    }
    ok.sort((a, b) => compareVersions(b.manifest, a.manifest));
    if (ok.length) passed.set(id, ok);
    // A refusal is reported for the current copy only when nothing of that id passed, so a kept previous is no fault.
    else refused.push(...bad.filter((b) => b.slot === 'current').map(({ id: rid, code, reason }) => ({ id: rid, code, reason })));
  };
  for (const id of engines) pick(id, null);
  const engine = engines.map((id) => passed.get(id)?.[0]).filter(Boolean).sort((a, b) => compareVersions(b.manifest, a.manifest))[0] || null;
  for (const id of ids.filter((x) => !engines.includes(x))) pick(id, engine?.manifest || null);

  // What each requires, among what passed: a pack whose need is missing waits for it, and so in turn does one that needs that.
  const use = new Map([...passed].map(([id, list]) => [id, list[0]]));
  for (let changed = true; changed; ) {
    changed = false;
    for (const [id, a] of [...use]) {
      const provides = (name) => [...use.values()].some((b) => b.id === name || (b.manifest.provides || []).includes(name));
      const missing = (a.manifest.requires || []).filter((r) => !provides(r));
      if (!missing.length) continue;
      use.delete(id);
      refused.push({ id, code: 'requires', reason: `${a.manifest.name || id} needs ${missing.join(', ')}, which ${missing.length > 1 ? 'are' : 'is'} not added.` });
      changed = true;
    }
  }

  const addons = [...use.values()].map((a) => ({
    id: a.id,
    kind: a.manifest.kind,
    dir: a.folder,
    envelopeFile: a.envelopeFile,
    manifest: a.manifest,
    slot: a.slot,
    previous: (passed.get(a.id) || []).find((b) => b !== a) ? { version: passed.get(a.id).find((b) => b !== a).manifest.version, revision: passed.get(a.id).find((b) => b !== a).manifest.revision } : null,
  })).sort((a, b) => (a.id < b.id ? -1 : 1));
  refused.sort((a, b) => (a.id < b.id ? -1 : 1));
  return registry({ dir: dir ? path.resolve(dir) : null, addons, refused, installed: ids, deferred: deferred.sort() });
}

/**
 * @typedef {object} Registry
 * @property {string|null} dir
 * @property {() => object[]} list        the add-ons that may be used: { id, kind, dir, manifest }
 * @property {(name: string) => object|null} find  by id, or by a name an add-on provides
 * @property {object[]} refused           { id, code, reason }
 * @property {string[]} installed         every id with an envelope here, used or not: what the update gate counts
 * @property {() => object|null} engine   the speech engine to load, if one may be used
 * @property {string[]} deferred          ids not checked this start, the time for add-ons having run out: tried again at the next
 */
function registry({ dir, addons, refused, installed, deferred = [] }) {
  const frozen = addons.map((a) => Object.freeze({ ...a }));
  return Object.freeze({
    dir,
    list: () => [...frozen],
    find: (name) => frozen.find((a) => a.id === name) || frozen.find((a) => (a.manifest.provides || []).includes(name)) || null,
    refused: Object.freeze(refused.map((r) => Object.freeze({ ...r }))),
    installed: Object.freeze([...installed]),
    deferred: Object.freeze([...deferred]),
    engine: () => frozen.find((a) => a.kind === ENGINE) || null,
  });
}

/** No add-ons: what a copy with no folder, or one started with --no-addons, uses. */
export const noAddons = ({ installed = [] } = {}) => registry({ dir: null, addons: [], refused: [], installed });

/** An envelope read from disk, or null: for the add-ons service, which lists what is installed. */
export const envelopeAt = readEnvelope;

/** Every current add-on's envelope in a folder, checked or not: { id, envelope|null }. */
export function envelopesIn(dir) {
  const out = [];
  for (const [id, list] of candidatesIn(path.resolve(dir))) {
    const c = list.find((x) => x.slot === 'current');
    if (c) out.push({ id, envelope: readEnvelope(c.envelopeFile) });
  }
  return out;
}
