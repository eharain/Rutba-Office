// The add-ons at each start, before the first window (private design:
// office-pack-compatibility-design.md 1.2, 1.3 and 5.4; C12 and C13).
//
// A start never fails for an add-on:
//   - each add-on is checked in its own try (addons.js), and the add-ons step
//     has a time budget: what is not checked in time waits for the next start;
//   - a count is kept of starts that did not reach the first window since the
//     add-ons last changed. After two, the next start skips every add-on and
//     says so, with Try again; changing the add-ons (one removed, one added)
//     starts the count again;
//   - `--no-addons` starts without any;
//   - every message is for the app's own windows, never a native dialog.
//
// Before the loader runs: an update's packs staged for this version are
// checked with this version's own rules and moved in (C12), and the removals
// asked for are done, where this process can write the folder: the
// person's own, for every copy in 2.26.0 (addons-folder.js).

import path from 'node:path';
import { createRequire } from 'node:module';
import { findAddons, noAddons, checkFolder, sha512Of, envelopeAt, envelopesIn } from '@rutba/office-shell/electron/addons';
import { usable, compareVersions } from '@rutba/office-shell/electron/addons-compat';
import { lockFolder, applyRemovals, letGo, readState, unmarkRemoval, putInPlace, stagingFor, canWrite, pathsOf, isLink, linkFree, removeTree, STAGING, PARTS } from './addons-folder.js';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');

/** Starts that may fail to reach the first window, after the add-ons changed, before add-ons are skipped. */
export const FAILED_STARTS = 2;
/** The add-ons step's time at start, beyond which what is left waits for the next. */
export const BUDGET_MS = 4000;

/** What the add-ons folder holds, as a short string: it changes when an add-on is added, updated or removed. */
export function fingerprint(dir) {
  try {
    return fs.readdirSync(dir).filter((n) => n.endsWith('.pack.json')).sort().map((n) => {
      const st = fs.statSync(path.join(dir, n));
      return `${n}:${st.size}:${Math.round(st.mtimeMs)}`;
    }).join('|');
  } catch {
    return '';
  }
}

function readMemory(file) {
  try {
    const m = JSON.parse(fs.readFileSync(file, 'utf8'));
    return { fp: String(m.fp ?? ''), unfinished: Number.isInteger(m.unfinished) ? m.unfinished : 0, skipping: Boolean(m.skipping) };
  } catch {
    return { fp: null, unfinished: 0, skipping: false };
  }
}
function writeMemory(file, m) {
  try {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify(m));
  } catch {
    // The count is a safeguard; a profile folder that cannot be written loses it, never the start.
  }
}

/** A staged file's bytes found where they already are: the current folder or the kept previous, when the hash is the same. */
function reuse(dir, id, rel, f) {
  const p = pathsOf(dir, id);
  for (const base of [p.folder, path.join(p.previous, id)]) {
    const at = path.join(base, ...rel.split('/'));
    try {
      if (fs.statSync(at).size === f.size && sha512Of(at) === f.sha512) return at;
    } catch {
      // not there
    }
  }
  return null;
}

/**
 * The packs staged for this version, checked with its rules and moved in.
 * The staged folder for an id holds the envelope and the files that changed;
 * each file it lacks is copied from the add-on in place when its hash is the
 * same. A pack whose only fault is that it needs another (requires) is moved
 * in all the same: it waits there for its need, as an added one does.
 * @param {object} o
 * @param {string} o.dir       the add-ons folder
 * A model is checked as every start checks it, by size: its bytes were
 * hashed as they came, and the folder is the person's own to change at any
 * time, so hashing them again would hold the first window for nothing.
 * A pack that is fine but could not be moved in now (a file of the one in
 * place open) stays staged, in `kept`, and is tried again at the next start.
 * @returns {{ moved: string[], refused: { id, code, reason }[], kept: string[] }}
 */
export function applyStaged({ dir, contract, office, platform = `${process.platform}-${process.arch}`, touch = () => {} }) {
  const own = stagingFor(dir, office);
  const moved = [];
  const refused = [];
  const kept = [];
  if (!fs.existsSync(own)) return { moved, refused, kept };
  const staged = envelopesIn(own);
  const keys = contract.keys || {};
  // The engine first: which speech data may be used depends on it.
  const order = staged.sort((a, b) => (a.envelope?.manifest?.kind === 'engine' ? -1 : 0) - (b.envelope?.manifest?.kind === 'engine' ? -1 : 0));
  let engine = null;
  for (const { id, envelope } of order) {
    touch();
    try {
      const u = usable(envelope, { contract, office, keys, engine });
      if (!u.usable && u.code !== 'requires') { refused.push({ id, code: u.code, reason: u.reason }); continue; }
      const m = envelope.manifest;
      if (m.id !== id) { refused.push({ id, code: 'envelope', reason: `${id}.pack.json is for ${m.id}.` }); continue; }
      if (m.kind === 'engine' && m.platform !== platform) { refused.push({ id, code: 'version', reason: `${m.name || id} is for ${m.platform}.` }); continue; }
      const folder = path.join(own, id);
      fs.mkdirSync(folder, { recursive: true });
      for (const [rel, f] of Object.entries(m.files || {})) {
        const at = path.join(folder, ...String(rel).split('/'));
        if (fs.existsSync(at)) continue;
        const from2 = reuse(dir, id, rel, f);
        if (!from2) continue; // checkFolder reports it missing
        fs.mkdirSync(path.dirname(at), { recursive: true });
        fs.copyFileSync(from2, at);
      }
      const disk = checkFolder(folder, m);
      if (disk) { refused.push({ id, ...disk }); continue; }
      try {
        putInPlace(dir, own, id);
      } catch (err) {
        kept.push(id);
        refused.push({ id, code: 'busy', reason: `${id} could not be moved in now (${err?.code || err?.message || err}); it is tried again at the next start.` });
        continue;
      }
      moved.push(id);
      if (m.kind === 'engine') engine = m;
    } catch (err) {
      refused.push({ id, code: 'files', reason: `${id} could not be moved in (${err?.code || err?.message || err}).` });
    }
  }
  const keep = new Set(kept.flatMap((id) => [id, `${id}.pack.json`]));
  // What is left of this version's staging, but downloads in progress, and every older version's.
  // A newer version's is kept: it is staged for an update not yet installed.
  // Never through a link: a .staging or a version folder that is one is left alone, and what it points at with it.
  try {
    const root = path.join(dir, STAGING);
    for (const name of linkFree(dir, root) ? fs.readdirSync(root) : []) {
      const at = path.join(root, name);
      if (isLink(at)) continue;
      if (name === office) {
        for (const inner of fs.readdirSync(at)) if (inner !== PARTS && !keep.has(inner) && !isLink(path.join(at, inner))) removeTree(path.join(at, inner));
      } else if (/^\d+\.\d+\.\d+$/.test(name) && compareVersions({ version: name }, { version: office }) < 0) {
        removeTree(at);
      }
    }
  } catch {
    // nothing staged
  }
  return { moved, refused, kept };
}

/**
 * The add-ons for this start.
 * @param {object} o
 * @param {{ dir: string, scope: string } | null} o.folder   addons-folder.js addonsFolder()
 * @param {object} o.contract   addons-compat contractFor()
 * @param {string} o.office
 * @param {string} o.memoryFile where the count of failed starts is kept (the profile's folder: written as the person in every install)
 * @param {string[]} [o.argv]
 * @param {string} [o.platform]
 * @param {number} [o.budgetMs]
 * @param {() => number} [o.now]
 * @param {string[]} [o.off]   ids turned off (withdrawn by their makers with no version kept)
 * @param {string[]} [o.keep]  removals taken back while another Office held the folder (the profile's addons.keepAtRestart)
 * @returns {{ registry, skipped: null|'switch'|'failed-starts'|'fault', said: string|null, applied, removed, takenBack: string[], writable: boolean, reached: () => void, tryAgain: () => void }}
 *   `takenBack`: the ids of `keep` whose marks are gone from the folder now, for the profile to let go of
 */
export function startAddons({ folder, contract, office, memoryFile, argv = process.argv, platform = `${process.platform}-${process.arch}`, budgetMs = BUDGET_MS, now = Date.now, off = [], keep = [] }) {
  const none = (skipped, said, installed = []) => ({ registry: noAddons({ installed }), skipped, said, applied: { moved: [], refused: [] }, removed: [], takenBack: [], writable: false, reached: () => {}, tryAgain: () => {} });
  if (!folder?.dir) return none(null, null);
  const dir = folder.dir;
  if (argv.includes('--no-addons')) return none('switch', 'Started without voices and languages, as asked.');

  let applied = { moved: [], refused: [] };
  let removed = [];
  let takenBack = [];
  const writable = fs.existsSync(dir) ? canWrite(dir) : false;
  if (writable) {
    const release = lockFolder(dir);
    if (release) {
      try {
        applied = applyStaged({ dir, contract, office, platform, touch: () => release.refresh?.() });
        release.refresh?.();
        // A removal taken back is never done: its mark goes before the removals are.
        for (const id of keep) if (readState(dir).removals.includes(id)) unmarkRemoval(dir, id);
        takenBack = [...keep];
        removed = applyRemovals(dir).removed;
        release.refresh?.();
        letGo(dir, { office, now: now() });
      } catch {
        // Left for the next start; the add-ons in place still load.
      } finally {
        release();
      }
    }
  }

  const fp = fingerprint(dir);
  const memory = readMemory(memoryFile);
  if (memory.fp !== fp) Object.assign(memory, { fp, unfinished: 0, skipping: false });
  if (memory.skipping || memory.unfinished >= FAILED_STARTS) {
    memory.skipping = true;
    writeMemory(memoryFile, memory);
    const installed = envelopesIn(dir).map((e) => e.id);
    return {
      ...none('failed-starts', 'Started without voices and languages: Rutba Office did not start the last two times after they changed.', installed),
      applied, removed, takenBack, writable,
      tryAgain: () => writeMemory(memoryFile, { fp, unfinished: 0, skipping: false }),
    };
  }
  memory.unfinished += 1;
  writeMemory(memoryFile, memory);

  let registry;
  try {
    const skip = [...readState(dir).removals.filter((x) => !keep.includes(x)), ...off];
    registry = findAddons({ dir, contract, office, platform, skip, deadline: now() + budgetMs, now });
  } catch (err) {
    return { ...none('fault', `Started without voices and languages (${err?.message || err}).`), applied, removed, takenBack, writable, reached: () => writeMemory(memoryFile, { fp, unfinished: 0, skipping: false }) };
  }
  return {
    registry,
    skipped: null,
    said: registry.deferred.length ? `Some voices and languages were not ready in time and will be at the next start: ${registry.deferred.join(', ')}.` : null,
    applied,
    removed,
    takenBack,
    writable,
    // The first window came up: this set of add-ons starts.
    reached: () => writeMemory(memoryFile, { fp, unfinished: 0, skipping: false }),
    tryAgain: () => writeMemory(memoryFile, { fp, unfinished: 0, skipping: false }),
  };
}

/** For the add-ons service: what an envelope in the folder says, or null. */
export const envelopeOf = (dir, id) => envelopeAt(pathsOf(dir, id).envelope);
