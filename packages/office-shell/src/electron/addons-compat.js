// What this Office can use of an add-on, and whether an update may be
// offered to someone who has add-ons: the one rule, shared by the loader, the
// update gate and the packages repository's own checks.
//
// Add-ons are the packs a person adds on top of Office from the packages
// repository (github.com/eharain/rutba-office-packages): languages,
// dictionaries, fonts, speech data (model, pronunciation, voice) and the
// speech engine. What this build of Office reads is written once, in
// ../addons-contract.json, and this module decides from it and from nothing
// else. It has no Electron and no file system: the loader (addons.js) runs
// the on-disk checks, with the same refusal codes.
//
//   usable(envelope, { contract, office, keys, engine })
//     -> { usable: true, id, kind, version, revision }
//     -> { usable: false, id, code, reason }
//   readiness({ installed, feed, target, contract })
//     -> { offer, ready: [id], waiting: [{ id, why }], retired: [{ id, reason }] }
//
// Refusal codes: envelope, signature, key-for-kind, kind, format, version,
// files, link, requires. Tests compare the code, never the reason, which is
// a sentence for the Voices and languages page.

import crypto from 'node:crypto';
import { canonical, publicKeyOf } from './packs.js';

/** The contract versions this module reads. */
export const CONTRACT = 1;
/** What an add-on's envelope says it is, as an app pack's does. */
export const ENVELOPE_FORMAT = 'rutba-office-pack';
/** Kinds whose data only the speech engine reads, and so need a usable engine that reads their format. */
export const SPEECH_DATA = ['model', 'pronunciation', 'voice'];
/** English is carried by every Office release, so it never holds an update. */
export const NEVER_HOLDS = ['lang-en'];
/** The words a feed's signature covers before its canonical JSON. */
export const FEED_WORDS = 'Rutba Office packages feed\n';
export const FEED_STATES = ['partial', 'complete'];
const ID = /^[a-z][a-z0-9-]{0,31}$/;
const VERSION = /^\d+\.\d+\.\d+$/;

/** Order two of one id: version, then revision. Below 0, 0 or above 0. */
export function compareVersions(a, b) {
  const pa = String(a?.version).split('.').map(Number);
  const pb = String(b?.version).split('.').map(Number);
  for (let i = 0; i < 3; i += 1) if ((pa[i] || 0) !== (pb[i] || 0)) return (pa[i] || 0) - (pb[i] || 0);
  return (a?.revision || 0) - (b?.revision || 0);
}

const keysOf = (list) => (Array.isArray(list) ? list : []).filter(Boolean).map((k) => publicKeyOf(k));

function signedBy(envelope, keys) {
  const bytes = Buffer.from(canonical(envelope.manifest), 'utf8');
  const signature = Buffer.from(envelope.signature, 'base64');
  return keys.some((k) => {
    try {
      return crypto.verify(null, bytes, k, signature);
    } catch {
      return false;
    }
  });
}

const KIND_WORDS = {
  language: 'language', dictionary: 'dictionary', font: 'font', model: 'speech model', pronunciation: 'pronunciation', voice: 'voice', engine: 'speech engine',
};
const nameOf = (m) => (typeof m?.name === 'string' && m.name ? m.name : m?.id || 'This add-on');

/**
 * Whether this Office can use an add-on, from its envelope and the contract.
 * @param {object} envelope  { format, keyId, manifest, signature }
 * @param {object} o
 * @param {object} o.contract   this Office's addons-contract.json
 * @param {string} o.office     this Office's version
 * @param {{ packages?: *[], release?: *[] }} o.keys  the public keys for each role
 * @param {object|null} [o.engine]  the manifest of the engine this Office uses, if any: it decides the speech data kinds
 */
export function usable(envelope, { contract, office, keys = {}, engine = null }) {
  const m = envelope?.manifest;
  const id = typeof m?.id === 'string' ? m.id : null;
  const no = (code, reason) => ({ usable: false, id, code, reason });
  if (contract?.contract !== CONTRACT) return no('envelope', 'This Rutba Office cannot read its own list of add-ons.');
  if (!envelope || envelope.format !== ENVELOPE_FORMAT || !m || typeof m !== 'object' || typeof envelope.signature !== 'string' || !ID.test(String(id))) {
    return no('envelope', 'This is not a Rutba Office add-on.');
  }
  const name = nameOf(m);
  const kind = contract.kinds?.[m.kind];
  if (!kind) return no('kind', `${name} is a kind of add-on this Rutba Office does not use (${m.kind}).`);

  // The signature, under the keys trusted for this kind only.
  const roles = { packages: keysOf(keys.packages), release: keysOf(keys.release) };
  if (!signedBy(envelope, roles[kind.trust] || [])) {
    const other = kind.trust === 'release' ? roles.packages : roles.release;
    if (signedBy(envelope, other)) return no('key-for-kind', `${name} is signed with a key Rutba does not use for a ${KIND_WORDS[m.kind]}.`);
    return no('signature', `${name} is not signed by Rutba, or was changed after it was signed.`);
  }

  if (!(kind.formats || []).includes(m.format)) {
    return no('format', `${name} is in a format this Rutba Office does not read (${m.format}). Add its version for this Rutba Office.`);
  }
  if (!VERSION.test(String(m.version))) return no('envelope', `${name} does not say which Rutba Office it is for.`);
  if (kind.exactVersion && m.version !== office) {
    return no('version', `${name} is for Rutba Office ${m.version}, and this is ${office}. Add the ${KIND_WORDS[m.kind]} for this version.`);
  }
  if (m.kind === 'engine') {
    const sp = contract.speech || {};
    const host = contract.engineHost || {};
    const same = (a, b) => JSON.stringify([...(a || [])].sort()) === JSON.stringify([...(b || [])].sort());
    if (m.speechApi !== sp.api || m.host !== host.api || !same(m.ipc?.[sp.namespace], sp.calls) || !same(m.events, sp.events)) {
      return no('version', `${name} was made for another Rutba Office. Add the speech engine for this version.`);
    }
    if (!(host.platforms || []).includes(m.platform)) return no('version', `${name} is for ${m.platform}, which this Rutba Office does not run a speech engine on.`);
  }
  if (SPEECH_DATA.includes(m.kind)) {
    const reads = engine?.reads?.[m.kind];
    if (!Array.isArray(reads) || !reads.includes(m.format)) {
      return no('requires', `${name} needs the speech engine for this version of Rutba Office.`);
    }
  }
  return { usable: true, id, kind: m.kind, version: m.version, revision: m.revision };
}

/**
 * Whether an update to `target` may be offered, given the add-ons installed
 * here and the packages feed for `target`: each installed pack is ready (the
 * feed lists it at `target`, not withdrawn, with what it requires ready too),
 * waiting (why: unpublished, withdrawn, needs, no-feed) or retired. English
 * never counts. No add-ons: offer at once, needing no feed.
 * @param {object} o
 * @param {{ id: string, kind?: string, version?: string, revision?: number }[]} o.installed
 * @param {object|null} o.feed  the verified feed for `target`, or null when none could be had
 * @param {string} o.target
 * @param {object} [o.contract]
 */
export function readiness({ installed = [], feed = null, target, contract = null }) {
  const counted = [...new Map(installed.filter((p) => p?.id && !NEVER_HOLDS.includes(p.id)).map((p) => [p.id, p])).values()];
  const byId = (a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
  if (!counted.length) return { offer: true, ready: [], waiting: [], retired: [] };
  if (!feed || feed.office !== target || (contract && !(contract.feed?.formats || [1]).includes(feed.format))) {
    return { offer: false, ready: [], waiting: counted.map((p) => ({ id: p.id, why: 'no-feed' })).sort(byId), retired: [] };
  }
  const packs = Array.isArray(feed.packs) ? feed.packs : [];
  const retiredList = Array.isArray(feed.retired) ? feed.retired : [];
  const listed = new Map(packs.map((p) => [p.id, p]));
  const provider = (name) => packs.find((p) => p.id === name) || packs.find((p) => (p.provides || []).includes(name));
  const published = (p) => p && p.version === target && !p.withdrawn;

  // Whether a pack in the feed is there for `target` with everything it requires, followed through the feed.
  const memo = new Map();
  const whole = (p, seen = new Set()) => {
    if (memo.has(p.id)) return memo.get(p.id);
    if (seen.has(p.id)) return true;
    seen.add(p.id);
    const ok = (p.requires || []).every((r) => {
      const q = provider(r);
      return published(q) && whole(q, seen);
    });
    memo.set(p.id, ok);
    return ok;
  };

  const ready = [];
  const waiting = [];
  const retired = [];
  for (const mine of counted) {
    const gone = retiredList.find((r) => r?.id === mine.id);
    if (gone) { retired.push({ id: mine.id, reason: String(gone.reason || '') }); continue; }
    const p = listed.get(mine.id);
    if (!p) {
      if (feed.state === 'complete') retired.push({ id: mine.id, reason: `Rutba Office ${target} has no ${mine.id}.` });
      else waiting.push({ id: mine.id, why: 'unpublished' });
      continue;
    }
    if (p.withdrawn) { waiting.push({ id: mine.id, why: 'withdrawn' }); continue; }
    if (p.version !== target) { waiting.push({ id: mine.id, why: 'unpublished' }); continue; }
    if (!whole(p)) { waiting.push({ id: mine.id, why: 'needs' }); continue; }
    ready.push(mine.id);
  }
  ready.sort();
  waiting.sort(byId);
  retired.sort(byId);
  return { offer: waiting.length === 0, ready, waiting, retired };
}

/** Whether a feed's signature holds under one of the packages keys. */
export function verifyFeed(feed, sig, keys) {
  const bytes = Buffer.from(FEED_WORDS + canonical(feed), 'utf8');
  const signature = Buffer.from(String(sig?.signature || ''), 'base64');
  return keysOf(keys).some((k) => {
    try {
      return crypto.verify(null, bytes, k, signature);
    } catch {
      return false;
    }
  });
}

/**
 * What Office does with a feed it has fetched and verified: take it for the
 * version it names (the one running, or an update's target), or refuse it.
 * `seen` keeps the highest sequence for each Office version, so the target's
 * feed, read for the update gate, never makes a later feed for the running
 * version look old.
 * @returns {{ ok: true, seen: Record<string, number> } | { ok: false, why: string }}
 */
export function acceptFeed(feed, { office, seen = {}, contract = null }) {
  const formats = contract?.feed?.formats || [1];
  if (!formats.includes(feed?.format)) return { ok: false, why: `The list of add-ons is in a format this Rutba Office does not read (${feed?.format}).` };
  if (feed.office !== office) return { ok: false, why: `The list of add-ons is for Rutba Office ${feed.office}, not ${office}.` };
  if (!FEED_STATES.includes(feed.state) || !Array.isArray(feed.retired) || !Array.isArray(feed.packs)) return { ok: false, why: 'The list of add-ons is not complete.' };
  const highest = Number(seen[office]) || 0;
  if (!(Number.isInteger(feed.sequence) && feed.sequence > highest)) return { ok: false, why: 'The list of add-ons is older than one this Rutba Office has already read.' };
  return { ok: true, seen: { ...seen, [office]: feed.sequence } };
}

/**
 * The contract a running Office decides with: the file as built, its version
 * set to this Office's, and the release key the core already holds for app
 * packs added to the keys trusted for the engine. The packages keys are the
 * file's own.
 * @param {object} file  addons-contract.json
 * @param {{ office: string, releaseKeys?: *[] }} o
 */
export function contractFor(file, { office, releaseKeys = [] }) {
  const release = [...(file.keys?.release || []), ...releaseKeys.map((k) => publicKeyOf(k).export({ type: 'spki', format: 'pem' }).toString())];
  return { ...file, office, keys: { packages: [...(file.keys?.packages || [])], release: [...new Set(release)] } };
}
