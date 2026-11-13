// The add-ons service (C4): what the packages feed offers this Office, adding
// and removing add-ons, keeping them up to date, and the two questions the
// update service asks of them (C11, C12). Generic for every kind; it never
// names a language or a voice. Private design: office-pack-compatibility-design.md
// 1.3, 2.4, 5; office-speech-core-changes.md C4.
//
// The feed of this Office's own version is
//   github.com/eharain/rutba-office-packages/releases/download/v<version>/feed.json
// and feed.json.sig beside it, verified with the packages keys of the
// contract, taken only for the version it names and only with a sequence
// above the highest this folder has seen for that version. A copy of the feed
// built into the release lists what can be added when the network cannot be
// reached. An administrator's share can stand in for the releases
// (RUTBA_ADDONS_SOURCE: v<version>/feed.json and the files beside it).
//
// Adding brings the envelope and every file whose hash is not already on this
// computer into this version's staging folder, each checked as it arrives;
// the envelope goes in last, so nothing half-fetched is ever loaded. An
// add-on is used from the next start, as an app pack is. An install for
// everyone cannot write its folder as the person: its add-ons are staged in
// the person's own folder and moved in by the administrator's prompt at
// Restart, checked again there with this version's rules.
//
// With Office's updates on, a newer revision of an installed add-on is taken
// by itself (per person and portable); for everyone it waits for an
// administrator's click. A revision withdrawn by its makers is rolled back to
// the one kept; with none kept the add-on is turned off. Nothing new is ever
// added without a click.

import path from 'node:path';
import { createRequire } from 'node:module';
import { readiness as decide, verifyFeed, acceptFeed, usable, signedForKind, compareVersions, isAddonId, NEVER_HOLDS } from '@rutba/office-shell/electron/addons-compat';
import { envelopesIn, envelopeAt, checkFolder, sha512Of } from '@rutba/office-shell/electron/addons';
import { readState, writeState, markRemoval, unmarkRemoval, putInPlace, rollBack, stagingFor, pathsOf, PARTS } from './addons-folder.js';
import { fetchToFile, releaseUrl, fromSource, allowedUrl, freeSpace, DownloadError } from './addons-download.js';
import { readZip } from '@rutba/ooxml/zip';
import { whatStopsRestart } from './update-error.js';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');

const ID = /^[a-z][a-z0-9-]{0,31}$/;
/** A call into another part that must never break the list: its answer, or null. */
function safe(fn) {
  try {
    return fn();
  } catch {
    return null;
  }
}

/** Room kept free beyond an add-on's size, so a download never fills the disk. */
export const SPARE_BYTES = 200 * 1024 * 1024;
/** The setting that turns off add-ons taking new revisions by themselves. */
export const AUTO_KEY = 'addons.autoUpdate';
const SEEN_KEY = 'addons.seen';
/** The hash of the feed taken at the highest sequence, per version, where the folder cannot be written. */
const SEEN_FEEDS_KEY = 'addons.seenFeeds';
const OFF_KEY = 'addons.off';
/** The version whose first start looked for the fonts its languages need (L8). */
const FONTS_KEY = 'addons.fontsChecked';

/** Every pack a set of ids needs, through `requires`, in the feed: the ids, then what they need. */
export function closure(ids, packs) {
  const byName = (name) => packs.find((p) => p.id === name) || packs.find((p) => (p.provides || []).includes(name));
  const out = [];
  const visit = (name) => {
    const p = byName(name);
    if (!p || out.includes(p.id)) return;
    out.push(p.id);
    for (const r of p.requires || []) visit(r);
  };
  for (const id of ids) visit(id);
  return out;
}

/**
 * @param {object} o
 * @param {{ dir: string, scope: string } | null} o.folder   addonsFolder()
 * @param {object} o.start        startAddons() for this run
 * @param {object} o.contract     contractFor()
 * @param {string} o.office
 * @param {object} o.stores       settings
 * @param {Function} [o.broadcast]
 * @param {Function} [o.fetchImpl]
 * @param {string|null} [o.source]        RUTBA_ADDONS_SOURCE
 * @param {{ feed, sig } | null} [o.bundled]  the feed built into this release
 * @param {string|null} [o.personStaging] where an install for everyone stages as the person
 * @param {() => boolean} [o.updatesOn]   Office's own updates are on
 * @param {object} [o.app]                { relaunch, quit } for Restart
 * @param {Function} [o.elevate]          (args, { wait }) -> Promise<{ ok, code }>: runs this Office once as administrator with those arguments
 * @param {() => { dirty?: boolean, name?: string }[]} [o.openWindows]  the open windows, for Restart's question on unsaved work
 * @param {number} [o.pid]                this process, which the administrator's step waits to quit
 * @param {string} [o.platform]
 * @param {() => ({ engine: string|null, problem: string|null } | null)} [o.engineState]  the speech engine this run, once loaded (engine-host.js)
 * @param {object} [o.languages]  the language packs (L4, L6, L7, L8): known, coverageOf, inUse, builtInInUse() -> [lang-<tag>], and fontsFor({ tags, installed, offered }) -> font pack ids
 */
export function createAddonsService({
  folder, start, contract, office, stores, broadcast = () => {}, fetchImpl = (...a) => fetch(...a), source = process.env.RUTBA_ADDONS_SOURCE || null,
  bundled = null, personStaging = null, updatesOn = () => true, app = null, elevate = null, platform = `${process.platform}-${process.arch}`,
  engineState = () => null,
  languages = null,
  openWindows = null,
  pid = process.pid,
}) {
  const dir = folder?.dir || null;
  const perMachine = folder?.scope === 'machine';
  const writable = Boolean(start?.writable) && !perMachine;
  const registry = start?.registry;
  /** Where this run stages: the folder's own for this version, or the person's for an install for everyone. */
  const stagingRoot = (version) => (writable || !personStaging ? stagingFor(dir, version) : path.join(personStaging, version));
  const jobs = new Map(); // id -> { state, done, total, controller, error }
  const feeds = new Map(); // version -> { feed, from }
  let restart = false;

  // The highest sequence seen per version: the folder's, and the profile's where the folder cannot be written.
  const seen = () => {
    const a = dir ? readState(dir).seen : {};
    const b = stores.settings.get(SEEN_KEY, {}) || {};
    const out = { ...a };
    for (const [v, n] of Object.entries(b)) if (Number.isInteger(n) && !(out[v] >= n)) out[v] = n;
    return out;
  };
  // The hash of the feed taken at that sequence, so the same feed is taken again at the next start.
  const hashes = () => {
    const s = seen();
    const a = dir ? readState(dir).feeds : {};
    const b = stores.settings.get(SEEN_FEEDS_KEY, {}) || {};
    const out = {};
    for (const v of Object.keys(s)) {
      const fromFolder = dir && readState(dir).seen[v] === s[v] ? a[v] : undefined;
      const fromProfile = (stores.settings.get(SEEN_KEY, {}) || {})[v] === s[v] ? b[v] : undefined;
      if (typeof fromFolder === 'string') out[v] = fromFolder;
      else if (typeof fromProfile === 'string') out[v] = fromProfile;
    }
    return out;
  };
  const remember = (next, nextHashes) => {
    if (writable) {
      const s = readState(dir);
      writeState(dir, { ...s, seen: next, feeds: nextHashes });
    } else {
      stores.settings.set(SEEN_KEY, next);
      stores.settings.set(SEEN_FEEDS_KEY, nextHashes);
    }
  };
  const off = () => (stores.settings.get(OFF_KEY, []) || []).filter((id) => ID.test(id));

  const changed = () => broadcast('addons:changed', { restart });
  const progress = (id) => {
    const j = jobs.get(id);
    broadcast('addons:progress', { id, state: j?.state || 'idle', done: j?.done || 0, total: j?.total || 0, error: j?.error || null });
  };

  const fetchText = async (url) => {
    const address = fromSource(url, source);
    if (!allowedUrl(address, { source })) throw new DownloadError('address', 'The list of add-ons is at an address Rutba Office does not fetch from.');
    if (address.startsWith('file:')) {
      const { fileURLToPath } = await import('node:url');
      return fs.readFileSync(fileURLToPath(address), 'utf8');
    }
    const r = await fetchImpl(address, { headers: { 'user-agent': 'Rutba-Office' }, redirect: 'follow', signal: AbortSignal.timeout(60_000) });
    if (!r.ok) throw new DownloadError('http', `The list of add-ons could not be fetched (${r.status}).`);
    return r.text();
  };

  /**
   * The verified feed for a version: fetched, checked with the packages keys,
   * taken only for that version and when newer than any seen. For the
   * running version, the release's built-in copy when the network fails.
   * @returns {Promise<{ feed: object, from: 'network'|'bundled' } | { feed: null, why: string }>}
   */
  async function feedFor(version, { refresh = true } = {}) {
    if (!refresh && feeds.has(version)) return feeds.get(version);
    const keys = contract.keys?.packages || [];
    try {
      const base = releaseUrl(version);
      const [text, sigText] = await Promise.all([fetchText(`${base}feed.json`), fetchText(`${base}feed.json.sig`)]);
      const feed = JSON.parse(text);
      if (!verifyFeed(feed, JSON.parse(sigText), keys)) throw new Error('The list of add-ons is not signed by Rutba.');
      const known = feeds.get(version);
      // The same feed fetched again this run is the one already taken.
      if (known?.feed && known.feed.sequence === feed.sequence && known.from === 'network') return known;
      const taken = acceptFeed(feed, { office: version, seen: seen(), hashes: hashes(), contract });
      if (!taken.ok) {
        if (known?.feed) return known;
        throw new Error(taken.why);
      }
      remember(taken.seen, taken.hashes);
      const got = { feed, from: 'network' };
      feeds.set(version, got);
      return got;
    } catch (err) {
      if (version === office && bundled?.feed && verifyFeed(bundled.feed, bundled.sig, keys) && bundled.feed.office === office) {
        const got = { feed: bundled.feed, from: 'bundled', why: err?.message || String(err) };
        if (!feeds.has(version)) feeds.set(version, got);
        return feeds.get(version);
      }
      return { feed: null, why: err?.message || String(err) };
    }
  }

  /**
   * What is installed, as the update gate counts it: every add-on here the
   * person has not asked to remove, used or not, that is signed for its kind
   * under its own id. An envelope that is not (a file put in the folder by
   * hand) counts for nothing, so it can never hold an update.
   */
  function installedSet({ withRemovals = false } = {}) {
    if (!dir || !fs.existsSync(dir)) return [];
    const removing = [...readState(dir).removals, ...(stores.settings.get('addons.removeAtRestart', []) || [])];
    const signed = (e) => e.envelope?.manifest?.id === e.id && signedForKind(e.envelope, { contract });
    return envelopesIn(dir).filter((e) => (withRemovals || !removing.includes(e.id)) && signed(e)).map(({ id, envelope }) => ({
      id, kind: envelope?.manifest?.kind, version: envelope?.manifest?.version, revision: envelope?.manifest?.revision,
    }));
  }

  function stagedIds() {
    try {
      return fs.readdirSync(stagingRoot(office)).filter((n) => n.endsWith('.pack.json')).map((n) => n.slice(0, -'.pack.json'.length)).sort();
    } catch {
      return [];
    }
  }

  /**
   * What the update gate counts as here: the add-ons installed, and the
   * built-in language the person uses as its lang-<tag> (L7), so an update
   * waits for that pack and brings it, for the release that leaves it out
   * of the core.
   */
  function gateSet() {
    const here = installedSet();
    const builtIn = (safe(() => languages?.builtInInUse?.()) || []).filter((id) => ID.test(id) && !here.some((p) => p.id === id));
    return [...here, ...builtIn.map((id) => ({ id, kind: 'language', version: office, revision: 0, builtIn: true }))];
  }

  /** A language added here (L6): its tag and review, and how much of this Office's words it has, counted against this Office's own keys. */
  function installedLanguage(a, id) {
    const m = a?.manifest || (dir ? envelopeAt(pathsOf(dir, id).envelope)?.manifest : null) || {};
    const counted = safe(() => languages?.coverageOf(id)) || null;
    return {
      tag: m.tag || null, nativeName: m.nativeName || null, direction: m.direction || null, variantOf: m.variantOf || null, review: m.review || null,
      licence: m.licence?.spdx || null, coverage: counted, inUse: (safe(() => languages?.inUse()) || []).includes(id),
    };
  }

  /** A language the feed offers (L6): what it says of itself before it is added; its coverage is of that version's English. */
  const offeredLanguage = (p) => ({ tag: p.tag || null, script: p.script || null, direction: p.direction || null, variantOf: p.variantOf || null, review: p.review || null, coverage: p.coverage || null, packStatus: p.status || null });

  const languagesKnown = (packs) => safe(() => languages?.known(packs)) || [];

  function list() {
    const known = feeds.get(office)?.feed || null;
    const removing = dir ? [...readState(dir).removals, ...(stores.settings.get('addons.removeAtRestart', []) || [])] : [];
    const used = new Map((registry?.list() || []).map((a) => [a.id, a]));
    const refused = new Map((registry?.refused || []).map((r) => [r.id, r]));
    const installed = installedSet({ withRemovals: true }).map((p) => {
      const a = used.get(p.id);
      const r = refused.get(p.id);
      const state = removing.includes(p.id) ? 'removing'
        : off().includes(p.id) ? 'off'
          : a ? 'loaded'
            : r ? (r.code === 'requires' ? 'waiting' : 'refused')
              : (registry?.deferred || []).includes(p.id) ? 'deferred' : 'added';
      return { ...p, name: a?.manifest?.name || p.id, state, reason: r?.reason || null, code: r?.code || null, ...(p.kind === 'language' ? installedLanguage(a, p.id) : {}) };
    });
    const mine = new Map(installed.map((p) => [p.id, p]));
    const offered = (known?.packs || []).filter((p) => isAddonId(p?.id)).map((p) => {
      const have = mine.get(p.id);
      const job = jobs.get(p.id);
      const status = job ? job.state
        : !have ? 'add'
          : compareVersions(p, have) > 0 ? 'update' : 'installed';
      return {
        id: p.id, kind: p.kind, name: p.name, nativeName: p.nativeName || null, version: p.version, revision: p.revision, size: p.size || 0,
        licence: p.licence?.spdx || null, requires: p.requires || [], provides: p.provides || [], withdrawn: Boolean(p.withdrawn), tags: p.tags || [], status,
        ...(p.kind === 'language' ? offeredLanguage(p) : {}),
      };
    });
    return {
      office, dir, scope: folder?.scope || null, writable, perMachine, restart,
      skipped: start?.skipped || null, said: start?.said || null,
      // The speech engine this run: loaded, or why not.
      engine: engineState(),
      automatic: stores.settings.get(AUTO_KEY, true) !== false,
      feed: known ? { office: known.office, sequence: known.sequence, state: known.state, from: feeds.get(office).from } : null,
      installed, offered,
      // Every language there is (L4): the built-in list with the feed's entries over it, for the languages not offered yet.
      languages: languagesKnown(known?.packs || []),
      // Staged and waiting to be moved in: at the next start, or for everyone at Restart with the administrator's permission.
      staged: stagedIds(),
      jobs: [...jobs].map(([id, j]) => ({ id, state: j.state, done: j.done, total: j.total, error: j.error || null })),
    };
  }

  /**
   * One add-on brought into a staging folder from a feed entry: the files it
   * lacks fetched, the rest copied from the add-on in place, the whole
   * checked with this Office's rules, the envelope last.
   * @returns {Promise<string>}  the staging folder holding <id>/ and <id>.pack.json
   */
  async function stageOne(p, { version, root, signal, engine = null, job = { state: 'adding', done: 0, total: 0 } }) {
    // The feed's id names folders here: checked before any is made of it.
    if (!isAddonId(p?.id)) throw new DownloadError('envelope', 'The list of add-ons names one with a name an add-on cannot have.');
    const parts = path.join(root, PARTS, p.id);
    fs.mkdirSync(parts, { recursive: true });
    const envFile = path.join(parts, `${p.id}.pack.json`);
    await fetchToFile({ ...p.envelope, to: envFile, fetchImpl, signal, source });
    const envelope = JSON.parse(fs.readFileSync(envFile, 'utf8'));
    const m = envelope.manifest;
    const u = usable(envelope, { contract: { ...contract, office: version }, office: version, keys: contract.keys, engine });
    // For an update's version only who signed it is decided here: its format and APIs are that version's to judge, at its first start.
    const theirs = version !== office ? ['requires', 'format', 'version'] : ['requires'];
    if (!u.usable && !theirs.includes(u.code)) throw new DownloadError(u.code, u.reason);
    if (m?.id !== p.id || compareVersions(m, p) !== 0) throw new DownloadError('envelope', `The list of add-ons and ${p.id}'s own envelope disagree.`);
    const folderIn = path.join(parts, p.id);
    const todo = [];
    for (const [rel, f] of Object.entries(m.files || {})) {
      const at = path.join(folderIn, ...rel.split('/'));
      if (fs.existsSync(at) && fs.statSync(at).size === f.size) continue;
      const here = dir ? [pathsOf(dir, p.id).folder, path.join(pathsOf(dir, p.id).previous, p.id)].map((b) => path.join(b, ...rel.split('/'))).find((x) => {
        try { return fs.statSync(x).size === f.size && sha512Of(x) === f.sha512; } catch { return false; }
      }) : null;
      if (here) {
        fs.mkdirSync(path.dirname(at), { recursive: true });
        fs.copyFileSync(here, at);
        continue;
      }
      const listed = p.files?.[rel];
      if (!listed || listed.sha512 !== f.sha512 || listed.size !== f.size) throw new DownloadError('envelope', `The list of add-ons does not give ${p.id}'s file ${rel}.`);
      todo.push({ rel, ...listed, to: at });
    }
    job.total = todo.reduce((n, t) => n + t.size, 0);
    job.done = 0;
    const free = freeSpace(root);
    if (free !== null && free < job.total + SPARE_BYTES) {
      throw new DownloadError('space', `There is not room for ${p.name || p.id} on this disk: it needs ${Math.ceil((job.total + SPARE_BYTES) / 1048576)} MB free.`);
    }
    for (const t of todo) {
      await fetchToFile({ url: t.url, size: t.size, sha512: t.sha512, to: t.to, fetchImpl, signal, source, onBytes: (n) => { job.done += n; progress(p.id); } });
    }
    const disk = checkFolder(folderIn, m, { hashAll: true });
    if (disk) throw new DownloadError(disk.code, disk.reason);
    // Whole and checked: out of the parts, into the version's staging, envelope last.
    fs.rmSync(path.join(root, p.id), { recursive: true, force: true });
    fs.renameSync(folderIn, path.join(root, p.id));
    fs.renameSync(envFile, path.join(root, `${p.id}.pack.json`));
    fs.rmSync(parts, { recursive: true, force: true });
    return root;
  }

  /** Brings packs of the running version in: staged, and put in place where this process may write. */
  async function bring(ids, feed) {
    const root = stagingRoot(office);
    const engineEntry = feed.packs.find((p) => p.kind === 'engine' && ids.includes(p.id));
    const order = [...ids].sort((a, b) => (a === engineEntry?.id ? -1 : b === engineEntry?.id ? 1 : 0));
    let engine = registry?.engine()?.manifest || null;
    for (const id of order) {
      const p = feed.packs.find((x) => x.id === id);
      if (!p || !isAddonId(id)) continue;
      // A download stopped a moment ago finishes stopping before this one takes its files.
      await jobs.get(id)?.running?.catch(() => {});
      const job = { state: 'adding', done: 0, total: p.size || 0, controller: new AbortController(), running: null };
      jobs.set(id, job);
      progress(id);
      job.running = stageOne(p, { version: office, root, signal: job.controller.signal, engine, job });
      try {
        await job.running;
        if (p.kind === 'engine') engine = JSON.parse(fs.readFileSync(path.join(root, `${id}.pack.json`), 'utf8')).manifest;
        if (writable) {
          putInPlace(dir, root, id);
          if (jobs.get(id) === job) jobs.delete(id);
        } else if (jobs.get(id) === job) {
          jobs.set(id, { state: 'staged', done: p.size || 0, total: p.size || 0 });
        }
        restart = true;
        unmarkOff(id);
        takeBackRemoval(id);
      } catch (err) {
        // Only this job's own record: a newer one for the same add-on is left alone.
        if (jobs.get(id) === job) {
          if (job.state === 'paused') job.running = null;
          else if (job.controller.signal.aborted) jobs.delete(id);
          else jobs.set(id, { state: 'failed', done: job.done, total: job.total, error: err?.message || String(err) });
        }
      }
      progress(id);
    }
    changed();
  }
  const unmarkOff = (id) => stores.settings.set(OFF_KEY, off().filter((x) => x !== id));
  /** A removal asked for and not yet done, taken back: the add-on stays. */
  const takeBackRemoval = (id) => {
    if (dir && writable && readState(dir).removals.includes(id)) unmarkRemoval(dir, id);
    const atRestart = stores.settings.get('addons.removeAtRestart', []) || [];
    if (atRestart.includes(id)) stores.settings.set('addons.removeAtRestart', atRestart.filter((x) => x !== id));
  };

  async function add({ id } = {}) {
    if (!dir) throw new Error('This copy of Rutba Office has no add-ons folder.');
    if (!isAddonId(id)) throw new Error(`${id} is not an add-on.`);
    const got = await feedFor(office, { refresh: !feeds.has(office) });
    if (!got.feed) throw new Error(got.why);
    if (!got.feed.packs.some((p) => p.id === id)) throw new Error(`${id} is not offered for Rutba Office ${office}.`);
    const mine = new Map(installedSet({ withRemovals: true }).map((p) => [p.id, p]));
    const wanted = closure([id], got.feed.packs);
    // Removed and then added again before a restart: the removal is taken back.
    for (const x of wanted) if (mine.has(x)) takeBackRemoval(x);
    // The add-on and what it needs that is not here at its revision: one click, and the page says what came with it.
    const ids = wanted.filter((x) => {
      const p = got.feed.packs.find((q) => q.id === x);
      return !p.withdrawn && (!mine.has(x) || compareVersions(p, mine.get(x)) > 0) && !(jobs.get(x)?.state === 'adding');
    });
    // A language brings its font when this computer has none for it (L8): looked for after the language is on its way.
    const languageIds = ids.filter((x) => got.feed.packs.find((q) => q.id === x)?.kind === 'language');
    bring(ids, got.feed)
      .then(() => (languageIds.length ? fontsFor(languageIds, got.feed) : []))
      .then((fonts) => (fonts.length ? bring(fonts, got.feed) : null))
      .catch(() => {});
    return { ...list(), adding: ids };
  }

  /**
   * The font packs some languages need on this computer and the feed offers
   * (L8), less any here, staged or on their way.
   */
  async function fontsFor(languageIds, feed) {
    if (!languages?.fontsFor) return [];
    const here = [...installedSet().map((p) => p.id), ...stagedIds(), ...[...jobs.entries()].filter(([, j]) => j.state === 'adding' || j.state === 'staged').map(([id]) => id)];
    const offered = feed.packs.filter((p) => p.kind === 'font' && !p.withdrawn).map((p) => p.id);
    try {
      return (await languages.fontsFor({ tags: languageIds, installed: here, offered })) || [];
    } catch {
      return [];
    }
  }

  /**
   * At the first start of a version: the fonts the languages already here
   * need (L8). Once per version; nothing when the feed cannot be had.
   */
  async function fontsCheck() {
    if (!dir || !languages?.fontsFor) return [];
    if (stores.settings.get(FONTS_KEY, null) === office) return [];
    const got = await feedFor(office, { refresh: !feeds.has(office) });
    if (!got.feed) return [];
    const tags = [...new Set([...installedSet().filter((p) => p.kind === 'language').map((p) => p.id), ...stagedIds().filter((id) => id.startsWith('lang-'))])];
    const fonts = tags.length ? await fontsFor(tags, got.feed) : [];
    stores.settings.set(FONTS_KEY, office);
    if (fonts.length) bring(fonts, got.feed).catch(() => {});
    return fonts;
  }

  /**
   * An add-on from a file (L5), for a computer with no network: a .zip as the
   * packages releases publish one beside the feed, holding <id>.pack.json and
   * the <id> folder and nothing else. Checked as a fetched add-on is, with
   * this Office's rules: signed by the packages key, for this Office, every
   * file as the envelope says. Then staged, and put in place where this
   * process may write, as Add does; used from the next start.
   *
   * Nothing is written until the envelope, read from the file in memory, is
   * usable here; then only the files its signed manifest lists are written,
   * each under a name it allows. A version older than the one here is
   * refused, since a file could otherwise put back a revision withdrawn for a fault.
   * @param {{ file: string }} o
   */
  async function addFromFile({ file } = {}) {
    if (!dir) throw new Error('This copy of Rutba Office has no add-ons folder.');
    let entries;
    try {
      ({ entries } = readZip(fs.readFileSync(String(file || ''))));
    } catch {
      throw new Error('This file is not an add-on: open a .zip published with the Rutba Office add-ons.');
    }
    const files = entries.filter((e) => !e.name.replace(/\\/g, '/').endsWith('/'));
    const names = files.map((e) => e.name.replace(/\\/g, '/'));
    const envelopes = names.filter((n) => /^[a-z][a-z0-9-]{0,31}\.pack\.json$/.test(n));
    if (envelopes.length !== 1) throw new Error('This file is not an add-on: it should hold one add-on and its folder.');
    const id = envelopes[0].slice(0, -'.pack.json'.length);
    // The envelope from memory, and checked, before anything is written.
    let envelope;
    try {
      envelope = JSON.parse(Buffer.from(files[names.indexOf(envelopes[0])].data).toString('utf8'));
    } catch {
      throw new Error(`${id}.pack.json in this file cannot be read.`);
    }
    const u = usable(envelope, { contract, office, keys: contract.keys, engine: registry?.engine?.()?.manifest || null });
    if (!u.usable && u.code !== 'requires') throw new Error(u.reason);
    const m = envelope.manifest;
    if (m.id !== id) throw new Error(`${id}.pack.json in this file is for ${m.id}.`);
    if (m.kind === 'engine' && m.platform !== platform) throw new Error(`${m.name || id} is for ${m.platform}.`);
    const have = installedSet({ withRemovals: true }).find((p) => p.id === id);
    if (have && compareVersions(m, have) < 0) throw new Error(`${m.name || id} here is newer than the one in this file, so it was kept.`);
    // Only what the signed manifest lists, once each, under its own folder.
    const listed = new Set(Object.keys(m.files || {}).map((rel) => `${id}/${rel}`));
    const seen = new Set();
    for (const n of names) {
      if (n === envelopes[0] && !seen.has(n)) { seen.add(n); continue; }
      if (!listed.has(n) || seen.has(n)) throw new Error(`This file holds something an add-on may not: ${n}.`);
      seen.add(n);
    }
    const root = stagingRoot(office);
    const parts = path.join(root, PARTS, `from-file-${id}`);
    fs.rmSync(parts, { recursive: true, force: true });
    try {
      for (const [i, n] of names.entries()) {
        if (n === envelopes[0]) continue;
        const to = path.join(parts, ...n.split('/'));
        fs.mkdirSync(path.dirname(to), { recursive: true });
        fs.writeFileSync(to, files[i].data);
      }
      fs.writeFileSync(path.join(parts, envelopes[0]), files[names.indexOf(envelopes[0])].data);
      const disk = checkFolder(path.join(parts, id), m, { hashAll: true });
      if (disk) throw new Error(disk.reason);
      fs.rmSync(path.join(root, id), { recursive: true, force: true });
      fs.renameSync(path.join(parts, id), path.join(root, id));
      fs.renameSync(path.join(parts, envelopes[0]), path.join(root, envelopes[0]));
    } finally {
      fs.rmSync(parts, { recursive: true, force: true });
    }
    if (writable) putInPlace(dir, root, id);
    else jobs.set(id, { state: 'staged', done: 0, total: 0 });
    restart = true;
    unmarkOff(id);
    takeBackRemoval(id);
    changed();
    return { ...list(), added: id };
  }

  function pause({ id } = {}) {
    const j = jobs.get(id);
    if (j?.state === 'adding') {
      j.state = 'paused';
      j.controller?.abort();
      progress(id);
    }
    return list();
  }

  /** Stops an addition (its partial files go) or takes back a removal not yet done. */
  function cancel({ id } = {}) {
    const j = jobs.get(id);
    if (j) {
      j.controller?.abort();
      jobs.delete(id);
      const root = stagingRoot(office);
      const clear = () => {
        fs.rmSync(path.join(root, PARTS, id), { recursive: true, force: true });
        fs.rmSync(path.join(root, id), { recursive: true, force: true });
        fs.rmSync(path.join(root, `${id}.pack.json`), { force: true });
      };
      clear();
      // And again once the download has stopped, in case it wrote after the first.
      j.running?.catch(() => {}).finally(clear);
      progress(id);
    }
    takeBackRemoval(id);
    changed();
    return list();
  }

  /** Gone at the next start, with what depends on it named. The person's documents are never touched. */
  function remove({ id } = {}) {
    if (!ID.test(String(id))) throw new Error(`${id} is not an add-on.`);
    if (NEVER_HOLDS.includes(id)) throw new Error('English is part of Rutba Office.');
    // The language a window is in, or the menus are, is in use until those windows close (L6).
    if ((safe(() => languages?.inUse()) || []).includes(id)) throw new Error('A window is open in this language. Close it, or open the windows in another language, before you remove it.');
    const dependents = (registry?.list() || []).filter((a) => (a.manifest.requires || []).some((r) => r === id || (registry.find(r)?.id === id))).map((a) => a.id);
    if (writable) markRemoval(dir, id);
    else stores.settings.set('addons.removeAtRestart', [...new Set([...(stores.settings.get('addons.removeAtRestart', []) || []), id])]);
    restart = true;
    changed();
    return { ...list(), dependents };
  }

  /**
   * The running version's feed read again: new revisions taken (per person
   * and portable, with updates on), withdrawn ones rolled back or turned off.
   * @returns {Promise<{ updated: string[], rolledBack: string[], turnedOff: string[], waiting: string[] }>}
   */
  async function check({ manual = false } = {}) {
    const out = { updated: [], rolledBack: [], turnedOff: [], waiting: [] };
    if (!dir || !installedSet().length) return out;
    const got = await feedFor(office);
    if (!got.feed) return out;
    const mine = installedSet();
    for (const p of got.feed.packs) {
      const have = mine.find((x) => x.id === p.id);
      if (!have) continue;
      if (p.withdrawn && compareVersions(p, have) === 0) {
        if (writable && rollBack(dir, p.id)) out.rolledBack.push(p.id);
        else { stores.settings.set(OFF_KEY, [...new Set([...off(), p.id])]); out.turnedOff.push(p.id); }
        restart = true;
        continue;
      }
      if (!p.withdrawn && p.version === office && compareVersions(p, have) > 0) {
        if (perMachine || !(manual || (updatesOn() && stores.settings.get(AUTO_KEY, true) !== false))) out.waiting.push(p.id);
        else out.updated.push(p.id);
      }
    }
    if (out.updated.length) await bring(out.updated, got.feed);
    if (out.rolledBack.length || out.turnedOff.length) changed();
    return out;
  }

  /**
   * Whether an update to `target` may be offered (C11): no add-ons, at once
   * and with no request; otherwise target's feed, verified and taken for
   * target, through the one rule (addons-compat readiness).
   */
  async function readiness(target) {
    const installed = gateSet();
    if (!installed.filter((p) => !NEVER_HOLDS.includes(p.id)).length) return decide({ installed: [], target });
    const got = await feedFor(target);
    const r = decide({ installed, feed: got.feed, target, contract });
    const name = (id) => got.feed?.packs?.find((p) => p.id === id)?.name || (registry?.list() || []).find((a) => a.id === id)?.manifest?.name || id;
    const kind = (id) => got.feed?.packs?.find((p) => p.id === id)?.kind || installed.find((p) => p.id === id)?.kind || null;
    return {
      ...r,
      waitingFor: r.waiting.map((w) => ({ ...w, name: name(w.id), kind: kind(w.id) })),
      retired: r.retired.map((x) => ({ ...x, name: name(x.id), kind: kind(x.id) })),
      feedError: got.feed ? null : got.why,
    };
  }

  /**
   * X's packs staged before "Restart and update" (C12): for each installed
   * add-on X's feed lists, its envelope and the files that changed, into
   * X's staging folder; X's first start checks and moves them in. A pack
   * that cannot be staged now is fetched by X itself.
   */
  async function stageFor(target) {
    const got = await feedFor(target, { refresh: !feeds.has(target) });
    const staged = [];
    const failed = [];
    if (!got.feed || !dir) return { staged, failed };
    const root = stagingRoot(target);
    const ids = gateSet().map((p) => p.id).filter((id) => got.feed.packs.some((p) => p.id === id && p.version === target && !p.withdrawn));
    const engineId = got.feed.packs.find((p) => p.kind === 'engine' && ids.includes(p.id))?.id;
    let engine = null;
    for (const id of [...ids].sort((a, b) => (a === engineId ? -1 : b === engineId ? 1 : 0))) {
      const p = got.feed.packs.find((x) => x.id === id);
      try {
        await stageOne(p, { version: target, root, signal: null, engine });
        if (id === engineId) engine = JSON.parse(fs.readFileSync(path.join(root, `${id}.pack.json`), 'utf8')).manifest;
        staged.push(id);
      } catch (err) {
        failed.push({ id, reason: err?.message || String(err) });
      }
    }
    return { staged, failed, root };
  }

  /**
   * Started again to finish, as the update service restarts: not while a
   * window has unsaved work, and through app.quit(), so every window's own
   * question and the quit steps (the Locker's seal, the sweep) run. An
   * install for everyone asks for the administrator first: the step it runs
   * waits for Office to quit before it moves anything in or does the
   * removals, then starts Office again as the person. Said no to, nothing
   * is forgotten: the removals and the staged add-ons wait for the next try.
   */
  async function restartNow() {
    const stopped = whatStopsRestart({ ready: true, windows: safe(() => openWindows?.()) || [] });
    if (stopped) return { restarting: false, ...stopped };
    if (perMachine && elevate) {
      const removals = stores.settings.get('addons.removeAtRestart', []) || [];
      // The person's staging root: the run takes this version's folder inside it.
      const args = [`--addons-apply-staged=${personStaging || ''}`, `--addons-after=${pid}`, '--addons-relaunch'];
      if (removals.length) args.push(`--addons-remove=${removals.join(',')}`);
      const r = await elevate(args, { wait: false });
      if (!r?.ok) return { restarting: false, declined: true, reason: 'Windows did not give permission to change the add-ons of a copy installed for everyone. Nothing was changed; try again, or ask whoever looks after this computer.' };
      stores.settings.set('addons.removeAtRestart', []);
      // The administrator's step starts Office again once it has done its work.
      app?.quit?.();
      return { restarting: true, elevated: true };
    }
    app?.relaunch?.();
    app?.quit?.();
    return { restarting: true };
  }

  /**
   * An install for everyone, at the first start after its staged add-ons
   * could not be moved in by the installer: the administrator is asked once,
   * the add-ons moved in for the next start, nothing restarted.
   */
  async function applyNow() {
    if (!perMachine || !elevate || !stagedIds().length) return { asked: false };
    const r = await elevate([`--addons-apply-staged=${personStaging || ''}`]);
    restart = true;
    changed();
    return { asked: true, ok: Boolean(r?.ok) };
  }

  /** After a start that skipped add-ons (two starts in a row did not reach the first window): try them again. */
  function tryAgain() {
    start?.tryAgain?.();
    return restartNow();
  }

  return {
    list: () => list(),
    tryAgain,
    applyNow,
    feed: async ({ refresh = true } = {}) => {
      await feedFor(office, { refresh });
      return list();
    },
    add,
    addFromFile,
    fontsCheck,
    pause,
    cancel,
    remove,
    check: (o) => check(o),
    restart: restartNow,
    // For the update service, never a window.
    readiness,
    stageFor,
    installed: () => installedSet(),
    off,
  };
}
