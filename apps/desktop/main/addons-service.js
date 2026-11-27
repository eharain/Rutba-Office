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
// add-on is used from the next start, as an app pack is. Every copy, an
// install for everyone too, keeps its add-ons in the person's own folder in
// 2.26.0 (addons-folder.js, D-2026-11-13-35), so nothing here asks for the
// administrator.
//
// With Office's updates on, a newer revision of an installed add-on is taken
// by itself. A revision withdrawn by its makers is rolled back to
// the one kept; with none kept the add-on is turned off. Nothing new is ever
// added without a click.

import path from 'node:path';
import { createRequire } from 'node:module';
import { readiness as decide, verifyFeed, acceptFeed, usable, signedForKind, compareVersions, isAddonId, NEVER_HOLDS } from '@rutba/office-shell/electron/addons-compat';
import { envelopesIn, envelopeAt, checkFolder, sha512Of } from '@rutba/office-shell/electron/addons';
import { readState, writeState, markRemoval, unmarkRemoval, putInPlace, rollBack, stagingFor, pathsOf, whileLocked, linkFree, removeTree, PARTS } from './addons-folder.js';
import { fetchToFile, releaseUrl, fromSource, allowedUrl, freeSpace, DownloadError } from './addons-download.js';
import { readZip } from '@rutba/ooxml/zip';
import { whatStopsRestart } from './update-error.js';
import { listedLanguage } from '@rutba/office-shell/electron/language-packs';

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
/** Said when another Office holds the folder (voices-dialog.js FOLDER_WORDS). */
export const FOLDER_BUSY = 'Another Rutba Office is changing the voices and languages just now. Try again in a moment.';
/** Removals taken back while another Office held the folder: the next start takes their marks back before it removes anything (addons-start.js). */
export const KEEP_KEY = 'addons.keepAtRestart';

/** The version whose first start looked for the fonts its languages need (L8). */
const FONTS_KEY = 'addons.fontsChecked';

/** Add-ons set going and not here yet: brought again at the next start if Office closed first (2.26.1). */
const WANTED_KEY = 'addons.onTheWay';
/** Add-ons the person removed by hand: never brought back as something another add-on needs. */
const BY_HAND_KEY = 'addons.removedByHand';

/**
 * Every pack a set of ids needs, through `requires`, in the feed: the ids, then what they need.
 * A need given as a name several packs provide (a speech model) is met by one already here (`have`),
 * else by the first the feed lists: 2.26.0 named the first always, so each voice added brought the
 * large model again over a small one the person had chosen.
 */
export function closure(ids, packs, { have = [] } = {}) {
  const provides = (p, name) => (p.provides || []).includes(name);
  const byName = (name) => packs.find((p) => p.id === name) || packs.find((p) => provides(p, name) && have.includes(p.id)) || packs.find((p) => provides(p, name));
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
 * @param {() => boolean} [o.updatesOn]   Office's own updates are on
 * @param {object} [o.app]                { relaunch, quit } for Restart
 * @param {() => { dirty?: boolean, name?: string }[]} [o.openWindows]  the open windows, for Restart's question on unsaved work
 * @param {string} [o.platform]
 * @param {() => ({ engine: string|null, problem: string|null } | null)} [o.engineState]  the speech engine this run, once loaded (engine-host.js)
 * @param {object} [o.languages]  the language packs (L4, L6, L7, L8): known, coverageOf, inUse, builtInInUse() -> [lang-<tag>], and fontsFor({ tags, installed, offered }) -> font pack ids
 */
export function createAddonsService({
  folder, start, contract, office, stores, broadcast = () => {}, fetchImpl = (...a) => fetch(...a), source = process.env.RUTBA_ADDONS_SOURCE || null,
  bundled = null, updatesOn = () => true, app = null, platform = `${process.platform}-${process.arch}`,
  engineState = () => null,
  languages = null,
  openWindows = null,
}) {
  const dir = folder?.dir || null;
  const writable = Boolean(start?.writable);
  const registry = start?.registry;
  /** Where this run stages: the folder's own, for this version. */
  const stagingRoot = (version) => stagingFor(dir, version);
  const jobs = new Map(); // id -> { state, done, total, controller, error }
  const feeds = new Map(); // version -> { feed, from }
  // version -> the languages its own feed (latest.yml) says it leaves out of the core, as readiness was told: staged by the same set.
  const releasedLeft = new Map();
  let restart = false;
  // Set going by a click and not yet reached by the run that brings them: counted as on their way, so the next
  // click does not bring them a second time, and no restart is offered over them.
  const queued = new Set();
  const onTheWay = () => queued.size > 0 || [...jobs.values()].some((j) => j.state === 'adding');
  /**
   * A restart is offered once everything set going has arrived. 2.26.0 offered it when the first small
   * add-on landed, with the speech model (311 MB) still coming: restarting then cut the model off, and the
   * voices were added and never spoke.
   */
  const restartOffered = () => restart && !onTheWay();
  const idList = (key) => (stores.settings.get(key, []) || []).filter((id) => typeof id === 'string' && ID.test(id));
  const want = (ids) => stores.settings.set(WANTED_KEY, [...new Set([...idList(WANTED_KEY), ...ids])]);
  const unwant = (id) => { if (idList(WANTED_KEY).includes(id)) stores.settings.set(WANTED_KEY, idList(WANTED_KEY).filter((x) => x !== id)); };

  /**
   * A change to the add-ons folder, made holding it (addons-folder.js
   * whileLocked) so that another Office's start never moves add-ons in or
   * out half-way through it. When another Office holds it: `otherwise`, or
   * the person is told to try again.
   */
  const held = (fn, otherwise = null) => {
    const r = whileLocked(dir, fn);
    if (r.done) return r.value;
    if (otherwise) return otherwise();
    throw new Error(FOLDER_BUSY);
  };

  const keptBack = () => stores.settings.get(KEEP_KEY, []) || [];
  /** The add-ons asked to go at the next start: the folder's marks but those taken back, and the profile's where the folder cannot be written. */
  const removing = () => {
    if (!dir) return [];
    const back = keptBack();
    return [...readState(dir).removals.filter((x) => !back.includes(x)), ...(stores.settings.get('addons.removeAtRestart', []) || [])];
  };

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
      // Another Office writing the folder this moment: the feed is still taken this run, and remembered at the next.
      held(() => writeState(dir, { ...readState(dir), seen: next, feeds: nextHashes }), () => null);
    } else {
      stores.settings.set(SEEN_KEY, next);
      stores.settings.set(SEEN_FEEDS_KEY, nextHashes);
    }
  };
  const off = () => (stores.settings.get(OFF_KEY, []) || []).filter((id) => ID.test(id));

  const changed = () => broadcast('addons:changed', { restart: restartOffered() });
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
    const going = removing();
    const signed = (e) => e.envelope?.manifest?.id === e.id && signedForKind(e.envelope, { contract });
    return envelopesIn(dir).filter((e) => (withRemovals || !going.includes(e.id)) && signed(e)).map(({ id, envelope }) => ({
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
   * The add-ons here, staged, or on their way (being fetched or waiting their
   * turn), by id, as list() names them, with no language's words counted: a
   * first count reads every language pack through on the main process.
   */
  function presentIds() {
    return [...new Set([
      ...installedSet().map((p) => p.id),
      ...stagedIds(),
      ...[...jobs].filter(([, j]) => j.state === 'adding' || j.state === 'staged').map(([id]) => id),
      ...queued,
    ])];
  }

  /**
   * What the update gate counts as here: the add-ons installed, and the
   * built-in language the person uses as its lang-<tag> (L7), so an update
   * waits for that pack and brings it, for the release that leaves it out
   * of the core.
   */
  function gateSet(feed = null, released = null) {
    const here = installedSet();
    // The built-in language the windows are in counts as its pack only for a
    // version that no longer carries it built in: as that version's packages
    // feed says (leftCore), or as the update's own feed says (latest.yml's
    // leftCore, build/write-left-core.js), which comes even when the packages
    // feed cannot be had (D-2026-11-13-36). With neither saying so, it counts
    // for nothing: it never makes an update wait for a feed while the version
    // still carries the language (L1).
    const tags = (v) => (Array.isArray(v) ? v.filter((t) => typeof t === 'string' && /^[a-z]{2,3}$/.test(t)) : []);
    const left = [...tags(feed?.leftCore), ...tags(released)];
    const builtIn = (safe(() => languages?.builtInInUse?.()) || []).filter((id) => ID.test(id) && left.includes(id.slice('lang-'.length)) && !here.some((p) => p.id === id));
    return [...here, ...builtIn.map((id) => ({ id, kind: 'language', version: office, revision: 0, builtIn: true }))];
  }
  /**
   * Whether anything here could hold an update: an add-on, or a built-in
   * language in use, which a version may leave out of the core. With
   * neither (English, nothing added), no update asks for a feed.
   */
  const holdsAny = () => installedSet().some((p) => !NEVER_HOLDS.includes(p.id)) || (safe(() => languages?.builtInInUse?.()) || []).some((id) => ID.test(id));

  /** A language added here (L6): its tag and review, and how much of this Office's words it has, counted against this Office's own keys. */
  function installedLanguage(a, id, { coverage = true } = {}) {
    const m = a?.manifest || (dir ? envelopeAt(pathsOf(dir, id).envelope)?.manifest : null) || {};
    const counted = coverage ? safe(() => languages?.coverageOf(id)) || null : null;
    return {
      tag: m.tag || null, nativeName: m.nativeName || null, direction: m.direction || null, variantOf: m.variantOf || null, review: m.review || null,
      licence: m.licence?.spdx || null, coverage: counted, inUse: (safe(() => languages?.inUse()) || []).includes(id),
    };
  }

  /** A language the feed offers (L6): what it says of itself before it is added; its coverage is of that version's English. */
  const offeredLanguage = (p) => ({ tag: p.tag || null, script: p.script || null, direction: p.direction || null, variantOf: p.variantOf || null, review: p.review || null, coverage: p.coverage || null, packStatus: p.status || null });

  const languagesKnown = (packs) => safe(() => languages?.known(packs)) || [];

  /**
   * Everything the add-ons page shows. `coverage: false` leaves each added
   * language's count of words out (null), for the start's own looks, which
   * need only what is here and offered: the count is made when a window asks.
   */
  function list({ coverage = true } = {}) {
    const known = feeds.get(office)?.feed || null;
    const going = removing();
    const used = new Map((registry?.list() || []).map((a) => [a.id, a]));
    const refused = new Map((registry?.refused || []).map((r) => [r.id, r]));
    const installed = installedSet({ withRemovals: true }).map((p) => {
      const a = used.get(p.id);
      const r = refused.get(p.id);
      const state = going.includes(p.id) ? 'removing'
        : off().includes(p.id) ? 'off'
          : a ? 'loaded'
            : r ? (r.code === 'requires' ? 'waiting' : 'refused')
              : (registry?.deferred || []).includes(p.id) ? 'deferred' : 'added';
      return { ...p, name: a?.manifest?.name || p.id, state, reason: r?.reason || null, code: r?.code || null, ...(p.kind === 'language' ? installedLanguage(a, p.id, { coverage }) : {}) };
    });
    const mine = new Map(installed.map((p) => [p.id, p]));
    const offered = (known?.packs || []).filter((p) => isAddonId(p?.id)).map((p) => {
      const have = mine.get(p.id);
      const job = jobs.get(p.id);
      const status = job ? job.state
        : queued.has(p.id) ? 'adding'
          : !have ? 'add'
          : compareVersions(p, have) > 0 ? 'update' : 'installed';
      return {
        id: p.id, kind: p.kind, name: p.name, nativeName: p.nativeName || null, version: p.version, revision: p.revision, size: p.size || 0,
        licence: p.licence?.spdx || null, requires: p.requires || [], provides: p.provides || [], withdrawn: Boolean(p.withdrawn), tags: p.tags || [], status,
        ...(p.kind === 'language' ? offeredLanguage(p) : {}),
      };
    });
    return {
      office, dir, scope: folder?.scope || null, writable, restart: restartOffered(),
      skipped: start?.skipped || null, said: start?.said || null,
      // The speech engine this run: loaded, or why not.
      engine: engineState(),
      automatic: stores.settings.get(AUTO_KEY, true) !== false,
      feed: known ? { office: known.office, sequence: known.sequence, state: known.state, from: feeds.get(office).from } : null,
      installed, offered,
      // Every language there is (L4): the built-in list with the feed's entries over it, for the languages not offered yet.
      languages: languagesKnown(known?.packs || []),
      // Staged and waiting to be moved in at the next start.
      staged: stagedIds(),
      jobs: [
        ...[...jobs].map(([id, j]) => ({ id, state: j.state, done: j.done, total: j.total, error: j.error || null })),
        // Waiting their turn behind the one being fetched.
        ...[...queued].filter((id) => !jobs.has(id)).map((id) => ({ id, state: 'adding', done: 0, total: known?.packs?.find((p) => p.id === id)?.size || 0, error: null })),
      ],
    };
  }

  /**
   * One add-on brought into a staging folder from a feed entry: the files it
   * lacks fetched, the rest copied from the add-on in place, the whole
   * checked with this Office's rules, the envelope last.
   * @returns {Promise<string>}  the staging folder holding <id>/ and <id>.pack.json
   */
  async function stageOne(p, { version, root, signal, engine = null, job = { state: 'adding', done: 0, total: 0 }, report = () => progress(p.id) }) {
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
      await fetchToFile({ url: t.url, size: t.size, sha512: t.sha512, to: t.to, fetchImpl, signal, source, onBytes: (n) => { job.done += n; report(); } });
    }
    const disk = checkFolder(folderIn, m, { hashAll: true });
    if (disk) throw new DownloadError(disk.code, disk.reason);
    // Whole and checked: out of the parts, into the version's staging, envelope last.
    removeTree(path.join(root, p.id));
    fs.renameSync(folderIn, path.join(root, p.id));
    fs.renameSync(envFile, path.join(root, `${p.id}.pack.json`));
    removeTree(parts);
    return root;
  }

  // Cancelled before the run that brings them had reached them.
  const dropped = new Set();
  /** Set going: marked as on their way before the first is fetched, and remembered until each is here. */
  const setGoing = (ids, feed) => {
    for (const x of ids) { queued.add(x); dropped.delete(x); }
    want(ids);
    changed();
    return bring(ids, feed).finally(() => {
      for (const x of ids) queued.delete(x);
      changed();
    });
  };

  /** Brings packs of the running version in: staged, and put in place where this process may write. */
  async function bring(ids, feed) {
    const root = stagingRoot(office);
    const engineEntry = feed.packs.find((p) => p.kind === 'engine' && ids.includes(p.id));
    const order = [...ids].sort((a, b) => (a === engineEntry?.id ? -1 : b === engineEntry?.id ? 1 : 0));
    let engine = registry?.engine()?.manifest || null;
    for (const id of order) {
      const p = feed.packs.find((x) => x.id === id);
      if (!p || !isAddonId(id)) { queued.delete(id); continue; }
      // Cancelled while it waited its turn.
      if (dropped.delete(id)) { queued.delete(id); continue; }
      // A download stopped a moment ago finishes stopping before this one takes its files.
      await jobs.get(id)?.running?.catch(() => {});
      queued.delete(id);
      const job = { state: 'adding', done: 0, total: p.size || 0, controller: new AbortController(), running: null };
      jobs.set(id, job);
      progress(id);
      job.running = stageOne(p, { version: office, root, signal: job.controller.signal, engine, job });
      try {
        await job.running;
        if (p.kind === 'engine') engine = JSON.parse(fs.readFileSync(path.join(root, `${id}.pack.json`), 'utf8')).manifest;
        // Another Office holding the folder: it stays staged, and the next start moves it in.
        if (writable && held(() => (putInPlace(dir, root, id), true), () => false)) {
          if (jobs.get(id) === job) jobs.delete(id);
        } else if (jobs.get(id) === job) {
          jobs.set(id, { state: 'staged', done: p.size || 0, total: p.size || 0 });
        }
        restart = true;
        unmarkOff(id);
        takeBackRemoval(id);
        unwant(id);
      } catch (err) {
        // Only this job's own record: a newer one for the same add-on is left alone.
        if (jobs.get(id) === job) {
          if (job.state === 'paused') job.running = null;
          else if (job.controller.signal.aborted) jobs.delete(id);
          else jobs.set(id, { state: 'failed', done: job.done, total: job.total, error: err?.message || String(err) });
        }
        // Stopped by the person: not brought again at the next start. One that failed is tried again then.
        if (job.state === 'paused' || job.controller.signal.aborted) unwant(id);
      }
      progress(id);
    }
    changed();
  }
  const unmarkOff = (id) => stores.settings.set(OFF_KEY, off().filter((x) => x !== id));
  /** A removal asked for and not yet done, taken back: the add-on stays. */
  const takeBackRemoval = (id) => {
    // Another Office holding the folder (the add-on may be in place already): the profile keeps it, and the next start takes the mark back.
    if (dir && writable && readState(dir).removals.includes(id)) held(() => unmarkRemoval(dir, id), () => stores.settings.set(KEEP_KEY, [...new Set([...keptBack(), id])]));
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
    // What is here or on its way meets a need before anything new does.
    const have = [...mine.keys(), ...stagedIds(), ...queued, ...[...jobs].filter(([, j]) => j.state === 'adding').map(([x]) => x)];
    const wanted = closure([id], got.feed.packs, { have });
    // Removed and then added again before a restart: the removal is taken back.
    for (const x of wanted) if (mine.has(x)) takeBackRemoval(x);
    // Asked for by name: no longer one the person took away.
    if (idList(BY_HAND_KEY).includes(id)) stores.settings.set(BY_HAND_KEY, idList(BY_HAND_KEY).filter((x) => x !== id));
    // The add-on and what it needs that is not here at its revision: one click, and the page says what came with it.
    // One staged at that revision is here too, waiting for the next start. On a first use the folder is not there
    // when Office starts, so everything added that run is staged: 2.26.0 counted none of it, and each further
    // voice fetched the engine and the speech model (311 MB) again.
    const stagedAt = (x) => {
      try { return JSON.parse(fs.readFileSync(path.join(stagingRoot(office), `${x}.pack.json`), 'utf8')).manifest; } catch { return null; }
    };
    const ids = wanted.filter((x) => {
      const p = got.feed.packs.find((q) => q.id === x);
      const s = stagedAt(x);
      return !p.withdrawn && (!mine.has(x) || compareVersions(p, mine.get(x)) > 0) && !(s && compareVersions(p, s) <= 0) && !(jobs.get(x)?.state === 'adding') && !queued.has(x);
    });
    // A language brings its font when this computer has none for it (L8): looked for after the language is on its way.
    const languageIds = ids.filter((x) => got.feed.packs.find((q) => q.id === x)?.kind === 'language');
    setGoing(ids, got.feed)
      .then(() => (languageIds.length ? fontsFor(languageIds, got.feed) : []))
      .then((fonts) => (fonts?.length ? setGoing(fonts, got.feed) : null))
      .catch(() => {});
    return { ...list(), adding: ids };
  }

  /**
   * The font packs some languages need on this computer and the feed offers
   * (L8), less any here, staged or on their way; null when this computer's
   * fonts could not be looked through, so that the look is not counted as done.
   */
  async function fontsFor(languageIds, feed) {
    if (!languages?.fontsFor) return [];
    const here = [...installedSet().map((p) => p.id), ...stagedIds(), ...[...jobs.entries()].filter(([, j]) => j.state === 'adding' || j.state === 'staged').map(([id]) => id)];
    const offered = feed.packs.filter((p) => p.kind === 'font' && !p.withdrawn).map((p) => p.id);
    try {
      return (await languages.fontsFor({ tags: languageIds, installed: here, offered })) || [];
    } catch {
      return null;
    }
  }

  /**
   * At a start: what was set going and never arrived (Office closed or was restarted first), and what the
   * add-ons here need and lack, is brought again. 2.26.0 left a voice whose speech model had not finished
   * added and silent, with nothing brought again. Nothing is asked of the network unless something is
   * lacking, and nothing the person removed by hand is brought back.
   * @returns {Promise<string[]>} the add-ons set going
   */
  async function needsCheck() {
    if (!dir) return [];
    const here = new Set([...installedSet({ withRemovals: true }).map((p) => p.id), ...stagedIds()]);
    const interrupted = idList(WANTED_KEY).filter((x) => !here.has(x));
    // Remembered and here already: nothing left to bring.
    for (const x of idList(WANTED_KEY)) if (here.has(x)) unwant(x);
    const lacking = (registry?.refused || []).some((r) => r.code === 'requires');
    if (!interrupted.length && !lacking) return [];
    const got = await feedFor(office, { refresh: !feeds.has(office) });
    if (!got.feed) return [];
    const offered = new Map(got.feed.packs.filter((p) => isAddonId(p?.id) && !p.withdrawn).map((p) => [p.id, p]));
    const kept = installedSet().map((p) => p.id).filter((x) => offered.has(x));
    const needed = closure(kept, got.feed.packs, { have: [...here] }).filter((x) => !here.has(x));
    const byHand = idList(BY_HAND_KEY);
    const ids = [...new Set([...interrupted, ...needed])].filter((x) => offered.has(x) && !byHand.includes(x) && !queued.has(x) && jobs.get(x)?.state !== 'adding');
    if (!ids.length) return [];
    setGoing(ids, got.feed).catch(() => {});
    return ids;
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
    // Looked through only when it could be: a font that could not be read leaves it for the next start.
    if (fonts === null) return [];
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
    removeTree(parts);
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
      removeTree(path.join(root, id));
      fs.renameSync(path.join(parts, id), path.join(root, id));
      fs.renameSync(path.join(parts, envelopes[0]), path.join(root, envelopes[0]));
    } finally {
      removeTree(parts);
    }
    if (!(writable && held(() => (putInPlace(dir, root, id), true), () => false))) jobs.set(id, { state: 'staged', done: 0, total: 0 });
    restart = true;
    unmarkOff(id);
    takeBackRemoval(id);
    changed();
    return { ...list(), added: id };
  }

  /** One waiting its turn is taken off the run that would have brought it. */
  const dropQueued = (id) => {
    if (!queued.has(id) || jobs.has(id)) return false;
    queued.delete(id);
    dropped.add(id);
    unwant(id);
    changed();
    return true;
  };

  function pause({ id } = {}) {
    if (dropQueued(id)) return list();
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
    if (dropQueued(id)) return list();
    unwant(id);
    // A removal taken back: kept after all, so no longer one the person took away.
    if (idList(BY_HAND_KEY).includes(id)) stores.settings.set(BY_HAND_KEY, idList(BY_HAND_KEY).filter((x) => x !== id));
    const j = jobs.get(id);
    if (j) {
      j.controller?.abort();
      jobs.delete(id);
      const root = stagingRoot(office);
      const clear = () => {
        // Never through a link out of the add-ons folder.
        if (!linkFree(dir, root)) return;
        removeTree(path.join(root, PARTS, id));
        removeTree(path.join(root, id));
        removeTree(path.join(root, `${id}.pack.json`));
      };
      // The download just stopped may still hold its part open: what cannot
      // go now goes once it has stopped, and Cancel itself never fails on it.
      try {
        clear();
      } catch {
        // cleared again below
      }
      j.running?.catch(() => {}).finally(() => {
        try {
          clear();
        } catch {
          // left for letGo at a later start
        }
      });
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
    if (writable) held(() => markRemoval(dir, id));
    else stores.settings.set('addons.removeAtRestart', [...new Set([...(stores.settings.get('addons.removeAtRestart', []) || []), id])]);
    // Asked to go again after a take-back the folder has not seen yet.
    if (keptBack().includes(id)) stores.settings.set(KEEP_KEY, keptBack().filter((x) => x !== id));
    // Taken away by the person: never brought back as something another add-on needs (needsCheck).
    unwant(id);
    stores.settings.set(BY_HAND_KEY, [...new Set([...idList(BY_HAND_KEY), id])]);
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
        if (writable && held(() => rollBack(dir, p.id), () => false)) out.rolledBack.push(p.id);
        else { stores.settings.set(OFF_KEY, [...new Set([...off(), p.id])]); out.turnedOff.push(p.id); }
        restart = true;
        continue;
      }
      if (!p.withdrawn && p.version === office && compareVersions(p, have) > 0) {
        if (!(manual || (updatesOn() && stores.settings.get(AUTO_KEY, true) !== false))) out.waiting.push(p.id);
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
   * @param {string} target
   * @param {object} [o]
   * @param {string[]|null} [o.leftCore]  the languages the update's own feed (latest.yml) says it leaves out of the core
   */
  async function readiness(target, { leftCore = null } = {}) {
    releasedLeft.set(target, leftCore);
    if (!holdsAny()) return decide({ installed: [], target });
    const got = await feedFor(target);
    const installed = gateSet(got.feed, leftCore);
    const r = decide({ installed, feed: got.feed, target, contract });
    // A built-in language has no pack to name it, and may have no feed: its name in its own words, from the list.
    const ownName = (id) => (id.startsWith('lang-') ? (({ nativeName, name } = {}) => nativeName || name)(listedLanguage(id.slice('lang-'.length)) || {}) : null);
    const name = (id) => got.feed?.packs?.find((p) => p.id === id)?.name || (registry?.list() || []).find((a) => a.id === id)?.manifest?.name || ownName(id) || id;
    const kind = (id) => got.feed?.packs?.find((p) => p.id === id)?.kind || installed.find((p) => p.id === id)?.kind || null;
    const builtIn = (id) => installed.some((p) => p.id === id && p.builtIn);
    return {
      ...r,
      waitingFor: r.waiting.map((w) => ({ ...w, name: name(w.id), kind: kind(w.id), ...(builtIn(w.id) ? { builtIn: true } : {}) })),
      retired: r.retired.map((x) => ({ ...x, name: name(x.id), kind: kind(x.id) })),
      feedError: got.feed ? null : got.why,
    };
  }

  /**
   * X's packs staged before "Restart and update" (C12): for each installed
   * add-on X's feed lists, its envelope and the files that changed, into
   * X's staging folder; X's first start checks and moves them in. A pack
   * that cannot be staged now is fetched by X itself.
   * @param {string} target
   * @param {object} [o]
   * @param {AbortSignal|null} [o.signal]  stops it (a newer version found): what came is kept, to be resumed
   * @param {(p: { done: number, total: number }) => void} [o.onProgress]  bytes fetched of all the packs, as they come
   * @param {string[]|null} [o.leftCore]  as readiness was given it for target (latest.yml), unless said again here
   */
  async function stageFor(target, { signal = null, onProgress = null, leftCore = releasedLeft.get(target) ?? null } = {}) {
    const staged = [];
    const failed = [];
    if (!dir || !holdsAny()) return { staged, failed };
    const got = await feedFor(target, { refresh: !feeds.has(target) });
    if (!got.feed) return { staged, failed };
    const root = stagingRoot(target);
    // The same set readiness decided by: a built-in language only the release names is staged too.
    const ids = gateSet(got.feed, leftCore).map((p) => p.id).filter((id) => got.feed.packs.some((p) => p.id === id && p.version === target && !p.withdrawn));
    const engineId = got.feed.packs.find((p) => p.kind === 'engine' && ids.includes(p.id))?.id;
    let engine = null;
    const total = ids.reduce((n, id) => n + (got.feed.packs.find((x) => x.id === id)?.size || 0), 0);
    let before = 0;
    const tell = (done) => {
      try {
        onProgress?.({ done: Math.min(done, total), total });
      } catch {
        // the one told has gone
      }
    };
    for (const id of [...ids].sort((a, b) => (a === engineId ? -1 : b === engineId ? 1 : 0))) {
      if (signal?.aborted) return { staged, failed, root, stopped: true };
      const p = got.feed.packs.find((x) => x.id === id);
      const job = { state: 'adding', done: 0, total: 0 };
      try {
        await stageOne(p, { version: target, root, signal, engine, job, report: () => tell(before + job.done) });
        if (id === engineId) engine = JSON.parse(fs.readFileSync(path.join(root, `${id}.pack.json`), 'utf8')).manifest;
        staged.push(id);
      } catch (err) {
        if (signal?.aborted) return { staged, failed, root, stopped: true };
        failed.push({ id, reason: err?.message || String(err) });
      }
      before += p?.size || 0;
      tell(before);
    }
    return { staged, failed, root };
  }

  /**
   * Started again to finish, as the update service restarts: not while a
   * window has unsaved work, and through app.quit(), so every window's own
   * question and the quit steps (the Locker's seal, the sweep) run.
   */
  async function restartNow() {
    const stopped = whatStopsRestart({ ready: true, windows: safe(() => openWindows?.()) || [] });
    if (stopped) return { restarting: false, ...stopped };
    app?.relaunch?.();
    app?.quit?.();
    return { restarting: true };
  }

  /** After a start that skipped add-ons (two starts in a row did not reach the first window): try them again. */
  async function tryAgain() {
    // The count of failed starts is let go of only when the restart goes ahead: unsaved work stops both.
    const stopped = whatStopsRestart({ ready: true, windows: safe(() => openWindows?.()) || [] });
    if (stopped) return { restarting: false, ...stopped };
    start?.tryAgain?.();
    return restartNow();
  }

  return {
    list: ({ coverage = true } = {}) => list({ coverage: coverage !== false }),
    tryAgain,
    feed: async ({ refresh = true, coverage = true } = {}) => {
      await feedFor(office, { refresh });
      return list({ coverage: coverage !== false });
    },
    present: () => presentIds(),
    add,
    addFromFile,
    fontsCheck,
    needsCheck,
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
