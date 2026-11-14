// Words from a language pack (L1, L2): the table a window, and the main
// process, show their words from, made from an installed lang-<tag> add-on.
//
// The code names every label by its English as written in it, the t() or
// tn() argument (catalogues/messages.json lists them all, the parts that keep
// their words apart among them: codeLabels). A pack holds each
// label by its id, with `source`: the first 12 hex digits of the SHA-256 of
// that key (for a counted label, of the JSON text {"one":"...","other":"..."}).
// So a key's hash finds its entry in any pack directly, and an entry whose
// `source` is not the hash of a key this Office has is never shown: it was
// made from English the code no longer uses (L2, never a stale translation).
// What comes out is the table messages.js has always read, keyed by the
// English, so t(), tn(), tFilled() and addToCatalogue() work unchanged.
//
// English is a pack like the others (C15). Office carries a copy of a
// released lang-en in resources/languages (tools/take-lang-en.mjs), and an
// added lang-en is used instead when its (version, revision) is higher. Its
// wording is shown where it differs from the code's key; a key it does not
// have shows as the code writes it. In another language's window, a label
// that language lacks is shown in that English too.
//
// Built once per language and kept: a window asks for its own language as it
// starts (WORDS_CHANNEL, one synchronous request), and nothing else is read.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { compareVersions } from './addons-compat.js';
// Called only inside functions: language-choice.js reads this module's list, so neither runs the other at load.
import { packForTag } from './language-choice.js';

/**
 * The languages this release carries in its own bundle, and those an earlier
 * release carried that now come only as packs (L7), from the add-ons
 * contract. A language leaves the core only in a release whose packages feed
 * lists its pack: the release before counts a built-in language in use as an
 * installed lang-<tag>, so its update waits for that pack and brings it.
 */
const CONTRACT_LANGUAGES = createRequire(import.meta.url)('../addons-contract.json').languages || {};
export const BUILT_IN = Object.freeze([...(CONTRACT_LANGUAGES.builtIn || [])]);
export const LEFT_CORE = Object.freeze([...(CONTRACT_LANGUAGES.leftCore || [])]);

/** The built-in language a person uses, as the pack the update gate counts: [lang-<tag>] or []. */
export function builtInInUse(tag) {
  const base = String(tag || '').toLowerCase().replace(/_/g, '-').split('-')[0];
  return BUILT_IN.includes(base) ? [`lang-${base}`] : [];
}

/** The pack formats this core reads, from the add-ons contract: the current one and the one before (L2). */
export const LANGUAGE_FORMATS = Object.freeze([...(createRequire(import.meta.url)('../addons-contract.json').kinds?.language?.formats || [])]);

/** The languages written right to left, by base tag: for a tag neither a pack nor the list below speaks for. */
export const RIGHT_TO_LEFT = new Set(['ar', 'ckb', 'dv', 'fa', 'he', 'ps', 'sd', 'ug', 'ur', 'yi']);

const FORMS = ['zero', 'one', 'two', 'few', 'many', 'other'];

/** A tag as the windows write it: lower case, `-` for `_`. */
export const tagOf = (locale) => String(locale || '').toLowerCase().replace(/_/g, '-');

/**
 * Every language there is a pack for (L4): tag, pack, name, nativeName,
 * script, direction, region, variantOf. Taken from the packages repository
 * by tools/take-languages.mjs; the feed adds to it at run time (knownLanguages).
 */
export const LANGUAGE_LIST = Object.freeze((createRequire(import.meta.url)('../languages.json').languages || []).map((r) => Object.freeze({ ...r })));
const listed = new Map(LANGUAGE_LIST.map((r) => [tagOf(r.tag), r]));
/** The list's row for a tag, exact or by its base. */
export const listedLanguage = (tag) => listed.get(tagOf(tag)) || listed.get(tagOf(tag).split('-')[0]) || null;

/**
 * The languages Voices and languages can offer: the list, with the feed's
 * language entries over it (a newer name, its coverage and review) and any
 * language the feed has that the list does not. By tag, in the list's order.
 */
export function knownLanguages(feedPacks = []) {
  const out = new Map(LANGUAGE_LIST.map((r) => [tagOf(r.tag), { ...r }]));
  for (const p of feedPacks || []) {
    if (p?.kind !== 'language' || !p.tag || p.id === 'lang-en') continue;
    const tag = tagOf(p.tag);
    const fields = Object.fromEntries(['name', 'nativeName', 'script', 'direction', 'variantOf', 'review', 'coverage', 'status'].filter((k) => p[k] !== undefined).map((k) => [k, p[k]]));
    out.set(tag, { ...(out.get(tag) || { tag: p.tag }), ...fields, pack: p.id });
  }
  return [...out.values()];
}

/** The key the code uses for a label in catalogues/messages.json: its English, or its two forms as JSON. */
export const keyOf = (m) => (m.plural ? JSON.stringify({ one: m.message, other: m.plural }) : m.message);

/** A key's hash, as a pack's `source` holds it. */
export const hashOf = (key) => crypto.createHash('sha256').update(String(key), 'utf8').digest('hex').slice(0, 12);

/**
 * The labels this Office uses, from catalogues/messages.json: the shared
 * messages and each part's own (the Locker, Talk, the colour palette, Sheets'
 * Find), every key once. `leaveOut` names the parts this build does not carry
 * (a build made without the Locker), whose labels it never shows.
 * @param {{ messages?: object[], parts?: Record<string, { labels?: object[] }> }} list
 * @param {{ leaveOut?: string[] }} [o]
 */
export function codeLabels(list, { leaveOut = [] } = {}) {
  const out = [];
  const seen = new Set();
  const add = (m) => {
    if (!m || typeof m.message !== 'string') return;
    const key = keyOf(m);
    if (seen.has(key)) return;
    seen.add(key);
    out.push(m);
  };
  for (const m of Array.isArray(list?.messages) ? list.messages : []) add(m);
  for (const [part, p] of Object.entries(list?.parts || {})) {
    if (leaveOut.includes(part)) continue;
    for (const m of Array.isArray(p?.labels) ? p.labels : []) add(m);
  }
  return out;
}

const NAME = /\{(\w+)\}/g;
/** The `{names}` in a text, or none when it has no brace at all (most labels). */
const namesIn = (s) => (String(s).includes('{') ? [...String(s).matchAll(NAME)].map((m) => m[1]) : []);
/** A key's names, made once per key. */
const keyNames = new WeakMap();
const namesOfKey = (m) => {
  let n = keyNames.get(m);
  if (!n) keyNames.set(m, (n = new Set(namesIn(m.plural ? `${m.message} ${m.plural}` : m.message))));
  return n;
};

/**
 * An entry's words, as the table holds them for this key, or null when it may
 * not be shown: a string for a plain label, the forms for a counted one, and
 * no `{name}` the key does not have, which would be shown as it is written.
 */
export function wordsOf(m, e) {
  if (!e || typeof e !== 'object') return null;
  const fits = (s) => typeof s === 'string' && s !== '' && (!s.includes('{') || namesIn(s).every((n) => namesOfKey(m).has(n)));
  if (m.plural) {
    const f = e.forms;
    if (e.text !== undefined || !f || typeof f !== 'object' || !fits(f.other)) return null;
    const out = {};
    for (const name of FORMS) {
      if (f[name] === undefined) continue;
      if (!fits(f[name])) return null;
      out[name] = f[name];
    }
    return out;
  }
  if (e.forms !== undefined || !fits(e.text)) return null;
  return e.text;
}

/**
 * The table for one pack's messages.json: { key: words }, only for entries
 * whose `source` is a key's hash here. `hashes` is the keys' hashes in the
 * same order, made once.
 */
export function tableOf(keys, hashes, entries) {
  const bySource = new Map();
  for (const e of Object.values(entries || {})) if (e && typeof e.source === 'string') bySource.set(e.source, e);
  const table = {};
  let used = 0;
  keys.forEach((m, i) => {
    const words = wordsOf(m, bySource.get(hashes[i]));
    if (words === null) return;
    table[m.message] = words;
    used += 1;
  });
  return { table, used };
}

/** lang-en's table: only the labels it words otherwise than the code does. */
export function englishOf(keys, hashes, entries) {
  const { table } = tableOf(keys, hashes, entries);
  const out = {};
  for (const m of keys) {
    const w = table[m.message];
    if (w === undefined) continue;
    if (m.plural ? (w.one !== m.message || w.other !== m.plural || Object.keys(w).length > 2) : w !== m.message) out[m.message] = w;
  }
  return out;
}

/**
 * The versions a label has beside its medium one (L3, format 2): { key:
 * { short?, complete? } }, only for entries whose `source` is a key's hash
 * here, each version the shape of the medium (a string, or the forms of a
 * counted label) with no `{name}` the key does not have. A version that
 * reads as the medium does (`medium`, the label's words in this pack, else
 * the code's English) is left out: it adds nothing, and the windows are
 * sent less.
 */
export function versionsTableOf(keys, hashes, entries, medium = {}) {
  const out = {};
  if (!entries || typeof entries !== 'object') return out;
  const bySource = new Map();
  for (const e of Object.values(entries)) if (e && typeof e.source === 'string') bySource.set(e.source, e);
  const same = (a, b) => (typeof a === 'string' ? a === b : JSON.stringify(a) === JSON.stringify(b));
  keys.forEach((m, i) => {
    const e = bySource.get(hashes[i]);
    if (!e) return;
    const shown = medium[m.message] ?? (m.plural ? { one: m.message, other: m.plural } : m.message);
    const v = {};
    for (const which of ['short', 'complete']) {
      if (e[which] === undefined) continue;
      const words = wordsOf(m, m.plural ? { forms: e[which] } : { text: e[which] });
      if (words !== null && !same(words, shown)) v[which] = words;
    }
    if (v.short !== undefined || v.complete !== undefined) out[m.message] = v;
  });
  return out;
}

/**
 * A pack's manifest and messages, from its folder, or null when it cannot be
 * read or its format is not one this core reads. A format 2 pack's
 * versions.json comes too, as `versions`; without one, or unreadable, it is
 * null and the pack's labels have their medium version alone.
 */
export function readLanguagePack(dir, manifest = null) {
  try {
    const man = manifest || JSON.parse(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8'));
    if (man.kind !== 'language' || !LANGUAGE_FORMATS.includes(man.format)) return null;
    const entries = JSON.parse(fs.readFileSync(path.join(dir, 'messages.json'), 'utf8'));
    if (!entries || typeof entries !== 'object') return null;
    let versions = null;
    if (man.format >= 2) {
      try {
        const v = JSON.parse(fs.readFileSync(path.join(dir, 'versions.json'), 'utf8'));
        versions = v && typeof v === 'object' && !Array.isArray(v) ? v : null;
      } catch {
        versions = null;
      }
    }
    return { manifest: man, entries, versions };
  } catch {
    return null;
  }
}

/**
 * The English copy Office carries (C15): resources/languages/lang-en.pack.json,
 * the released envelope as it was, and the pack's files in lang-en/. Its
 * manifest is the envelope's; the program's own files are trusted as the
 * program is, and the gate checks the signature before a release (L9).
 */
export function readEnglishCopy(dir) {
  if (!dir) return null;
  try {
    const envelope = JSON.parse(fs.readFileSync(path.join(dir, 'lang-en.pack.json'), 'utf8'));
    if (envelope?.manifest?.id !== 'lang-en') return null;
    const pack = readLanguagePack(path.join(dir, 'lang-en'), envelope.manifest);
    return pack ? { ...pack, from: 'copy' } : null;
  } catch {
    return null;
  }
}

/** The direction a language is written in: its pack's word, else by its base tag. */
export function directionOf(tag, manifest = null) {
  if (manifest?.direction === 'rtl' || manifest?.direction === 'ltr') return manifest.direction;
  const row = listedLanguage(tag);
  if (row) return row.direction;
  return RIGHT_TO_LEFT.has(tagOf(tag).split('-')[0]) ? 'rtl' : 'ltr';
}

/** The list's directions by tag, small enough to hand every window: for a language with no pack in it. */
export const listedDirections = () => Object.fromEntries(LANGUAGE_LIST.map((r) => [tagOf(r.tag), r.direction]));

/**
 * @param {object} o
 * @param {{ message: string, plural?: string }[]} o.keys   catalogues/messages.json's messages
 * @param {string|null} [o.copyDir]                        resources/languages
 * @param {() => object[]} [o.addons]                      this run's add-ons ({ id, kind, dir, manifest })
 * @param {(m: string) => void} [o.log]
 * @returns {{ wordsFor: (tag: string) => object, languages: () => object[], english: () => object|null }}
 */
export function createLanguageWords({ keys, copyDir = null, addons = () => [], log = () => {} }) {
  const list = (Array.isArray(keys) ? keys : []).filter((m) => m && typeof m.message === 'string');
  let hashes = null;
  // messages.json holds each key's hash (tools/extract-messages.mjs); one without is hashed here.
  const hashesNow = () => (hashes ||= list.map((m) => (typeof m.source === 'string' && /^[0-9a-f]{12}$/.test(m.source) ? m.source : hashOf(keyOf(m)))));
  const made = new Map();
  let englishMade;

  const languagePacks = () => {
    try {
      return (addons() || []).filter((a) => a.kind === 'language' && a.manifest);
    } catch {
      return [];
    }
  };

  /** The English used this run: the copy, or an added lang-en above it. */
  function english() {
    if (englishMade !== undefined) return englishMade;
    const copy = readEnglishCopy(copyDir);
    const added = languagePacks().find((a) => a.id === 'lang-en');
    const addedPack = added ? readLanguagePack(added.dir, added.manifest) : null;
    const pick = addedPack && (!copy || compareVersions(addedPack.manifest, copy.manifest) > 0) ? { ...addedPack, from: 'added' } : copy;
    if (pick) {
      const table = englishOf(list, hashesNow(), pick.entries);
      englishMade = { manifest: pick.manifest, from: pick.from, table, versions: versionsTableOf(list, hashesNow(), pick.versions, table) };
    } else englishMade = null;
    if (!pick) log('no English pack: the code\'s own English is shown');
    return englishMade;
  }

  /**
   * The added pack for a window's tag: the exact tag; else the pack the
   * installer and a first start choose for it (language-choice.js: zh-CN is
   * lang-zh-hans, pt-AO lang-pt-pt, en-US lang-en-us); else its base. English
   * has no base pack here: an English window with no regional pack is lang-en's.
   */
  function packFor(tag) {
    const packs = languagePacks().filter((a) => a.id !== 'lang-en');
    const tagged = (t) => packs.find((a) => tagOf(a.manifest.tag) === t) || null;
    const chosen = (() => {
      try {
        const id = packForTag(tag);
        return (id && packs.find((a) => a.id === id)) || null;
      } catch {
        return null;
      }
    })();
    const base = tag.split('-')[0];
    return tagged(tag) || chosen || (base === 'en' ? null : tagged(base)) || null;
  }

  /**
   * A pack's table and its labels' other versions, a regional one laid over
   * its base (fr-CA over fr). A label the regional pack words itself takes
   * only its own versions: the base's short belongs to the base's medium.
   */
  function wordsOfPack(a, seen = new Set()) {
    if (!a || seen.has(a.id)) return null;
    seen.add(a.id);
    const pack = readLanguagePack(a.dir, a.manifest);
    if (!pack) {
      log(`${a.id}: its words could not be read, or its format is not one this Office reads`);
      return null;
    }
    const own = tableOf(list, hashesNow(), pack.entries).table;
    const ownVersions = versionsTableOf(list, hashesNow(), pack.versions, own);
    const baseId = a.manifest.variantOf;
    const base = baseId ? wordsOfPack(languagePacks().find((b) => b.id === baseId), seen) : null;
    if (!base) return { table: own, versions: ownVersions };
    const baseVersions = Object.fromEntries(Object.entries(base.versions).filter(([k]) => own[k] === undefined));
    return { table: { ...base.table, ...own }, versions: { ...baseVersions, ...ownVersions } };
  }
  const tableFor = (a, seen) => wordsOfPack(a, seen)?.table || null;

  const counted = new Map();
  return {
    english,
    /** The languages Voices and languages can offer: the list with the feed's entries over it. */
    known: (feedPacks = []) => knownLanguages(feedPacks),
    /**
     * How much of this Office's words an added language has (L6): { found, of },
     * found counted against this Office's own keys, a regional pack with its base.
     */
    coverageOf(id) {
      if (counted.has(id)) return counted.get(id);
      const a = languagePacks().find((x) => x.id === id);
      let answer = null;
      if (a) {
        const pack = id === 'lang-en' ? readLanguagePack(a.dir, a.manifest) : null;
        const table = id === 'lang-en' ? (pack ? tableOf(list, hashesNow(), pack.entries).table : null) : tableFor(a);
        if (table) answer = { found: Object.keys(table).length, of: list.length };
      }
      counted.set(id, answer);
      return answer;
    },
    /** The installed languages, English first: { tag, name, nativeName, direction, status }, status the pack's own (draft, reviewed). */
    languages() {
      const out = [{ tag: 'en', name: 'English', nativeName: 'English', direction: 'ltr' }];
      for (const a of languagePacks()) {
        if (a.id === 'lang-en' || !a.manifest.tag) continue;
        const row = listedLanguage(a.manifest.tag);
        out.push({ tag: tagOf(a.manifest.tag), name: a.manifest.name || row?.name || a.manifest.tag, nativeName: a.manifest.nativeName || row?.nativeName || a.manifest.name || a.manifest.tag, direction: directionOf(a.manifest.tag, a.manifest), ...(a.manifest.status === 'draft' || a.manifest.status === 'reviewed' ? { status: a.manifest.status } : {}) });
      }
      return out;
    },
    /**
     * A language's words: { tag, direction, table, english, pack, uses }
     * (`uses`: the ids of the packs its words came from, a regional one's base among them). `tag` is
     * the pack's when one is used, else null, and the window keeps its own
     * tables (the built-in ones, or English). `english` is the English
     * wording to show for a label the language lacks.
     *
     * `versions` (L3): the language's own short and complete versions of its
     * labels, { key: { short?, complete? } }; `englishVersions`: lang-en's,
     * for a label shown in English for want of a translation, and only for
     * that. An English window has lang-en's as both.
     */
    wordsFor(asked) {
      const tag = tagOf(asked) || 'en';
      if (made.has(tag)) return made.get(tag);
      const en = english();
      let answer;
      const regional = tag.split('-')[0] === 'en' ? packFor(tag) : null;
      if (tag.split('-')[0] === 'en' && !regional) {
        answer = { uses: en?.from === 'added' ? ['lang-en'] : [], tag: 'en', direction: 'ltr', table: en?.table || {}, english: null, versions: en?.versions || {}, englishVersions: en?.versions || {}, pack: en ? { id: 'lang-en', version: en.manifest.version, revision: en.manifest.revision || 0, from: en.from } : null };
      } else {
        const a = packFor(tag);
        const seen = new Set();
        const own = a ? wordsOfPack(a, seen) : null;
        const table = own?.table || null;
        answer = {
          uses: table ? [...seen] : [],
          tag: table ? tagOf(a.manifest.tag) : null,
          direction: directionOf(table ? a.manifest.tag : tag, table ? a.manifest : null),
          table,
          english: en?.table || null,
          versions: table ? own.versions : {},
          englishVersions: en?.versions || {},
          pack: table ? { id: a.id, version: a.manifest.version, revision: a.manifest.revision || 0, from: 'added' } : null,
        };
      }
      made.set(tag, answer);
      return answer;
    },
  };
}
