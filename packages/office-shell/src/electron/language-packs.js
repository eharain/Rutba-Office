// Words from a language pack (L1, L2): the table a window, and the main
// process, show their words from, made from an installed lang-<tag> add-on.
//
// The code names every label by its English as written in it, the t() or
// tn() argument (catalogues/messages.json lists them all). A pack holds each
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

/** A pack's manifest and messages, from its folder, or null when it cannot be read or its format is not one this core reads. */
export function readLanguagePack(dir, manifest = null) {
  try {
    const man = manifest || JSON.parse(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8'));
    if (man.kind !== 'language' || !LANGUAGE_FORMATS.includes(man.format)) return null;
    const entries = JSON.parse(fs.readFileSync(path.join(dir, 'messages.json'), 'utf8'));
    return entries && typeof entries === 'object' ? { manifest: man, entries } : null;
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
    englishMade = pick ? { manifest: pick.manifest, from: pick.from, table: englishOf(list, hashesNow(), pick.entries) } : null;
    if (!pick) log('no English pack: the code\'s own English is shown');
    return englishMade;
  }

  /** The pack for a tag: the exact tag, else its base. */
  function packFor(tag) {
    const packs = languagePacks().filter((a) => a.id !== 'lang-en');
    const tagged = (t) => packs.find((a) => tagOf(a.manifest.tag) === t) || null;
    return tagged(tag) || tagged(tag.split('-')[0]) || null;
  }

  /** A pack's table, a regional one laid over its base (fr-CA over fr). */
  function tableFor(a, seen = new Set()) {
    if (!a || seen.has(a.id)) return null;
    seen.add(a.id);
    const pack = readLanguagePack(a.dir, a.manifest);
    if (!pack) {
      log(`${a.id}: its words could not be read, or its format is not one this Office reads`);
      return null;
    }
    const own = tableOf(list, hashesNow(), pack.entries).table;
    const baseId = a.manifest.variantOf;
    const base = baseId ? tableFor(languagePacks().find((b) => b.id === baseId), seen) : null;
    return base ? { ...base, ...own } : own;
  }

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
    /** The installed languages, English first: { tag, name, nativeName, direction }. */
    languages() {
      const out = [{ tag: 'en', name: 'English', nativeName: 'English', direction: 'ltr' }];
      for (const a of languagePacks()) {
        if (a.id === 'lang-en' || !a.manifest.tag) continue;
        const row = listedLanguage(a.manifest.tag);
        out.push({ tag: tagOf(a.manifest.tag), name: a.manifest.name || row?.name || a.manifest.tag, nativeName: a.manifest.nativeName || row?.nativeName || a.manifest.name || a.manifest.tag, direction: directionOf(a.manifest.tag, a.manifest) });
      }
      return out;
    },
    /**
     * A language's words: { tag, direction, table, english, pack, uses }
     * (`uses`: the ids of the packs its words came from, a regional one's base among them). `tag` is
     * the pack's when one is used, else null, and the window keeps its own
     * tables (the built-in ones, or English). `english` is the English
     * wording to show for a label the language lacks.
     */
    wordsFor(asked) {
      const tag = tagOf(asked) || 'en';
      if (made.has(tag)) return made.get(tag);
      const en = english();
      let answer;
      if (tag.split('-')[0] === 'en') {
        answer = { uses: en?.from === 'added' ? ['lang-en'] : [], tag: 'en', direction: 'ltr', table: en?.table || {}, english: null, pack: en ? { id: 'lang-en', version: en.manifest.version, revision: en.manifest.revision || 0, from: en.from } : null };
      } else {
        const a = packFor(tag);
        const seen = new Set();
        const table = a ? tableFor(a, seen) : null;
        answer = {
          uses: table ? [...seen] : [],
          tag: table ? tagOf(a.manifest.tag) : null,
          direction: directionOf(table ? a.manifest.tag : tag, table ? a.manifest : null),
          table,
          english: en?.table || null,
          pack: table ? { id: a.id, version: a.manifest.version, revision: a.manifest.revision || 0, from: 'added' } : null,
        };
      }
      made.set(tag, answer);
      return answer;
    },
  };
}
