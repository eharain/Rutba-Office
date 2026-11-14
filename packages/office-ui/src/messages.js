// The message catalogue: the words a window shows, in the language it is set to.
//
// Each message is written in English where it is used and passed through
// `t`, which is also its key: `t('Close')`, `t('Saved {name}', { name })`.
// A catalogue for another language maps each English message to its
// translation; a message it has none for shows in English, so a catalogue
// half done is a window half translated, never a broken one.
//
// English is a language pack too (L1): its table holds only the labels it
// words otherwise than the code, and a window in another language shows that
// English for a label its own table lacks (setFallback). With no table at
// all, the code's English is shown as it is written.
//
// Counted messages go through `tn(count, one, other, values)`: the
// language's own plural rules (Intl.PluralRules) choose the form — Arabic
// has six, Japanese one — and a catalogue gives the forms it needs by those
// rules' names: { one, two, few, many, other, zero }.
//
// A label can have two more versions than the one `t` gives (L3): a short
// one for a ribbon with too little room, and a complete one for its tooltip.
// They come from the language's pack (setVersions), and `versionsOf(shown)`
// finds them for a text `t` or `tn` gave lately, filled with the same names
// and in the same plural form, so no call changes: a button asks for its own
// label's. A label's versions are its language's own; lang-en's are used
// only for a label shown in English for want of a translation, never in
// place of a version the language lacks.
//
// A call may name its label by its meaning key instead (D-2026-11-14-10 to
// -16): `t('common.close')`, `tn(count, key, values)` for a counted one. The
// key's definition (catalogues/labels.json) gives its English. Every table
// is held by key (stage 5): one given by English (a pack, a part's words) is
// put by key as it is registered, and a call by English, or words that
// arrive as data, find their label's key through its English.
//
// tools/extract-messages.mjs keeps catalogues/labels.json to every label the
// code passes to `t` and `tn`, the list a translator works through; its
// literal strings are what makes that possible, so a label is always written
// out in the call, never built from pieces.

import KEYS from './catalogues/keys.js';

/** The keys known here: the catalogue's, and those a part adds for itself (addKeys). */
const known = new Map(Object.entries(KEYS));

/**
 * More keys and their English, from a part of the suite a build can leave out
 * (the Locker, locker-keys.js), so a build without it carries none of them.
 */
export function addKeys(more) {
  for (const [key, english] of Object.entries(more || {})) known.set(key, english);
  byEnglish = null;
  // A table registered before these keys were known may name their labels by English: by key now.
  for (const tag of new Set([...bases.keys(), ...additions.keys(), ...packs.keys()])) compose(tag);
  fallback = byKey(fallbackGiven);
  versionsOwn = byKey(versionsOwnGiven);
  versionsEnglish = byKey(versionsEnglishGiven);
  recent.clear();
}

/** Each key by its English (its singular for a counted one), made when first asked for: a call by English still names its key. */
let byEnglish = null;
const keyOfEnglish = (message) => {
  if (!byEnglish) {
    byEnglish = new Map();
    for (const [key, e] of known) byEnglish.set(Array.isArray(e) ? e[0] : e, key);
  }
  return byEnglish.get(message);
};

/**
 * A label's key: itself when it is one, else the key whose English it is (a
 * call, or a table, naming it by its English, or words that arrive as data:
 * a template's name), else null.
 */
const keyOfLabel = (label) => (typeof label !== 'string' ? null : known.has(label) ? label : keyOfEnglish(label) ?? null);

/**
 * A table by key (stage 5 of the meaning keys): an entry named by its English
 * is moved to its label's key; one whose words name no label (an engine's
 * message the definitions do not hold) is kept as it is. An entry named by
 * its key wins over the same label named by its English.
 */
function byKey(given) {
  if (!given || typeof given !== 'object') return null;
  const out = {};
  for (const [k, v] of Object.entries(given)) if (known.has(k)) out[k] = v;
  for (const [k, v] of Object.entries(given)) {
    if (known.has(k)) continue;
    const key = keyOfEnglish(k) ?? k;
    if (!Object.prototype.hasOwnProperty.call(out, key)) out[key] = v;
  }
  return out;
}

/**
 * The texts t() and tn() gave lately, each with the key it was made from
 * (or MIXED, when two keys give that text): what labelOf reads, so a
 * component given a label's words can mark the element with its key
 * (D-2026-11-14-11). Bounded; the one given or read longest ago goes first,
 * so a label a window keeps drawing, or keeps marking, keeps its key while
 * texts made with values (a file's name in a question) pass through.
 */
const SHOWN = 12000;
const shownKeys = new Map();
function shownAs(shown, key) {
  if (!key) return;
  const was = shownKeys.get(shown);
  shownKeys.delete(shown);
  shownKeys.set(shown, was === undefined || was === key ? key : 'mixed');
  if (shownKeys.size > SHOWN) shownKeys.delete(shownKeys.keys().next().value);
}

/**
 * The key a text t() or tn() gave lately was made from, in its two parts:
 * { key, scope, name } (word.ribbon.multiple-pages: scope word.ribbon, name
 * multiple-pages), or null when it was not given lately, has no key, or two
 * keys give the same text. The scope says where the key applies and is never
 * shown; the name is what a window with no pack shows.
 */
export function labelOf(shown) {
  const text = String(shown ?? '');
  const key = shownKeys.get(text);
  if (!key || key === 'mixed') return null;
  // A text read is one in use: it goes to the back, as if given again.
  shownKeys.delete(text);
  shownKeys.set(text, key);
  const dot = key.lastIndexOf('.');
  return { key, scope: key.slice(0, dot), name: key.slice(dot + 1) };
}

/** A meaning key's English: a string, or [one, other] for a counted label; undefined for words that are no key. */
const englishOfKey = (key) => (typeof key === 'string' ? known.get(key) : undefined);

/** A label as the tables hold it, its English: a key's own, else the words as given. */
const asWords = (message) => {
  const e = englishOfKey(message);
  if (e === undefined) return message;
  return Array.isArray(e) ? e[0] : e;
};

const catalogues = new Map();
const bases = new Map();
const additions = new Map();
const packs = new Map();
let current = 'en';
let table = null;
let fallback = null;
/** The tables as given, before they are put by key: kept so keys added later (addKeys) still find them. */
let fallbackGiven = null;
let versionsOwnGiven = null;
let versionsEnglishGiven = null;
const asked = new Set();
/** The window's language's versions of its labels and lang-en's: { key: { short?, complete? } } (L3). */
let versionsOwn = null;
let versionsEnglish = null;
/**
 * The texts t() and tn() gave lately, each with the key, names and count it
 * was made from, and whether it came in the window's language: what
 * versionsOf reads. Bounded; the oldest goes first. A text two labels give
 * alike whose versions differ is remembered as having none.
 */
const RECENT = 4000;
const recent = new Map();
const MIXED = 'mixed';

/** English as the source language words it (lang-en's table, by key), a label the window's language lacks shown in it. */
const english = (at, message) => {
  const w = fallback?.[at];
  return typeof w === 'string' && w ? w : message;
};

/**
 * A language tag as BCP 47 writes it, lower-cased. A system locale such as
 * pt_BR carries an underscore, which Intl refuses with a RangeError, so the
 * first counted message would have thrown in a window set to it.
 */
const tagOf = (locale) => String(locale).toLowerCase().replace(/_/g, '-');

/** A language's table: its catalogue, a part's own words over it, an installed pack's words over all (D-2026-11-14-16). */
function compose(tag) {
  const whole = { ...byKey(bases.get(tag)), ...byKey(additions.get(tag)), ...byKey(packs.get(tag)) };
  catalogues.set(tag, whole);
  if (tag === current) table = whole;
}

/** Register a language's catalogue: { 'English message': 'translation' | { one, other, … } }. */
export function registerCatalogue(locale, messages) {
  const tag = tagOf(locale);
  bases.set(tag, messages || {});
  compose(tag);
}

/**
 * More messages for a language, from a part of the suite a build can leave out
 * and so has no place in the shared catalogue. Kept whichever is registered
 * first; an installed pack's words for the same labels stay above them.
 */
export function addToCatalogue(locale, messages) {
  const tag = tagOf(locale);
  additions.set(tag, { ...(additions.get(tag) || {}), ...messages });
  compose(tag);
}

/**
 * An installed language pack's table for its language (L1): above the
 * built-in catalogue and above the words a part adds for itself, whichever
 * comes first, so a pack's words are the ones shown (D-2026-11-14-16).
 */
export function registerPack(locale, messages) {
  const tag = tagOf(locale);
  packs.set(tag, messages || {});
  compose(tag);
}

/**
 * Show the windows in a language, by its BCP 47 tag ("fr-CA", "ar"): its
 * own catalogue, else its base language's ("fr"), else English. Answers
 * the language chosen.
 */
export function setLanguage(locale) {
  const tag = tagOf(locale || 'en');
  const base = tag.split('-')[0];
  current = catalogues.has(tag) ? tag : catalogues.has(base) ? base : 'en';
  table = catalogues.get(current) || null;
  asked.clear();
  return current;
}

/**
 * The English shown for a label the window's language has no words for: the
 * source language's own table (lang-en), keyed like any other, or none for
 * the code's English as written.
 */
export function setFallback(messages) {
  fallbackGiven = messages && typeof messages === 'object' ? messages : null;
  fallback = byKey(fallbackGiven);
}

/**
 * The short and complete versions of the window's labels (L3): `own`, the
 * language's, for a label shown in it; `english`, lang-en's, for one shown in
 * English for want of a translation (in an English window, both are lang-en's).
 */
export function setVersions(own, english) {
  const usable = (v) => (v && typeof v === 'object' && Object.keys(v).length ? v : null);
  versionsOwnGiven = usable(own);
  versionsEnglishGiven = usable(english);
  versionsOwn = byKey(versionsOwnGiven);
  versionsEnglish = byKey(versionsEnglishGiven);
  recent.clear();
}

const sameVersions = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
const versionsAt = (r) => (r.own ? versionsOwn : versionsEnglish)?.[r.key];

/** Remember a text t() or tn() gave, for versionsOf; nothing to do when there are no versions at all. */
function remember(shown, key, values, count, own) {
  if (!versionsOwn && !versionsEnglish) return;
  const entry = { key, values, count, own };
  const was = recent.get(shown);
  if (was !== undefined) {
    recent.delete(shown);
    if (was === MIXED || ((was.key !== key || was.own !== own) && !sameVersions(versionsAt(was), versionsAt(entry)))) {
      recent.set(shown, MIXED);
      return;
    }
  }
  recent.set(shown, entry);
  if (recent.size > RECENT) recent.delete(recent.keys().next().value);
}

/** A counted version's form for `count`, by the rules of the language it is in, as tn() picks the medium's. */
function formOf(forms, count, own) {
  if (typeof forms === 'string') return forms;
  if (!forms || typeof forms !== 'object') return '';
  const n = Number(count);
  if (!own) return (n === 1 ? forms.one : forms.other) || forms.other || '';
  let rule = 'other';
  try { rule = new Intl.PluralRules(current).select(n); } catch { /* an unknown tag: other */ }
  return (n === 0 && forms.zero) || forms[rule] || forms.other || '';
}

/**
 * The other versions of a label t() or tn() gave lately: { short?, complete? }
 * filled as it was, or null when it has none, it was not given lately, or two
 * labels give that text with versions that differ (a wrong short is worse
 * than none).
 */
export function versionsOf(shown) {
  const r = recent.get(String(shown ?? ''));
  if (!r || r === MIXED) return null;
  const v = versionsAt(r);
  if (!v) return null;
  const out = {};
  for (const which of ['short', 'complete']) {
    const words = v[which] === undefined ? '' : formOf(v[which], r.count, r.own);
    if (words) out[which] = fill(words, r.values);
  }
  return out.short || out.complete ? out : null;
}

/** The language the windows are shown in. */
export const language = () => current;

/**
 * The digits the window writes its numbers in: a numbering system ('latn'
 * for 0 to 9, 'arab' for Arabic's ٠ to ٩, 'arabext' for Urdu's and Persian's
 * ۰ to ۹), or none for the language's own. Arabic windows write 0 to 9
 * unless the digits setting says otherwise.
 */
let digits = null;
export function setDigits(system) {
  digits = /^[a-z]{4,8}$/.test(String(system || '')) ? String(system) : null;
  return digits;
}

/** The tag numbers and dates are formatted with: the language, and its digits when they are set. */
export const numberTag = () => (digits ? `${current}-u-nu-${digits}` : current);

/** The languages there is a catalogue for, English first. */
export const languages = () => ['en', ...[...catalogues.keys()].filter((k) => k !== 'en')];

/** `{name}` in a message filled from `values`; a name with no value is left as written. */
function fill(text, values) {
  if (!values) return text;
  return String(text).replace(/\{(\w+)\}/g, (m, name) => (values[name] === undefined || values[name] === null ? m : String(values[name])));
}

/**
 * A message marked for the catalogue where it is written down — in a table
 * made before the language is set — and shown through `t` where it is used.
 * Named by its key, it gives the key's English, as a message written in
 * English gives itself: what reads the table (an engine's words, the
 * installer's) reads the words.
 */
export const msg = (message) => asWords(message);

/** A message in the current language, its `{names}` filled. */
export function t(label, values) {
  const message = asWords(label);
  const key = keyOfLabel(label);
  // Where the tables hold it: its key, or the words as given for words no label is.
  const at = key ?? message;
  const own = table?.[at];
  if (typeof own === 'string' && own) {
    const shown = fill(own, values);
    remember(shown, at, values, null, true);
    shownAs(shown, key);
    return shown;
  }
  if (table && current !== 'en') asked.add(message);
  const shown = fill(english(at, message), values);
  remember(shown, at, values, null, current === 'en');
  shownAs(shown, key);
  return shown;
}

/**
 * A counted message: `one` or `other` in English, the language's own forms
 * in a catalogue. `{count}` is the count, as the language writes numbers.
 * Named by its key, the call is `tn(count, key, values)`.
 */
export function tn(count, label, ...rest) {
  const e = englishOfKey(label);
  const keyed = e !== undefined;
  const one = keyed ? (Array.isArray(e) ? e[0] : e) : label;
  const other = keyed ? (Array.isArray(e) ? e[1] : e) : rest[0];
  const key = keyed ? label : keyOfLabel(one);
  const at = key ?? one;
  const values = (keyed ? rest[0] : rest[1]) || {};
  const n = Number(count);
  const shown = Number.isFinite(n) ? n.toLocaleString(numberTag()) : String(count);
  const all = { ...values, count: shown };
  const source = fallback?.[at] && typeof fallback[at] === 'object' ? fallback[at] : null;
  const inEnglish = (n === 1 ? source?.one || one : source?.other || other);
  const forms = table?.[at];
  if (!forms || typeof forms !== 'object') {
    if (table && current !== 'en') asked.add(one);
    const text = fill(inEnglish, all);
    remember(text, at, all, n, current === 'en');
    shownAs(text, key);
    return text;
  }
  let rule = 'other';
  try { rule = new Intl.PluralRules(current).select(n); } catch { /* an unknown tag: other */ }
  const form = (n === 0 && forms.zero) || forms[rule] || forms.other;
  const text = fill(form || inEnglish, all);
  remember(text, at, all, n, Boolean(form));
  shownAs(text, key);
  return text;
}

/**
 * A message that arrives already written — an engine's or a service's,
 * thrown in English with its names filled in ('There is no paragraph style
 * "Quote".') — in the window's language: itself when the catalogue has it as
 * it is, else the message whose pattern it fits ('There is no paragraph style
 * "{name}".'), translated with the names put back. Unknown, it stays as it came.
 */
export function tFilled(message) {
  const text = String(asWords(message) ?? '');
  if (!table || !text) return text;
  const own = table[keyOfLabel(text) ?? text];
  if (typeof own === 'string' && own) return own;
  for (const { re, names, key } of patternsFor(table)) {
    const m = re.exec(text);
    if (!m) continue;
    const values = Object.fromEntries(names.map((name, i) => [name, m[i + 1]]));
    const translation = table[key];
    if (typeof translation === 'string' && translation) return fill(translation, values);
  }
  return text;
}

/** The catalogue's messages with names in them, as patterns to fit a filled message to; made once per catalogue. */
const patterns = new WeakMap();
function patternsFor(catalogue) {
  if (patterns.has(catalogue)) return patterns.get(catalogue);
  const list = [];
  for (const at of Object.keys(catalogue)) {
    // The table is by key; the pattern is the label's English, as an engine writes it.
    const english = asWords(at);
    // A message of names and marks alone ('{a} × {b}') would fit almost anything: only one with words in it is a pattern.
    if (typeof catalogue[at] !== 'string' || typeof english !== 'string' || !/\{\w+\}/.test(english) || !/[A-Za-z]{2,}/.test(english.replace(/\{\w+\}/g, ''))) continue;
    const names = [];
    const source = english.split(/(\{\w+\})/).map((part) => {
      const name = /^\{(\w+)\}$/.exec(part)?.[1];
      if (name) { names.push(name); return '(.+?)'; }
      return part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }).join('');
    // The most literal patterns first, so a specific message wins over a loose one.
    list.push({ re: new RegExp(`^${source}$`, 's'), names, key: at, weight: english.replace(/\{\w+\}/g, '').length });
  }
  list.sort((a, b) => b.weight - a.weight);
  patterns.set(catalogue, list);
  return list;
}

/** The messages shown in English since the language was set for want of a translation: what a translator has left. */
export const untranslated = () => [...asked].sort();
