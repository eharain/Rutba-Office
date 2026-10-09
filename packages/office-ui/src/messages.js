// The message catalogue: the words a window shows, in the language it is set to.
//
// Each message is written in English where it is used and passed through
// `t`, which is also its key: `t('Close')`, `t('Saved {name}', { name })`.
// A catalogue for another language maps each English message to its
// translation; a message it has none for shows in English, so a catalogue
// half done is a window half translated, never a broken one. English needs
// no catalogue at all.
//
// Counted messages go through `tn(count, one, other, values)`: the
// language's own plural rules (Intl.PluralRules) choose the form — Arabic
// has six, Japanese one — and a catalogue gives the forms it needs by those
// rules' names: { one, two, few, many, other, zero }.
//
// tools/extract-messages.mjs lists every message the code passes to `t` and
// `tn` in catalogues/messages.json, the list a translator works through;
// its literal strings are what makes that possible, so a message is always
// written out in the call, never built from pieces.

const catalogues = new Map();
let current = 'en';
let table = null;
const asked = new Set();

/**
 * A language tag as BCP 47 writes it, lower-cased. A system locale such as
 * pt_BR carries an underscore, which Intl refuses with a RangeError, so the
 * first counted message would have thrown in a window set to it.
 */
const tagOf = (locale) => String(locale).toLowerCase().replace(/_/g, '-');

/** Register a language's catalogue: { 'English message': 'translation' | { one, other, … } }. */
export function registerCatalogue(locale, messages) {
  const tag = tagOf(locale);
  catalogues.set(tag, messages || {});
  if (tag === current) table = messages || {};
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
  table = current === 'en' ? null : catalogues.get(current);
  asked.clear();
  return current;
}

/** The language the windows are shown in. */
export const language = () => current;

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
 */
export const msg = (message) => message;

/** A message in the current language, its `{names}` filled. */
export function t(message, values) {
  if (!table) return fill(message, values);
  const own = table[message];
  if (typeof own === 'string' && own) return fill(own, values);
  asked.add(message);
  return fill(message, values);
}

/**
 * A counted message: `one` or `other` in English, the language's own forms
 * in a catalogue. `{count}` is the count, as the language writes numbers.
 */
export function tn(count, one, other, values = {}) {
  const n = Number(count);
  const shown = Number.isFinite(n) ? n.toLocaleString(current) : String(count);
  const all = { ...values, count: shown };
  const english = n === 1 ? one : other;
  if (!table) return fill(english, all);
  const forms = table[one];
  if (!forms || typeof forms !== 'object') {
    asked.add(one);
    return fill(english, all);
  }
  let rule = 'other';
  try { rule = new Intl.PluralRules(current).select(n); } catch { /* an unknown tag: other */ }
  const form = (n === 0 && forms.zero) || forms[rule] || forms.other;
  return fill(form || english, all);
}

/**
 * A message that arrives already written — an engine's or a service's,
 * thrown in English with its names filled in ('There is no paragraph style
 * "Quote".') — in the window's language: itself when the catalogue has it as
 * it is, else the message whose pattern it fits ('There is no paragraph style
 * "{name}".'), translated with the names put back. Unknown, it stays as it came.
 */
export function tFilled(message) {
  const text = String(message ?? '');
  if (!table || !text) return text;
  const own = table[text];
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
  for (const key of Object.keys(catalogue)) {
    // A message of names and marks alone ('{a} × {b}') would fit almost anything: only one with words in it is a pattern.
    if (typeof catalogue[key] !== 'string' || !/\{\w+\}/.test(key) || !/[A-Za-z]{2,}/.test(key.replace(/\{\w+\}/g, ''))) continue;
    const names = [];
    const source = key.split(/(\{\w+\})/).map((part) => {
      const name = /^\{(\w+)\}$/.exec(part)?.[1];
      if (name) { names.push(name); return '(.+?)'; }
      return part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }).join('');
    // The most literal patterns first, so a specific message wins over a loose one.
    list.push({ re: new RegExp(`^${source}$`, 's'), names, key, weight: key.replace(/\{\w+\}/g, '').length });
  }
  list.sort((a, b) => b.weight - a.weight);
  patterns.set(catalogue, list);
  return list;
}

/** The messages shown in English since the language was set for want of a translation: what a translator has left. */
export const untranslated = () => [...asked].sort();
