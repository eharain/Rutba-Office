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

/** Register a language's catalogue: { 'English message': 'translation' | { one, other, … } }. */
export function registerCatalogue(locale, messages) {
  catalogues.set(String(locale).toLowerCase(), messages || {});
  if (String(locale).toLowerCase() === current) table = messages || {};
}

/**
 * Show the windows in a language, by its BCP 47 tag ("fr-CA", "ar"): its
 * own catalogue, else its base language's ("fr"), else English. Answers
 * the language chosen.
 */
export function setLanguage(locale) {
  const tag = String(locale || 'en').toLowerCase();
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

/** The messages shown in English since the language was set for want of a translation: what a translator has left. */
export const untranslated = () => [...asked].sort();
