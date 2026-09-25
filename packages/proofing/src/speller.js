// The dictionaries.
//
// Chromium's spell checker has no JavaScript API that answers "is this word
// right?", so the pass checks against Hunspell dictionaries of its own —
// SCOWL's English word lists, read by nspell. Two ship: English (United
// Kingdom) and English (United States). A dictionary is read the first time
// a pass asks for it, not when a window opens: building one takes the best
// part of a second.

/** Which dictionary a language tag is spelt with — or null for a language we have none for. */
export function dictionaryFor(tag) {
  const t = String(tag || '').trim().replace('_', '-').toLowerCase();
  if (!t) return null;
  if (!t.startsWith('en')) return null;
  if (t === 'en-us' || t === 'en-ph' || t === 'en-pr' || t === 'en-um' || t === 'en-vi' || t === 'en-as' || t === 'en-gu' || t === 'en-mp') return 'en-US';
  if (t === 'en') return null;
  return 'en-GB';
}

/**
 * The language a document is checked in: its own, where it says one we have
 * a dictionary for; otherwise the system's — English (United States) on a
 * system set to it, English (United Kingdom) everywhere else.
 */
export function chooseLanguage(documentLang, systemLocale) {
  const own = dictionaryFor(documentLang);
  if (own) return { lang: own, from: 'document' };
  const system = String(systemLocale || '').replace('_', '-').toLowerCase();
  return { lang: system === 'en-us' ? 'en-US' : 'en-GB', from: documentLang ? 'fallback' : 'system', documentLang: documentLang || null };
}

export const LANGUAGE_NAMES = { 'en-GB': 'English (United Kingdom)', 'en-US': 'English (United States)' };

const loading = new Map();

/** A dictionary, read and built once per process. */
export function loadDictionary(lang) {
  const which = lang === 'en-US' ? 'en-US' : 'en-GB';
  if (!loading.has(which)) {
    loading.set(which, (async () => {
      const [{ default: nspell }, dict] = await Promise.all([
        import('nspell'),
        which === 'en-US' ? import('dictionary-en') : import('dictionary-en-gb'),
      ]);
      const started = Date.now();
      const spell = nspell(dict.default);
      return { lang: which, spell, builtMs: Date.now() - started };
    })());
  }
  return loading.get(which);
}

/** Whether each word is spelt right, from a built dictionary. */
export function checkWords(spell, words) {
  const wrong = [];
  for (const w of words) {
    if (!spell.correct(w)) {
      // A possessive the dictionary does not list ("Northwind's" when it has
      // no Northwind) is as right as the word it is made from.
      const base = /^(.+)'s$/i.exec(w)?.[1];
      if (base && spell.correct(base)) continue;
      wrong.push(w);
    }
  }
  return wrong;
}

/** Suggestions, best first, at most `limit`. */
export function suggestWords(spell, word, limit = 6) {
  try {
    return spell.suggest(word).slice(0, limit);
  } catch {
    return [];
  }
}
