/**
 * Find, as every editing surface does it: a needle in a text, matched as
 * Office matches it.
 *
 * Case aside, Office in Arabic finds a word whatever marks it is written
 * with: Find's Match diacritics, Match kashida and Match alef hamza are off
 * unless a person turns them on. So "كتب" finds "كَتَبَ" and "كتـــب", and "احمد"
 * finds "أحمد". Here the same is true by default:
 *
 *   - the vowel marks (harakat: fatha, damma, kasra, sukun, shadda, tanwin),
 *     the superscript alef and the Quran's marks are passed over;
 *   - tatweel (kashida, the stretch in "كتـــب") is passed over;
 *   - alef with hamza or madda, and alef wasla, are alef; alef maqsura and
 *     Urdu's and Persian's yeh are yeh; Urdu's and Persian's kaf is kaf.
 *
 * `matchDiacritics: true` turns all of that off. A match is reported in the
 * text's own positions, from its first letter to the end of the marks on its
 * last, so replacing it takes the marks with it.
 */

/** Marks passed over: harakat, superscript alef, Quranic annotation, tatweel. */
const PASSED = /[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06DC\u06DF-\u06E4\u06E7\u06E8\u06EA-\u06ED\u0640]/;

/** Letters read as one. */
const SAME = new Map([
  ['\u0622', '\u0627'], ['\u0623', '\u0627'], ['\u0625', '\u0627'], ['\u0671', '\u0627'], // alef forms
  ['\u0649', '\u064A'], ['\u06CC', '\u064A'], // alef maqsura, Farsi yeh
  ['\u06A9', '\u0643'], // keheh
]);

/**
 * A text as a search compares it, with `map[i]` the position in the text of
 * folded character i (and `map[length]` the text's length).
 */
export function foldForSearch(text, { matchCase = false, matchDiacritics = false } = {}) {
  const source = String(text ?? '');
  let folded = '';
  const map = [];
  for (let i = 0; i < source.length; i++) {
    let ch = source[i];
    if (!matchDiacritics) {
      if (PASSED.test(ch)) continue;
      ch = SAME.get(ch) ?? ch;
    }
    if (!matchCase) {
      const lower = ch.toLowerCase();
      // A letter whose lower case is longer (İ) is kept as it is, so each
      // folded character still stands for one of the text's.
      if (lower.length === 1) ch = lower;
    }
    folded += ch;
    map.push(i);
  }
  map.push(source.length);
  return { text: folded, map };
}

/** Whether a needle folds to nothing: only marks, which find nothing on their own. */
const empty = (needle, options) => foldForSearch(needle, options).text === '';

/**
 * Every match of a needle in a text, as { start, end } in the text's own
 * positions, left to right and not overlapping.
 */
export function findAll(text, needle, options = {}) {
  const want = String(needle ?? '');
  if (!want) return [];
  // A needle of marks alone is looked for as it is written.
  const opts = empty(want, options) ? { ...options, matchDiacritics: true } : options;
  const hay = foldForSearch(text, opts);
  const target = foldForSearch(want, opts).text;
  const found = [];
  for (let at = hay.text.indexOf(target); at !== -1; at = hay.text.indexOf(target, at + target.length)) {
    const end = at + target.length;
    // The match runs to just before the next letter kept, so the marks on its last letter go with it.
    found.push({ start: hay.map[at], end: hay.map[end] });
  }
  return found;
}

/** Whether a text holds the needle. */
export const includesFolded = (text, needle, options = {}) => findAll(text, needle, options).length > 0;

/** A text with every match of the needle replaced: { text, count }. */
export function replaceAllIn(text, needle, replacement, options = {}) {
  const source = String(text ?? '');
  const matches = findAll(source, needle, options);
  if (!matches.length) return { text: source, count: 0 };
  let out = '';
  let from = 0;
  for (const m of matches) {
    out += source.slice(from, m.start) + String(replacement ?? '');
    from = m.end;
  }
  return { text: out + source.slice(from), count: matches.length };
}
