// For tests that read the labels a file passes to t() and tn() from its
// source: a call may name its label by its meaning key (D-2026-11-14-16), and
// these give the English the key's definition holds (catalogues/labels.json),
// the words those tests compare with a part's translations.
import fs from 'node:fs';

const LABELS = JSON.parse(fs.readFileSync(new URL('../packages/office-ui/src/catalogues/labels.json', import.meta.url), 'utf8')).labels;

/** A call's first argument as English: a key's English, else the words as written. */
export const englishOf = (s) => (Object.prototype.hasOwnProperty.call(LABELS, s) ? LABELS[s].english : s);

/** A key's English plural, or null. */
export const pluralOf = (s) => (Object.prototype.hasOwnProperty.call(LABELS, s) ? LABELS[s].plural || null : null);

/** Whether a string is a defined key. */
export const isLabelKey = (s) => Object.prototype.hasOwnProperty.call(LABELS, s);

const STR = String.raw`'((?:[^'\\\n]|\\.)*)'`;
const unescape = (s) => s.replace(/\\(n|.)/g, (m, c) => (c === 'n' ? '\n' : c));

/**
 * The labels a source text passes to t() and tn(), as English: a Map of the
 * English to { plural } (null for a plain label), whether the call names the
 * key or the English.
 */
export function wordsInSource(text) {
  const words = new Map();
  for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])t\(\s*${STR}`, 'g'))) {
    const s = unescape(m[1]);
    words.set(englishOf(s), { plural: null });
  }
  for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])tn\([^']*?,\s*${STR}\s*,\s*${STR}`, 'g'))) {
    const one = unescape(m[1]);
    if (!isLabelKey(one)) words.set(one, { plural: unescape(m[2]) });
  }
  for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])tn\([^']*?,\s*${STR}`, 'g'))) {
    const s = unescape(m[1]);
    if (isLabelKey(s)) words.set(englishOf(s), { plural: pluralOf(s) });
  }
  return words;
}
