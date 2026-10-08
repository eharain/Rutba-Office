/**
 * Review → Thesaurus: words of like meaning, from a word list of the suite's
 * own (thesaurus-en.js).
 *
 * The list is sets of words that mean the same thing in one sense, each with
 * its part of speech — "adj: happy, glad, cheerful, joyful, …". A word in
 * several sets has several meanings, and the pane shows each as a group of
 * its own, as Word's does. A word typed with an ending (running, cities,
 * happier) is looked up by the word it was made from when it is not in the
 * list itself, and the pane says so.
 */

import { THESAURUS_EN } from './thesaurus-en.js';

const PARTS = { n: 'noun', v: 'verb', a: 'adjective', r: 'adverb' };

let index = null;

/** word -> [set], built the first time a word is looked up. */
function sets() {
  if (index) return index;
  index = new Map();
  for (const line of THESAURUS_EN.split('\n')) {
    const m = /^\s*([nvar])\s*:\s*(.+)$/.exec(line);
    if (!m) continue;
    const words = m[2].split(',').map((w) => w.trim()).filter(Boolean);
    if (words.length < 2) continue;
    const set = { pos: m[1], words };
    for (const w of words) {
      const key = w.toLowerCase();
      if (!index.has(key)) index.set(key, []);
      index.get(key).push(set);
    }
  }
  return index;
}

/** The words a form may have been made from, most likely first. */
function bases(word) {
  const w = word.toLowerCase();
  const out = [];
  const add = (b, how) => { if (b && b.length > 1 && b !== w) out.push({ base: b, how }); };
  if (w.endsWith('ies')) add(w.slice(0, -3) + 'y', 'plural');
  if (w.endsWith('es')) add(w.slice(0, -2), 'plural');
  if (w.endsWith('s') && !w.endsWith('ss')) add(w.slice(0, -1), 'plural');
  if (w.endsWith('ied')) add(w.slice(0, -3) + 'y', 'past');
  if (w.endsWith('ed')) {
    add(w.slice(0, -2), 'past');
    add(w.slice(0, -1), 'past');
    if (/([bdfgklmnprtvz])\1ed$/.test(w)) add(w.slice(0, -3), 'past');
  }
  if (w.endsWith('ing')) {
    add(w.slice(0, -3), 'ing');
    add(w.slice(0, -3) + 'e', 'ing');
    if (/([bdfgklmnprtvz])\1ing$/.test(w)) add(w.slice(0, -4), 'ing');
    if (w.endsWith('ying')) add(w.slice(0, -4) + 'ie', 'ing');
  }
  if (w.endsWith('ier')) add(w.slice(0, -3) + 'y', 'comparative');
  if (w.endsWith('iest')) add(w.slice(0, -4) + 'y', 'superlative');
  if (w.endsWith('er')) { add(w.slice(0, -2), 'comparative'); add(w.slice(0, -1), 'comparative'); if (/([bdgmnpt])\1er$/.test(w)) add(w.slice(0, -3), 'comparative'); }
  if (w.endsWith('est')) { add(w.slice(0, -3), 'superlative'); add(w.slice(0, -2), 'superlative'); if (/([bdgmnpt])\1est$/.test(w)) add(w.slice(0, -4), 'superlative'); }
  if (w.endsWith('ily')) add(w.slice(0, -3) + 'y', 'adverb');
  if (w.endsWith('ly')) add(w.slice(0, -2), 'adverb');
  return out;
}

/**
 * The meanings of `word`: `{ word, base, how, meanings: [{ pos, part,
 * words }] }`, `base` the word it was found by (itself, or what an ending
 * was made from) and `how` the ending; no meanings when the list has none.
 */
export function lookUp(word) {
  const query = String(word || '').trim().replace(/^[^\p{L}]+|[^\p{L}'-]+$/gu, '');
  const out = { word: query, base: null, how: null, meanings: [] };
  if (!query) return out;
  const table = sets();
  let found = table.get(query.toLowerCase());
  let base = query.toLowerCase();
  let how = null;
  if (!found) {
    for (const b of bases(query)) {
      if (table.has(b.base)) { found = table.get(b.base); base = b.base; how = b.how; break; }
    }
  }
  if (!found) return out;
  out.base = base;
  out.how = how;
  // A set's part of speech decides which endings fit it: a plural looks for nouns, a past for verbs.
  const fits = { plural: ['n', 'v'], past: ['v', 'a'], ing: ['v', 'n', 'a'], comparative: ['a', 'r'], superlative: ['a', 'r'], adverb: ['a', 'r'] };
  const chosen = how ? found.filter((s) => (fits[how] || []).includes(s.pos)) : found;
  for (const s of (chosen.length ? chosen : found)) {
    const words = s.words.filter((w) => w.toLowerCase() !== base);
    if (words.length) out.meanings.push({ pos: s.pos, part: PARTS[s.pos], words });
  }
  return out;
}

/** How many words the list knows: for the pane's footnote and the tests. */
export function thesaurusSize() {
  return sets().size;
}
