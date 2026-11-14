// The checks name a control by its label's English, as the code wrote it
// before the meaning keys (D-2026-11-14-16). With an English copy the window
// may word that label its own way (lang-en's medium, short or complete), and a
// button's tip is its complete version, so a check that matched the words or
// the tip misses the control. The window marks each labelled element with its
// key (stage 3: data-scope and data-key), so a check finds the control by the
// key that English names, whatever words the window draws for it.
//
// For the code a check sends into a window, as a JavaScript expression:
//
//   const b = list.find(${labelKeyed('Multiple Pages')}) || list.find(old match);
//
// labelKeyed() is a predicate that holds for an element carrying that label's key,
// or one that holds for none when the English names no label (a check's own
// words, a file name), so the check falls back to the words as before. The
// key is tried first: a prefix of the words can match a longer label too
// ("Group" begins "Group by"), the key never does.

import { createRequire } from 'node:module';

let byEnglish = null;
function keys() {
  if (byEnglish) return byEnglish;
  const { labels } = createRequire(import.meta.url)('@rutba/office-ui/catalogues/labels.json');
  byEnglish = new Map();
  for (const [key, d] of Object.entries(labels)) byEnglish.set(d.english, key);
  return byEnglish;
}

/**
 * The key of the label whose English is these words, or null. A check names
 * a control by the start of its words as often as by all of them (the start
 * of a tip that runs on, as the TOA check names Update Table): words no label is, but
 * that begin exactly one label's English, give that label's key; words that
 * begin more than one give none, as a guess could be the wrong control.
 */
export function keyFor(english) {
  const words = String(english ?? '');
  if (!words) return null;
  const exact = keys().get(words);
  if (exact) return exact;
  let found = null;
  for (const [e, key] of keys()) {
    if (!e.startsWith(words)) continue;
    if (found) return null;
    found = key;
  }
  return found;
}

/** The key's two parts as the window marks them: { scope, name }, or null. */
export function keyParts(english) {
  const key = keyFor(english);
  if (!key) return null;
  const dot = key.lastIndexOf('.');
  return { scope: key.slice(0, dot), name: key.slice(dot + 1) };
}

/**
 * A predicate, as JavaScript source for a window, that holds for an element
 * carrying the key of the label with this English, as its own label
 * (data-scope, data-key) or as its tip (data-tip-scope, data-tip-key: a check
 * may name a button by its tip, as the spelling check names Change):
 * `(n) => ...`. Words that name no label give one that holds for none.
 */
export function labelKeyed(english) {
  const p = keyParts(english);
  if (!p) return '(() => false)';
  const scope = JSON.stringify(p.scope);
  const name = JSON.stringify(p.name);
  return `((n) => Boolean(n && n.dataset && ((n.dataset.scope === ${scope} && n.dataset.key === ${name}) || (n.dataset.tipScope === ${scope} && n.dataset.tipKey === ${name}))))`;
}

/** A CSS selector for the elements carrying that label's key, as their label or their tip, or null. */
export function keySelector(english) {
  const p = keyParts(english);
  if (!p) return null;
  const scope = JSON.stringify(p.scope);
  const name = JSON.stringify(p.name);
  return `[data-scope=${scope}][data-key=${name}], [data-tip-scope=${scope}][data-tip-key=${name}]`;
}
