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
//
// A finder says what its words matched before, and the key matches no more
// than that (the integrator's review of stage 4):
//   - by default the element's own label (data-scope, data-key), as a finder
//     that read textContent matched the words drawn; a button's tip key is
//     not its label, and the quick-access bar's icon buttons carry their
//     ribbon twins' keys as tips only;
//   - { tip: true } its tip too (data-tip-scope, data-tip-key), where the
//     finder read the title or data-tip;
//   - { start: true } words that begin exactly one label name that label,
//     where the finder matched the start of the words (startsWith); an exact
//     finder's words name a label only when they are all of its English.

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
 * The key of the label whose English is these words, or null. With
 * { start: true }, for a finder that matched the start of the words (the start
 * of a tip that runs on, as the TOA check names Update Table): words no label
 * is, but that begin exactly one label's English, give that label's key;
 * words that begin more than one give none, as a guess could be the wrong
 * control.
 */
export function keyFor(english, { start = false } = {}) {
  const words = String(english ?? '');
  if (!words) return null;
  const exact = keys().get(words);
  if (exact) return exact;
  if (!start) return null;
  let found = null;
  for (const [e, key] of keys()) {
    if (!e.startsWith(words)) continue;
    if (found) return null;
    found = key;
  }
  return found;
}

/** The key's two parts as the window marks them: { scope, name }, or null. */
export function keyParts(english, o = {}) {
  const key = keyFor(english, o);
  if (!key) return null;
  const dot = key.lastIndexOf('.');
  return { scope: key.slice(0, dot), name: key.slice(dot + 1) };
}

/**
 * A predicate, as JavaScript source for a window, that holds for an element
 * carrying the key of the label with this English as its own label
 * (data-scope, data-key), and with { tip: true } as its tip too
 * (data-tip-scope, data-tip-key: a check that named a button by its tip, as
 * the spelling check names Change): `(n) => ...`. Words that name no label
 * give one that holds for none.
 * @param {string} english
 * @param {{ tip?: boolean, start?: boolean }} [o]
 */
export function labelKeyed(english, { tip = false, start = false } = {}) {
  const p = keyParts(english, { start });
  if (!p) return '(() => false)';
  const scope = JSON.stringify(p.scope);
  const name = JSON.stringify(p.name);
  const own = `(n.dataset.scope === ${scope} && n.dataset.key === ${name})`;
  const asTip = `(n.dataset.tipScope === ${scope} && n.dataset.tipKey === ${name})`;
  return `((n) => Boolean(n && n.dataset && (${tip ? `${own} || ${asTip}` : own})))`;
}

/** A CSS selector for the elements carrying that label's key, as their label (and with { tip: true } their tip), or null. */
export function keySelector(english, { tip = false, start = false } = {}) {
  const p = keyParts(english, { start });
  if (!p) return null;
  const scope = JSON.stringify(p.scope);
  const name = JSON.stringify(p.name);
  const own = `[data-scope=${scope}][data-key=${name}]`;
  return tip ? `${own}, [data-tip-scope=${scope}][data-tip-key=${name}]` : own;
}
