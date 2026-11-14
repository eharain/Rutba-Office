// Run once (D-2026-11-14-16): gives every label Office uses its meaning key
// and writes the keys' definitions, packages/office-ui/src/catalogues/labels.json.
//
//   node tools/make-label-keys.mjs [<packages labels.json>]
//
// For each label in catalogues/messages.json (the shared messages and the
// parts' own):
//   - its scope is the scope of the id the packages repository gave it, found
//     in Office's English copy (lang-en) by the label's hash, else one made
//     from the files that use it (label-keys.mjs scopeOf);
//   - its name is new and short (label-keys.mjs nameIn), made from its
//     English, unique in its scope;
//   - its space is the one the packages repository measured for it
//     (languages/labels.json, default beside this checkout), else one made by
//     the same rule (spaceOf);
//   - `was` keeps the packages id it had, which the packages repository's own
//     tool maps to the new key in every pack, so nothing is translated again.
//
// After this, the definitions are kept by tools/extract-messages.mjs, and a
// key's name or space may be set by hand there.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { codeLabels, keyOf } from '../packages/office-shell/src/electron/language-packs.js';
import { nameIn, scopeOf, spaceOf, KEY } from './label-keys.mjs';
import { renderLabels, LABELS } from './extract-messages.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (file) => JSON.parse(fs.readFileSync(file, 'utf8'));

export function makeLabelKeys({ list, copy, spaces }) {
  const idBySource = new Map(Object.entries(copy).map(([id, e]) => [e.source, id]));
  const shared = new Set(list.messages.map(keyOf));
  const partOf = new Map();
  for (const [part, p] of Object.entries(list.parts || {})) for (const m of p.labels) if (!shared.has(keyOf(m)) && !partOf.has(keyOf(m))) partOf.set(keyOf(m), part);
  const labels = codeLabels(list)
    .map((m) => ({ m, was: idBySource.get(m.source) || null }))
    .sort((a, b) => a.m.message.length - b.m.message.length || (a.m.message < b.m.message ? -1 : a.m.message > b.m.message ? 1 : 0) || String(a.m.plural || '').localeCompare(String(b.m.plural || '')));
  const taken = new Set();
  const out = {};
  for (const { m, was } of labels) {
    const scope = was && KEY.test(was) ? was.slice(0, was.lastIndexOf('.')) : scopeOf(m.where || []);
    const space = (was && spaces[was]?.space) || spaceOf(m.message, m.plural);
    const name = nameIn(scope, m.message, taken, Math.max(10, Math.min(space.short, 24)));
    const part = partOf.get(keyOf(m));
    out[`${scope}.${name}`] = {
      english: m.message,
      ...(m.plural ? { plural: m.plural } : {}),
      ...(part ? { part } : {}),
      space: { short: space.short, medium: space.medium, complete: space.complete },
      source: m.source,
      ...(was ? { was } : {}),
    };
  }
  return out;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const list = read(path.join(ROOT, 'packages', 'office-ui', 'src', 'catalogues', 'messages.json'));
  const copy = read(path.join(ROOT, 'apps', 'desktop', 'resources', 'languages', 'lang-en', 'messages.json'));
  const spacesFile = process.argv[2] || path.join(ROOT, '..', 'Rutba-office-packages', 'languages', 'labels.json');
  const spaces = fs.existsSync(spacesFile) ? read(spacesFile).labels : {};
  const labels = makeLabelKeys({ list, copy, spaces });
  fs.writeFileSync(path.join(ROOT, LABELS), renderLabels(labels));
  const n = Object.keys(labels).length;
  const named = Object.values(labels).filter((l) => l.was).length;
  const measured = Object.values(labels).filter((l) => l.was && spaces[l.was]).length;
  console.log(`${LABELS}: ${n} keys, ${named} with the packages id they had, ${measured} with the space the packages repository measured`);
}
