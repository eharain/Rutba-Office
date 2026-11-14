// The labels' definitions, kept to the code: every label the windows pass to
// t(), tn() and msg(), by its meaning key.
//
// The parts of the suite that keep their words apart (the Locker, Talk's
// newer words, the colour palette and Sheets' Find: PARTS below) are read
// from the files that take t and tn from that part's words module, and their
// own labels carry the part's name. A pack words them as it words every
// other label, so Office counts them as labels it uses (language-packs.js
// codeLabels): a build without a part leaves its labels out.
//
// A call names its label by its meaning key (`t('common.close')`,
// `tn(count, 'mail.messages-count', values)`). The keys' definitions are in
// catalogues/labels.json (D-2026-11-14-10 to -16): each key's English, its
// plural, the part that carries it and its space. This tool keeps them: a
// label new in the code, written in English, gets a key (label-keys.mjs), and
// a key no call uses any more goes. A call with a key that has no definition
// is refused, and --check refuses a call that names its label by its English:
// node tools/to-keys.mjs moves it to the key this tool gave it. catalogues/
// keys.js is the keys' English, which finds the key of words that arrive as
// data (a template's name) and is shown until a window is told its packs.
//
//   node tools/extract-messages.mjs           writes labels.json, keys.js and locker-keys.js
//   node tools/extract-messages.mjs --check   says whether they are current (exit 1 when not)
//
// Only files that import t, tn or msg from @rutba/office-ui (or its own
// messages.js, or the shell's words.js, which hands the main process the
// same) are read, so another function called t is never mistaken for the
// catalogue's. A message is found only where it is written out as a
// literal string in the call, which is why the code always writes it so.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { keyOf, hashOf } from '../packages/office-shell/src/electron/language-packs.js';
import { isKey, nameIn, scopeOf, spaceOf } from './label-keys.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const SOURCES = ['packages/office-ui/src', 'apps/desktop/renderer', 'packages/office-shell/src/electron'];
export const LABELS = 'packages/office-ui/src/catalogues/labels.json';
export const KEYS_MODULE = 'packages/office-ui/src/catalogues/keys.js';
export const LOCKER_KEYS_MODULE = 'apps/desktop/renderer/apps/locker-keys.js';
/**
 * The parts whose words are kept apart (each adds its own translations with
 * addToCatalogue): its name, as the packages repository's extractor names it,
 * and the words module its files take t and tn from.
 */
export const PARTS = [
  { part: 'colour', words: 'colour-words.js' },
  { part: 'locker', words: 'locker-words.js' },
  { part: 'sheets-find', words: 'sheets-words.js' },
  { part: 'talk', words: 'talk-words.js' },
];

const STR = String.raw`'((?:[^'\\\n]|\\.)*)'|"((?:[^"\\\n]|\\.)*)"`;
const IMPORTS = /import\s*\{[^}]*\b(?:t|tn|msg)\b[^}]*\}\s*from\s*['"](?:@rutba\/office-ui(?:\/messages)?|\.\/messages\.js|\.\.\/messages\.js|\.\/words\.js)['"]/;
/**
 * A call's label as a literal: t() and msg() name it first, after any spaces;
 * tn() names it after its count, which may be any expression
 * (Number(form.reminder)), so its shape is the one listOf reads a tn() label by.
 */
const CALLS = [String.raw`(?<![\w.$])(?:t|msg)\(\s*(?:${STR})`, String.raw`(?<![\w.$])tn\([^'"]*?,\s*(?:${STR})`];
const unescape = (s) => s.replace(/\\(n|.)/g, (m, c) => (c === 'n' ? '\n' : c));

function files(dir, out = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    if (entry.isDirectory()) files(p, out);
    else if (entry.name.endsWith('.js')) out.push(p);
  }
  return out;
}

/** Every source file once: [{ rel, text }]. made() reads them once for all its lists. */
export function readSources(root = ROOT) {
  const out = [];
  for (const dir of SOURCES) {
    for (const file of files(path.join(root, dir))) out.push({ rel: path.relative(root, file).split(path.sep).join('/'), text: fs.readFileSync(file, 'utf8') });
  }
  return out;
}

/** The keys' definitions as written: { key: { english, plural?, part?, space, source, was? } }. */
export function readLabels(root = ROOT) {
  const file = path.join(root, LABELS);
  return fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')).labels || {} : {};
}

/**
 * Every message, sorted: { message, plural?, source, where: [files] }.
 * `source` is the key's hash, as a language pack's entries hold it (L1), so
 * a window finds its words without hashing every key as it starts. A call by
 * key is listed by its definition's English.
 */
export function extract(root = ROOT, labels = readLabels(root), sources = readSources(root)) {
  return listOf(sources, (text) => IMPORTS.test(text), ['t', 'tn', 'msg'], labels);
}

/**
 * Each part's own labels, the same shape as extract()'s, by part:
 * { colour: [...], locker: [...], ... }. A label a part shares with the
 * catalogue is on both lists; keepLabels joins them once.
 */
export function extractParts(root = ROOT, labels = readLabels(root), sources = readSources(root)) {
  const out = {};
  for (const p of PARTS) {
    const words = p.words.replace(/[.-]/g, (c) => `\\${c}`);
    const from = new RegExp(String.raw`import\s*\{[^}]*\b(?:t|tn)\b[^}]*\}\s*from\s*['"](?:\.\.?\/)+(?:[\w-]+\/)*${words}['"]`);
    out[p.part] = listOf(sources, (text) => from.test(text), ['t', 'tn'], labels);
  }
  return out;
}

/** The keys calls name that have no definition: [{ key, file }]. */
export function unknownKeys(root = ROOT, labels = readLabels(root), sources = readSources(root)) {
  const out = [];
  for (const { rel, text } of sources) {
    for (const shape of CALLS) {
      for (const m of text.matchAll(new RegExp(shape, 'g'))) {
        const s = unescape(m[1] ?? m[2]);
        if (isKey(s) && !labels[s]) out.push({ key: s, file: rel });
      }
    }
  }
  return out;
}

/**
 * The calls that name their label by its English rather than its key
 * (D-2026-11-14-16, stage 5): [{ words, file, line }], as unknownKeys reads
 * a call's label. Words that arrive as data (t(name)) are no literal, and so
 * none of these.
 */
export function englishCalls(root = ROOT, sources = readSources(root)) {
  const out = [];
  for (const { rel, text } of sources) {
    for (const shape of CALLS) {
      for (const m of text.matchAll(new RegExp(shape, 'g'))) {
        const s = unescape(m[1] ?? m[2]);
        if (!isKey(s)) out.push({ words: s, file: rel, line: text.slice(0, m.index).split('\n').length });
      }
    }
  }
  return out;
}

/** The messages passed to t(), tn() and msg() in the files `take` accepts. */
function listOf(sources, take, calls, labels) {
  const found = new Map();
  const note = (message, file, plural = null) => {
    const id = keyOf({ message, plural });
    const entry = found.get(id) || { message, ...(plural ? { plural } : {}), where: new Set() };
    entry.where.add(file);
    found.set(id, entry);
  };
  const byKey = (s, file) => {
    const def = labels[s];
    if (def) note(def.english, file, def.plural || null);
  };
  const single = calls.filter((c) => c !== 'tn').join('|');
  for (const { rel, text } of sources) {
    if (!take(text)) continue;
    for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])(?:${single})\(\s*(?:${STR})`, 'g'))) {
      const s = unescape(m[1] ?? m[2]);
      if (isKey(s)) byKey(s, rel);
      else note(s, rel);
    }
    if (!calls.includes('tn')) continue;
    for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])tn\([^'"]*?,\s*(?:${STR})`, 'g'))) {
      const s = unescape(m[1] ?? m[2]);
      if (isKey(s)) byKey(s, rel);
    }
    for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])tn\([^'"]*?,\s*(?:${STR})\s*,\s*(?:${STR})`, 'g'))) {
      const one = unescape(m[1] ?? m[2]);
      if (!isKey(one)) note(one, rel, unescape(m[3] ?? m[4]));
    }
  }
  return [...found.values()]
    .sort((a, b) => (a.message < b.message ? -1 : a.message > b.message ? 1 : 0) || String(a.plural || '').localeCompare(String(b.plural || '')))
    .map((e) => ({ message: e.message, ...(e.plural ? { plural: e.plural } : {}), source: hashOf(keyOf(e)), where: [...e.where].sort() }));
}

/**
 * The definitions for what the code uses now: a label kept has its key, name
 * and space as written; a label new in the code gets a key in the scope its
 * files give it; a key no call uses goes. `part` is said again each time.
 * @returns {{ labels: object, added: string[], retired: string[] }}
 */
export function keepLabels(before, messages, parts) {
  const shared = new Set(messages.map(keyOf));
  const partOf = new Map();
  for (const [part, list] of Object.entries(parts)) for (const m of list) if (!shared.has(keyOf(m)) && !partOf.has(keyOf(m))) partOf.set(keyOf(m), part);
  const keyByLabel = new Map(Object.entries(before).map(([key, d]) => [keyOf({ message: d.english, plural: d.plural }), key]));
  const taken = new Set(Object.keys(before));
  // The shared labels, then each part's, every label once.
  const used = [];
  const seen = new Set();
  for (const m of [...messages, ...Object.values(parts).flat()]) {
    if (!m || typeof m.message !== 'string' || seen.has(keyOf(m))) continue;
    seen.add(keyOf(m));
    used.push(m);
  }
  const labels = {};
  const added = [];
  for (const m of used) {
    const id = keyOf(m);
    let key = keyByLabel.get(id);
    let def = key ? before[key] : null;
    if (!key) {
      const space = spaceOf(m.message, m.plural);
      const scope = scopeOf(m.where || []);
      key = `${scope}.${nameIn(scope, m.message, taken, Math.max(10, Math.min(space.short, 24)))}`;
      def = { space };
      added.push(key);
    }
    const part = partOf.get(id);
    labels[key] = {
      english: m.message,
      ...(m.plural ? { plural: m.plural } : {}),
      ...(part ? { part } : {}),
      space: def.space,
      source: hashOf(id),
      ...(def.was ? { was: def.was } : {}),
    };
  }
  const retired = Object.keys(before).filter((k) => !labels[k]);
  return { labels, added, retired };
}

const sortedKeys = (labels) => Object.keys(labels).sort();

/**
 * JSON with the long dash and the invisible characters written as escapes
 * (a backslash-u code, as JSON allows): the same text once read, and no line a commit adds
 * carries them (the 2026-10-10 rule in .githooks/pre-commit).
 */
const MARKS = /[\u2014\u200b\u200c\u200e\u200f\u2060\u00a0\ufeff]/g;
const json = (value) => JSON.stringify(value).replace(MARKS, (c) => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`);

/** The definitions as written: one key per line, sorted, so a change is one line in a diff. */
export function renderLabels(labels) {
  const rows = sortedKeys(labels).map((k) => `  ${json(k)}: ${json(labels[k])}`);
  return [
    '{',
    `"about": ${JSON.stringify('Every label Office shows, by its meaning key (scope.name; the scope is never shown): its English, from which lang-en is made, its plural, the part that carries it, its space (the most display columns its short, medium and complete versions may take, a wide character 2 and a placeholder 8), the hash a pack entry records, and the packages id it had. Kept by tools/extract-messages.mjs; a name or a space may be set by hand. D-2026-11-14-10 to -16.')},`,
    `"count": ${rows.length},`,
    '"labels": {',
    rows.join(',\n'),
    '}',
    '}',
  ].join('\n') + '\n';
}

/** The keys' English for the windows, while calls by English remain (messages.js). The Locker's are in a module of its own. */
export function renderKeysModule(labels) {
  const rows = rowsOf(labels, (d) => d.part !== 'locker');
  return [
    "// Each meaning key's English (catalogues/labels.json), for t() and tn() while",
    '// some calls still name a label by its English. A part a build can leave out',
    '// has its keys in a module of its own. Made by tools/extract-messages.mjs.',
    'export default {',
    ...rows,
    '};',
  ].join('\n') + '\n';
}

/** The Locker's keys and their English, loaded by locker-words.js (addKeys) only in a build that has the Locker. */
export function renderLockerKeysModule(labels) {
  return [
    "// The Locker's meaning keys and their English (catalogues/labels.json), added",
    '// to the catalogue by locker-words.js, so a build without the Locker carries',
    '// none of them. Made by tools/extract-messages.mjs.',
    'export const LOCKER_KEYS = {',
    ...rowsOf(labels, (d) => d.part === 'locker'),
    '};',
  ].join('\n') + '\n';
}

const rowsOf = (labels, take) => sortedKeys(labels).filter((k) => take(labels[k])).map((k) => `  ${json(k)}: ${json(labels[k].plural ? [labels[k].english, labels[k].plural] : labels[k].english)},`);

/** Everything the tool writes, by file, made from the code and the definitions now. */
export function made(root = ROOT, sources = readSources(root)) {
  const before = readLabels(root);
  const unknown = unknownKeys(root, before, sources);
  const english = englishCalls(root, sources);
  const messages = extract(root, before, sources);
  const parts = extractParts(root, before, sources);
  const kept = keepLabels(before, messages, parts);
  return {
    unknown,
    english,
    added: kept.added,
    retired: kept.retired,
    files: { [LABELS]: renderLabels(kept.labels), [KEYS_MODULE]: renderKeysModule(kept.labels), [LOCKER_KEYS_MODULE]: renderLockerKeysModule(kept.labels) },
  };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const out = made();
  if (out.unknown.length) {
    for (const u of out.unknown.slice(0, 20)) console.error(`${u.file}: ${u.key} is not a defined key (${LABELS})`);
    process.exit(1);
  }
  const byEnglish = (u) => `${u.file}:${u.line}: ${JSON.stringify(u.words)} names its label by its English: run node tools/to-keys.mjs ${u.file}`;
  if (process.argv.includes('--check')) {
    for (const u of out.english.slice(0, 20)) console.error(byEnglish(u));
    if (out.english.length) process.exit(1);
    const stale = Object.entries(out.files).filter(([rel, text]) => {
      const file = path.join(ROOT, rel);
      return (fs.existsSync(file) ? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n') : '') !== text;
    });
    for (const [rel] of stale) console.error(`${rel} is out of date: run node tools/extract-messages.mjs`);
    if (stale.length) process.exit(1);
    console.log('labels.json, keys.js and locker-keys.js are current');
  } else {
    for (const [rel, text] of Object.entries(out.files)) fs.writeFileSync(path.join(ROOT, rel), text);
    for (const u of out.english.slice(0, 20)) console.log(byEnglish(u));
    console.log(`${LABELS}: ${JSON.parse(out.files[LABELS]).count} keys, ${out.added.length} new, ${out.retired.length} gone${out.added.length ? ` (new: ${out.added.slice(0, 5).join(', ')})` : ''}`);
  }
}
