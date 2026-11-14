// The message catalogue's list: every message the windows pass to t(), tn()
// and msg(), in English, with the files that use it.
//
// The parts of the suite that keep their words apart from the shared
// catalogues (the Locker, Talk's newer words, the colour palette and Sheets'
// Find: PARTS below) are listed too, each as its own list under `parts`, read
// from the files that take t and tn from that part's words module. A pack
// words them as it words every other label, so Office counts them as labels
// it uses (language-packs.js codeLabels): a build without a part leaves its
// list out.
//
//   node tools/extract-messages.mjs           writes packages/office-ui/src/catalogues/messages.json
//   node tools/extract-messages.mjs --check   says whether that file is current (exit 1 when not)
//
// Only files that import t, tn or msg from @rutba/office-ui (or its own
// messages.js, or the shell's words.js, which hands the main process the
// same) are read, so another function called t is never mistaken for the
// catalogue's. A message is found only where it is written out as a
// literal string in the call — which is why the code always writes it so.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { keyOf, hashOf } from '../packages/office-shell/src/electron/language-packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const SOURCES = ['packages/office-ui/src', 'apps/desktop/renderer', 'packages/office-shell/src/electron'];
export const OUTPUT = 'packages/office-ui/src/catalogues/messages.json';
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
const unescape = (s) => s.replace(/\\(n|.)/g, (m, c) => (c === 'n' ? '\n' : c));

function files(dir, out = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    if (entry.isDirectory()) files(p, out);
    else if (entry.name.endsWith('.js')) out.push(p);
  }
  return out;
}

/**
 * Every message, sorted: { message, plural?, source, where: [files] }.
 * `source` is the key's hash, as a language pack's entries hold it (L1), so
 * a window finds its words without hashing every key as it starts.
 */
export function extract(root = ROOT) {
  return listOf(root, (text) => IMPORTS.test(text));
}

/**
 * Each part's own labels, the same shape as extract()'s, by part:
 * { colour: [...], locker: [...], ... }. A label a part shares with the
 * catalogue is on both lists; codeLabels joins them once.
 */
export function extractParts(root = ROOT) {
  const out = {};
  for (const p of PARTS) {
    const words = p.words.replace(/[.-]/g, (c) => `\\${c}`);
    const from = new RegExp(String.raw`import\s*\{[^}]*\b(?:t|tn)\b[^}]*\}\s*from\s*['"](?:\.\.?\/)+(?:[\w-]+\/)*${words}['"]`);
    out[p.part] = listOf(root, (text) => from.test(text), ['t', 'tn']);
  }
  return out;
}

/** The messages passed to t(), tn() and msg() in the files `take` accepts. */
function listOf(root, take, calls = ['t', 'tn', 'msg']) {
  const found = new Map();
  const note = (message, file, plural = null) => {
    const entry = found.get(message) || { message, ...(plural ? { plural } : {}), where: new Set() };
    if (plural && !entry.plural) entry.plural = plural;
    entry.where.add(file);
    found.set(message, entry);
  };
  const single = calls.filter((c) => c !== 'tn').join('|');
  for (const dir of SOURCES) {
    for (const file of files(path.join(root, dir))) {
      const text = fs.readFileSync(file, 'utf8');
      if (!take(text)) continue;
      const rel = path.relative(root, file).split(path.sep).join('/');
      for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])(?:${single})\(\s*(?:${STR})`, 'g'))) note(unescape(m[1] ?? m[2]), rel);
      for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])tn\([^'"]*?,\s*(?:${STR})\s*,\s*(?:${STR})`, 'g'))) note(unescape(m[1] ?? m[2]), rel, unescape(m[3] ?? m[4]));
    }
  }
  return [...found.values()]
    .sort((a, b) => (a.message < b.message ? -1 : a.message > b.message ? 1 : 0))
    .map((e) => ({ message: e.message, ...(e.plural ? { plural: e.plural } : {}), source: hashOf(keyOf(e)), where: [...e.where].sort() }));
}

/** The file as written: the messages, each part's own, and a line on what it is. */
export function render(messages, parts = {}) {
  return JSON.stringify({
    about: 'Every message the windows show, in English, and the files that use it: what a catalogue translates. Under parts, the labels of each part of the suite that keeps its words apart (its own translations, added with addToCatalogue), which a language pack words like any other. Made by tools/extract-messages.mjs; see catalogues/index.js.',
    count: messages.length,
    messages,
    parts: Object.fromEntries(Object.entries(parts).map(([name, labels]) => [name, { count: labels.length, labels }])),
  }, null, 2) + '\n';
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const text = render(extract(), extractParts());
  const out = path.join(ROOT, OUTPUT);
  if (process.argv.includes('--check')) {
    const current = fs.existsSync(out) ? fs.readFileSync(out, 'utf8').replace(/\r\n/g, '\n') : '';
    if (current !== text) {
      console.error(`${OUTPUT} is out of date: run node tools/extract-messages.mjs`);
      process.exit(1);
    }
    console.log(`${OUTPUT} is current`);
  } else {
    fs.writeFileSync(out, text);
    console.log(`${OUTPUT}: ${JSON.parse(text).count} messages, ${Object.entries(JSON.parse(text).parts).map(([n, p]) => `${n} ${p.count}`).join(', ')} in the parts`);
  }
}
