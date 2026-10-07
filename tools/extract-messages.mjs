// The message catalogue's list: every message the windows pass to t(), tn()
// and msg(), in English, with the files that use it.
//
//   node tools/extract-messages.mjs           writes packages/office-ui/src/catalogues/messages.json
//   node tools/extract-messages.mjs --check   says whether that file is current (exit 1 when not)
//
// Only files that import t, tn or msg from @rutba/office-ui (or its own
// messages.js) are read, so another function called t is never mistaken for
// the catalogue's. A message is found only where it is written out as a
// literal string in the call — which is why the code always writes it so.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const SOURCES = ['packages/office-ui/src', 'apps/desktop/renderer'];
export const OUTPUT = 'packages/office-ui/src/catalogues/messages.json';

const STR = String.raw`'((?:[^'\\\n]|\\.)*)'|"((?:[^"\\\n]|\\.)*)"`;
const IMPORTS = /import\s*\{[^}]*\b(?:t|tn|msg)\b[^}]*\}\s*from\s*['"](?:@rutba\/office-ui|\.\/messages\.js|\.\.\/messages\.js)['"]/;
const unescape = (s) => s.replace(/\\(n|.)/g, (m, c) => (c === 'n' ? '\n' : c));

function files(dir, out = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    if (entry.isDirectory()) files(p, out);
    else if (entry.name.endsWith('.js')) out.push(p);
  }
  return out;
}

/** Every message, sorted: { message, plural?, where: [files] }. */
export function extract(root = ROOT) {
  const found = new Map();
  const note = (message, file, plural = null) => {
    const entry = found.get(message) || { message, ...(plural ? { plural } : {}), where: new Set() };
    if (plural && !entry.plural) entry.plural = plural;
    entry.where.add(file);
    found.set(message, entry);
  };
  for (const dir of SOURCES) {
    for (const file of files(path.join(root, dir))) {
      const text = fs.readFileSync(file, 'utf8');
      if (!IMPORTS.test(text)) continue;
      const rel = path.relative(root, file).split(path.sep).join('/');
      for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])(?:t|msg)\(\s*(?:${STR})`, 'g'))) note(unescape(m[1] ?? m[2]), rel);
      for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])tn\([^'"]*?,\s*(?:${STR})\s*,\s*(?:${STR})`, 'g'))) note(unescape(m[1] ?? m[2]), rel, unescape(m[3] ?? m[4]));
    }
  }
  return [...found.values()]
    .sort((a, b) => (a.message < b.message ? -1 : a.message > b.message ? 1 : 0))
    .map((e) => ({ ...e, where: [...e.where].sort() }));
}

/** The file as written: the messages, and a line on what it is. */
export function render(messages) {
  return JSON.stringify({
    about: 'Every message the windows show, in English, and the files that use it — what a catalogue translates. Made by tools/extract-messages.mjs; see catalogues/index.js.',
    count: messages.length,
    messages,
  }, null, 2) + '\n';
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const text = render(extract());
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
    console.log(`${OUTPUT}: ${JSON.parse(text).count} messages`);
  }
}
