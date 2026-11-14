// Moves calls that name a label by its English to its meaning key
// (D-2026-11-14-10 to -16), one area at a time:
//
//   node tools/to-keys.mjs <file or folder> [...]        rewrites them
//   node tools/to-keys.mjs --dry <file or folder> [...]  says what it would change
//
// t('Close')                       becomes t('common.close')
// msg('Close')                     becomes msg('common.close')
// tn(n, '{count} file', '{count} files', v)  becomes  tn(n, 'common.count-file', v)
//
// The key is the one catalogues/labels.json defines for that English (and
// plural). Only files that take t, tn or msg from the catalogue or from a
// part's words module are touched, so another function called t is left
// alone. A literal with no definition is left as it is and reported: run
// tools/extract-messages.mjs first, which gives every new label its key.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { keyOf } from '../packages/office-shell/src/electron/language-packs.js';
import { isKey } from './label-keys.mjs';
import { readLabels, PARTS } from './extract-messages.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const STR = String.raw`'((?:[^'\\\n]|\\.)*)'|"((?:[^"\\\n]|\\.)*)"`;
const unescape = (s) => s.replace(/\\(n|.)/g, (m, c) => (c === 'n' ? '\n' : c));
const IMPORTS = new RegExp(String.raw`import\s*\{[^}]*\b(?:t|tn|msg)\b[^}]*\}\s*from\s*['"](?:@rutba\/office-ui(?:\/messages)?|\.\/messages\.js|\.\.\/messages\.js|\.\/words\.js|(?:\.\.?\/)+(?:[\w-]+\/)*(?:${PARTS.map((p) => p.words.replace(/[.-]/g, (c) => `\\${c}`)).join('|')}))['"]`);

/** Whether a file's t, tn and msg are the catalogue's. */
export const usesCatalogue = (text) => IMPORTS.test(text);

/**
 * A file's text with its calls by English moved to keys.
 * @returns {{ text: string, moved: number, unknown: string[] }}
 */
export function toKeys(text, labels) {
  const byEnglish = new Map(Object.entries(labels).map(([k, d]) => [keyOf({ message: d.english, plural: d.plural }), k]));
  let moved = 0;
  const unknown = [];
  // Counted first: its two strings become one key.
  let out = text.replace(new RegExp(String.raw`(?<![\w.$])tn\(([^'"]*?),\s*(?:${STR})\s*,\s*(?:${STR})`, 'g'), (whole, count, a1, a2, b1, b2) => {
    const one = unescape(a1 ?? a2);
    if (isKey(one)) return whole;
    const key = byEnglish.get(keyOf({ message: one, plural: unescape(b1 ?? b2) }));
    if (!key) {
      unknown.push(one);
      return whole;
    }
    moved += 1;
    return `tn(${count}, '${key}'`;
  });
  out = out.replace(new RegExp(String.raw`(?<![\w.$])(t|msg)\((\s*)(?:${STR})`, 'g'), (whole, fn, space, a1, a2) => {
    const english = unescape(a1 ?? a2);
    if (isKey(english)) return whole;
    const key = byEnglish.get(keyOf({ message: english }));
    if (!key) {
      unknown.push(english);
      return whole;
    }
    moved += 1;
    return `${fn}(${space}'${key}'`;
  });
  return { text: out, moved, unknown };
}

function filesUnder(p, out = []) {
  const stat = fs.statSync(p);
  if (stat.isFile()) {
    if (p.endsWith('.js')) out.push(p);
    return out;
  }
  for (const e of fs.readdirSync(p, { withFileTypes: true })) {
    if (e.name === 'node_modules') continue;
    filesUnder(path.join(p, e.name), out);
  }
  return out;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2);
  const dry = args.includes('--dry');
  const targets = args.filter((a) => a !== '--dry');
  if (!targets.length) {
    console.error('usage: node tools/to-keys.mjs [--dry] <file or folder> [...]');
    process.exit(2);
  }
  const labels = readLabels(ROOT);
  let files = 0;
  let calls = 0;
  const unknown = new Set();
  for (const target of targets) {
    for (const file of filesUnder(path.resolve(ROOT, target))) {
      const text = fs.readFileSync(file, 'utf8');
      if (!usesCatalogue(text)) continue;
      const r = toKeys(text, labels);
      for (const u of r.unknown) unknown.add(`${path.relative(ROOT, file)}: ${u}`);
      if (!r.moved) continue;
      files += 1;
      calls += r.moved;
      if (!dry) fs.writeFileSync(file, r.text);
    }
  }
  console.log(`${dry ? 'would move' : 'moved'} ${calls} calls in ${files} files to their keys`);
  if (unknown.size) {
    console.error(`${unknown.size} literals have no key (run node tools/extract-messages.mjs first):`);
    for (const u of [...unknown].slice(0, 20)) console.error(`  ${u}`);
    process.exit(1);
  }
}
