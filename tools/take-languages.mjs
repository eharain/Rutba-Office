// Takes the packages repository's list of languages into Office (L4).
//
//   node tools/take-languages.mjs <rutba-office-packages>/languages/languages.json
//
// Every language there is, with its native name, script, direction and
// region, so Voices and languages can offer them before any pack is here.
// At run time the feed's language entries add to it and correct it
// (office-shell electron/language-packs.js knownLanguages), so a new language
// needs no core release; this copy is the list an offline Office has.
// Written to packages/office-shell/src/languages.json.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const OUT = path.join(ROOT, 'packages', 'office-shell', 'src', 'languages.json');
const FIELDS = ['tag', 'pack', 'name', 'nativeName', 'script', 'direction', 'region', 'variantOf'];

/** The rows as Office keeps them, or an error naming the first bad one. */
export function languageRows(list) {
  const rows = Array.isArray(list?.languages) ? list.languages : null;
  if (!rows?.length) throw new Error('That file has no languages.');
  const seen = new Set();
  return rows.map((r, i) => {
    const where = `row ${i + 1} (${r?.tag})`;
    if (typeof r?.tag !== 'string' || !/^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$/.test(r.tag)) throw new Error(`${where}: the tag is not a language tag.`);
    if (seen.has(r.tag.toLowerCase())) throw new Error(`${where}: listed twice.`);
    seen.add(r.tag.toLowerCase());
    if (r.pack !== `lang-${r.tag.toLowerCase()}`) throw new Error(`${where}: its pack should be lang-${r.tag.toLowerCase()}.`);
    if (!r.name || !r.nativeName) throw new Error(`${where}: it needs its name and its native name.`);
    if (r.direction !== 'ltr' && r.direction !== 'rtl') throw new Error(`${where}: the direction is ltr or rtl.`);
    return Object.fromEntries(FIELDS.filter((k) => r[k] !== undefined).map((k) => [k, r[k]]));
  });
}

export const render = (rows) => `${JSON.stringify({
  about: 'Every language there is a pack for, from rutba-office-packages languages/languages.json, taken by tools/take-languages.mjs. The feed adds to it at run time.',
  languages: rows,
}, null, 2)}\n`;

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const from = process.argv[2];
  if (!from) {
    console.error('node tools/take-languages.mjs <rutba-office-packages>/languages/languages.json');
    process.exit(2);
  }
  try {
    const rows = languageRows(JSON.parse(fs.readFileSync(from, 'utf8')));
    fs.writeFileSync(OUT, render(rows));
    console.log(`${path.relative(ROOT, OUT)}: ${rows.length} languages`);
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}
