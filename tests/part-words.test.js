// The parts that keep their words apart (the Locker, Talk's newer words, the
// colour palette and Sheets' Find) carry no translations of their own since
// D-2026-11-14-16: their words, as every label's, are a language pack's. A
// part's files still take t from its own words module, so its labels carry
// the part's name in the definitions (catalogues/labels.json) and a build
// without the part does not count them (language-packs.js codeLabels).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { PARTS } from '../tools/extract-messages.mjs';

const read = (rel) => fs.readFileSync(new URL(`../apps/desktop/renderer/apps/${rel}`, import.meta.url), 'utf8');
const { labels } = JSON.parse(fs.readFileSync(new URL('../packages/office-ui/src/catalogues/labels.json', import.meta.url), 'utf8'));

/** Each part's files, which take t (and tn) from its words module and not from the shared catalogue. */
const FILES = {
  colour: ['colour/palette.js'],
  locker: ['locker.js', 'locker-tile.js'],
  'sheets-find': ['sheets/find.js'],
  talk: ['talk/invite-email.js', 'talk/connect.js', 'talk/address.js'],
};

test("each part's files take their words from its words module, and the module carries no translations", () => {
  assert.deepEqual(Object.keys(FILES).sort(), PARTS.map((p) => p.part).sort());
  for (const { part, words } of PARTS) {
    for (const f of FILES[part]) {
      const text = read(f);
      assert.match(text, new RegExp(`from '\\.\\.?/${words.replace('.', '\\.')}'`), `${f} takes its words from ${words}`);
      assert.doesNotMatch(text, /import\s*\{[^}]*\bt\b[^}]*\}\s*from\s*'@rutba\/office-ui'/, `${f} takes no t() from the shared catalogue`);
    }
    const module = read(words);
    assert.doesNotMatch(module, /addToCatalogue|translations\.js/, `${words} adds no words of its own`);
    assert.ok(!fs.existsSync(new URL(`../apps/desktop/renderer/apps/${words.replace('-words.js', '-translations.js')}`, import.meta.url)), `${part}: its translations file is gone`);
  }
});

test("each part's own labels carry its name in the definitions", () => {
  const count = (part) => Object.values(labels).filter((d) => d.part === part).length;
  assert.ok(count('locker') >= 250, 'the Locker');
  assert.ok(count('talk') >= 60, 'Talk');
  assert.ok(count('colour') >= 10, 'the colour palette');
  assert.ok(count('sheets-find') >= 7, "Sheets' Find");
});

test('the Locker still adds its keys, so a build without it carries none of them', () => {
  const module = read('locker-words.js');
  assert.match(module, /addKeys\(LOCKER_KEYS\);/);
  assert.match(module, /import \{ LOCKER_KEYS \} from '\.\/locker-keys\.js';/);
});
