// Urdu, Arabic and Hindi in the core (L7): the add-ons contract says which
// languages this release carries built in, and which have left the core for
// their packs; the bundle's catalogues are exactly the built-in ones.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { CATALOGUES } from '../packages/office-ui/src/catalogues/index.js';
import { BUILT_IN, LEFT_CORE, builtInInUse } from '../packages/office-shell/src/electron/language-packs.js';

const contract = JSON.parse(fs.readFileSync(new URL('../packages/office-shell/src/addons-contract.json', import.meta.url), 'utf8'));

test('the contract\'s built-in languages are the catalogues the bundle carries', () => {
  assert.deepEqual([...contract.languages.builtIn].sort(), Object.keys(CATALOGUES).sort());
  assert.deepEqual(BUILT_IN, contract.languages.builtIn);
  assert.deepEqual(LEFT_CORE, contract.languages.leftCore);
  for (const tag of LEFT_CORE) assert.ok(!BUILT_IN.includes(tag), `${tag} is either built in or has left`);
  assert.ok(!BUILT_IN.includes('en'), 'English is the lang-en copy, not a catalogue');
});

test('the built-in language a person uses is counted as its pack, by its base', () => {
  assert.deepEqual(builtInInUse('ur'), ['lang-ur']);
  assert.deepEqual(builtInInUse('ur-PK'), ['lang-ur']);
  assert.deepEqual(builtInInUse('ar_SA'), ['lang-ar']);
  assert.deepEqual(builtInInUse('hi-IN'), ['lang-hi']);
  assert.deepEqual(builtInInUse('en-GB'), []);
  assert.deepEqual(builtInInUse('fr'), []);
  assert.deepEqual(builtInInUse(null), []);
});

test('the menu lists the built-in languages from the catalogues, not a list of its own', () => {
  const shell = fs.readFileSync(new URL('../apps/desktop/renderer/shell.js', import.meta.url), 'utf8');
  assert.match(shell, /const BUILT_IN_LANGUAGES = \['ur', 'ar', 'hi', \.\.\.Object\.keys\(CATALOGUES\)\.sort\(\)\]\s*\.filter\(\(tag, i, all\) => tag in CATALOGUES/);
});
