// The Locker's words are translated in a file of their own, so a build without it carries none of them.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { LOCKER_WORDS } from '../apps/desktop/renderer/apps/locker-translations.js';
import ur from '../packages/office-ui/src/catalogues/ur.js';
import ar from '../packages/office-ui/src/catalogues/ar.js';
import hi from '../packages/office-ui/src/catalogues/hi.js';
import { englishOf } from './label-english.js';
import { t, setLanguage, registerCatalogue, addToCatalogue, registerPack } from '../packages/office-ui/src/messages.js';

// Words the rest of the suite uses too are in the shared catalogues.
const SHARED = { ur, ar, hi };
const STR = String.raw`'((?:[^'\\\n]|\\.)*)'`;
const unescape = (s) => s.replace(/\\(n|.)/g, (m, c) => (c === 'n' ? '\n' : c));
const used = () => {
  const out = new Set();
  for (const f of ['locker.js', 'locker-tile.js']) {
    const text = fs.readFileSync(new URL(`../apps/desktop/renderer/apps/${f}`, import.meta.url), 'utf8');
    for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])t\(\s*${STR}`, 'g'))) out.add(englishOf(unescape(m[1])));
  }
  return [...out];
};

test('every word the Locker shows is translated, and nothing is kept that it no longer shows', () => {
  const words = used();
  assert.ok(words.length > 100, 'the words were read');
  for (const [tag, list] of Object.entries(LOCKER_WORDS)) {
    assert.deepEqual(words.filter((w) => !list[w] && !SHARED[tag][w]), [], `${tag}: not translated`);
    assert.deepEqual(Object.keys(list).filter((w) => !words.includes(w)), [], `${tag}: no longer used`);
  }
});

test('words added to a language are there whichever is registered first, and the language stays current', () => {
  addToCatalogue('xx', { One: 'uno' });
  registerCatalogue('xx', { Two: 'dos' });
  setLanguage('xx');
  try {
    assert.equal(t('One'), 'uno');
    assert.equal(t('Two'), 'dos');
    addToCatalogue('xx', { Three: 'tres' });
    assert.equal(t('Three'), 'tres');
  } finally {
    setLanguage('en');
  }
});

test("an installed pack's words are above a part's own built-in words, whichever comes first (D-2026-11-14-16)", () => {
  registerCatalogue('yy', { Shared: 'built-in shared' });
  registerPack('yy', { Shared: 'pack shared', Part: 'pack part' });
  addToCatalogue('yy', { Part: 'built-in part', Own: 'built-in own' });
  setLanguage('yy');
  try {
    assert.equal(t('Shared'), 'pack shared');
    assert.equal(t('Part'), 'pack part', 'the part loads after the pack and still does not cover it');
    assert.equal(t('Own'), 'built-in own', 'a word the pack lacks keeps the built-in one');
    registerCatalogue('yy', { Shared: 'built-in again' });
    assert.equal(t('Shared'), 'pack shared', 'nor does a catalogue registered later');
  } finally {
    setLanguage('en');
  }
});
