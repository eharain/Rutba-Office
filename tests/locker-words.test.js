// The layers of words a window's table is made of: a language's catalogue, a
// part's own words above it (addToCatalogue), an installed pack's above all,
// whichever is registered first (D-2026-11-14-16).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { t, setLanguage, registerCatalogue, addToCatalogue, registerPack } from '../packages/office-ui/src/messages.js';

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

test("a window registers a pack's own words alone, so a part's words stay above Office's built-in ones and a label the pack lacks keeps them", () => {
  // As renderer/language.js does it: the built-in catalogue first, the part's
  // words as it loads, the pack's table as the main process made it from the pack.
  registerCatalogue('zw', { Close: 'built-in close', 'Try again': 'built-in try again', 'New in this Office': 'built-in new' });
  addToCatalogue('zw', { 'Try again': 'locker try again' });
  registerPack('zw', { Close: 'pack close' });
  setLanguage('zw');
  try {
    assert.equal(t('Close'), 'pack close');
    assert.equal(t('Try again'), 'locker try again', "the part's own word, not the built-in one the pack does not word");
    assert.equal(t('New in this Office'), 'built-in new', 'a label new in this Office keeps the built-in word');
  } finally {
    setLanguage('en');
  }
  const language = fs.readFileSync(new URL('../apps/desktop/renderer/language.js', import.meta.url), 'utf8');
  assert.match(language, /registerPack\(words\.tag, words\.table\)/, "the pack's table as it came");
  assert.doesNotMatch(language, /registerPack\([^)]*CATALOGUES/, 'never one with the built-in catalogue inside it');
  // The main process lays a pack the same way: over the built-in words, not in their place.
  const main = fs.readFileSync(new URL('../packages/office-shell/src/electron/words.js', import.meta.url), 'utf8');
  assert.match(main, /registerPack\(w\.tag, w\.table\)/);
  assert.doesNotMatch(main, /registerCatalogue\(w\.tag/);
});
