// The message catalogue: a window's words looked up by their English, filled,
// counted by the language's own plural rules, English wherever a catalogue
// has no word yet — and the list a translator works from kept current.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { t, tn, msg, tFilled, setLanguage, language, languages, registerCatalogue, untranslated } from '@rutba/office-ui/messages';
import { CATALOGUES } from '@rutba/office-ui/catalogues';
import { extract, extractParts, render, OUTPUT, PARTS } from '../tools/extract-messages.mjs';
import { codeLabels, keyOf } from '../packages/office-shell/src/electron/language-packs.js';
import { LOCKER_WORDS } from '../apps/desktop/renderer/apps/locker-translations.js';
import { TALK_WORDS } from '../apps/desktop/renderer/apps/talk-translations.js';
import { COLOUR_WORDS } from '../apps/desktop/renderer/apps/colour-translations.js';
import { SHEETS_WORDS } from '../apps/desktop/renderer/apps/sheets-translations.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('in English a message is itself, its names filled and a name with no value left as written', () => {
  setLanguage('en-GB');
  assert.equal(language(), 'en');
  assert.equal(t('Close'), 'Close');
  assert.equal(t('Save changes to {name}?', { name: 'Budget.xlsx' }), 'Save changes to Budget.xlsx?');
  assert.equal(t('Released {date}.'), 'Released {date}.');
  assert.equal(tn(1, '{count} word', '{count} words'), '1 word');
  assert.equal(tn(2500, '{count} word', '{count} words'), '2,500 words');
  assert.equal(msg('Errors'), 'Errors', 'a mark, not a lookup');
});

test('a catalogue translates what it has, its base language serves a regional tag, and the rest shows in English', () => {
  registerCatalogue('fr', {
    Close: 'Fermer',
    'Save changes to {name}?': 'Enregistrer les modifications de {name} ?',
    '{count} word': { one: '{count} mot', other: '{count} mots' },
  });
  try {
    assert.equal(setLanguage('fr-CA'), 'fr');
    assert.ok(languages().includes('fr'));
    assert.equal(t('Close'), 'Fermer');
    assert.equal(t('Save changes to {name}?', { name: 'Budget.xlsx' }), 'Enregistrer les modifications de Budget.xlsx ?');
    assert.equal(t('Minimise'), 'Minimise', 'no word yet: English');
    assert.deepEqual(untranslated(), ['Minimise'], 'and noted for the translator');
    assert.equal(tn(1, '{count} word', '{count} words'), '1 mot');
    assert.equal(tn(0, '{count} word', '{count} words'), '0 mot', 'French counts nought as one');
    assert.equal(tn(1500, '{count} word', '{count} words').replace(/\s/g, ' '), '1 500 mots', 'the number as French writes it');
    assert.equal(setLanguage('de'), 'en', 'a language with no catalogue is English');
  } finally {
    setLanguage('en');
  }
});

test('plural forms by the language\'s own rules: Arabic\'s two, few and many', () => {
  registerCatalogue('ar', { '{count} word': { zero: 'لا كلمات', one: 'كلمة واحدة', two: 'كلمتان', few: '{count} كلمات', many: '{count} كلمة', other: '{count} كلمة' } });
  try {
    setLanguage('ar');
    assert.equal(tn(0, '{count} word', '{count} words'), 'لا كلمات');
    assert.equal(tn(2, '{count} word', '{count} words'), 'كلمتان');
    assert.match(tn(5, '{count} word', '{count} words'), /كلمات$/);
    assert.match(tn(11, '{count} word', '{count} words'), /كلمة$/);
  } finally {
    setLanguage('en');
  }
});

test('the list of messages is current, every catalogue keys only messages on it, and a placeholder kept', () => {
  const messages = extract(ROOT);
  const written = fs.readFileSync(path.join(ROOT, OUTPUT), 'utf8').replace(/\r\n/g, '\n');
  assert.equal(written, render(messages, extractParts(ROOT)), `${OUTPUT} is out of date: run node tools/extract-messages.mjs`);
  assert.ok(messages.length > 150, 'the shell and the shared panes go through the catalogue');
  for (const want of ['Minimise', 'Save changes to {name}?', 'Change All', 'That password is not right — passwords are case-sensitive.']) {
    assert.ok(messages.some((m) => m.message === want), want);
  }
  const counted = messages.find((m) => m.message === '{count} word');
  assert.equal(counted.plural, '{count} words');
  assert.ok(counted.where.includes('packages/office-ui/src/proofing.js'), 'a counted message lists the files that use it');
  const known = new Set(messages.map((m) => m.message));
  for (const [tag, catalogue] of Object.entries(CATALOGUES)) {
    for (const [message, translation] of Object.entries(catalogue)) {
      assert.ok(known.has(message), `${tag}: "${message}" is not a message the windows show`);
      // {count} may be left out of a form that says the number in words (Arabic's one, two).
      const names = (s) => [...String(s).matchAll(/\{(\w+)\}/g)].map((m) => m[1]).filter((n) => n !== 'count').sort().join(',');
      for (const form of typeof translation === 'string' ? [translation] : Object.values(translation)) {
        assert.equal(names(form), names(message), `${tag}: "${message}" keeps its placeholders`);
      }
    }
  }
});

test('a system locale written with an underscore is the same language as its hyphenated tag', () => {
  // Linux hands out pt_BR; Intl.PluralRules and toLocaleString refuse that
  // spelling, so a counted message in such a window used to throw.
  registerCatalogue('pt_BR', { '{count} word': { one: '{count} palavra', other: '{count} palavras' } });
  try {
    assert.equal(setLanguage('pt_BR'), 'pt-br');
    assert.equal(tn(1, '{count} word', '{count} words'), '1 palavra');
    assert.equal(tn(3, '{count} word', '{count} words'), '3 palavras');
  } finally {
    setLanguage('en');
  }
});

test("an engine's message, thrown already filled, finds its translation as it is or by its pattern", () => {
  registerCatalogue('xx', {
    'The password is not right.': 'Le mot de passe est faux.',
    'There is no paragraph style "{name}".': 'Il n’y a pas de style « {name} ».',
    'Row {row} of "{sheet}" is past the end.': 'La ligne {row} de « {sheet} » est après la fin.',
    'Row {row} of {what}.': 'Ligne {row} de {what}.',
    '{a} × {b}': '{b} × {a}',
  });
  try {
    setLanguage('xx');
    assert.equal(tFilled('The password is not right.'), 'Le mot de passe est faux.');
    assert.equal(tFilled('There is no paragraph style "Quote".'), 'Il n’y a pas de style « Quote ».');
    // The most literal pattern wins over a looser one the message also fits.
    assert.equal(tFilled('Row 12 of "Budget" is past the end.'), 'La ligne 12 de « Budget » est après la fin.');
    // A message of names and marks alone is no pattern: it would reorder any sentence with a × in it.
    assert.equal(tFilled('Paper 3 × 5'), 'Paper 3 × 5');
    assert.equal(tFilled('Something no catalogue knows.'), 'Something no catalogue knows.');
    assert.equal(tFilled(''), '');
    setLanguage('en');
    assert.equal(tFilled('There is no paragraph style "Quote".'), 'There is no paragraph style "Quote".', 'English is left as it came');
  } finally {
    setLanguage('en');
  }
});

test('each part that keeps its words apart has its own list of labels, read from the files that take its t', () => {
  const parts = extractParts(ROOT);
  assert.deepEqual(Object.keys(parts).sort(), PARTS.map((p) => p.part).sort());
  const shared = extract(ROOT);
  const sharedWords = new Set(shared.map((m) => m.message));
  const own = (name) => parts[name].filter((m) => !sharedWords.has(m.message));
  // Their own labels when lang-en 2.26.0 r1 was built: locker 260, talk 63, colour 10, sheets 7.
  assert.ok(own('locker').length >= 250, 'the Locker');
  assert.ok(own('talk').length >= 60, 'Talk');
  assert.ok(own('colour').length >= 10, 'the colour palette');
  assert.ok(own('sheets-find').length >= 7, "Sheets' Find");
  assert.ok(parts.locker.find((m) => m.message === 'Lock')?.where.includes('apps/desktop/renderer/apps/locker.js'), 'a label lists the files that use it');
  for (const [name, list] of Object.entries(parts)) {
    for (const m of list) assert.ok(m.where.every((f) => f.startsWith('apps/desktop/renderer/')), `${name}: "${m.message}" is used only in a part's own files`);
  }
  // Every word a part translates for itself is on its list: what the packs word for it is what Office counts.
  const TRANSLATED = { locker: LOCKER_WORDS, talk: TALK_WORDS, colour: COLOUR_WORDS, 'sheets-find': SHEETS_WORDS };
  for (const [name, byTag] of Object.entries(TRANSLATED)) {
    const listed = new Set(parts[name].map((m) => m.message));
    const words = new Set(Object.values(byTag).flatMap((w) => Object.keys(w)));
    assert.deepEqual([...words].filter((w) => !listed.has(w)), [], `${name}: translated but not on its list`);
  }
});

test("Office's labels are the shared ones and each part's, every key once, and a build without a part leaves its labels out", () => {
  const list = {
    messages: [{ message: 'Close' }, { message: '{count} file', plural: '{count} files' }],
    parts: {
      locker: { labels: [{ message: 'Close' }, { message: 'Lock' }, { message: '{count} file', plural: '{count} files' }] },
      talk: { labels: [{ message: 'Call' }, { message: '{count} file', plural: 'some files' }] },
    },
  };
  assert.deepEqual(codeLabels(list).map(keyOf), ['Close', keyOf({ message: '{count} file', plural: '{count} files' }), 'Lock', 'Call', keyOf({ message: '{count} file', plural: 'some files' })]);
  assert.deepEqual(codeLabels(list, { leaveOut: ['locker'] }).map((m) => m.message), ['Close', '{count} file', 'Call', '{count} file']);
  assert.deepEqual(codeLabels({}), []);
  assert.deepEqual(codeLabels({ messages: [{ message: 'Close' }, null, { text: 'x' }] }).map((m) => m.message), ['Close']);

  // The file as written: every part's labels are counted, and none twice.
  const written = JSON.parse(fs.readFileSync(path.join(ROOT, OUTPUT), 'utf8'));
  const all = codeLabels(written);
  assert.equal(new Set(all.map(keyOf)).size, all.length);
  const withoutLocker = codeLabels(written, { leaveOut: ['locker'] });
  assert.ok(all.length - withoutLocker.length >= 250, 'the Locker\'s own labels go with it');
  assert.ok(withoutLocker.length > written.messages.length, 'Talk, the palette and Find stay');
});
