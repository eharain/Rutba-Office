// The message catalogue: a window's words looked up by their English, filled,
// counted by the language's own plural rules, English wherever a catalogue
// has no word yet — and the list a translator works from kept current.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { t, tn, msg, setLanguage, language, languages, registerCatalogue, untranslated } from '@rutba/office-ui/messages';
import { CATALOGUES } from '@rutba/office-ui/catalogues';
import { extract, render, OUTPUT } from '../tools/extract-messages.mjs';

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
  assert.equal(written, render(messages), `${OUTPUT} is out of date: run node tools/extract-messages.mjs`);
  assert.ok(messages.length > 150, 'the shell and the shared panes go through the catalogue');
  for (const want of ['Minimise', 'Save changes to {name}?', 'Change All', 'That password is not right — passwords are case-sensitive.']) {
    assert.ok(messages.some((m) => m.message === want), want);
  }
  assert.deepEqual(messages.find((m) => m.message === '{count} word'), { message: '{count} word', plural: '{count} words', where: ['packages/office-ui/src/proofing.js'] });
  const known = new Set(messages.map((m) => m.message));
  for (const [tag, catalogue] of Object.entries(CATALOGUES)) {
    for (const [message, translation] of Object.entries(catalogue)) {
      assert.ok(known.has(message), `${tag}: "${message}" is not a message the windows show`);
      const names = (s) => [...String(s).matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort().join(',');
      for (const form of typeof translation === 'string' ? [translation] : Object.values(translation)) {
        assert.equal(names(form).replace('count', '') || '', names(message).replace('count', '') || '', `${tag}: "${message}" keeps its placeholders`);
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
