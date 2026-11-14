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
import { extract, extractParts, render, made, keepLabels, readLabels, readSources, OUTPUT, LABELS, PARTS } from '../tools/extract-messages.mjs';
import { isKey, partsOf, namesFor, nameIn, spaceOf, scopeOf, widthOf } from '../tools/label-keys.mjs';
import { codeLabels, keyOf, hashOf } from '../packages/office-shell/src/electron/language-packs.js';
import { LOCKER_WORDS } from '../apps/desktop/renderer/apps/locker-translations.js';
import { TALK_WORDS } from '../apps/desktop/renderer/apps/talk-translations.js';
import { COLOUR_WORDS } from '../apps/desktop/renderer/apps/colour-translations.js';
import { SHEETS_WORDS } from '../apps/desktop/renderer/apps/sheets-translations.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
/** The source files, read once for every test here that lists the labels. */
let sources;
const src = () => (sources ||= readSources(ROOT));

test('in English a message is itself, its names filled and a name with no value left as written', () => {
  setLanguage('en-GB');
  assert.equal(language(), 'en');
  assert.equal(t('Close'), 'Close');
  assert.equal(t('Save changes to {name}?', { name: 'Budget.xlsx' }), 'Save changes to Budget.xlsx?');
  assert.equal(t('Released {date}.'), 'Released {date}.');
  assert.equal(tn(1, '{count} word', '{count} words'), '1 word');
  assert.equal(tn(2500, '{count} word', '{count} words'), '2,500 words');
  assert.equal(msg('Errors'), 'Errors', 'a mark, not a lookup');
  assert.equal(msg('ui.proofing.errors'), 'Errors', "named by its key, the key's English");
  assert.equal(t(msg('common.close')), 'Close');
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

test('--check finds a key no label defines wherever a call names it: a tn() whose count is any expression, a t() with a space before its key', () => {
  const labels = readLabels(ROOT);
  const sources = [{ rel: 'apps/desktop/renderer/apps/calendar.js', text: [
    "const a = tn(Number(form.reminder), 'calendar.count-minute-before-renamed', { at });",
    "const b = tn(list.filter((x) => x.on).length, 'mail.no-such-count');",
    "const c = t( 'common.no-such-label');",
    "const d = msg(\n  'common.no-such-message');",
    "const e = tn(Number(form.reminder), 'calendar.count-minute-before', { at });",
  ].join('\n') }];
  const now = made(ROOT, [...src(), ...sources]);
  assert.deepEqual(now.unknown.map((u) => u.key).sort(), ['calendar.count-minute-before-renamed', 'common.no-such-label', 'common.no-such-message', 'mail.no-such-count']);
  assert.ok(labels['calendar.count-minute-before'], 'a defined key is not named');
  assert.ok(now.unknown.every((u) => u.file === 'apps/desktop/renderer/apps/calendar.js'));
});

test('the list of messages is current, every catalogue keys only messages on it, and a placeholder kept', () => {
  const messages = extract(ROOT, readLabels(ROOT), src());
  const now = made(ROOT, src());
  assert.deepEqual(now.unknown, [], 'every key a call names is defined');
  for (const [rel, text] of Object.entries(now.files)) {
    const written = fs.readFileSync(path.join(ROOT, rel), 'utf8').replace(/\r\n/g, '\n');
    assert.equal(written, text, `${rel} is out of date: run node tools/extract-messages.mjs`);
  }
  assert.equal(now.files[OUTPUT], render(messages, extractParts(ROOT, readLabels(ROOT), src())));
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
  const parts = extractParts(ROOT, readLabels(ROOT), src());
  assert.deepEqual(Object.keys(parts).sort(), PARTS.map((p) => p.part).sort());
  const shared = extract(ROOT, readLabels(ROOT), src());
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

test('a label named by its meaning key reads as its English does: the same table, the same names, the same plural forms', async () => {
  const { default: KEYS } = await import('../packages/office-ui/src/catalogues/keys.js');
  const key = (english) => Object.keys(KEYS).find((k) => KEYS[k] === english);
  const counted = Object.keys(KEYS).find((k) => Array.isArray(KEYS[k]) && KEYS[k][0] === '{count} word');
  assert.ok(key('Close') && counted, 'the keys are defined');
  setLanguage('en');
  assert.equal(t(key('Close')), 'Close');
  assert.equal(t(key('Save changes to {name}?'), { name: 'Budget.xlsx' }), 'Save changes to Budget.xlsx?');
  assert.equal(tn(1, counted), '1 word');
  assert.equal(tn(2500, counted, { unused: 'x' }), '2,500 words');
  registerCatalogue('fr', { Close: 'Fermer', '{count} word': { one: '{count} mot', other: '{count} mots' } });
  try {
    setLanguage('fr');
    assert.equal(t(key('Close')), 'Fermer');
    assert.equal(tn(3, counted), '3 mots');
    assert.equal(t('Close'), 'Fermer', 'a call by English is the same label');
    assert.equal(t('not.a-defined-key'), 'not.a-defined-key', 'a key with no definition is shown as written');
  } finally {
    setLanguage('en');
  }
});

test('the words t and tn gave lately name the key they came from, in its scope and name, so an element can carry them', async () => {
  const { labelOf } = await import('@rutba/office-ui/messages');
  setLanguage('en');
  assert.equal(t('common.save'), 'Save');
  assert.deepEqual(labelOf('Save'), { key: 'common.save', scope: 'common', name: 'save' });
  // A call by English names the same key.
  assert.equal(t('Close'), 'Close');
  assert.deepEqual(labelOf('Close'), { key: 'common.close', scope: 'common', name: 'close' });
  // Counted, and a scope of more than one part.
  assert.equal(tn(4, 'common.count-word'), '4 words');
  assert.equal(labelOf('4 words')?.key, 'common.count-word');
  assert.equal(tn(5, '{count} word', '{count} words'), '5 words');
  assert.equal(labelOf('5 words')?.key, 'common.count-word');
  assert.equal(t('ui.proofing.errors'), 'Errors');
  assert.deepEqual(labelOf('Errors'), { key: 'ui.proofing.errors', scope: 'ui.proofing', name: 'errors' });
  // In another language, the words it shows name the key.
  registerCatalogue('fr', { Save: 'Enregistrer' });
  try {
    setLanguage('fr');
    assert.equal(t('common.save'), 'Enregistrer');
    assert.equal(labelOf('Enregistrer')?.key, 'common.save');
  } finally {
    setLanguage('en');
  }
  // Words no key gave, words never given, and words two keys give have none.
  assert.equal(t('Words that are no label at all'), 'Words that are no label at all');
  assert.equal(labelOf('Words that are no label at all'), null);
  assert.equal(labelOf('Never shown by t'), null);
  assert.equal(labelOf(''), null);
  registerCatalogue('fr', { Save: 'Pareil', Close: 'Pareil' });
  try {
    setLanguage('fr');
    t('common.save');
    t('common.close');
    assert.equal(labelOf('Pareil'), null, 'two keys, one text: no key rather than a wrong one');
  } finally {
    setLanguage('en');
  }
  // Bounded, the one given longest ago going first: a label drawn again keeps its key.
  const { default: KEYS } = await import('../packages/office-ui/src/catalogues/keys.js');
  const named = Object.keys(KEYS).find((k) => KEYS[k] === 'Save changes to {name}?');
  t(named, { name: 'first.docx' });
  t('common.save');
  for (let i = 0; i < 13000; i += 1) {
    t(named, { name: `${i}.docx` });
    if (i % 1000 === 0) t('common.save');
  }
  assert.equal(labelOf('Save')?.key, 'common.save', 'drawn again and again, kept');
  assert.equal(labelOf('Save changes to first.docx?'), null, 'given once, long ago: gone');
  assert.equal(labelOf('Save changes to 12999.docx?')?.key, named);
  // A label given once and then only read, as a component marking it reads it
  // each time it draws, keeps its key too: a read counts as use.
  t('common.close');
  for (let i = 0; i < 13000; i += 1) {
    t(named, { name: `again-${i}.docx` });
    if (i % 1000 === 0) assert.equal(labelOf('Close')?.key, 'common.close', `read at ${i}`);
  }
  assert.equal(labelOf('Close')?.key, 'common.close', 'read again and again, kept');
  assert.equal(labelOf('Save changes to 5.docx?'), null);
});

test('every label has a meaning key: scope and short name, its English, its space, and the hash a pack records', () => {
  const { labels, count } = JSON.parse(fs.readFileSync(path.join(ROOT, LABELS), 'utf8'));
  const keys = Object.keys(labels);
  assert.equal(keys.length, count);
  const used = codeLabels(JSON.parse(fs.readFileSync(path.join(ROOT, OUTPUT), 'utf8')));
  assert.equal(keys.length, used.length, "one key for each label the code uses, the parts' among them");
  const english = new Set(keys.map((k) => keyOf({ message: labels[k].english, plural: labels[k].plural })));
  assert.deepEqual(used.filter((m) => !english.has(keyOf(m))).map((m) => m.message), []);
  for (const k of keys) {
    const d = labels[k];
    assert.ok(isKey(k), `${k} is written as a key`);
    const { scope, name } = partsOf(k);
    assert.ok(scope && name, k);
    assert.ok(name.length <= 64, `${k}: the name is short`);
    assert.ok(d.space.short <= d.space.medium && d.space.medium <= d.space.complete, `${k}: short, medium, complete in order`);
    assert.equal(d.source, hashOf(keyOf({ message: d.english, plural: d.plural })), `${k}: the hash of its English`);
  }
  assert.ok(keys.filter((k) => labels[k].part === 'locker').length >= 250, "the Locker's own are marked as its part");
  // The tables hold a label by its English until every call names its key
  // (language-packs.js tableOf), so two labels with the same English, one of
  // them counted, would overwrite each other there.
  const byEnglish = new Map();
  for (const k of keys) byEnglish.set(labels[k].english, [...(byEnglish.get(labels[k].english) || []), k]);
  assert.deepEqual([...byEnglish].filter(([, ks]) => ks.length > 1), [], 'no two labels share their English');
});

test("a key is made from the first words of a label's meaning, short enough to fit, and unique in its scope", () => {
  assert.ok(isKey('locker.names-apply-later') && isKey('word.ribbon.page-borders') && isKey('common.close'));
  for (const no of ['Close', 'close', 'Save As', 'e.g.', 'www.', '.close', 'common.', 'Common.close', 'common.Close', 'a..b']) assert.ok(!isKey(no), no);
  assert.deepEqual(partsOf('word.ribbon.page-borders'), { scope: 'word.ribbon', name: 'page-borders' });
  assert.equal(namesFor('Close')[0], 'close');
  assert.equal(namesFor('Save As')[0], 'save-as', 'three words or fewer keep every word');
  assert.equal(namesFor('Horizontal In')[0], 'horizontal-in');
  assert.equal(namesFor('A change to names applies to files locked from now on.', 24)[0], 'change-names-applies');
  assert.equal(namesFor('Saved {name}')[0], 'saved-name', 'a placeholder is a word of the name');
  assert.equal(namesFor('Résumé')[0], 'resume');
  assert.equal(namesFor('{a} × {b}')[0], 'a-b');
  assert.equal(namesFor('…')[0], 'label');
  const taken = new Set(['common.save']);
  assert.equal(nameIn('common', 'Save', taken), 'save-2', 'taken: the next free number');
  assert.equal(nameIn('common', 'Save the copy elsewhere now please', taken, 12), 'save-copy');
  assert.equal(nameIn('common', 'Save the copy elsewhere', taken, 12), 'save-copy-elsewhere', 'taken: one word more');
  assert.equal(nameIn('word', 'Save', taken), 'save', 'another scope is free');
  assert.equal(scopeOf(['apps/desktop/renderer/apps/locker.js']), 'locker');
  assert.equal(scopeOf(['apps/desktop/renderer/apps/word-ribbon.js']), 'word.ribbon');
  assert.equal(scopeOf(['apps/desktop/renderer/apps/locker.js', 'packages/office-ui/src/index.js']), 'common');
  assert.deepEqual(spaceOf('Close'), { short: 8, medium: 15, complete: 23 });
  assert.ok(spaceOf('{count} file', '{count} files').short >= 8, 'a short keeps its placeholders');
});

test('the definitions follow the code: a new label gets a key, one no longer used goes, a kept key keeps its name and space', () => {
  const before = {
    'common.close': { english: 'Close', space: { short: 9, medium: 15, complete: 30 }, source: 'x', was: 'common.close' },
    'common.gone': { english: 'Gone', space: { short: 8, medium: 15, complete: 23 }, source: 'y' },
  };
  const messages = [{ message: 'Close', where: ['packages/office-ui/src/index.js'] }, { message: 'Open the file', where: ['apps/desktop/renderer/apps/word.js'] }];
  const parts = { locker: [{ message: 'Lock', where: ['apps/desktop/renderer/apps/locker.js'] }, { message: 'Close', where: ['apps/desktop/renderer/apps/locker.js'] }] };
  const { labels, added, retired } = keepLabels(before, messages, parts);
  assert.deepEqual(Object.keys(labels).sort(), ['common.close', 'locker.lock', 'word.open-the-file']);
  assert.deepEqual(added.sort(), ['locker.lock', 'word.open-the-file']);
  assert.deepEqual(retired, ['common.gone']);
  assert.deepEqual(labels['common.close'].space, before['common.close'].space, 'the space as written');
  assert.equal(labels['common.close'].was, 'common.close');
  assert.equal(labels['common.close'].part, undefined, 'shared, so no part');
  assert.equal(labels['locker.lock'].part, 'locker');
  assert.equal(labels['common.close'].source, hashOf(keyOf({ message: 'Close' })));
});
