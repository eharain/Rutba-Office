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
import { extract, extractParts, made, keepLabels, readLabels, readSources, englishCalls, LABELS, PARTS } from '../tools/extract-messages.mjs';
import { isKey, partsOf, namesFor, nameIn, spaceOf, scopeOf, widthOf } from '../tools/label-keys.mjs';
import { codeLabels, keyOf, hashOf } from '../packages/office-shell/src/electron/language-packs.js';

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

test('a table is by key whether it names its labels by key or by English, a key winning, and a key added later moves its English entries (D-2026-11-14-16, stage 5)', async () => {
  const { addKeys, registerPack, setFallback, setVersions, versionsOf } = await import('@rutba/office-ui/messages');
  registerCatalogue('zk', {
    'common.save': 'by key',
    Save: 'by English, under the key',
    Close: 'close by English',
    'common.count-word': { one: '{count} key word', other: '{count} key words' },
    'engine.no-paragraph-style': 'no style {style} here',
    'Words no label is': 'kept as they are',
    'Only the Part says this {name}': 'part {name} by English',
  });
  setLanguage('zk');
  try {
    assert.equal(t('common.save'), 'by key', 'the entry named by its key');
    assert.equal(t('Save'), 'by key', 'a call by English finds it under the key');
    assert.equal(t('common.close'), 'close by English', 'an entry named by its English is under its key');
    assert.equal(tn(2, 'common.count-word'), '2 key words');
    assert.equal(tn(1, '{count} word', '{count} words'), '1 key word');
    // An engine's message, written in English, fits the label's English and takes the key's words.
    assert.equal(tFilled('There is no paragraph style "Quote" to link a level to.'), 'no style Quote here');
    assert.equal(t('Words no label is'), 'kept as they are', 'words that name no label stay as given');
    // A part's keys arrive after its table: the table's English entry moves to the new key.
    assert.equal(t('Only the Part says this {name}', { name: 'x' }), 'part x by English');
    addKeys({ 'zz.part-only': 'Only the Part says this {name}' });
    assert.equal(t('zz.part-only', { name: 'y' }), 'part y by English');
    // A pack above, by key, wins; lang-en's table and the versions are by key too.
    registerPack('zk', { Save: 'pack save' });
    assert.equal(t('common.save'), 'pack save');
    setFallback({ 'Saved {name}': 'Stored {name}' });
    assert.equal(t('common.saved-name', { name: 'A' }), 'Stored A', 'a label the language lacks, in lang-en words by key');
    setVersions({ Close: { short: 'cl' } }, null);
    assert.deepEqual(versionsOf(t('common.close')), { short: 'cl' });
  } finally {
    setFallback(null);
    setVersions(null, null);
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
    'const f = t(`common.no-such-template`);',
    "const g = tn(fn(a, 'common.save'), 'common.no-such-after-a-call');",
  ].join('\n') }];
  const now = made(ROOT, [...src(), ...sources]);
  assert.deepEqual(now.unknown.map((u) => u.key).sort(), ['calendar.count-minute-before-renamed', 'common.no-such-after-a-call', 'common.no-such-label', 'common.no-such-message', 'common.no-such-template', 'mail.no-such-count']);
  assert.ok(labels['calendar.count-minute-before'], 'a defined key is not named');
  assert.ok(now.unknown.every((u) => u.file === 'apps/desktop/renderer/apps/calendar.js'));
});

test('--check refuses a call that names its label by its English, with its file and line; a key, or words that arrive as data, it takes (stage 5)', () => {
  const sources = [{ rel: 'apps/desktop/renderer/apps/calendar.js', text: [
    "const a = t('Close');",
    "const b = tn(n, '{count} word', '{count} words');",
    "const c = msg(\n  'Saved {name}');",
    "const d = t('common.close');",
    "const e = t(name);",
    "const f = tn(Number(x), 'common.count-word');",
    // In a template with nothing filled in, or with a space before the bracket: English all the same.
    'const g = t(`Open`);',
    "const h = t ('Print');",
    // A count that holds a string of its own: the label is the one after it, not the count's.
    "const i = tn(sizeOf(a, 'x'), 'Item', 'Items');",
    "const j = tn(sizeOf(a, 'x'), 'common.count-word');",
    // A template that fills something in is words made at run time, no literal.
    'const k = t(`Saved ${name}`);',
  ].join('\n') }];
  const now = made(ROOT, [...src(), ...sources]);
  const at = (words, line) => ({ words, file: 'apps/desktop/renderer/apps/calendar.js', line });
  assert.deepEqual(now.english, [at('Close', 1), at('{count} word', 2), at('Saved {name}', 3), at('Open', 8), at('Print', 9), at('Item', 10)]);
  assert.deepEqual(englishCalls(ROOT, src()), [], 'the code names every label by its key');
  const tool = fs.readFileSync(path.join(ROOT, 'tools/extract-messages.mjs'), 'utf8');
  assert.match(tool, /if \(out\.english\.length\) process\.exit\(1\);/);
});

test('the list of messages is current, every catalogue keys only messages on it, and a placeholder kept', () => {
  const messages = extract(ROOT, readLabels(ROOT), src());
  const now = made(ROOT, src());
  assert.deepEqual(now.unknown, [], 'every key a call names is defined');
  for (const [rel, text] of Object.entries(now.files)) {
    const written = fs.readFileSync(path.join(ROOT, rel), 'utf8').replace(/\r\n/g, '\n');
    assert.equal(written, text, `${rel} is out of date: run node tools/extract-messages.mjs`);
  }
  assert.ok(!fs.existsSync(path.join(ROOT, 'packages/office-ui/src/catalogues/messages.json')), 'the definitions are the list: messages.json is gone (stage 5)');
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
});

test("Office's labels are its definitions, each key with its English and hash, and a build without a part leaves its labels out", () => {
  const defs = {
    labels: {
      'common.close': { english: 'Close', source: 'aaaaaaaaaaaa' },
      'common.count-file': { english: '{count} file', plural: '{count} files' },
      'locker.lock': { english: 'Lock', part: 'locker' },
      'talk.call': { english: 'Call', part: 'talk' },
      'common.broken': { space: {} },
    },
  };
  assert.deepEqual(codeLabels(defs), [
    { key: 'common.close', message: 'Close', source: 'aaaaaaaaaaaa' },
    { key: 'common.count-file', message: '{count} file', plural: '{count} files' },
    { key: 'locker.lock', message: 'Lock' },
    { key: 'talk.call', message: 'Call' },
  ]);
  assert.deepEqual(codeLabels(defs, { leaveOut: ['locker'] }).map((m) => m.key), ['common.close', 'common.count-file', 'talk.call']);
  assert.deepEqual(codeLabels({}), []);
  assert.deepEqual(codeLabels(null), []);

  // The file as written: every label once, the Locker's own going with it.
  const written = JSON.parse(fs.readFileSync(path.join(ROOT, LABELS), 'utf8'));
  const all = codeLabels(written);
  assert.equal(all.length, written.count);
  assert.equal(new Set(all.map(keyOf)).size, all.length, 'no two labels with the same English');
  const withoutLocker = codeLabels(written, { leaveOut: ['locker'] });
  assert.ok(all.length - withoutLocker.length >= 250, "the Locker's own labels go with it");
  assert.ok(withoutLocker.some((m) => written.labels[m.key].part === 'talk'), 'Talk, the palette and Find stay');
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
  assert.deepEqual(labelOf('Save'), { key: 'common.save', scope: 'common', name: 'save', bare: false });
  // A call by English names the same key.
  assert.equal(t('Close'), 'Close');
  assert.deepEqual(labelOf('Close'), { key: 'common.close', scope: 'common', name: 'close', bare: false });
  // Counted, and a scope of more than one part.
  assert.equal(tn(4, 'common.count-word'), '4 words');
  assert.equal(labelOf('4 words')?.key, 'common.count-word');
  assert.equal(tn(5, '{count} word', '{count} words'), '5 words');
  assert.equal(labelOf('5 words')?.key, 'common.count-word');
  assert.equal(t('ui.proofing.errors'), 'Errors');
  assert.deepEqual(labelOf('Errors'), { key: 'ui.proofing.errors', scope: 'ui.proofing', name: 'errors', bare: false });
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

test('a bare text gives each element the key of its own call, in the order of the calls, and only in the task that made it (labelOfCall)', async () => {
  const { labelOfCall, setBare } = await import('@rutba/office-ui/messages');
  setLanguage('en');
  setBare(true);
  try {
    assert.equal(t('common.close'), 'close');
    assert.equal(t('mail.close'), 'close');
    assert.equal(t('word.ribbon.close'), 'close');
    assert.equal(labelOfCall('close').key, 'common.close');
    assert.equal(labelOfCall('close').key, 'mail.close');
    assert.deepEqual(labelOfCall('close'), { key: 'word.ribbon.close', scope: 'word.ribbon', name: 'close', bare: true });
    assert.equal(labelOfCall('close'), null, 'every call taken: the words alone are no label');
    t('common.close');
    await new Promise((resolve) => setTimeout(resolve, 5));
    assert.equal(labelOfCall('close'), null, 'a call no element took marks nothing once its task is over');
  } finally {
    setBare(null);
  }
  // Words a pack gives are read by their words, as labelOf reads them.
  assert.equal(t('common.save'), 'Save');
  assert.deepEqual(labelOfCall('Save'), { key: 'common.save', scope: 'common', name: 'save', bare: false });
  assert.deepEqual(labelOfCall('Save'), { key: 'common.save', scope: 'common', name: 'save', bare: false });
});

test('with no pack wording a label, it is shown by its key\'s name with spaces for its hyphens, and named bare (D-2026-11-14-16 (2))', async () => {
  const { labelOf, setBare, setFallback, bareName } = await import('@rutba/office-ui/messages');
  setLanguage('en');
  try {
    // No English pack at all: every label.
    setBare(true);
    assert.equal(t('word.ribbon.multiple'), 'multiple');
    assert.equal(t('common.save'), 'save');
    assert.equal(t('Close'), 'close', 'a call by English names the same key');
    assert.equal(tn(4, 'common.count-word'), 'count word', 'a counted label by its name alone');
    assert.deepEqual(labelOf('count word'), { key: 'common.count-word', scope: 'common', name: 'count-word', bare: true });
    assert.equal(t('Words that are no label at all'), 'Words that are no label at all', 'words no key gave are shown as given');
    assert.equal(bareName('ui.proofing.change-this-one'), 'change this one');
    // A table that words it still wins: another language's, and lang-en's own wording.
    registerCatalogue('fr', { Save: 'Enregistrer' });
    setLanguage('fr');
    assert.equal(t('common.save'), 'Enregistrer');
    assert.equal(t('common.close'), 'close', 'a label French lacks, with no English pack, bare');
    setFallback({ 'common.close': 'Close the window' });
    assert.equal(t('common.close'), 'Close the window', 'lang-en\'s words when it words the label otherwise');
    setFallback(null);
    setLanguage('en');
    // The keys an English pack lacks, and only those.
    setBare(['common.close']);
    assert.equal(t('common.save'), 'Save');
    assert.equal(t('common.close'), 'close');
    assert.equal(labelOf('Save').bare, false);
    // Not told (a test, a window before its words): the definitions' English.
    setBare(null);
    assert.equal(t('common.close'), 'Close');
  } finally {
    setBare(null);
    setFallback(null);
    setLanguage('en');
  }
  // The windows and the main process are told which, from the words the shell made.
  const window = fs.readFileSync(path.join(ROOT, 'apps/desktop/renderer/language.js'), 'utf8').replace(/\r\n/g, '\n');
  assert.match(window, /setBare\(words \? words\.bare \?\? null : null\)/);
  const main = fs.readFileSync(path.join(ROOT, 'packages/office-shell/src/electron/words.js'), 'utf8').replace(/\r\n/g, '\n');
  assert.match(main, /setBare\(w\.bare \?\? null\)/);
  assert.match(main, /\{ tag: null, table: null, english: null, bare: true,/, 'no words at all: every label bare');
  // The source failing is no reason to show every label bare: English, as the definitions give it.
  assert.match(main, /const IN_ENGLISH = Object\.freeze\(\{ tag: null, table: null, english: null, bare: null \}\)/);
  assert.match(main, /: failed \? \{ \.\.\.IN_ENGLISH,/);
  assert.match(main, /if \(failed\) \{\n {4}setFallback\(null\);\n {4}setBare\(null\);/);
});

test('every label has a meaning key: scope and short name, its English, its space, and the hash a pack records', () => {
  const { labels, count } = JSON.parse(fs.readFileSync(path.join(ROOT, LABELS), 'utf8'));
  const keys = Object.keys(labels);
  assert.equal(keys.length, count);
  // What the code uses: the shared labels and each part's, every one once.
  const usedIds = new Set([...extract(ROOT, readLabels(ROOT), src()), ...Object.values(extractParts(ROOT, readLabels(ROOT), src())).flat()].map(keyOf));
  const used = [...usedIds].map((id) => (id.startsWith('{"one":') ? (({ one, other }) => ({ message: one, plural: other }))(JSON.parse(id)) : { message: id }));
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
