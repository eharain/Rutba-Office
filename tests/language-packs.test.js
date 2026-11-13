// Words from a language pack (L1, L2) and Office's English copy (C15): the
// key's hash finds a label in any pack, an entry made from other English is
// never shown, English is a pack too, a regional pack lies over its base,
// and the copy is taken only as Office would use it.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { keyOf, hashOf, wordsOf, tableOf, englishOf, createLanguageWords, readEnglishCopy, directionOf, LANGUAGE_FORMATS } from '../packages/office-shell/src/electron/language-packs.js';
import { canonical, keyIdOf } from '../packages/office-shell/src/electron/packs.js';
import { takeLangEn } from '../tools/take-lang-en.mjs';
import * as M from '../packages/office-ui/src/messages.js';

const CELL = '"{address}" is not a cell or a range';
const CANCEL = { message: 'A cancellation goes to the {count} person invited, from the account the invitation went from.', plural: 'A cancellation goes to the {count} people invited, from the account the invitation went from.' };
const KEYS = [{ message: CELL }, CANCEL, { message: 'Close' }, { message: 'Saved {name}' }, { message: 'New since the pack' }];
const H = (m) => hashOf(keyOf(m));

const temp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-langpacks-'));
const write = (file, value) => {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, typeof value === 'string' ? value : JSON.stringify(value));
};

test('a key hashes as the packages build hashes it: its English, or its two forms as JSON', () => {
  // The two examples in the languages session's note, from lang-en at 0a3a0ae.
  assert.equal(hashOf(keyOf({ message: CELL })), 'e9ba0af2c14e');
  assert.equal(keyOf(CANCEL), JSON.stringify({ one: CANCEL.message, other: CANCEL.plural }));
  assert.equal(hashOf(keyOf(CANCEL)), '0051c76ed18b');
  assert.deepEqual(LANGUAGE_FORMATS, [1, 2], 'the current format and the one before, from the contract');
  // messages.json carries each key's hash, so a window start hashes nothing.
  const list = JSON.parse(fs.readFileSync(new URL('../packages/office-ui/src/catalogues/messages.json', import.meta.url), 'utf8')).messages;
  for (const m of list) assert.equal(m.source, hashOf(keyOf(m)), m.message);
});

test('an entry is shown only when it fits its key: text or forms, other present, no name the key lacks', () => {
  assert.equal(wordsOf({ message: 'Saved {name}' }, { text: '{name} enregistré', source: 'x' }), '{name} enregistré');
  assert.equal(wordsOf({ message: 'Saved {name}' }, { text: '{file} enregistré' }), null, 'a name the key does not have');
  assert.equal(wordsOf({ message: 'Close' }, { text: '' }), null, 'empty');
  assert.equal(wordsOf({ message: 'Close' }, { forms: { other: 'x' } }), null, 'forms for a plain label');
  assert.equal(wordsOf(CANCEL, { text: 'x' }), null, 'text for a counted label');
  assert.equal(wordsOf(CANCEL, { forms: { one: 'une {count}' } }), null, 'no other');
  assert.deepEqual(wordsOf(CANCEL, { forms: { zero: 'aucune', one: 'une', other: '{count} personnes', extra: 'x' } }), { zero: 'aucune', one: 'une', other: '{count} personnes' }, 'a form may leave out {count}, and unknown forms are dropped');
});

test('a pack becomes the table keyed by the English, never with a stale entry', () => {
  const hashes = KEYS.map(H);
  const entries = {
    'engine.cell': { text: '« {address} » n\'est pas une cellule', source: H({ message: CELL }) },
    'calendar.cancel': { forms: { one: 'une annulation', other: '{count} annulations' }, source: H(CANCEL) },
    'shell.close': { text: 'Fermer', source: H({ message: 'Close' }) },
    // Made from English the code no longer has: never shown (L2).
    'shell.saved': { text: '{name} enregistré', source: hashOf('Saved {name}.') },
    'gone.label': { text: 'Retiré', source: 'aaaaaaaaaaaa' },
  };
  const { table, used } = tableOf(KEYS, hashes, entries);
  assert.equal(used, 3);
  assert.deepEqual(Object.keys(table).sort(), [CELL, CANCEL.message, 'Close'].sort());
  assert.deepEqual(table[CANCEL.message], { one: 'une annulation', other: '{count} annulations' }, 'a counted label keyed by its singular, as tn() reads it');
});

test('English holds only the labels lang-en words otherwise than the code', () => {
  const hashes = KEYS.map(H);
  const en = {
    a: { text: 'Close', source: H({ message: 'Close' }) },
    b: { text: '"{address}" is not a cell or range', source: H({ message: CELL }) },
    c: { forms: { one: CANCEL.message, other: CANCEL.plural }, source: H(CANCEL) },
  };
  assert.deepEqual(englishOf(KEYS, hashes, en), { [CELL]: '"{address}" is not a cell or range' });
});

/** A language pack in an add-ons folder, as the registry lists it. */
function addon(root, manifest, entries) {
  const dir = path.join(root, manifest.id);
  write(path.join(dir, 'messages.json'), entries);
  return { id: manifest.id, kind: 'language', dir, manifest: { kind: 'language', format: 2, revision: 1, ...manifest } };
}

function englishCopy(root, revision, closeWord) {
  const manifest = { id: 'lang-en', kind: 'language', format: 2, version: '2.25.0', revision, tag: 'en', direction: 'ltr' };
  write(path.join(root, 'lang-en.pack.json'), { format: 'rutba-office-pack', keyId: 'x', manifest, signature: 'x' });
  write(path.join(root, 'lang-en', 'messages.json'), { 'shell.close': { text: closeWord, source: H({ message: 'Close' }) } });
}

test('the window\'s words: its pack, a regional one over its base, English as the copy or a newer lang-en', () => {
  const root = temp();
  try {
    const copy = path.join(root, 'resources');
    englishCopy(copy, 1, 'Close the window');
    const fr = addon(path.join(root, 'addons'), { id: 'lang-fr', tag: 'fr', name: 'French', nativeName: 'Français' }, {
      'shell.close': { text: 'Fermer', source: H({ message: 'Close' }) },
      'shell.saved': { text: '{name} enregistré', source: H({ message: 'Saved {name}' }) },
    });
    const frCa = addon(path.join(root, 'addons'), { id: 'lang-fr-ca', tag: 'fr-CA', variantOf: 'lang-fr', requires: ['lang-fr'] }, {
      'shell.close': { text: 'Fermer la fenêtre', source: H({ message: 'Close' }) },
    });
    const ur = addon(path.join(root, 'addons'), { id: 'lang-ur', tag: 'ur', direction: 'rtl' }, { 'shell.close': { text: 'بند کریں', source: H({ message: 'Close' }) } });
    const odd = addon(path.join(root, 'addons'), { id: 'lang-de', tag: 'de', format: 3 }, { 'shell.close': { text: 'Schließen', source: H({ message: 'Close' }) } });
    let installed = [fr, frCa, ur, odd];
    const lines = [];
    const words = createLanguageWords({ keys: KEYS, copyDir: copy, addons: () => installed, log: (l) => lines.push(l) });

    const en = words.wordsFor('en-GB');
    assert.equal(en.tag, 'en');
    assert.deepEqual(en.table, { Close: 'Close the window' });
    assert.equal(en.pack.from, 'copy');

    const frWords = words.wordsFor('fr');
    assert.equal(frWords.tag, 'fr');
    assert.deepEqual(frWords.table, { Close: 'Fermer', 'Saved {name}': '{name} enregistré' });
    assert.deepEqual(frWords.english, { Close: 'Close the window' }, 'a label French lacks is shown in lang-en\'s English');

    assert.deepEqual(words.wordsFor('fr-CA').table, { Close: 'Fermer la fenêtre', 'Saved {name}': '{name} enregistré' }, 'fr-CA over fr');
    assert.equal(words.wordsFor('fr_BE').tag, 'fr', 'a region with no pack takes its base');
    assert.equal(words.wordsFor('ur').direction, 'rtl');
    assert.equal(words.wordsFor('ja').tag, null, 'no pack: the window keeps its own tables');
    assert.equal(words.wordsFor('ar').direction, 'rtl', 'with no pack, by the built-in list');
    assert.equal(words.wordsFor('de').table, null, 'a format this core does not read is not used');
    assert.ok(lines.some((l) => l.startsWith('lang-de')));
    assert.deepEqual(words.languages().map((l) => l.tag), ['en', 'fr', 'fr-ca', 'ur', 'de']);

    // An added lang-en above the copy is used; one at or below it is not.
    const later = createLanguageWords({ keys: KEYS, copyDir: copy, addons: () => [addon(path.join(root, 'more'), { id: 'lang-en', tag: 'en', version: '2.25.0', revision: 2 }, { 'shell.close': { text: 'Close it', source: H({ message: 'Close' }) } })] });
    assert.deepEqual(later.wordsFor('en').table, { Close: 'Close it' });
    assert.equal(later.wordsFor('en').pack.from, 'added');
    const same = createLanguageWords({ keys: KEYS, copyDir: copy, addons: () => [addon(path.join(root, 'same'), { id: 'lang-en', tag: 'en', version: '2.25.0', revision: 1 }, { 'shell.close': { text: 'Close it', source: H({ message: 'Close' }) } })] });
    assert.equal(same.wordsFor('en').pack.from, 'copy', 'the copy is the floor');

    // No English at all: the code's own words, and nothing breaks.
    const none = createLanguageWords({ keys: KEYS, copyDir: null, addons: () => [] });
    assert.deepEqual(none.wordsFor('en').table, {});
    assert.equal(none.wordsFor('fr').english, null);
    installed = [];
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
  assert.equal(readEnglishCopy(null), null);
  assert.equal(directionOf('he'), 'rtl');
  assert.equal(directionOf('he', { direction: 'ltr' }), 'ltr', 'the pack\'s word first');
});

test('the windows read the table as before: English as lang-en words it, a missing label in that English, untranslated() kept', () => {
  M.registerCatalogue('en', { Close: 'Close the window', [CANCEL.message]: { one: 'One cancellation', other: '{count} cancellations' } });
  M.setFallback(null);
  assert.equal(M.setLanguage('en'), 'en');
  assert.equal(M.t('Close'), 'Close the window');
  assert.equal(M.t('Saved {name}', { name: 'a.docx' }), 'Saved a.docx', 'a label lang-en lacks, as the code writes it');
  assert.equal(M.tn(2, CANCEL.message, CANCEL.plural), '2 cancellations');
  assert.deepEqual(M.untranslated(), [], 'English is never counted as untranslated');

  M.registerCatalogue('fr', { 'Saved {name}': '{name} enregistré' });
  M.setFallback({ Close: 'Close the window', [CANCEL.message]: { one: 'One cancellation', other: '{count} cancellations' } });
  assert.equal(M.setLanguage('fr-FR'), 'fr');
  assert.equal(M.t('Saved {name}', { name: 'a.docx' }), 'a.docx enregistré');
  assert.equal(M.t('Close'), 'Close the window', 'missing in French: lang-en\'s English');
  assert.equal(M.tn(1, CANCEL.message, CANCEL.plural), 'One cancellation');
  assert.equal(M.t('New since the pack'), 'New since the pack');
  assert.deepEqual(M.untranslated(), [CANCEL.message, 'Close', 'New since the pack'].sort());
  M.setFallback(null);
  M.setLanguage('en');
});

/** A signed lang-en release folder, as the packages repository's sign.mjs leaves it. */
function release(root, { privateKey, publicKey }, { version = '2.25.0', revision = 1 } = {}) {
  const files = {
    'messages.json': JSON.stringify({ 'shell.close': { text: 'Close', source: H({ message: 'Close' }) }, 'engine.cell': { text: CELL, source: H({ message: CELL }) } }),
    'LICENSE.txt': 'Licence',
  };
  const listed = {};
  for (const [name, text] of Object.entries(files)) {
    write(path.join(root, 'lang-en', name), text);
    listed[name] = { size: Buffer.byteLength(text), sha512: crypto.createHash('sha512').update(text).digest('base64') };
  }
  const manifest = { id: 'lang-en', kind: 'language', format: 2, version, revision, tag: 'en', name: 'English', files: listed };
  const signature = crypto.sign(null, Buffer.from(canonical(manifest), 'utf8'), privateKey).toString('base64');
  write(path.join(root, 'lang-en.pack.json'), `${JSON.stringify({ format: 'rutba-office-pack', keyId: keyIdOf(publicKey), manifest, signature }, null, 2)}\n`);
}

test('taking lang-en: checked as Office would, copied byte for byte, with where it came from', () => {
  const root = temp();
  try {
    const pair = crypto.generateKeyPairSync('ed25519');
    const pem = pair.publicKey.export({ type: 'spki', format: 'pem' }).toString();
    const from = path.join(root, 'v2.25.0');
    const to = path.join(root, 'copy');
    release(from, pair);
    write(path.join(to, 'stale.txt'), 'left from before');

    const r = takeLangEn({ from, to, keys: [pem], labels: KEYS, source: 'https://example.invalid/v2.25.0/', office: '2.25.0' });
    assert.equal(r.checked, true);
    assert.deepEqual(r.labels, { used: KEYS.length, found: 2 });
    assert.ok(fs.readFileSync(path.join(to, 'lang-en.pack.json')).equals(fs.readFileSync(path.join(from, 'lang-en.pack.json'))), 'the envelope byte for byte');
    assert.ok(fs.readFileSync(path.join(to, 'lang-en', 'messages.json')).equals(fs.readFileSync(path.join(from, 'lang-en', 'messages.json'))));
    assert.ok(!fs.existsSync(path.join(to, 'stale.txt')), 'nothing of the old copy is left');
    const record = JSON.parse(fs.readFileSync(path.join(to, 'lang-en.source.json'), 'utf8'));
    assert.equal(record.from, 'https://example.invalid/v2.25.0/');
    assert.deepEqual(Object.keys(record.files).sort(), ['LICENSE.txt', 'messages.json']);
    // And the windows read it as their English.
    assert.equal(readEnglishCopy(to).manifest.revision, 1);

    // Another key, a changed file, no keys at all: refused.
    const other = crypto.generateKeyPairSync('ed25519').publicKey.export({ type: 'spki', format: 'pem' }).toString();
    assert.throws(() => takeLangEn({ from, to: path.join(root, 'b'), keys: [other], labels: KEYS }), /not signed by Rutba/);
    assert.throws(() => takeLangEn({ from, to: path.join(root, 'c'), keys: [], labels: KEYS }), /no packages keys/);
    assert.equal(takeLangEn({ from, to: path.join(root, 'd'), keys: [], unchecked: true, labels: KEYS }).checked, false);
    fs.appendFileSync(path.join(from, 'lang-en', 'LICENSE.txt'), '!');
    assert.throws(() => takeLangEn({ from, to: path.join(root, 'e'), keys: [pem], labels: KEYS }), /changed since it was signed/);
    assert.ok(!fs.existsSync(path.join(root, 'e')), 'nothing is copied when the check fails');
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});
