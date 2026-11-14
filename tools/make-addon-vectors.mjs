// The shared add-on cases (private design: office-pack-compatibility-design.md 3.1):
// signed envelopes and feeds, with what Office's addons-compat.js must
// answer for each, run here by tests/addons-contract.test.js and by the
// packages repository's tools/contract.test.mjs against the same files.
//
//   node tools/make-addon-vectors.mjs
//
// Writes packages/office-shell/test-vectors/addons/:
//   vectors.json          { vectors, contract, from, cases }
//   test-key.pub          the public half of a key made for this run; the
//                         private half is never written, so the cases are
//                         made again whole, never patched
//   packs/<name>.pack.json, feeds/<name>.json
//   languages/            small language packs and the words a window gets
//                         from each (tests/language-vectors.test.js, L9)
//
// `from` is the SHA-256 of the contract's canonical JSON the cases were made from:
// the test fails when the contract changes and the cases were not made again.
// A case names `signedAs`: the role (packages or release) the test key is
// given, the other role holding no key.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { canonical, keyIdOf } from '../packages/office-shell/src/electron/packs.js';
import { ENVELOPE_FORMAT } from '../packages/office-shell/src/electron/addons-compat.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const CONTRACT_FILE = path.join(ROOT, 'packages/office-shell/src/addons-contract.json');
export const VECTORS_DIR = path.join(ROOT, 'packages/office-shell/test-vectors/addons');
export const VECTORS = 1;

/** The contract's SHA-256, over its canonical JSON, so line endings on checkout never change it. */
export const contractHash = (file = CONTRACT_FILE) => crypto.createHash('sha256').update(canonical(JSON.parse(fs.readFileSync(file, 'utf8')))).digest('hex');

/** A file entry as a manifest lists one. */
export const fileEntry = (bytes) => ({ size: Buffer.byteLength(bytes), sha512: crypto.createHash('sha512').update(bytes).digest('base64') });

/** An add-on's envelope: the manifest signed with an Ed25519 key. */
export function envelopeOf(manifest, privateKey) {
  const signature = crypto.sign(null, Buffer.from(canonical(manifest), 'utf8'), privateKey).toString('base64');
  return { format: ENVELOPE_FORMAT, keyId: keyIdOf(crypto.createPublicKey(privateKey)), manifest, signature };
}

/** The manifests the cases use, for the contract's Office version. */
export function manifests(contract) {
  const office = contract.office;
  const files = { 'data.txt': fileEntry('test\n') };
  const engine = {
    id: 'speech-engine', name: 'Test speech engine', kind: 'engine', version: office, revision: 1, format: 1,
    speechApi: contract.speech.api, host: contract.engineHost.api, platform: 'win32-x64',
    ipc: { [contract.speech.namespace]: [...contract.speech.calls] }, events: [...contract.speech.events],
    reads: { model: [1], pronunciation: [1], voice: [1] }, main: 'main/index.js',
    files: { 'main/index.js': fileEntry('export function register() { return {}; }\n') },
  };
  return {
    lang: { id: 'lang-xx', name: 'Test language', kind: 'language', version: office, revision: 1, format: 2, provides: ['language:xx'], files },
    dict: { id: 'dict-xx', name: 'Test dictionary', kind: 'dictionary', version: office, revision: 1, format: 1, requires: ['lang-xx'], files },
    voice: { id: 'voice-test', name: 'Test voice', kind: 'voice', version: office, revision: 1, format: 1, requires: ['model-test'], files },
    model: { id: 'model-test', name: 'Test model', kind: 'model', version: office, revision: 1, format: 1, requires: ['speech-engine'], files },
    engine,
  };
}

/** Every case, with the envelopes and feeds it reads. */
export function cases(contract, privateKey) {
  const office = contract.office;
  const m = manifests(contract);
  const packs = {};
  const sign = (name, manifest) => { packs[name] = envelopeOf(manifest, privateKey); return `packs/${name}.pack.json`; };
  const load = (name, pack, signedAs, expect, extra = {}) => ({ type: 'load', name, pack, office, signedAs, ...extra, expect });
  const yes = (mf) => ({ usable: true, id: mf.id, kind: mf.kind });
  const no = (id, code) => ({ usable: false, id, code });

  const lang = sign('lang-xx', m.lang);
  const engine = sign('speech-engine', m.engine);
  const tampered = { ...envelopeOf(m.lang, privateKey), manifest: { ...m.lang, revision: 2 } };
  packs['lang-xx-tampered'] = tampered;
  packs['not-an-addon'] = { format: 'something-else', manifest: m.lang, signature: packs['lang-xx'].signature };

  const load_cases = [
    load('a language signed for packages is usable', lang, 'packages', yes(m.lang)),
    load('a language signed with the release key is refused for its kind', lang, 'release', no('lang-xx', 'key-for-kind')),
    load('a dictionary is usable', sign('dict-xx', m.dict), 'packages', yes(m.dict)),
    load('a manifest changed after signing is refused', 'packs/lang-xx-tampered.pack.json', 'packages', no('lang-xx', 'signature')),
    load('something that is not an add-on is refused', 'packs/not-an-addon.pack.json', 'packages', no('lang-xx', 'envelope')),
    load('a language in a format this Office does not read is refused', sign('lang-xx-format-9', { ...m.lang, format: 9 }), 'packages', no('lang-xx', 'format')),
    load('a kind this Office does not use is refused', sign('macro-xx', { ...m.lang, id: 'macro-xx', kind: 'macro' }), 'packages', no('macro-xx', 'kind')),
    load('the engine signed with the release key is usable', engine, 'release', yes(m.engine)),
    load('the engine signed for packages is refused for its kind', engine, 'packages', no('speech-engine', 'key-for-kind')),
    load('an engine for another Office version is refused', sign('speech-engine-old', { ...m.engine, version: '0.0.1' }), 'release', no('speech-engine', 'version')),
    load('an engine for another speech API is refused', sign('speech-engine-api-9', { ...m.engine, speechApi: 9 }), 'release', no('speech-engine', 'version')),
    load('an engine whose calls are not the contract\'s is refused', sign('speech-engine-calls', { ...m.engine, ipc: { speech: ['voices', 'say'] } }), 'release', no('speech-engine', 'version')),
    load('an engine for a platform Office hosts none on is refused', sign('speech-engine-linux', { ...m.engine, platform: 'linux-x64' }), 'release', no('speech-engine', 'version')),
    load('a voice with the engine is usable', sign('voice-test', m.voice), 'packages', yes(m.voice), { engine }),
    load('a voice with no engine needs it', 'packs/voice-test.pack.json', 'packages', no('voice-test', 'requires')),
    load('a voice in a format this Office does not read is refused', sign('voice-test-format-2', { ...m.voice, format: 2 }), 'packages', no('voice-test', 'format')),
    load('a model with the engine is usable', sign('model-test', m.model), 'packages', yes(m.model), { engine }),
    // Language packs (L9).
    load('a language in the format before is usable', sign('lang-xx-format-1', { ...m.lang, format: 1 }), 'packages', yes(m.lang)),
    load('a language that holds a program file is refused', sign('lang-xx-program', { ...m.lang, files: { ...m.lang.files, 'words.js': fileEntry('export default {};\n') } }), 'packages', no('lang-xx', 'files')),
    load('a language signed with an app pack\'s key is refused', 'packs/lang-xx.pack.json', 'apps', no('lang-xx', 'signature')),
    load('a language with its version dropped is refused', sign('lang-xx-no-version', (({ version, ...rest }) => rest)(m.lang)), 'packages', no('lang-xx', 'envelope')),
    load('a language with its kind dropped is refused', sign('lang-xx-no-kind', (({ kind, ...rest }) => rest)(m.lang)), 'packages', no('lang-xx', 'kind')),
  ];

  // The update gate: what an update to TARGET may be offered for.
  const TARGET = '99.0.0';
  const feeds = {};
  const feedOf = (name, packsIn, { state = 'partial', retired = [], office: forOffice = TARGET } = {}) => {
    feeds[name] = { format: 1, readiness: 1, office: forOffice, sequence: 1, state, retired, packs: packsIn };
    return `feeds/${name}.json`;
  };
  const p = (id, more = {}) => ({ id, version: TARGET, revision: 1, ...more });
  // An installed entry is an id, or { id, builtIn: true } for the built-in language in use that the target leaves out of the core (L7).
  const ready = (name, installed, feed, expect) => ({ type: 'ready', name, installed: installed.map((x) => (typeof x === 'string' ? { id: x } : x)), feed, target: TARGET, expect });
  const ready_cases = [
    ready('no add-ons: offer at once, with no feed', [], null, { offer: true, ready: [], waiting: [], retired: [] }),
    ready('English alone never holds an update', ['lang-en'], null, { offer: true, ready: [], waiting: [], retired: [] }),
    ready('every add-on published for the target: offer', ['lang-xx', 'dict-xx'], feedOf('all-there', [p('lang-xx', { provides: ['language:xx'] }), p('dict-xx', { requires: ['lang-xx'] })]),
      { offer: true, ready: ['dict-xx', 'lang-xx'], waiting: [], retired: [] }),
    ready('no feed could be had: wait', ['lang-xx'], null, { offer: false, waiting: [{ id: 'lang-xx', why: 'no-feed' }] }),
    ready('a feed for another Office is no feed', ['lang-xx'], feedOf('other-office', [p('lang-xx')], { office: '98.0.0' }), { offer: false, waiting: [{ id: 'lang-xx', why: 'no-feed' }] }),
    ready('an add-on not yet published: wait', ['lang-xx', 'dict-xx'], feedOf('partial', [p('lang-xx')]), { offer: false, ready: ['lang-xx'], waiting: [{ id: 'dict-xx', why: 'unpublished' }] }),
    ready('an add-on listed only at an older version: wait', ['lang-xx'], feedOf('older', [p('lang-xx', { version: '98.0.0' })]), { offer: false, waiting: [{ id: 'lang-xx', why: 'unpublished' }] }),
    ready('an add-on withdrawn: wait', ['lang-xx'], feedOf('withdrawn', [p('lang-xx', { withdrawn: true })]), { offer: false, waiting: [{ id: 'lang-xx', why: 'withdrawn' }] }),
    ready('an add-on whose need is not published: wait', ['voice-test'], feedOf('needs', [p('voice-test', { requires: ['model-test'] })]), { offer: false, waiting: [{ id: 'voice-test', why: 'needs' }] }),
    ready('a complete feed without an add-on retires it', ['lang-xx', 'dict-xx'], feedOf('complete', [p('lang-xx')], { state: 'complete' }),
      { offer: true, ready: ['lang-xx'], waiting: [], retired: [{ id: 'dict-xx', reason: 'Rutba Office 99.0.0 has no dict-xx.' }] }),
    ready('a retired add-on never holds an update', ['lang-xx', 'dict-xx'], feedOf('retired', [p('lang-xx')], { retired: [{ id: 'dict-xx', reason: 'Joined into lang-xx.', replacedBy: 'lang-xx' }] }),
      { offer: true, ready: ['lang-xx'], waiting: [], retired: [{ id: 'dict-xx', reason: 'Joined into lang-xx.' }] }),
    // A built-in language the target leaves out of the core (D-2026-11-13-36): it waits for its pack even in a complete feed, never retired by the feed leaving it out.
    ready('a built-in language left out of the core waits for its pack, even in a complete feed', [{ id: 'lang-xx', builtIn: true }, 'dict-xx'], feedOf('complete-without-language', [p('dict-xx')], { state: 'complete' }),
      { offer: false, ready: ['dict-xx'], waiting: [{ id: 'lang-xx', why: 'unpublished' }], retired: [] }),
    ready('a built-in language left out of the core, its pack in a complete feed: offer', [{ id: 'lang-xx', builtIn: true }], 'feeds/complete.json',
      { offer: true, ready: ['lang-xx'], waiting: [], retired: [] }),
    ready('a built-in language a complete feed retires never holds an update', [{ id: 'lang-xx', builtIn: true }], feedOf('complete-retiring-language', [p('dict-xx')], { state: 'complete', retired: [{ id: 'lang-xx', reason: 'Not made for this version.' }] }),
      { offer: true, ready: [], waiting: [], retired: [{ id: 'lang-xx', reason: 'Not made for this version.' }] }),
  ];
  return { packs, feeds, cases: [...load_cases, ...ready_cases] };
}

/**
 * The language cases (L9): small packs in folders, as an add-ons folder holds
 * them, with the words a window must get from each, run by
 * tests/language-vectors.test.js through language-packs.js. Unsigned: the
 * signature, kind and format of a language are the load cases above; these
 * are what becomes of its words once it may be used.
 */
export function languageCases(office) {
  const keys = [
    { message: 'Close' },
    { message: 'Saved {name}' },
    { message: 'Open {name}' },
    { message: '{count} file', plural: '{count} files' },
    { message: 'New in this Office' },
  ];
  const src = (m) => crypto.createHash('sha256').update(m.plural ? JSON.stringify({ one: m.message, other: m.plural }) : m.message, 'utf8').digest('hex').slice(0, 12);
  const [close, saved, open, files] = keys;
  const lang = (id, tag, entries, more = {}, versions = null) => ({ manifest: { id, kind: 'language', format: 2, version: office, revision: 1, tag, name: id, ...more }, messages: entries, ...(versions ? { versions } : {}) });
  const folders = {
    copy: lang('lang-en', 'en', { 'shell.close': { text: 'Close the window', source: src(close) }, 'shell.files': { forms: { one: 'One file', other: '{count} files' }, source: src(files) } }, { revision: 1 }, {
      // Its short one differs from its medium one; its complete one is the medium again, so left out.
      'shell.close': { short: 'Close', complete: 'Close the window', source: src(close) },
      'shell.saved': { short: 'Saved', source: src(saved) },
    }),
    'en-newer': lang('lang-en', 'en', { 'shell.close': { text: 'Close it', source: src(close) } }, { revision: 2 }),
    'en-older': lang('lang-en', 'en', { 'shell.close': { text: 'Close, old', source: src(close) } }, { version: '0.0.1', revision: 9 }),
    'fr-format-1': lang('lang-fr', 'fr', { 'shell.close': { text: 'Fermer', source: src(close) } }, { format: 1 }),
    fr: lang('lang-fr', 'fr', {
      'shell.close': { text: 'Fermer', source: src(close) },
      'shell.saved': { text: '{name} enregistré', source: src(saved) },
      'shell.files': { forms: { one: '{count} fichier', other: '{count} fichiers' }, source: src(files) },
      // Extra: an id this Office has no key for.
      'gone.thing': { text: 'Disparu', source: 'aaaaaaaaaaaa' },
      // Stale: made from English the code no longer has.
      'shell.open': { text: 'Ouvrir {name}', source: src({ message: 'Open {name}…' }) },
    }),
    'fr-ca': lang('lang-fr-ca', 'fr-CA', { 'shell.close': { text: 'Fermer la fenêtre', source: src(close) } }, { variantOf: 'lang-fr', requires: ['lang-fr'] }),
    'ur-rtl': lang('lang-ur', 'ur', { 'shell.close': { text: 'بند کریں', source: src(close) } }, { direction: 'rtl' }),
    'he-ltr': lang('lang-he', 'he', { 'shell.close': { text: 'סגור', source: src(close) } }, { direction: 'ltr' }),
    'de-placeholder': lang('lang-de', 'de', {
      'shell.close': { text: 'Schließen', source: src(close) },
      // Adds a name the key does not have: never shown.
      'shell.saved': { text: '{name} in {folder} gespeichert', source: src(saved) },
    }),
    // L3: a format 2 pack's short and complete versions.
    'fr-format-2': lang('lang-fr', 'fr', {
      'shell.close': { text: 'Fermer', source: src(close) },
      'shell.saved': { text: '{name} enregistré', source: src(saved) },
      'shell.files': { forms: { one: '{count} fichier', other: '{count} fichiers' }, source: src(files) },
    }, {}, {
      'shell.close': { complete: 'Fermer la fenêtre', source: src(close) },
      'shell.saved': { short: '{name} enreg.', complete: '{name} enregistré sur le disque', source: src(saved) },
      'shell.files': { short: { one: '{count} fich.', other: '{count} fich.' }, source: src(files) },
      // Stale: made from English the code no longer has.
      'shell.open': { short: 'Ouvr. {name}', source: src({ message: 'Open {name}…' }) },
    }),
    'fr-format-2-bad-versions': lang('lang-fr', 'fr', {
      'shell.close': { text: 'Fermer', source: src(close) },
      'shell.saved': { text: '{name} enregistré', source: src(saved) },
      'shell.files': { forms: { one: '{count} fichier', other: '{count} fichiers' }, source: src(files) },
    }, {}, {
      // A name the key does not have, a string for a counted label, a short the same as the medium: none of them kept.
      'shell.saved': { short: '{name} dans {folder}', source: src(saved) },
      'shell.files': { short: '{count} fich.', source: src(files) },
      'shell.close': { short: 'Fermer', source: src(close) },
    }),
    'fr-format-1-versions': lang('lang-fr', 'fr', { 'shell.close': { text: 'Fermer', source: src(close) } }, { format: 1 }, {
      'shell.close': { short: 'Ferm.', source: src(close) },
    }),
    'fr-ca-format-2': lang('lang-fr-ca', 'fr-CA', { 'shell.close': { text: 'Fermer cette fenêtre', source: src(close) } }, { variantOf: 'lang-fr', requires: ['lang-fr'] }),
    'xx-format-9': lang('lang-xx', 'xx', { 'shell.close': { text: 'Xx', source: src(close) } }, { format: 9 }),
  };
  const ask = (name, installed, tag, expect, more = {}) => ({ name, copy: 'copy', installed, ask: tag, ...more, expect });
  const english = { Close: 'Close the window', [files.message]: { one: 'One file', other: '{count} files' } };
  const englishVersions = { Close: { short: 'Close' }, 'Saved {name}': { short: 'Saved' } };
  const cases = [
    ask('English with no pack added: the copy', [], 'en', { tag: 'en', from: 'copy', table: english }),
    ask('an added lang-en above the copy is used', ['en-newer'], 'en-GB', { tag: 'en', from: 'added', table: { Close: 'Close it' } }),
    ask('an added lang-en for an older Office: the copy', ['en-older'], 'en', { tag: 'en', from: 'copy', table: english }),
    ask('lang-en withdrawn and rolled back: the copy again', [], 'en', { tag: 'en', from: 'copy', table: english }),
    ask('no copy and no lang-en: the code\'s own English', [], 'en', { tag: 'en', from: null, table: {} }, { copy: null }),
    ask('the format before loads', ['fr-format-1'], 'fr', { tag: 'fr', table: { Close: 'Fermer' }, english }),
    ask('extra, missing and stale entries: only what fits the code now', ['fr'], 'fr', { tag: 'fr', table: { Close: 'Fermer', 'Saved {name}': '{name} enregistré', [files.message]: { one: '{count} fichier', other: '{count} fichiers' } }, english }),
    ask('a regional pack over its base, then English', ['fr', 'fr-ca'], 'fr-CA', { tag: 'fr-ca', table: { Close: 'Fermer la fenêtre', 'Saved {name}': '{name} enregistré', [files.message]: { one: '{count} fichier', other: '{count} fichiers' } }, english }),
    ask('a region with no pack takes its base', ['fr'], 'fr-BE', { tag: 'fr', table: { Close: 'Fermer', 'Saved {name}': '{name} enregistré', [files.message]: { one: '{count} fichier', other: '{count} fichiers' } }, english }),
    ask('a pack\'s direction sets the window\'s', ['ur-rtl'], 'ur', { tag: 'ur', direction: 'rtl', table: { Close: 'بند کریں' }, english }),
    ask('a pack\'s direction wins over the built-in list', ['he-ltr'], 'he', { tag: 'he', direction: 'ltr', table: { Close: 'סגור' }, english }),
    ask('an entry that adds a placeholder is not shown', ['de-placeholder'], 'de', { tag: 'de', table: { Close: 'Schließen' }, english }),
    ask('a format this Office does not read: no pack', ['xx-format-9'], 'xx', { tag: null, table: null, english }),
    // L3: the versions a window gets beside its words.
    ask('a format 2 pack gives its short and complete versions, a stale one left out', ['fr-format-2'], 'fr', {
      tag: 'fr',
      table: { Close: 'Fermer', 'Saved {name}': '{name} enregistré', [files.message]: { one: '{count} fichier', other: '{count} fichiers' } },
      versions: { Close: { complete: 'Fermer la fenêtre' }, 'Saved {name}': { short: '{name} enreg.', complete: '{name} enregistré sur le disque' }, [files.message]: { short: { one: '{count} fich.', other: '{count} fich.' } } },
      englishVersions,
    }),
    ask('a version that adds a name, has the wrong shape or reads as the medium is not kept', ['fr-format-2-bad-versions'], 'fr', {
      tag: 'fr',
      table: { Close: 'Fermer', 'Saved {name}': '{name} enregistré', [files.message]: { one: '{count} fichier', other: '{count} fichiers' } },
      versions: {},
      englishVersions,
    }),
    ask('a format 1 pack\'s versions.json is not read', ['fr-format-1-versions'], 'fr', { tag: 'fr', table: { Close: 'Fermer' }, versions: {}, englishVersions }),
    ask('a regional pack\'s own label takes none of its base\'s versions', ['fr-format-2', 'fr-ca-format-2'], 'fr-CA', {
      tag: 'fr-ca',
      table: { Close: 'Fermer cette fenêtre', 'Saved {name}': '{name} enregistré', [files.message]: { one: '{count} fichier', other: '{count} fichiers' } },
      versions: { 'Saved {name}': { short: '{name} enreg.', complete: '{name} enregistré sur le disque' }, [files.message]: { short: { one: '{count} fich.', other: '{count} fich.' } } },
      englishVersions,
    }),
    ask('an English window has lang-en\'s versions as its own', [], 'en', { tag: 'en', from: 'copy', table: english, versions: englishVersions, englishVersions }),
  ];
  return { keys, folders, cases };
}

/** Writes the language cases into `<dir>/languages/`. */
function writeLanguageCases(dir, office) {
  const made = languageCases(office);
  const root = path.join(dir, 'languages');
  for (const [name, { manifest, messages, versions }] of Object.entries(made.folders)) {
    fs.mkdirSync(path.join(root, name), { recursive: true });
    fs.writeFileSync(path.join(root, name, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
    fs.writeFileSync(path.join(root, name, 'messages.json'), `${JSON.stringify(messages, null, 2)}\n`);
    if (versions) fs.writeFileSync(path.join(root, name, 'versions.json'), `${JSON.stringify(versions, null, 2)}\n`);
  }
  fs.writeFileSync(path.join(root, 'cases.json'), `${JSON.stringify({ office, keys: made.keys, cases: made.cases }, null, 2)}\n`);
  return made.cases.length;
}

/** Makes the cases again, whole, into `dir`. */
export function makeVectors({ dir = VECTORS_DIR, contractFile = CONTRACT_FILE } = {}) {
  const contract = JSON.parse(fs.readFileSync(contractFile, 'utf8'));
  const { privateKey, publicKey } = crypto.generateKeyPairSync('ed25519');
  const made = cases(contract, privateKey);
  fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(path.join(dir, 'packs'), { recursive: true });
  fs.mkdirSync(path.join(dir, 'feeds'), { recursive: true });
  const write = (rel, value) => fs.writeFileSync(path.join(dir, rel), `${JSON.stringify(value, null, 2)}\n`);
  for (const [name, env] of Object.entries(made.packs)) write(`packs/${name}.pack.json`, env);
  for (const [name, feed] of Object.entries(made.feeds)) write(`feeds/${name}.json`, feed);
  fs.writeFileSync(path.join(dir, 'test-key.pub'), publicKey.export({ type: 'spki', format: 'pem' }));
  write('vectors.json', { vectors: VECTORS, contract: contract.contract, from: contractHash(contractFile), cases: made.cases });
  const languages = writeLanguageCases(dir, contract.office);
  return { dir, count: made.cases.length, languages };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { dir, count, languages } = makeVectors();
  console.log(`${count} cases and ${languages} language cases in ${path.relative(ROOT, dir)}`);
}
