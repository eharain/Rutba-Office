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
  ];

  // The update gate: what an update to TARGET may be offered for.
  const TARGET = '99.0.0';
  const feeds = {};
  const feedOf = (name, packsIn, { state = 'partial', retired = [], office: forOffice = TARGET } = {}) => {
    feeds[name] = { format: 1, readiness: 1, office: forOffice, sequence: 1, state, retired, packs: packsIn };
    return `feeds/${name}.json`;
  };
  const p = (id, more = {}) => ({ id, version: TARGET, revision: 1, ...more });
  const ready = (name, installed, feed, expect) => ({ type: 'ready', name, installed: installed.map((id) => ({ id })), feed, target: TARGET, expect });
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
  ];
  return { packs, feeds, cases: [...load_cases, ...ready_cases] };
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
  return { dir, count: made.cases.length };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { dir, count } = makeVectors();
  console.log(`${count} cases in ${path.relative(ROOT, dir)}`);
}
