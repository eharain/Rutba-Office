// The gate's language packs (apps/desktop/main/gate-packs.js): a gate run
// whose window checks open windows in Urdu, Arabic or Hindi is refused at its
// start without a release folder of signed packs (RUTBA_GATE_PACKS), and each
// check run adds lang-ur, lang-ar and lang-hi to its own add-ons folder
// through the add-ons service, as Add from file does: a stranger's signature,
// a file changed after signing or another version's folder is refused, and
// what is added is what the run's next start loads.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { envelopeOf, fileEntry, CONTRACT_FILE } from '../tools/make-addon-vectors.mjs';
import { GATE_PACKS_ENV, GATE_LANGUAGE_PACKS, GATE_SPEECH_PACKS, GATE_VOICE, PARTS_NEEDING_PACKS, PARTS_NEEDING_SPEECH, SPEECH_ONLY, partsNeedingPacks, partsNeedingSpeech, speechWanted, speechRefusal, speechReleaseKeys, languageSwitches, releaseProblem, packsRefusal, zipOfPack, installGatePacks } from '../apps/desktop/main/gate-packs.js';
import { startAddons } from '../apps/desktop/main/addons-start.js';
import { removeTree } from '../apps/desktop/main/addons-folder.js';
import { VERIFY_APPS } from '../tools/gate.js';
import { readZip } from '../packages/ooxml/src/zip.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-gate-packs-test-'));
test.after(() => removeTree(tmp));
let n = 0;
const fresh = (name) => {
  const d = path.join(tmp, `${name}-${(n += 1)}`);
  fs.mkdirSync(d, { recursive: true });
  return d;
};

const file = JSON.parse(fs.readFileSync(CONTRACT_FILE, 'utf8'));
const OFFICE = file.office;
const packagesKey = crypto.generateKeyPairSync('ed25519');
const pem = (k) => k.publicKey.export({ type: 'spki', format: 'pem' }).toString();
const contract = { ...file, keys: { ...file.keys, packages: [pem(packagesKey)] } };

/** A packages release folder, v<version>, as the packages build writes one: each pack's envelope and its folder. */
function releaseFolder({ version = OFFICE, ids = GATE_LANGUAGE_PACKS, key = packagesKey } = {}) {
  const dir = path.join(fresh('release'), `v${version}`);
  for (const id of ids) {
    const tag = id.slice('lang-'.length);
    const content = { 'manifest.json': `{"id":"${id}"}\n`, 'messages.json': `{"common.save":{"text":"${tag} save"}}\n` };
    const manifest = { id, name: `Test ${tag}`, kind: 'language', version, revision: 1, format: 2, provides: [`language:${tag}`], files: Object.fromEntries(Object.entries(content).map(([p, b]) => [p, fileEntry(b)])) };
    for (const [p, b] of Object.entries(content)) {
      fs.mkdirSync(path.join(dir, id), { recursive: true });
      fs.writeFileSync(path.join(dir, id, p), b);
    }
    fs.writeFileSync(path.join(dir, `${id}.pack.json`), JSON.stringify(envelopeOf(manifest, key.privateKey)));
  }
  return dir;
}

test('the parts that need the packs are the window checks that open windows in Urdu, Arabic or Hindi, each a part verify-apps.js runs', () => {
  const source = fs.readFileSync(path.join(ROOT, 'apps/desktop/main/verify-apps.js'), 'utf8');
  for (const id of PARTS_NEEDING_PACKS) {
    const [app, fn] = id.split('#');
    assert.ok(VERIFY_APPS.includes(app), id);
    assert.ok(source.includes(`part('${app}', () => ${fn}(`), `${id} is a part of verify-apps.js`);
  }
  // A whole run, and a run of those apps, includes them; a run of other apps or other parts does not.
  assert.deepEqual(partsNeedingPacks({ apps: VERIFY_APPS }), PARTS_NEEDING_PACKS);
  assert.deepEqual(partsNeedingPacks({ apps: ['mail', 'pdf'] }), ['pdf#verifyPdfApp', 'mail#verifySmime', 'mail#verifyMailAccounts', 'mail#verifyMailArchive', 'mail#verifyMailInvoice']);
  // A run of the shell alone (the chunk a change to the packs' tools gives) needs them too.
  assert.deepEqual(partsNeedingPacks({ apps: ['shell'] }), ['shell#verifyTemplates', 'shell#verifyAppMenu']);
  assert.deepEqual(partsNeedingPacks({ apps: ['word', 'sheets', 'slides'] }), []);
  assert.deepEqual(partsNeedingPacks({ apps: VERIFY_APPS, partIds: ['mail#reading', 'video#verifyVideoMp4'] }), []);
  assert.deepEqual(partsNeedingPacks({ apps: ['mail'], partIds: ['mail#reading', 'mail#verifySmime'] }), ['mail#verifySmime']);
});

/** Every window check source, by file name. */
const MAIN = path.join(ROOT, 'apps/desktop/main');
const verifySources = () => Object.fromEntries(fs.readdirSync(MAIN).filter((f) => /^verify-.*\.js$/.test(f)).map((f) => [f, fs.readFileSync(path.join(MAIN, f), 'utf8')]));

/**
 * The window checks that hold Urdu, Arabic or Hindi letters, or a right to
 * left direction, without setting the suite's language: what they read is a
 * document's own (its words and direction, which no pack gives), each with
 * why. A file that comes to hold them and is not here, nor a part that sets
 * the language, fails the test below until it is one or the other.
 */
const NOT_THE_WINDOWS_LANGUAGE = {
  'verify-apps.js': 'names the verifyWordRtl run (RUTBA_VERIFY_ONLY=rtl), no window in another language',
  'verify-deck-rtl.js': "a slide's own right to left text",
  'verify-image-text-edit.js': 'Urdu typed into a picture as its text',
  'verify-intro-video.js': "a film's captions and a brand's name in Urdu and Arabic, the person's own words",
  'verify-ocr-pdf.js': "Urdu words found in a PDF's pages by OCR",
  'verify-ocr-picture.js': "a picture's Urdu read by OCR, the lines right to left",
  'verify-pdf-find-mark.js': "Urdu and Arabic words found in a PDF's text",
  'verify-pdf-scripts.js': "Arabic in a PDF's text",
  'verify-rotate-frame-deck.js': 'the page set right to left by the check itself, to turn a frame both ways',
  'verify-rotate-frame-image.js': 'the page set right to left by the check itself, to turn a frame both ways',
  'verify-rotate-frame-word.js': 'the page set right to left by the check itself, to turn a frame both ways',
  'verify-sheet-rtl.js': "a worksheet's own right to left cells",
  'verify-word-rtl.js': "a document's own right to left paragraphs and table",
  'verify-word-table-tools.js': "a table's own direction, to click its last cell",
  'verify-word-table.js': "a table turned right to left in the document",
};
/** Urdu, Arabic or Hindi letters (or their escapes), or a right to left direction. */
const SCRIPT_OR_DIRECTION = /'rtl'|"rtl"|\\u06|\\u09[0-7]|[؀-ۿऀ-ॿ]/;

test('every part of the window checks that sets the suite\'s language is listed as needing the packs, and none is listed that does not', () => {
  const found = languageSwitches(verifySources());
  const parts = [...new Set(found.map((f) => f.part))];
  const missing = parts.filter((p) => !PARTS_NEEDING_PACKS.includes(p));
  assert.deepEqual(missing, [], `set the language but not in PARTS_NEEDING_PACKS: ${missing.map((p) => `${p} (${found.filter((f) => f.part === p).map((f) => f.at).join(', ')})`).join('; ')}`);
  assert.deepEqual(PARTS_NEEDING_PACKS.filter((p) => !parts.includes(p)), [], 'listed, but no longer sets the language');
  // Where the gate is documented, the same list.
  const releasing = fs.readFileSync(path.join(ROOT, 'docs/RELEASING.md'), 'utf8');
  const head = fs.readFileSync(path.join(ROOT, 'tools/gate.js'), 'utf8').split('\nimport ')[0];
  for (const p of PARTS_NEEDING_PACKS) {
    assert.ok(releasing.includes(`\`${p}\``), `RELEASING.md names ${p}`);
    assert.ok(head.includes(p), `the gate's head names ${p}`);
  }
  // The fifteen as found: the language checks through their one helper, a switch to a tag in a loop, Urdu by name.
  for (const [part, at] of [['languages#verifyHindi', 'verify-urdu.js:'], ['mail#verifyMailArchive', 'verify-mail-archive.js:'], ['shell#verifyAppMenu', 'verify-app-menu.js:'], ['recorder#verifyRecorder', 'verify-recorder.js:']]) {
    assert.ok(found.some((f) => f.part === part && f.at.startsWith(at)), part);
  }
});

test('the scan finds a switch to another language in a new check, in a helper, and passes over one that puts the language back', () => {
  const apps = "await part('pdf', () => verifyNew({ open }));\nawait part('pdf', () => verifyViaHelper({ open }));\nawait part('pdf', () => verifyBack({ open }));";
  const sources = {
    'verify-apps.js': apps,
    'verify-new.js': "export async function verifyNew() {\n  await js(`window.rutbaOffice.store.set({ key: 'language', value: 'ar' }).then(() => 1)`);\n}\n",
    'verify-helper.js': "async function inArabic(tag) {\n  await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(tag)} }).then(() => 1)`);\n}\nexport const verifyViaHelper = (h) => inArabic('ar');\n",
    'verify-back.js': "export async function verifyBack() {\n  await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(was ?? null)} }).then(() => 1)`);\n  await js(`window.rutbaOffice.store.set({ key: 'language', value: null }).then(() => 1)`);\n}\n",
    'verify-loose.js': "export async function verifyLoose() {\n  await js(`window.rutbaOffice.store.set({ key: 'language', value: 'hi' }).then(() => 1)`);\n}\n",
  };
  const found = languageSwitches(sources).map((f) => f.part).sort();
  // A check in no part is named by its file, so the list test says where it is.
  assert.deepEqual(found, ['pdf#verifyNew', 'pdf#verifyViaHelper', 'verify-loose.js#verifyLoose']);
});

test('every window check holding Urdu, Arabic or Hindi letters or a right to left direction either sets the language (and is listed) or reads a document\'s own, said why', () => {
  const sources = verifySources();
  const switching = new Set(languageSwitches(sources).map((f) => f.at.split(':')[0]));
  const holding = Object.entries(sources).filter(([, text]) => SCRIPT_OR_DIRECTION.test(text)).map(([f]) => f);
  const unexplained = holding.filter((f) => !switching.has(f) && !NOT_THE_WINDOWS_LANGUAGE[f]);
  assert.deepEqual(unexplained, [], 'holds another script or a right to left direction: set the language in a listed part, or say in NOT_THE_WINDOWS_LANGUAGE why it is a document\'s own');
  for (const f of Object.keys(NOT_THE_WINDOWS_LANGUAGE)) {
    assert.ok(holding.includes(f), `${f} no longer holds them: take it out of the list`);
    assert.ok(!switching.has(f), `${f} sets the language: it belongs in PARTS_NEEDING_PACKS`);
  }
});

test('a gate run whose window checks need the packs is refused at its start without a usable release folder, in plain words', () => {
  const good = releaseFolder();
  const all = { verifying: true, apps: VERIFY_APPS, partIds: [], office: OFFICE };
  // Not set.
  const none = packsRefusal({ ...all, release: undefined });
  assert.match(none, /^The gate will not start: its window checks include shell#verifyTemplates, .*mail#verifyMailInvoice, which open windows in Urdu, Arabic or Hindi\. RUTBA_GATE_PACKS is not set/);
  assert.match(none, new RegExp(`dist/v${OFFICE.replace(/\./g, '\\.')}`));
  // Not a folder; another version's; a pack missing.
  assert.match(packsRefusal({ ...all, release: path.join(tmp, 'nowhere') }), /which is not a folder/);
  assert.match(packsRefusal({ ...all, release: releaseFolder({ version: '9.9.9' }) }), /a release folder for another version: this Office is/);
  assert.match(packsRefusal({ ...all, release: releaseFolder({ ids: ['lang-ur', 'lang-ar'] }) }), /has no signed lang-hi/);
  // Usable, or not needed: the run starts.
  assert.equal(packsRefusal({ ...all, release: good }), null);
  assert.equal(packsRefusal({ ...all, verifying: false, release: undefined }), null, 'no window checks');
  assert.equal(packsRefusal({ ...all, apps: ['word', 'sheets'], release: undefined }), null, 'no part that needs them');
  for (const app of ['image', 'pdf', 'recorder', 'shell']) assert.ok(packsRefusal({ ...all, apps: [app], release: undefined }), `a run of ${app} alone is refused`);
  assert.equal(releaseProblem(good, { office: OFFICE }), null);
  assert.equal(GATE_PACKS_ENV, 'RUTBA_GATE_PACKS');
});

test('the gate refuses before it runs anything, and each window check run adds the packs before its Electron starts', () => {
  const gate = fs.readFileSync(path.join(ROOT, 'tools/gate.js'), 'utf8');
  const main = gate.slice(gate.indexOf('async function main()'));
  const refused = main.indexOf('packsRefusal(');
  assert.ok(refused > 0, 'the gate asks packsRefusal');
  for (const first of ['windowRun(', "run('node'", 'localTests(', 'budgetsReady(']) {
    const at = main.indexOf(first, main.indexOf('const want = '));
    assert.ok(at < 0 || refused < at, `refused before ${first}`);
  }
  const runner = fs.readFileSync(path.join(ROOT, 'apps/desktop/build/verify-apps-run.js'), 'utf8');
  assert.ok(runner.indexOf('installGatePacks(') > 0 && runner.indexOf('installGatePacks(') < runner.indexOf('spawn(electron'), 'added before Electron starts');
  assert.match(runner, /addonsEnv\.RUTBA_ADDONS_DIR = addonsDir/);
  assert.match(runner, /env: \{ \.\.\.process\.env, \.\.\.addonsEnv,/);
});

test('the check with a real voice is a part verify-apps.js runs, and a run is given the speech packs only when it includes that part', () => {
  const apps = fs.readFileSync(path.join(ROOT, 'apps/desktop/main/verify-apps.js'), 'utf8');
  for (const id of PARTS_NEEDING_SPEECH) {
    const [app, fn] = id.split('#');
    assert.ok(apps.includes(`part('${app}', () => ${fn}(`), `${id} is a part of the window checks`);
    assert.ok(VERIFY_APPS.includes(app), `${app} is an app the gate runs`);
  }
  assert.ok(apps.includes(`only.includes('${SPEECH_ONLY}')`), 'and it runs alone by its name');
  assert.ok(GATE_SPEECH_PACKS.includes('speech-engine') && GATE_SPEECH_PACKS.includes(GATE_VOICE));
  // By the environment of a run: everything, the part alone, its app, its id, or none of them.
  assert.equal(speechWanted({}), true, 'a run of everything');
  assert.equal(speechWanted({ RUTBA_VERIFY_ONLY: 'real-voice' }), true);
  assert.equal(speechWanted({ RUTBA_VERIFY_ONLY: 'read-aloud' }), false);
  assert.equal(speechWanted({ RUTBA_VERIFY_ONLY: 'read-aloud, real-voice' }), true);
  assert.equal(speechWanted({ RUTBA_VERIFY_APPS: 'word,mail' }), false);
  assert.equal(speechWanted({ RUTBA_VERIFY_APPS: 'word,recorder' }), true);
  assert.equal(speechWanted({ RUTBA_VERIFY_APPS: 'recorder', RUTBA_VERIFY_PART_IDS: 'recorder#verifyRecorder' }), false, 'another part of the same app');
  assert.equal(speechWanted({ RUTBA_VERIFY_APPS: 'recorder', RUTBA_VERIFY_PART_IDS: 'recorder#verifyRealVoice' }), true);
  assert.equal(speechWanted({ RUTBA_VERIFY_PART_IDS: 'recorder#*' }), true, 'the whole app, asked by id');
  assert.equal(speechWanted({ RUTBA_VERIFY_APPS: 'recorder', RUTBA_VERIFY_SKIP_IDS: 'recorder#verifyRealVoice' }), false, 'another piece of the app runs it');
  assert.deepEqual(partsNeedingSpeech({ apps: ['word'] }), []);
  assert.deepEqual(partsNeedingSpeech({ apps: VERIFY_APPS }), [...PARTS_NEEDING_SPEECH]);
  assert.deepEqual(partsNeedingSpeech({ apps: VERIFY_APPS, partIds: ['word#main'] }), []);
  assert.deepEqual(partsNeedingSpeech({ partIds: ['recorder#*'] }), [...PARTS_NEEDING_SPEECH]);
});

test('a gate run that includes the check with a real voice is refused at its start without the speech packs or the key the engine is checked with', () => {
  const all = { verifying: true, apps: VERIFY_APPS, partIds: [], office: OFFICE, releaseKeys: 1 };
  const withSpeech = releaseFolder({ ids: [...GATE_LANGUAGE_PACKS, ...GATE_SPEECH_PACKS] });
  assert.match(speechRefusal({ ...all, release: undefined }), /^The gate will not start: its window checks include recorder#verifyRealVoice, which speaks with a real added voice\. RUTBA_GATE_PACKS is not set/);
  assert.match(speechRefusal({ ...all, release: releaseFolder() }), /has no signed speech-engine, speech-en-kokoro, speech-kokoro-82m-q8, voice-kokoro-af-heart/);
  assert.match(speechRefusal({ ...all, release: withSpeech, releaseKeys: 0 }), /knows the public half of the release pack key: .*pack-keys\.json.*RUTBA_PACK_PUBLIC_KEY/);
  assert.equal(speechRefusal({ ...all, release: withSpeech }), null);
  assert.equal(speechRefusal({ ...all, verifying: false, release: undefined }), null, 'no window checks');
  assert.equal(speechRefusal({ ...all, apps: ['word', 'sheets'], release: undefined }), null, 'no part that speaks');
  assert.equal(speechRefusal({ ...all, partIds: ['recorder#verifyRecorder'], release: undefined }), null, 'another part of the Recorder');
  // The keys a run knows: a build's file, the .pem named, or none.
  const appDir = fresh('app');
  assert.deepEqual(speechReleaseKeys({ appDir, env: {} }), []);
  const named = path.join(appDir, 'release.pem');
  fs.writeFileSync(named, pem(packagesKey));
  assert.equal(speechReleaseKeys({ appDir, env: { RUTBA_PACK_PUBLIC_KEY: named } }).length, 1);
  assert.deepEqual(speechReleaseKeys({ appDir, env: { RUTBA_PACK_PUBLIC_KEY: path.join(appDir, 'nowhere.pem') } }), []);
  fs.mkdirSync(path.join(appDir, 'build', 'out'), { recursive: true });
  fs.writeFileSync(path.join(appDir, 'build', 'out', 'pack-keys.json'), JSON.stringify({ keys: [pem(packagesKey)] }));
  assert.equal(speechReleaseKeys({ appDir, env: {} }).length, 1);
  // The gate asks before it runs anything, and the runner adds them before Electron starts.
  const gate = fs.readFileSync(path.join(ROOT, 'tools/gate.js'), 'utf8');
  const main = gate.slice(gate.indexOf('async function main()'));
  const refused = main.indexOf('speechRefusal(');
  assert.ok(refused > 0, 'the gate asks speechRefusal');
  for (const first of ['windowRun(', "run('node'", 'localTests(', 'budgetsReady(']) {
    const at = main.indexOf(first, main.indexOf('const want = '));
    assert.ok(at < 0 || refused < at, `refused before ${first}`);
  }
  const runner = fs.readFileSync(path.join(ROOT, 'apps/desktop/build/verify-apps-run.js'), 'utf8');
  const added = runner.indexOf('ids: GATE_SPEECH_PACKS');
  assert.ok(added > runner.indexOf('if (speechWanted())') && added < runner.indexOf('spawn(electron'), 'added when wanted, before Electron starts');
});

test('a release pack is zipped as a person adds one: its envelope and the files its manifest lists, nothing else', () => {
  const dir = releaseFolder({ ids: ['lang-ur'] });
  fs.writeFileSync(path.join(dir, 'lang-ur', 'stray.txt'), 'not listed\n');
  const names = readZip(zipOfPack(dir, 'lang-ur')).entries.map((e) => e.name).sort();
  assert.deepEqual(names, ['lang-ur.pack.json', 'lang-ur/manifest.json', 'lang-ur/messages.json']);
});

test('the packs are added through the add-ons service and loaded by the next start; a stranger\'s signature or a changed file is refused', async () => {
  const addons = path.join(fresh('profile with a space'), 'addons');
  const added = await installGatePacks({ release: releaseFolder(), addonsDir: addons, office: OFFICE, contract });
  assert.deepEqual(added, GATE_LANGUAGE_PACKS);
  for (const id of GATE_LANGUAGE_PACKS) assert.ok(fs.existsSync(path.join(addons, `${id}.pack.json`)) && fs.existsSync(path.join(addons, id, 'messages.json')), id);
  // The check run's start, as addons-wiring.js makes it with RUTBA_ADDONS_DIR: it loads all three.
  const start = startAddons({ folder: { dir: addons, scope: 'named' }, contract, office: OFFICE, memoryFile: path.join(fresh('memory'), 's.json'), argv: [] });
  const loaded = [...start.registry.installed].sort();
  assert.deepEqual(loaded, [...GATE_LANGUAGE_PACKS].sort(), JSON.stringify(start.registry.refused));

  // Signed by another key: the service refuses it, and the run stops with the reason.
  await assert.rejects(
    installGatePacks({ release: releaseFolder({ key: crypto.generateKeyPairSync('ed25519') }), addonsDir: path.join(fresh('p'), 'addons'), office: OFFICE, contract }),
    /lang-ur from .* was refused: .*not signed by Rutba/,
  );
  // A file changed after signing.
  const changed = releaseFolder();
  fs.appendFileSync(path.join(changed, 'lang-ar', 'messages.json'), ' ');
  await assert.rejects(installGatePacks({ release: changed, addonsDir: path.join(fresh('p'), 'addons'), office: OFFICE, contract }), /lang-ar from .* was refused/);
  // No folder at all: refused before anything is written.
  const nowhere = path.join(fresh('p'), 'addons');
  await assert.rejects(installGatePacks({ release: undefined, addonsDir: nowhere, office: OFFICE, contract }), /RUTBA_GATE_PACKS is not set/);
  assert.ok(!fs.existsSync(nowhere));
});

test('with RUTBA_GATE_PACKS naming this release\'s packages folder, its signed packs are added with the contract\'s own key', { skip: !process.env.RUTBA_GATE_PACKS && 'RUTBA_GATE_PACKS not set' }, async () => {
  const addons = path.join(fresh('real'), 'addons');
  const office = JSON.parse(fs.readFileSync(path.join(ROOT, 'apps/desktop/package.json'), 'utf8')).version;
  assert.deepEqual(await installGatePacks({ release: process.env.RUTBA_GATE_PACKS, addonsDir: addons, office }), GATE_LANGUAGE_PACKS);
});
