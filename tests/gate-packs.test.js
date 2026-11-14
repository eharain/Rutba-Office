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
import { GATE_PACKS_ENV, GATE_LANGUAGE_PACKS, PARTS_NEEDING_PACKS, partsNeedingPacks, releaseProblem, packsRefusal, zipOfPack, installGatePacks } from '../apps/desktop/main/gate-packs.js';
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
  assert.deepEqual(partsNeedingPacks({ apps: ['mail', 'pdf'] }), ['mail#verifySmime', 'mail#verifyMailAccounts']);
  assert.deepEqual(partsNeedingPacks({ apps: ['word', 'sheets', 'pdf'] }), []);
  assert.deepEqual(partsNeedingPacks({ apps: VERIFY_APPS, partIds: ['mail#reading', 'video#verifyVideoMp4'] }), []);
  assert.deepEqual(partsNeedingPacks({ apps: ['mail'], partIds: ['mail#reading', 'mail#verifySmime'] }), ['mail#verifySmime']);
});

test('a gate run whose window checks need the packs is refused at its start without a usable release folder, in plain words', () => {
  const good = releaseFolder();
  const all = { verifying: true, apps: VERIFY_APPS, partIds: [], office: OFFICE };
  // Not set.
  const none = packsRefusal({ ...all, release: undefined });
  assert.match(none, /^The gate will not start: its window checks include languages#verifyUrdu, .*mail#verifyMailAccounts, which open windows in Urdu, Arabic or Hindi\. RUTBA_GATE_PACKS is not set/);
  assert.match(none, new RegExp(`dist/v${OFFICE.replace(/\./g, '\\.')}`));
  // Not a folder; another version's; a pack missing.
  assert.match(packsRefusal({ ...all, release: path.join(tmp, 'nowhere') }), /which is not a folder/);
  assert.match(packsRefusal({ ...all, release: releaseFolder({ version: '9.9.9' }) }), /a release folder for another version: this Office is/);
  assert.match(packsRefusal({ ...all, release: releaseFolder({ ids: ['lang-ur', 'lang-ar'] }) }), /has no signed lang-hi/);
  // Usable, or not needed: the run starts.
  assert.equal(packsRefusal({ ...all, release: good }), null);
  assert.equal(packsRefusal({ ...all, verifying: false, release: undefined }), null, 'no window checks');
  assert.equal(packsRefusal({ ...all, apps: ['word', 'pdf'], release: undefined }), null, 'no part that needs them');
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
