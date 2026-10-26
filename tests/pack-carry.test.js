// An app that carries its own code in its pack (docs/APP-PACKS.md 5.1.7, 5.2):
// what the build moves into the pack and leaves out of the core, read off the
// map and the renderer's build (apps/desktop/build/pack-contents.js); the pack
// made with Talk's code in it (build/make-app-packs.js); a module in a pack
// finding the core's packages (office-shell's pack-resolve.js); and the
// release check reading a chunk's owner off what went into it
// (tools/check-packs.mjs). The renderer's build is stood in for: a talk chunk
// and the list of what went into it.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { makeKeyPair } from '../tools/make-pack-key.mjs';
import { mapImports, ROOT } from '../tools/app-map.mjs';
import { packContents, rendererOwners, windowReach, npmOwners } from '../apps/desktop/build/pack-contents.js';
import { makeAppPacks } from '../apps/desktop/build/make-app-packs.js';
import { planPacks, readPackKeys } from '../apps/desktop/main/app-packs.js';
import { readArchive } from '../packages/office-shell/src/electron/packs.js';
import { installPackResolution, isBare } from '../packages/office-shell/src/electron/pack-resolve.js';
import { checkPacks } from '../tools/check-packs.mjs';

const VERSION = '9.9.9';
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-pack-carry-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));
const g = await mapImports();
const TALK_CHUNK = 'chunks/talk-TESTTALK.js';
const meta = {
  [TALK_CHUNK]: ['apps/desktop/renderer/apps/talk.js', 'apps/desktop/renderer/apps/talk/call.js'],
  'chunks/chunk-SHARED01.js': ['apps/desktop/renderer/protect.js'],
  'app.js': ['apps/desktop/renderer/start-page.js'],
};

// Talk's pack, made with its code, signed with a key thrown away after the test.
const rendererOut = path.join(tmp, 'renderer-out');
fs.mkdirSync(path.join(rendererOut, 'chunks'), { recursive: true });
fs.writeFileSync(path.join(rendererOut, ...TALK_CHUNK.split('/')), 'export const talk = 1;\n');
const root = path.join(tmp, 'packs');
const keysFile = path.join(tmp, 'core', 'build', 'out', 'pack-keys.json');
const privateKey = crypto.createPrivateKey(makeKeyPair().privatePem);
const built = await makeAppPacks({ privateKey, version: VERSION, only: ['talk', 'mail'], root, keysFile, installerFile: path.join(root, 'installer-packs.nsh'), carryCode: ['talk'], meta, rendererOut });

test('Talk carries its main-process files and its window chunk, and the core leaves them out with its tile', async () => {
  const { talk } = await packContents(['talk'], { g, meta });
  assert.equal(talk.register, 'apps/desktop/main/talk-pack.js');
  assert.ok(['apps/desktop/main/talk-pack.js', 'apps/desktop/main/talk.js', 'apps/desktop/main/talk-signin.js'].every((f) => talk.main.includes(f)), talk.main.join(', '));
  assert.ok(talk.main.every((f) => /^apps\/desktop\/main\/talk/.test(f)), 'nothing of another app or the core');
  assert.deepEqual(talk.chunks, [TALK_CHUNK]);
  assert.ok(talk.core.includes(`build/out/${TALK_CHUNK}`));
  assert.ok(talk.core.includes('resources/apps/talk.ico'));
  assert.ok(talk.core.includes('main/talk-pack.js'));
});

test('a move is refused, with the import to cut, when a chunk mixes an app\'s code with the core\'s', async () => {
  const mixed = { ...meta, 'chunks/chunk-MIXED001.js': ['apps/desktop/renderer/apps/talk/invite.js', 'apps/desktop/renderer/protect.js'] };
  await assert.rejects(packContents(['talk'], { g, meta: mixed }), /talk cannot move into its pack yet:[\s\S]*chunks\/chunk-MIXED001\.js mixes talk's code \(apps\/desktop\/renderer\/apps\/talk\/invite\.js\) with/);
  // And an output's owners are every owner of what went into it.
  const owners = rendererOwners(g, mixed);
  assert.deepEqual([...owners.get(TALK_CHUNK)], ['talk']);
  assert.ok(owners.get('chunks/chunk-MIXED001.js').has('talk') && owners.get('chunks/chunk-MIXED001.js').size > 1);
});

/** A renderer build with what imports what: the launcher, Talk's window, a chunk only Talk reaches and one the launcher shares. */
const build = {
  outputs: {
    'app.js': ['apps/desktop/renderer/index.js', 'apps/desktop/renderer/start-page.js'],
    'chunks/talk-REACHED1.js': ['apps/desktop/renderer/apps/talk.js'],
    'chunks/chunk-ONLYTALK.js': ['apps/desktop/renderer/apps/talk/call.js', 'apps/desktop/renderer/protect.js'],
    'chunks/chunk-SHARED02.js': ['apps/desktop/renderer/share.js'],
  },
  imports: {
    'app.js': [{ path: 'chunks/talk-REACHED1.js', kind: 'dynamic' }, { path: 'chunks/chunk-SHARED02.js', kind: 'static' }],
    'chunks/talk-REACHED1.js': [{ path: 'chunks/chunk-ONLYTALK.js', kind: 'static' }, { path: 'chunks/chunk-SHARED02.js', kind: 'static' }],
    'chunks/chunk-ONLYTALK.js': [],
    'chunks/chunk-SHARED02.js': [],
  },
  entries: { 'app.js': 'apps/desktop/renderer/index.js', 'chunks/talk-REACHED1.js': 'apps/desktop/renderer/apps/talk.js' },
};

test('a chunk only an app\'s window reaches goes with the app, a file of the core\'s in it too; one the launcher shares stays', async () => {
  const reach = windowReach(build, { talk: 'apps/desktop/renderer/apps/talk.js', mail: 'apps/desktop/renderer/apps/mail.js' });
  assert.deepEqual([...reach.get('app.js')], ['core']);
  assert.deepEqual([...reach.get('chunks/talk-REACHED1.js')], ['talk']);
  assert.deepEqual([...reach.get('chunks/chunk-ONLYTALK.js')], ['talk']);
  assert.deepEqual([...reach.get('chunks/chunk-SHARED02.js')].sort(), ['core', 'talk']);
  const { talk } = await packContents(['talk'], { g, meta: build });
  assert.deepEqual(talk.chunks.sort(), ['chunks/chunk-ONLYTALK.js', 'chunks/talk-REACHED1.js']);
  // Talk's code in the shared chunk is refused, with what reaches it.
  const shared = structuredClone(build);
  shared.outputs['chunks/chunk-SHARED02.js'].push('apps/desktop/renderer/apps/talk/invite.js');
  await assert.rejects(packContents(['talk'], { g, meta: shared }), /chunks\/chunk-SHARED02\.js holds talk's code \(apps\/desktop\/renderer\/apps\/talk\/invite\.js\) and is reached by core too/);
});

test('text recognition carries its engine and models and the code of its package; the core keeps the helpers the windows read', async () => {
  const { ocr } = await packContents(['ocr'], { g, meta: { outputs: {}, imports: {}, entries: {} } });
  assert.deepEqual(ocr.main, ['apps/desktop/main/ocr-pack.js', 'apps/desktop/main/ocr.js']);
  const [engine] = ocr.split;
  assert.equal(engine.name, '@rutba/ocr');
  assert.ok(['package.json', 'src/engine.js', 'src/worker.js', 'src/languages.js', 'vendor/tessdata/eng.traineddata'].every((f) => engine.files.includes(f)));
  assert.ok(ocr.core.includes('node_modules/@rutba/ocr/vendor/tessdata/eng.traineddata'));
  assert.ok(ocr.core.includes('node_modules/@rutba/ocr/src/engine.js'));
  for (const kept of ['languages.js', 'record.js', 'png.js']) assert.ok(!ocr.core.includes(`node_modules/@rutba/ocr/src/${kept}`), kept);
});

/** A renderer build of the named apps' window chunks alone, each reached from the launcher. */
function windowsBuild(ids) {
  const out = (id) => `chunks/${id}-WINDOW01.js`;
  return {
    outputs: Object.fromEntries(ids.map((id) => [out(id), [`apps/desktop/renderer/apps/${id}.js`]])),
    imports: Object.fromEntries(ids.map((id) => [out(id), []])),
    entries: Object.fromEntries(ids.map((id) => [out(id), `apps/desktop/renderer/apps/${id}.js`])),
  };
}

test('Worksheets carries the scripts worker its window starts by name; the image worker Image and Pictures share stays in the core', async () => {
  const { sheets, image } = await packContents(['sheets', 'image'], { g, meta: windowsBuild(['sheets', 'image']) });
  assert.deepEqual(sheets.workers, ['scripts-worker.js']);
  assert.ok(sheets.core.includes('build/out/scripts-worker.js'));
  assert.deepEqual(image.workers, []);
  assert.ok(!image.core.includes('build/out/image-worker.js'));
});

test('a move is refused when a file left in the core imports one the pack would take', async () => {
  const main = new Map(g.graphs.main);
  main.set('apps/desktop/main/main.js', [...(main.get('apps/desktop/main/main.js') || []), { to: 'apps/desktop/main/talk.js', spec: './talk.js', line: 1, kind: 'import-statement' }]);
  const reaching = { ...g, graphs: { ...g.graphs, main } };
  await assert.rejects(packContents(['talk'], { g: reaching, meta }), /apps\/desktop\/main\/main\.js imports apps\/desktop\/main\/talk\.js, which would go with talk's pack/);
  // The document service's own readers are the core's: Documents and Presentations take only their register modules and the show.
  const { word, slides } = await packContents(['word', 'slides'], { g, meta: windowsBuild(['word', 'slides']) });
  assert.deepEqual(word.main, ['apps/desktop/main/word-pack.js']);
  assert.deepEqual(slides.main, ['apps/desktop/main/present.js', 'apps/desktop/main/slides-pack.js']);
});

test('Data Files carries its engine and sample databases; the core keeps the file reader mail merge reads a database with', async () => {
  const { base } = await packContents(['base'], { g, meta: windowsBuild(['base']) });
  const engine = base.split.find((s) => s.name === '@rutba/database');
  assert.ok(engine && ['src/index.js', 'src/samples.js', 'src/base.js'].every((f) => engine.files.includes(f)));
  assert.ok(base.core.includes('node_modules/@rutba/database/src/samples.js'));
  assert.ok(base.core.includes('node_modules/@rutba/templates/data-files/active.js'));
  for (const kept of ['base.js', 'types.js', 'query-design.js']) assert.ok(!base.core.includes(`node_modules/@rutba/database/src/${kept}`), kept);
});

test('an npm package two apps need stays in the core; one only Data Files needs goes with it', () => {
  const owners = npmOwners(g);
  assert.equal(owners.get('pg'), 'base');
  assert.equal(owners.get('mysql2'), 'base');
  assert.equal(owners.get('iconv-lite'), 'core');
});

test('Talk\'s pack holds its code, its main module is its own register module, and moved.json lists what the core leaves out', () => {
  const archive = readArchive(path.join(built.out, 'talk.asar'));
  for (const f of ['main/talk-pack.js', 'main/talk.js', 'main/talk-signin.js', `renderer/${TALK_CHUNK}`]) assert.ok(archive.has(f), f);
  const plan = planPacks({ dir: built.out, keys: readPackKeys(keysFile), version: VERSION });
  assert.deepEqual(plan.refused, []);
  const talk = plan.packs.find((p) => p.id === 'talk');
  assert.equal(path.relative(talk.archive, talk.main).split(path.sep).join('/'), 'main/talk-pack.js');
  // Mail's code is still in the core: its main module registers the core's own.
  const mail = plan.packs.find((p) => p.id === 'mail');
  assert.equal(path.relative(mail.archive, mail.main).split(path.sep).join('/'), 'main/index.js');
  const moved = JSON.parse(fs.readFileSync(path.join(root, 'moved.json'), 'utf8'));
  assert.deepEqual(Object.keys(moved.packs), ['talk']);
  assert.ok(moved.packs.talk.includes('main/talk.js') && moved.packs.talk.includes(`build/out/${TALK_CHUNK}`));
});

/** A core laid out as its archive is, with the files listed. */
function coreWith(name, files) {
  const core = path.join(tmp, name);
  const put = (rel, text) => {
    fs.mkdirSync(path.dirname(path.join(core, rel)), { recursive: true });
    fs.writeFileSync(path.join(core, rel), text);
  };
  put('package.json', JSON.stringify({ version: VERSION }));
  put('build/out/pack-keys.json', fs.readFileSync(keysFile, 'utf8'));
  put('main/main.js', '// the core');
  for (const f of files) put(f, '//');
  return core;
}

test('the release check passes Talk carrying its code with none of it in the core, and fails a chunk of the core that mixes some in', async () => {
  const lean = coreWith('lean-core', ['build/out/app.js', 'build/out/chunks/chunk-SHARED01.js']);
  const r = await checkPacks({ core: lean, packsDir: built.out, outputs: meta });
  assert.deepEqual(r.failures, []);
  const talk = r.packs.find((p) => p.id === 'talk');
  assert.equal(talk.builtIn, false);
  assert.equal(talk.inCore.files, 0);

  // Talk's chunk left in the core, and a shared chunk holding some of Talk's code.
  const outputs = { ...meta, 'chunks/chunk-MIXED001.js': ['apps/desktop/renderer/apps/talk/invite.js', 'apps/desktop/renderer/protect.js'] };
  const fat = coreWith('fat-core', ['build/out/app.js', `build/out/${TALK_CHUNK}`, 'build/out/chunks/chunk-MIXED001.js']);
  const said = (await checkPacks({ core: fat, packsDir: built.out, outputs })).failures.join('\n');
  assert.match(said, /Talk carries its own code, yet the core still holds 1 of its files: build\/out\/chunks\/talk-TESTTALK\.js\./);
  assert.match(said, /Talk carries its own code, yet the core's chunks\/chunk-MIXED001\.js mixes some of it with /);
});

test('a module in a pack that imports a package it does not carry gets the core\'s; a file it should have carried stays an error', async () => {
  assert.ok(isBare('@rutba/office-shell/contract') && isBare('react'));
  assert.ok(!isBare('./talk.js') && !isBare('../x.js') && !isBare('node:fs') && !isBare('electron') && !isBare('C:\\x.js') && !isBare('file:///x.js'));
  const packsDir = path.join(tmp, 'resolve-packs');
  const main = path.join(packsDir, 'demo.asar.unpacked', 'main');
  fs.mkdirSync(main, { recursive: true });
  fs.writeFileSync(path.join(main, 'uses-core.mjs'), "export { PACKS } from '@rutba/office-shell/contract';\n");
  fs.writeFileSync(path.join(main, 'misses-file.mjs'), "export * from './not-here.mjs';\n");
  fs.writeFileSync(path.join(main, 'misses-package.mjs'), "export * from 'no-such-package-anywhere';\n");
  installPackResolution({ packsDir, coreFile: path.join(ROOT, 'apps', 'desktop', 'main', 'main.js') });
  const { PACKS } = await import(pathToFileURL(path.join(main, 'uses-core.mjs')).href);
  const core = await import('@rutba/office-shell/contract');
  assert.equal(PACKS, core.PACKS, 'the very module the core has, not a second copy');
  await assert.rejects(import(pathToFileURL(path.join(main, 'misses-file.mjs')).href), /not-here\.mjs/);
  await assert.rejects(import(pathToFileURL(path.join(main, 'misses-package.mjs')).href), /no-such-package-anywhere/);
});
