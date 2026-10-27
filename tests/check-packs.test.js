// The release check for app packs (tools/check-packs.mjs, docs/APP-PACKS.md
// 5.6): a core's files sorted by owner, each pack checked alone, and the
// installer's part for each. The packs are made for the test with a key
// thrown away after it, and the core is a folder laid out as its archive is.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { makeKeyPair } from '../tools/make-pack-key.mjs';
import { makePack } from '../tools/make-packs.mjs';
import { makeAppPacks } from '../apps/desktop/build/make-app-packs.js';
import { checkPacks, checkInstaller, installerParts } from '../tools/check-packs.mjs';
import { packFiles } from '../packages/office-shell/src/electron/pack-install.js';

const VERSION = '9.9.9';
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-check-packs-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));
const pair = makeKeyPair();
const privateKey = crypto.createPrivateKey(pair.privatePem);
const root = path.join(tmp, 'packs');
const keysFile = path.join(tmp, 'core', 'build', 'out', 'pack-keys.json');
const installer = path.join(tmp, 'installer-packs.nsh');
// Made with every app's code still in the core, as Mail's is in the core laid out below.
const built = await makeAppPacks({ privateKey, version: VERSION, only: ['mail', 'word', 'ai'], root, keysFile, installerFile: installer, carryCode: [] });

/** A core laid out as its archive is: the keys the build wrote, its version, and a few files of its own and of Mail's. */
const core = path.join(tmp, 'core');
const put = (rel, text) => {
  fs.mkdirSync(path.dirname(path.join(core, rel)), { recursive: true });
  fs.writeFileSync(path.join(core, rel), text);
};
put('package.json', JSON.stringify({ version: VERSION }));
put('main/main.js', '// the core');
put('main/mail.js', '// Mail');
put('build/out/app.js', '// the launcher');
put('build/out/chunks/mail-AB12CD34.js', '// Mail\'s window');
put('build/out/chunks/chunk-EF56GH78.js', '// shared');
put('resources/apps/mail.ico', 'i');

/** A pack made by hand from a built one's manifest: another main module, or other claims. */
async function handMade(id, { main = null, change = (m) => m } = {}) {
  const manifest = JSON.parse(fs.readFileSync(path.join(built.out, packFiles(id).manifest), 'utf8')).manifest;
  const src = path.join(tmp, `hand-${crypto.randomUUID()}`, id);
  fs.cpSync(path.join(built.out, packFiles(id).unpacked), src, { recursive: true });
  fs.mkdirSync(path.join(src, 'main'), { recursive: true });
  fs.writeFileSync(path.join(src, 'main', 'index.js'), main || 'export function register(core) {\n  return core.builtIn(core);\n}\n');
  const { name, needs, uses, ipc, events, tiles, fileTypes, tools, registry } = change(structuredClone(manifest));
  fs.writeFileSync(path.join(src, 'pack.json'), JSON.stringify({ name, needs, uses, main: 'main/index.js', ipc, events, tiles, fileTypes, tools, registry, unpack: ['tiles', 'filetypes'] }));
  const out = path.join(tmp, `packs-${crypto.randomUUID()}`);
  fs.cpSync(built.out, out, { recursive: true });
  for (const name of Object.values(packFiles(id))) fs.rmSync(path.join(out, name), { recursive: true, force: true });
  await makePack({ src, out, version: VERSION, privateKey });
  return out;
}

test('the packs the build makes pass: each alone holds nothing of another app, and an app still in the core is counted there', async () => {
  const r = await checkPacks({ core, packsDir: built.out, installer });
  assert.deepEqual(r.failures, []);
  assert.equal(r.version, VERSION);
  // The core's files by owner: Mail's source, its window chunk and its tile.
  assert.equal(r.owners.mail.files, 3);
  assert.ok(r.owners.core.files >= 4);
  const mail = r.packs.find((p) => p.id === 'mail');
  assert.equal(mail.builtIn, true);
  assert.equal(mail.inCore.files, 3);
  assert.deepEqual(mail.loaded, ['mail']);
  assert.ok(mail.fileTypes > 0 && mail.tiles === 1);
  assert.deepEqual(r.packs.find((p) => p.id === 'word').tools.sort(), ['create_document', 'export_pdf']);
  assert.ok(r.packs.find((p) => p.id === 'ai').tools.includes('list_folders'));
});

test('a pack that carries its own code fails while the core still holds any of its files', async () => {
  const out = await handMade('mail', { main: 'export function register() {\n  return {};\n}\n' });
  const r = await checkPacks({ core, packsDir: out });
  const mail = r.packs.find((p) => p.id === 'mail');
  assert.equal(mail.builtIn, false);
  assert.match(r.failures.join('\n'), /Mail carries its own code, yet the core still holds 3 of its files: /);
  // With Mail's files gone from the core, it passes.
  const lean = path.join(tmp, 'lean-core');
  fs.cpSync(core, lean, { recursive: true });
  for (const f of ['main/mail.js', 'build/out/chunks/mail-AB12CD34.js', 'resources/apps/mail.ico']) fs.rmSync(path.join(lean, f));
  assert.deepEqual((await checkPacks({ core: lean, packsDir: out })).failures, []);
});

test('a pack alone that claims another app\'s kind of file or AI tool fails, and so does a core that carries the window checks', async () => {
  const out = await handMade('word', {
    change: (m) => ({ ...m, fileTypes: [...m.fileTypes, { ...m.fileTypes[0], ext: 'eml' }], tiles: [...m.tiles, { app: 'mail', icon: m.tiles[0].icon }], tools: [...m.tools, 'create_workbook'] }),
  });
  const withCheck = path.join(tmp, 'core-with-check');
  fs.cpSync(core, withCheck, { recursive: true });
  fs.writeFileSync(path.join(withCheck, 'main', 'verify-mail.js'), '// a window check');
  const said = (await checkPacks({ core: withCheck, packsDir: out })).failures.join('\n');
  assert.match(said, /Documents alone: registers \.eml, which is not Documents's kind of file/);
  assert.match(said, /Documents alone: has a tile for Mail/);
  assert.match(said, /Documents alone: names the AI tool create_workbook, which is Worksheets's/);
  assert.match(said, /The core carries the window checks: main\/verify-mail\.js\./);
});

test('the installer\'s part for a pack writes that pack\'s files only and registers its kinds of file only', () => {
  const text = fs.readFileSync(installer, 'utf8');
  const parts = installerParts(text);
  assert.deepEqual(Object.keys(parts).sort(), ['ai', 'mail', 'word']);
  assert.deepEqual(parts.mail.files, ['mail.asar', 'mail.asar.unpacked', 'mail.pack.json']);
  assert.ok(parts.mail.types.includes('eml'));
  assert.deepEqual(checkInstaller(text, ['ai', 'mail', 'word']).failures, []);
  // Mail's part made to write Documents' archive and claim .docx.
  const at = text.indexOf('!macro rutbaWritePack_mail');
  const spoilt = `${text.slice(0, at)}!macro rutbaWritePack_mail\r\n  File "x\\word.asar"\r\n  !insertmacro rutbaAssociate "docx" "x" "x" "x"\r\n${text.slice(text.indexOf('\n', at) + 1)}`;
  const failures = checkInstaller(spoilt, ['ai', 'mail', 'word']).failures.join('\n');
  assert.match(failures, /The installer's part for Mail writes word\.asar\./);
  assert.match(failures, /The installer's part for Mail registers \.docx\./);
  assert.match(checkInstaller(text, ['ai', 'mail', 'word', 'talk']).failures.join('\n'), /The installer has no part for Talk\./);
});

test('a part that carries its pack as a 7-Zip file of its own carries that pack alone, and unpacks it', () => {
  const part = (id, source = `C:\\b\\installer\\${id}.7z`, extract = true) => [
    `!macro rutbaWritePack_${id}`,
    '  SetOutPath "$INSTDIR\\resources\\packs"',
    `  File "/oname=$PLUGINSDIR\\rutba-pack-${id}.7z" "${source}"`,
    ...(extract ? [`  Nsis7z::Extract "$PLUGINSDIR\\rutba-pack-${id}.7z"`] : []),
    `  Delete "$PLUGINSDIR\\rutba-pack-${id}.7z"`,
    `  File "C:\\b\\packs\\${packFiles(id).manifest}"`,
    '!macroend',
  ].join('\r\n');
  const text = `${part('word')}\r\n${part('mail')}\r\n`;
  assert.deepEqual(installerParts(text).mail.carried, [{ name: 'rutba-pack-mail.7z', source: 'C:\\b\\installer\\mail.7z' }]);
  const word = packFiles('word');
  const mail = packFiles('mail');
  const good = { 'C:\\b\\installer\\word.7z': [word.archive], 'C:\\b\\installer\\mail.7z': [mail.archive, mail.unpacked] };
  assert.deepEqual(checkInstaller(text, ['word', 'mail'], good).failures, []);
  assert.deepEqual(checkInstaller(text, ['word', 'mail']).failures, [], 'by its name where the file cannot be read');
  // Mail's 7-Zip file holding Documents' archive too.
  const spoilt = checkInstaller(text, ['word', 'mail'], { ...good, 'C:\\b\\installer\\mail.7z': [mail.archive, word.archive] }).failures.join('\n');
  assert.ok(spoilt.includes(`The installer's part for Mail unpacks ${word.archive}.`), spoilt);
  // Mail's part carrying Documents' 7-Zip file.
  const swapped = checkInstaller(`${part('word')}\r\n${part('mail', 'C:\\b\\installer\\word.7z')}\r\n`, ['word', 'mail']).failures.join('\n');
  assert.match(swapped, /The installer's part for Mail carries word\.7z as rutba-pack-mail\.7z\./);
  // Carried and never unpacked.
  const kept = checkInstaller(`${part('word')}\r\n${part('mail', undefined, false)}\r\n`, ['word', 'mail']).failures.join('\n');
  assert.match(kept, /carries rutba-pack-mail\.7z and never unpacks it/);
});
