// The app packs a release carries (apps/desktop/build/make-app-packs.js,
// docs/APP-PACKS.md 5.3): made and signed from the core's own description of
// each app, loaded by the core's loader with the key the build writes into
// the core, carried by the release as a file per pack that Add or remove
// apps unpacks and checks, and written by the installer a pack at a time.
// The key is made for the test and thrown away with its folder. Where the
// NSIS compiler electron-builder uses is on this computer, the installer's
// part is compiled too, for the installer and for the uninstaller.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { makeKeyPair } from '../tools/make-pack-key.mjs';
import { makeAppPacks, installerScript, installerArchives } from '../apps/desktop/build/make-app-packs.js';
import { describePack, planPacks, readPackKeys, readPackSizes } from '../apps/desktop/main/app-packs.js';
import { unpackPackFile, packFileName } from '../apps/desktop/main/packs-service.js';
import { stagePack } from '../packages/office-shell/src/electron/pack-install.js';
import { readZip, writeZip, ZipEntry } from '../packages/ooxml/src/zip.js';
import { PACKS } from '../packages/office-shell/src/contract.js';

const VERSION = '9.9.9';
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-app-packs-build-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));
const keys = makeKeyPair();
const privateKey = crypto.createPrivateKey(keys.privatePem);
const root = path.join(tmp, 'packs');
const keysFile = path.join(tmp, 'out', 'pack-keys.json');
const installerFile = path.join(root, 'installer-packs.nsh');
const release = path.join(tmp, 'release');
const built = await makeAppPacks({ privateKey, version: VERSION, only: ['mail', 'word', 'ai'], release, root, keysFile, installerFile });

test('each pack is made from the core\'s description of its app, and the core loads it with the key the build gives it', () => {
  const plan = planPacks({ dir: built.out, keys: readPackKeys(keysFile), version: VERSION });
  assert.deepEqual(plan.refused, []);
  assert.deepEqual(plan.ids().sort(), ['ai', 'mail', 'word']);
  const mail = plan.packs.find((p) => p.id === 'mail');
  const said = describePack('mail');
  assert.deepEqual(mail.ipc, said.ipc);
  assert.deepEqual(mail.fileTypes.map((t) => t.ext), said.fileTypes.map((t) => t.ext));
  assert.equal(path.relative(built.out, mail.tiles[0].icon), path.join('mail.asar.unpacked', 'tiles', 'mail.ico'), 'the tile as the taskbar reads it, on disk');
  assert.ok(fs.existsSync(path.join(built.out, 'mail.asar.unpacked', 'tiles', 'mail.ico')), 'the tile is on disk, where the taskbar reads it');
  assert.ok(fs.existsSync(path.join(built.out, 'mail.asar.unpacked', 'filetypes', 'mail.ico')));
  // AI Tools' own tools are named in its manifest, read from the tools' list.
  assert.ok(plan.packs.find((p) => p.id === 'ai').tools.includes('list_folders'));
  assert.ok(plan.packs.find((p) => p.id === 'word').tools.includes('create_document'));
  const sizes = readPackSizes(keysFile);
  assert.deepEqual(Object.keys(sizes).sort(), ['ai', 'mail', 'word']);
  assert.ok(sizes.mail > 0);
  // The core's file holds the public half only.
  assert.doesNotMatch(fs.readFileSync(keysFile, 'utf8'), /PRIVATE KEY/);
});

test('a release\'s file for a pack unpacks into what Add or remove apps checks, and a file that holds more is refused', () => {
  const file = path.join(release, packFileName(VERSION, 'mail'));
  assert.ok(fs.existsSync(file), 'Rutba-Office-<version>-<id>.pack');
  const into = unpackPackFile(fs.readFileSync(file), 'mail', path.join(tmp, 'unpacked-mail'));
  const pack = stagePack({ from: into, id: 'mail', staging: path.join(tmp, 'staged-mail'), publicKey: readPackKeys(keysFile), version: VERSION });
  assert.equal(pack.id, 'mail');

  // The same file with one more entry, out of the folder.
  const { entries } = readZip(fs.readFileSync(file));
  const extra = new ZipEntry({ name: '../evil.js', method: 8, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x21, flags: 0, externalAttrs: 0, comment: Buffer.alloc(0) });
  extra.data = Buffer.from('nothing');
  assert.throws(() => unpackPackFile(writeZip([...entries, extra]), 'mail', path.join(tmp, 'bad')), /holds something a pack may not: \.\.\/evil\.js\. It was not installed\./);
  // Another pack's file under this one's name is checked and refused.
  const word = fs.readFileSync(path.join(release, packFileName(VERSION, 'word')));
  assert.throws(() => unpackPackFile(word, 'mail', path.join(tmp, 'wrong')), /holds something a pack may not/);
});

test('the installer writes a pack\'s files only when it is chosen, its manifest last, and registers only its kinds of file', () => {
  const text = fs.readFileSync(installerFile, 'utf8');
  const mail = text.slice(text.indexOf('!macro rutbaWritePack_mail'), text.indexOf('!macroend', text.indexOf('!macro rutbaWritePack_mail')));
  const order = ['mail.asar"', 'mail.asar.unpacked"', 'mail.pack.json"'].map((s) => mail.indexOf(s));
  assert.ok(order.every((i) => i > 0) && order[0] < order[1] && order[1] < order[2], 'archive, its folder, then the manifest');
  for (const ext of describePack('mail').fileTypes.map((t) => t.ext)) assert.match(mail, new RegExp(`rutbaAssociate "${ext}" `), ext);
  assert.doesNotMatch(mail, /rutbaAssociate "docx"/);
  assert.match(text, /\$\{If\} \$rutbaPack_mail == "1"\r\n {4}!insertmacro rutbaWritePack_mail/);
  assert.match(text, /WriteRegStr HKCU "Software\\Rutba Office" "Packs" "\$rutbaPacksList"/);
  // The choice: the command line, else what is installed, else every pack.
  assert.ok(text.indexOf('"/PACKS="') < text.indexOf('"Packs"') && text.indexOf('"Packs"') < text.indexOf('${RUTBA_ALL_PACKS}'));
  // A pack new in this version is written on an update too: the packs the copy
  // before knew are read only when its Packs was, and each pack missing from them is chosen.
  const decide = text.slice(text.indexOf('!macro rutbaPacksDecide'), text.indexOf('!macroend', text.indexOf('!macro rutbaPacksDecide')));
  assert.match(decide, /\$\{Else\}\r\n {6}StrCpy \$R4 "1"/);
  assert.match(decide, /\$\{If\} \$R4 == "1"\r\n {4}ReadRegStr \$R3 HKCU "Software\\Rutba Office" "PacksKnown"/);
  for (const id of ['ai', 'mail', 'word']) assert.match(decide, new RegExp(`!insertmacro rutbaPackNew ${id}\\r\\n`));
  assert.ok(decide.lastIndexOf('rutbaPackDecide') < decide.indexOf('rutbaPackNew'), 'new packs added after the choice is read');
  // With no PacksKnown, every pack is taken as known.
  assert.match(decide, /"PacksKnown"\r\n {4}\$\{If\} \$\{Errors\}\r\n {4}\$\{OrIf\} \$R3 == ""\r\n {6}StrCpy \$R3 "\$\{RUTBA_ALL_PACKS\}"/);
  assert.match(text, /WriteRegStr HKCU "Software\\Rutba Office" "PacksKnown" "\$\{RUTBA_ALL_PACKS\}"/);
  assert.deepEqual(/!define RUTBA_ALL_PACKS "([a-z,]+)"/.exec(text)[1].split(',').sort(), ['ai', 'mail', 'word']);
  // The uninstaller lets go of every kind of file any pack registers.
  const forget = text.slice(text.indexOf('!macro rutbaForgetTypes'));
  assert.match(forget, /rutbaUnassociate "eml" /);
  assert.match(forget, /rutbaUnassociate "docx" /);
  // The installer's page and Choose apps read the packs chosen.
  const page = fs.readFileSync(path.join(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, '$1')), '..', 'apps', 'desktop', 'resources', 'installer.nsh'), 'utf8');
  assert.match(page, /!include "\$\{PROJECT_DIR\}\\build\\packs\\installer-packs\.nsh"/);
  assert.match(page, /Page custom rutbaPacksPageShow rutbaPacksPageLeave\r?\n {2}Page custom rutbaAppsPageShow rutbaAppsPageLeave/);
  assert.match(page, /!insertmacro rutbaPacksDecide/);
  assert.match(page, /!insertmacro rutbaWritePacks/);
  assert.match(page, /!macro customUnInstall\r?\n {2}!insertmacro rutbaForgetTypes/);
});

/** The NSIS compiler electron-builder downloaded, when it is on this computer. */
function makensis() {
  const cache = path.join(process.env.LOCALAPPDATA || '', 'electron-builder', 'Cache');
  if (!fs.existsSync(cache)) return null;
  for (const d of fs.readdirSync(cache).filter((n) => /^nsis-\d/.test(n))) {
    for (const sub of fs.readdirSync(path.join(cache, d))) {
      const exe = path.join(cache, d, sub, 'makensis.exe');
      if (fs.existsSync(exe)) return exe;
    }
  }
  return null;
}

test('the installer\'s part compiles, for the installer and the uninstaller, with no warning', { skip: process.platform !== 'win32' || !makensis() ? 'no NSIS compiler here' : false }, () => {
  // Every pack, so every macro is compiled: stand-in files where the build puts them.
  const dir = path.join(tmp, 'every');
  for (const { id } of PACKS) {
    fs.mkdirSync(path.join(dir, `${id}.asar.unpacked`, 'tiles'), { recursive: true });
    fs.writeFileSync(path.join(dir, `${id}.asar`), 'a');
    fs.writeFileSync(path.join(dir, `${id}.asar.unpacked`, 'tiles', `${id}.ico`), 'i');
    fs.writeFileSync(path.join(dir, `${id}.pack.json`), '{}');
  }
  const script = installerScript({ dir, packs: PACKS.map(({ id }) => ({ ...describePack(id), id })) });
  const include = path.join(tmp, 'all-packs.nsh');
  fs.writeFileSync(include, script);
  const ids = PACKS.map((p) => p.id);
  const head = ['Unicode true', '!include "LogicLib.nsh"', '!include "FileFunc.nsh"', '!include "WordFunc.nsh"', '!include "nsDialogs.nsh"', 'Name "check"', 'RequestExecutionLevel user', 'InstallDir "$TEMP\\rutba-packs-check"'];
  const installer = [
    ...head,
    `OutFile "${path.join(tmp, 'installer.exe')}"`,
    'Var appExe',
    ...ids.map((id) => `Var rutbaW_${id}`),
    `!include "${include}"`,
    'Page custom show leave',
    'Page instfiles',
    'Function .onInit',
    ...ids.map((id) => `  StrCpy $rutbaW_${id} "${id}"`),
    '  !insertmacro rutbaPacksDecide',
    'FunctionEnd',
    'Function show',
    '  nsDialogs::Create 1018',
    '  Pop $0',
    '  !insertmacro rutbaPackBoxes',
    '  nsDialogs::Show',
    'FunctionEnd',
    'Function leave',
    '  StrCpy $2 0',
    '  !insertmacro rutbaPackReads',
    'FunctionEnd',
    'Section',
    '  StrCpy $appExe "$INSTDIR\\Rutba Office.exe"',
    '  !insertmacro rutbaWritePacks',
    'SectionEnd',
  ];
  const uninstaller = [
    '!define BUILD_UNINSTALLER',
    ...head,
    `OutFile "${path.join(tmp, 'uninstaller-maker.exe')}"`,
    `!include "${include}"`,
    'Section',
    '  WriteUninstaller "$TEMP\\rutba-packs-check-uninstall.exe"',
    'SectionEnd',
    'Section "Uninstall"',
    '  !insertmacro rutbaForgetTypes',
    'SectionEnd',
  ];
  for (const [name, lines] of [['installer', installer], ['uninstaller', uninstaller]]) {
    const file = path.join(tmp, `${name}.nsi`);
    fs.writeFileSync(file, `${lines.join('\r\n')}\r\n`);
    try {
      execFileSync(makensis(), ['/V2', '/WX', '/INPUTCHARSET', 'UTF8', file], { encoding: 'utf8', windowsHide: true });
    } catch (err) {
      assert.fail(`${name}: ${err.stdout || ''}${err.stderr || err.message}`);
    }
  }
});

test('a pack carried 7-Zipped is stored as it is, unpacked into the packs folder, and its manifest still written last', () => {
  const dir = path.join(tmp, 'seven');
  const packs = ['mail', 'word'].map((id) => ({ ...describePack(id), id }));
  const script = installerScript({ dir: built.out, packs, archives: { mail: path.join(dir, 'mail.7z') } });
  const mail = script.slice(script.indexOf('!macro rutbaWritePack_mail'), script.indexOf('!macroend', script.indexOf('!macro rutbaWritePack_mail')));
  const steps = ['SetCompress off', `File "/oname=$PLUGINSDIR\\rutba-pack-mail.7z" "${path.join(dir, 'mail.7z')}"`, 'SetCompress "${COMPRESS}"', 'Nsis7z::Extract "$PLUGINSDIR\\rutba-pack-mail.7z"', 'Delete "$PLUGINSDIR\\rutba-pack-mail.7z"', 'mail.pack.json"'].map((s) => mail.indexOf(s));
  assert.ok(steps.every((i) => i > 0) && steps.every((i, n) => n === 0 || i > steps[n - 1]), JSON.stringify(steps));
  assert.doesNotMatch(mail, /mail\.asar"/, 'the archive is not carried a second time');
  // A pack with no 7-Zip archive is carried as before.
  const word = script.slice(script.indexOf('!macro rutbaWritePack_word'), script.indexOf('!macroend', script.indexOf('!macro rutbaWritePack_word')));
  assert.match(word, /File ".*word\.asar"/);
  assert.doesNotMatch(word, /Nsis7z/);
});

/** The folder of NSIS plugins electron-builder downloaded, when it is on this computer. */
function nsisPlugins() {
  const cache = path.join(process.env.LOCALAPPDATA || '', 'electron-builder', 'Cache');
  if (!fs.existsSync(cache)) return null;
  for (const d of fs.readdirSync(cache).filter((n) => /^nsis-resources-/.test(n))) {
    for (const sub of fs.readdirSync(path.join(cache, d))) {
      const plugins = path.join(cache, d, sub, 'plugins', 'x86-unicode');
      if (fs.existsSync(path.join(plugins, 'nsis7z.dll'))) return plugins;
    }
  }
  return null;
}

test('each pack is 7-Zipped whole, and the installer part that unpacks it compiles', { skip: process.platform !== 'win32' || !makensis() || !nsisPlugins() ? 'no NSIS compiler or 7-Zip here' : false }, async () => {
  const into = path.join(tmp, 'seven-zipped');
  const archives = await installerArchives(built.out, ['mail', 'word'], into);
  assert.deepEqual(Object.keys(archives).sort(), ['mail', 'word']);
  for (const id of ['mail', 'word']) {
    assert.ok(fs.statSync(archives[id]).size > 0);
    assert.ok(fs.statSync(archives[id]).size < fs.statSync(path.join(built.out, `${id}.asar`)).size, `${id}: smaller than its archive`);
  }
  const packs = ['mail', 'word', 'ai'].map((id) => ({ ...describePack(id), id }));
  const include = path.join(tmp, 'seven-packs.nsh');
  fs.writeFileSync(include, installerScript({ dir: built.out, packs, archives }));
  const ids = packs.map((p) => p.id);
  const lines = [
    'Unicode true', '!include "LogicLib.nsh"', '!include "FileFunc.nsh"', '!include "WordFunc.nsh"', '!include "nsDialogs.nsh"',
    `!addplugindir /x86-unicode "${nsisPlugins()}"`, '!define COMPRESS auto', 'SetCompressor zlib',
    'Name "check"', 'RequestExecutionLevel user', 'InstallDir "$TEMP\\rutba-packs-seven"', `OutFile "${path.join(tmp, 'seven.exe')}"`,
    'Var appExe', ...ids.map((id) => `Var rutbaW_${id}`), `!include "${include}"`,
    'Page custom show leave', 'Page instfiles',
    'Function .onInit', ...ids.map((id) => `  StrCpy $rutbaW_${id} "${id}"`), '  !insertmacro rutbaPacksDecide', 'FunctionEnd',
    'Function show', '  nsDialogs::Create 1018', '  Pop $0', '  !insertmacro rutbaPackBoxes', '  nsDialogs::Show', 'FunctionEnd',
    'Function leave', '  StrCpy $2 0', '  !insertmacro rutbaPackReads', 'FunctionEnd',
    'Section', '  StrCpy $appExe "$INSTDIR\\Rutba Office.exe"', '  !insertmacro rutbaWritePacks', 'SectionEnd',
  ];
  const file = path.join(tmp, 'seven.nsi');
  fs.writeFileSync(file, `${lines.join('\r\n')}\r\n`);
  try {
    execFileSync(makensis(), ['/V2', '/WX', '/INPUTCHARSET', 'UTF8', file], { encoding: 'utf8', windowsHide: true });
  } catch (err) {
    assert.fail(`${err.stdout || ''}${err.stderr || err.message}`);
  }
  // The installer holds the 7-Zip archives as they are, not compressed again.
  const exe = fs.statSync(path.join(tmp, 'seven.exe')).size;
  assert.ok(exe >= fs.statSync(archives.mail).size + fs.statSync(archives.word).size, `${exe} bytes`);
});
