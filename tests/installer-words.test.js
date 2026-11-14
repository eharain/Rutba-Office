// The installer's Choose apps page: a box for every app the suite has, an
// Office app ticked on a first install and a further one not, and its words
// in the suite's languages made from their language packs.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { APP_ORDER, OFF_UNTIL_CHOSEN } from '../packages/office-formats/src/registry.js';
import { makeInstallerWords, packsCopy, wordHashes, COPY, OUT, UNINSTALLER_WORDS, WORDS } from '../tools/make-installer-words.mjs';
import { keyOf, hashOf } from '../packages/office-shell/src/electron/language-packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const script = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'resources', 'installer.nsh'), 'utf8');
const messages = new Set(Object.values(JSON.parse(fs.readFileSync(path.join(ROOT, 'packages', 'office-ui', 'src', 'catalogues', 'labels.json'), 'utf8')).labels).map((d) => d.english));

test('the page has a box for every app, read when it is left, ticked on a first install for an Office app only', () => {
  for (const key of APP_ORDER) {
    const box = new RegExp(`!insertmacro rutbaAppBox ${key} "\\$rutbaW_${key}" \\S+ \\S+ (\\d)`).exec(script);
    assert.ok(box, `${key} has a box`);
    assert.equal(box[1], OFF_UNTIL_CHOSEN.includes(key) ? '1' : '0', `${key} is ${OFF_UNTIL_CHOSEN.includes(key) ? 'a further app' : 'an Office app'}`);
    assert.ok(script.includes(`!insertmacro rutbaAppRead ${key}\n`) || script.includes(`!insertmacro rutbaAppRead ${key}\r\n`), `${key} is read`);
    assert.ok(script.includes(`Var rutbaBox_${key}`), `${key} has its variable`);
    assert.ok(WORDS.some(([v]) => v === key), `${key} has its word`);
  }
});

test('the page opens on the apps on now: AppsOn read, and written back with AppsOff', () => {
  assert.match(script, /ReadRegStr \$rutbaAppsOnBefore HKCU "Software\\Rutba Office" "AppsOn"/);
  assert.match(script, /WriteRegStr HKCU "Software\\Rutba Office" "AppsOn" "\$rutbaAppsOn"/);
  assert.match(script, /WriteRegStr HKCU "Software\\Rutba Office" "AppsOff" "\$rutbaAppsOff"/);
  // Ticked from AppsOn when it is there, before the older AppsOff.
  const box = script.slice(script.indexOf('!macro rutbaAppBox'), script.indexOf('!macroend', script.indexOf('!macro rutbaAppBox')));
  assert.ok(box.indexOf('$rutbaAppsOnKnown') < box.indexOf('$rutbaAppsKnown'), 'AppsOn decides first');
});

test('every word on the page is a label, and the words file is what the packs this release ships make (D-2026-11-14-16)', async (t) => {
  for (const [, en] of [...WORDS, ...UNINSTALLER_WORDS]) assert.ok(messages.has(en), `"${en}" is a label`);
  const kept = fs.readFileSync(OUT, 'utf8').replace(/\r\n/g, '\n');
  // The copy Office carries of the packs the file was made from: the file is what it makes, word for word.
  const copy = JSON.parse(fs.readFileSync(COPY, 'utf8'));
  assert.equal(await makeInstallerWords(copy), kept, 'the file is not what its packs make: run node tools/make-installer-words.mjs <packages>/languages/packs');
  // Those packs are this release's, and every word in them was made from the English the label has now.
  const version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
  const now = new Set(wordHashes().values());
  for (const [tag, pack] of Object.entries(copy.packs)) {
    assert.equal(pack.id, `lang-${tag}`);
    assert.equal(pack.version, version, `lang-${tag} is the ${pack.version} pack, and this is Office ${version}: make the file again from the packs ${version} ships`);
    assert.match(pack.messages, /^[0-9a-f]{64}$/);
    for (const hash of Object.keys(pack.words)) assert.ok(now.has(hash), `lang-${tag} has a word for English the page no longer shows (${hash}): make the file again`);
    assert.ok(Object.keys(pack.words).length >= WORDS.length - 6, `lang-${tag}: the copy has the page's words`);
  }
  // A pack read from its folder makes the same file: each word found by its source hash, never by its id.
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'installer-words-'));
  try {
    for (const [tag, pack] of Object.entries(copy.packs)) {
      const entries = Object.fromEntries(Object.entries(pack.words).map(([source, text], i) => [`pack.word-${i}`, { text, source }]));
      entries['pack.stale'] = { text: 'never shown', source: 'ffffffffffff' };
      fs.mkdirSync(path.join(dir, `lang-${tag}`), { recursive: true });
      fs.writeFileSync(path.join(dir, `lang-${tag}`, 'messages.json'), JSON.stringify(entries));
      fs.writeFileSync(path.join(dir, `lang-${tag}`, 'manifest.json'), JSON.stringify({ id: pack.id, version: pack.version, revision: pack.revision }));
    }
    assert.equal(await makeInstallerWords(dir), kept, 'the packs make the file as it is');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
  // The packs themselves, when named (RELEASING.md): the copy is theirs, file for file.
  if (process.env.RUTBA_PACKS) {
    assert.deepEqual(packsCopy(process.env.RUTBA_PACKS), copy, 'run node tools/make-installer-words.mjs <packs>');
    assert.equal(await makeInstallerWords(process.env.RUTBA_PACKS), kept, 'run node tools/make-installer-words.mjs <packs>');
  } else t.diagnostic('RUTBA_PACKS names no packs: the file is held to the copy of the packs it was made from');
});

test('the words file is in Urdu, Arabic and Hindi, with nothing NSIS would read as a variable', async () => {
  const text = fs.readFileSync(OUT, 'utf8').replace(/\r\n/g, '\n');
  assert.match(text, /\$0 == 1 ; Arabic\n {4}StrCpy \$rutbaW_title "اختيار التطبيقات"/);
  assert.match(text, /\$0 == 32 ; Urdu/);
  assert.match(text, /\$0 == 57 ; Hindi/);
  for (const line of text.split('\n').filter((l) => l.includes('StrCpy $rutbaW_'))) {
    const value = /"(.*)"$/.exec(line)[1];
    assert.doesNotMatch(value.replace(/\$\$/g, '').replace(/\$\\./g, ''), /\$/, line);
  }
});

test('removing Rutba Office asks whether the voices and languages added go too, never on an update or silently', () => {
  const un = script.slice(script.indexOf('!macro customUnInstall'), script.indexOf('!macroend', script.indexOf('!macro customUnInstall')));
  assert.match(un, /\$\{IfNot\} \$\{isUpdated\}/, 'not on an update');
  assert.match(un, /MessageBox MB_YESNO\|MB_ICONQUESTION\|MB_DEFBUTTON2 "\$rutbaW_udrop[^"]*\$R4" \/SD IDNO IDYES/, 'No unless asked, and No when silent');
  assert.match(un, /\$\{GetSize\} "\$R0"/, 'the room they take is said');
  assert.match(un, /WordReplace\} "\$rutbaW_ukeep" "\{size\}"/, 'in the sentence');
  assert.ok(un.indexOf('RMDir /r "$R0"') > un.indexOf('MessageBox'), 'removed only after Yes');
  assert.match(script, /!ifdef BUILD_UNINSTALLER\r?\n {2}!insertmacro rutbaUninstallerWords/, 'its words are in the uninstaller');
  assert.ok(UNINSTALLER_WORDS.some(([, en]) => en.includes('{size}')), 'the size has its place in the sentence');
});
