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
import { makeInstallerWords, OUT, UNINSTALLER_WORDS, WORDS } from '../tools/make-installer-words.mjs';
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

/** The words file's words, by language and English, read back from its StrCpy lines. */
function wordsInFile(text) {
  const out = { ar: {}, ur: {}, hi: {} };
  const english = Object.fromEntries([...WORDS, ...UNINSTALLER_WORDS]);
  let lang = null;
  for (const line of text.split('\n')) {
    const at = / == \d+ ; (Arabic|Urdu|Hindi)$/.exec(line);
    if (at) { lang = { Arabic: 'ar', Urdu: 'ur', Hindi: 'hi' }[at[1]]; continue; }
    if (/\$\{EndIf\}/.test(line)) lang = null;
    const w = /^ {4}StrCpy \$rutbaW_(\w+) "(.*)"$/.exec(line);
    if (lang && w) out[lang][english[w[1]]] = w[2].replace(/\$\\r\$\\n/g, '\n').replace(/\$\\"/g, '"').replace(/\$\$/g, '$');
  }
  return out;
}

test('every word on the page is a label, and the words file is what the packs make (D-2026-11-14-16)', async (t) => {
  for (const [, en] of [...WORDS, ...UNINSTALLER_WORDS]) assert.ok(messages.has(en), `"${en}" is a label`);
  const kept = fs.readFileSync(OUT, 'utf8').replace(/\r\n/g, '\n');
  // Packs that carry the file's own words, each entry under an id of its own and found by its source hash, make the file again.
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'installer-words-'));
  try {
    for (const [tag, words] of Object.entries(wordsInFile(kept))) {
      assert.ok(Object.keys(words).length >= WORDS.length - 2, `${tag}: the file has its words`);
      const entries = Object.fromEntries(Object.entries(words).map(([en, text], i) => [`pack.word-${i}`, { text, source: hashOf(keyOf({ message: en })) }]));
      entries['pack.stale'] = { text: 'never shown', source: 'ffffffffffff' };
      fs.mkdirSync(path.join(dir, `lang-${tag}`), { recursive: true });
      fs.writeFileSync(path.join(dir, `lang-${tag}`, 'messages.json'), JSON.stringify(entries));
    }
    assert.equal(await makeInstallerWords(dir), kept, 'the packs make the file as it is');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
  // The packs a release ships, when named: the file is made from them (RELEASING.md).
  if (process.env.RUTBA_PACKS) assert.equal(await makeInstallerWords(process.env.RUTBA_PACKS), kept, 'run node tools/make-installer-words.mjs <packs>');
  else t.diagnostic('RUTBA_PACKS names no packs: the file is held to its shape and to the packs made from its own words');
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
