// The installer's Choose apps page: a box for every app the suite has, an
// Office app ticked on a first install and a further one not, and its words
// in the suite's languages made from the catalogues as they are now.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { APP_ORDER, OFF_UNTIL_CHOSEN } from '../packages/office-formats/src/registry.js';
import { makeInstallerWords, OUT, WORDS } from '../tools/make-installer-words.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const script = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'resources', 'installer.nsh'), 'utf8');
const messages = new Set(JSON.parse(fs.readFileSync(path.join(ROOT, 'packages', 'office-ui', 'src', 'catalogues', 'messages.json'), 'utf8')).messages.map((m) => m.message));

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

test('every word on the page is in the catalogue, and the words file is what the catalogues make now', async () => {
  for (const [, en] of WORDS) assert.ok(messages.has(en), `"${en}" is in the message list`);
  const kept = fs.readFileSync(OUT, 'utf8').replace(/\r\n/g, '\n');
  assert.equal(kept, await makeInstallerWords(), 'run node tools/make-installer-words.mjs');
});

test('the words file is in Urdu, Arabic and Hindi, with nothing NSIS would read as a variable', async () => {
  const text = await makeInstallerWords();
  assert.match(text, /\$0 == 1 ; Arabic\n {4}StrCpy \$rutbaW_title "اختيار التطبيقات"/);
  assert.match(text, /\$0 == 32 ; Urdu/);
  assert.match(text, /\$0 == 57 ; Hindi/);
  for (const line of text.split('\n').filter((l) => l.includes('StrCpy $rutbaW_'))) {
    const value = /"(.*)"$/.exec(line)[1];
    assert.doesNotMatch(value.replace(/\$\$/g, '').replace(/\$\\./g, ''), /\$/, line);
  }
});
