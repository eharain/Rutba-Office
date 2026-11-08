// The Locker's keys: what opens them, what a wrong one does (nothing), and
// that the folder's own key file is enough on another computer. And its
// rules: gitignore's wording, deciding what is locked and saying why.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createKeys, FOLDER_KEYS, recoveryBytes, passwordBytes } from '../apps/desktop/main/extras/locker/keys.js';
import { parseRules, decide } from '../apps/desktop/main/extras/locker/rules.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rl-keys-'));

test('a locker made with a password opens with it, and with the recovery key, and with nothing else', async () => {
  const dir = tmp();
  const k = createKeys({ dir, cost: CHEAP });
  const { recovery } = await k.setup('correct horse');
  assert.ok(k.isOpen());
  k.lock();
  assert.ok(!k.isOpen());
  await k.unlock({ password: 'correct horse' });
  assert.ok(k.isOpen());
  k.lock();
  await k.unlock({ recovery });
  assert.ok(k.isOpen(), 'the recovery key opens it');
  k.lock();
  await assert.rejects(() => k.unlock({ password: 'wrong' }), (e) => e.code === 'bad-password');
  await assert.rejects(() => k.unlock({ recovery: 'AAAA-AAAA' }), (e) => e.code === 'bad-recovery');
  assert.ok(!k.isOpen());
});

test('a wrong password changes nothing on disk', async () => {
  const dir = tmp();
  const k = createKeys({ dir, cost: CHEAP });
  await k.setup('secret');
  k.lock();
  const snapshot = () => fs.readFileSync(path.join(dir, 'keyring.json'));
  const before = snapshot();
  for (const guess of ['a', 'b', 'c']) await assert.rejects(() => k.unlock({ password: guess }), (e) => e.code === 'bad-password');
  assert.ok(snapshot().equals(before), 'the keyring is byte for byte the same');
  await k.unlock({ password: 'secret' });
});

test('the keyring holds no key in the clear, and a password is the same bytes however it is typed', async () => {
  const dir = tmp();
  const k = createKeys({ dir, cost: CHEAP });
  await k.setup('p@ss');
  const text = fs.readFileSync(path.join(dir, 'keyring.json'), 'utf8');
  assert.ok(!text.includes('p@ss'));
  assert.ok(passwordBytes('é').equals(passwordBytes('é')), 'one letter, two ways of typing it');
  const parsed = JSON.parse(text);
  for (const u of parsed.unlockers) assert.ok(u.iv && u.tag && u.data, `${u.kind} is sealed`);
  assert.deepEqual(parsed.unlockers.map((u) => u.kind), ['password', 'recovery']);
});

test('changing the password: the old one stops, the new one works, and a failed change leaves it as it was', async () => {
  const dir = tmp();
  const k = createKeys({ dir, cost: CHEAP });
  const { recovery } = await k.setup('one');
  await assert.rejects(() => k.changePassword('wrong', 'two'), (e) => e.code === 'bad-password');
  k.lock();
  await k.unlock({ password: 'one' });
  await k.changePassword('one', 'two');
  k.lock();
  await assert.rejects(() => k.unlock({ password: 'one' }), (e) => e.code === 'bad-password');
  await k.unlock({ password: 'two' });
  k.lock();
  await k.unlock({ recovery });
  assert.ok(k.isOpen(), 'the recovery key still opens it after a password change');
});

test('the recovery key is a text a person can type, with or without its dashes, and replacing it retires the old one', async () => {
  const dir = tmp();
  const k = createKeys({ dir, cost: CHEAP });
  const first = await k.setup('pw');
  assert.match(first.recovery, /^[A-Z2-9]{4}(-[A-Z2-9]{1,4})+$/);
  assert.equal(recoveryBytes(first.recovery.toLowerCase().replace(/-/g, ' ')).length, 32);
  const next = await k.replaceRecovery('pw');
  k.lock();
  await assert.rejects(() => k.unlock({ recovery: first.recovery }), (e) => e.code === 'bad-recovery');
  await k.unlock({ recovery: next.recovery });
});

test('Windows Hello is a way in only with its secret from this computer', async () => {
  const dir = tmp();
  const k = createKeys({ dir, cost: CHEAP });
  await k.setup('pw');
  const secret = Buffer.alloc(32, 5);
  await k.enableHello(secret);
  k.lock();
  await assert.rejects(() => k.unlock({ secret: Buffer.alloc(32, 6) }), (e) => e.code === 'bad-hello');
  await k.unlock({ secret });
  assert.deepEqual((await k.ways()).sort(), ['hello', 'password', 'recovery']);
  await k.disableHello();
  k.lock();
  await assert.rejects(() => k.unlock({ secret }), (e) => e.code === 'bad-hello');
});

test('a folder key survives in the profile\'s store and in the folder\'s own file, and either alone is enough', async () => {
  const dir = tmp();
  const k = createKeys({ dir, cost: CHEAP });
  const { recovery } = await k.setup('pw');
  const id = 'aa'.repeat(16);
  const key = Buffer.alloc(32, 9);
  k.rememberFolderKey(id, key);
  const file = k.wrapForFolder(key, id);
  const raw = JSON.stringify(file);
  assert.ok(!raw.includes(key.toString('hex')), 'no key in the clear');
  k.lock();
  await k.unlock({ password: 'pw' });
  assert.ok(k.folderKey(id).equals(key), 'from the profile\'s store');

  // Another computer: a different profile, only the file and the password.
  const other = createKeys({ dir: tmp(), cost: CHEAP });
  await other.setup('someone else');
  assert.ok((await other.openFolderKeys(file, { password: 'pw' })).equals(key), 'from the file, with the password');
  assert.ok((await other.openFolderKeys(file, { recovery })).equals(key), 'from the file, with the recovery key');
  assert.equal(await other.openFolderKeys(file, { password: 'nope' }), null);
  assert.equal(FOLDER_KEYS, '.rutbalocker-keys');
});

// --- rules ---

const sets = (...files) => files.map(([dir, text]) => ({ dir, rules: parseRules(text) }));

test('with no rules, every file is locked but the temporary and system ones, and each says why', () => {
  assert.deepEqual(decide('a/b.pdf', []), { lock: true, why: 'Every file', kind: 'default' });
  for (const name of ['~$budget.docx', 'x.tmp', 'desktop.ini', 'Thumbs.db']) {
    const d = decide(name, []);
    assert.equal(d.lock, false, name);
    assert.ok(d.why, name);
  }
});

test('gitignore wording: anchors, folders, stars, double stars, classes, escapes and comments', () => {
  const R = sets(['', `
# a comment
*.pdf
/top.txt
docs/*.docx
drafts/
photos/**/*.jpg
report[0-9].xlsx
\\#hash.txt
lit\\*eral.txt
trailing.txt
`]);
  const lock = (p) => decide(p, R).lock;
  assert.equal(lock('a/b/c.pdf'), true, 'no slash matches anywhere');
  assert.equal(lock('top.txt'), true);
  assert.equal(lock('sub/top.txt'), false, 'a leading slash anchors');
  assert.equal(lock('docs/a.docx'), true);
  assert.equal(lock('docs/x/a.docx'), false, 'one level only');
  assert.equal(lock('drafts/a/b.bin'), true, 'a folder rule covers everything in it');
  assert.equal(lock('drafts'), false, 'a folder rule is not a file of that name');
  assert.equal(lock('photos/a.jpg'), true);
  assert.equal(lock('photos/x/y/z.jpg'), true);
  assert.equal(lock('report7.xlsx'), true);
  assert.equal(lock('reportx.xlsx'), false);
  assert.equal(lock('#hash.txt'), true, 'an escaped hash');
  assert.equal(lock('lit*eral.txt'), true, 'an escaped star');
  assert.equal(lock('literalxx.txt'), false);
  assert.equal(lock('other.bin'), false, 'with rules, a file no rule names is not locked');
});

test('exceptions: the last rule that matches decides, and they work under a locked folder too', () => {
  const R = sets(['', `
*
!drafts/
!*.tmp2
drafts/keep.docx
`]);
  assert.equal(decide('a.docx', R).lock, true);
  assert.equal(decide('drafts/x.docx', R).lock, false);
  assert.equal(decide('drafts/x.docx', R).why, '!drafts/');
  assert.equal(decide('b.tmp2', R).lock, false);
  assert.equal(decide('drafts/keep.docx', R).lock, true, 'an exception to the exception, unlike git\'s rule');
  assert.equal(decide('drafts/keep.docx', R).why, 'drafts/keep.docx');
});

test('a subfolder\'s file is read after its parent\'s and wins; another folder is not affected', () => {
  const R = sets(['', '*.pdf'], ['sub', '!*.pdf\n*.txt']);
  assert.equal(decide('a.pdf', R).lock, true);
  assert.equal(decide('sub/a.pdf', R).lock, false, 'the subfolder took it back');
  assert.equal(decide('sub/a.txt', R).lock, true);
  assert.equal(decide('other/a.txt', R).lock, false);
  // A rule's folder is relative to its own file.
  const S = sets(['sub', '/top.txt\nd/*.x']);
  assert.equal(decide('sub/top.txt', S).lock, true);
  assert.equal(decide('sub/x/top.txt', S).lock, false);
  assert.equal(decide('sub/d/a.x', S).lock, true);
});

test('the tell-tale files are never locked whatever the rules say', () => {
  const R = sets(['', '*\n~$x.docx']);
  assert.equal(decide('~$x.docx', R).lock, false);
  assert.equal(decide('desktop.ini', R).lock, false);
});
