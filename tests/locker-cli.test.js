// The Locker's command line: the password only ever on standard input, the
// verbs doing what the app does, a wrong password changing nothing.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const CLI = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'apps', 'desktop', 'main', 'extras', 'locker', 'cli.js');
const tmp = (p) => fs.mkdtempSync(path.join(os.tmpdir(), p));

function run(args, input) {
  return spawnSync(process.execPath, [CLI, '--dir', globalThis.__lockerDir, '--json', ...args], { input, encoding: 'utf8', timeout: 90000, env: { ...process.env } });
}

const PW = 'cli password\n';

test('setup, add, lock, status, why, export, import, log, restore, unlock, change-password', { timeout: 180000 }, () => {
  const dir = tmp('rl-cli-');
  globalThis.__lockerDir = dir;
  const folder = tmp('rl-cli-folder-');
  const out = tmp('rl-cli-out-');

  const setup = run(['setup'], PW);
  assert.equal(setup.status, 0, setup.stderr);
  const { recovery } = JSON.parse(setup.stdout);
  assert.match(recovery, /^[A-Z2-9-]+$/);

  const add = run(['add', folder, '--visible', '--delete'], PW);
  assert.equal(add.status, 0, add.stderr);
  assert.ok(JSON.parse(add.stdout).id);

  const bytes = crypto.randomBytes(70000);
  fs.writeFileSync(path.join(folder, 'a.pdf'), bytes);
  fs.writeFileSync(path.join(folder, '~$a.pdf'), 'x');

  const status1 = JSON.parse(run(['status', folder], PW).stdout);
  assert.equal(status1[0].waiting, 1);

  const lock = run(['lock', folder], PW);
  assert.equal(lock.status, 0, lock.stderr);
  assert.equal(JSON.parse(lock.stdout).locked.length, 1);
  assert.ok(!fs.existsSync(path.join(folder, 'a.pdf')));
  assert.ok(fs.existsSync(path.join(folder, 'a.pdf.rlock')));

  const why = JSON.parse(run(['why', folder, '~$a.pdf'], PW).stdout);
  assert.equal(why.state, 'excluded');
  assert.match(why.why, /temporary/);

  const dest = path.join(out, 'back.pdf');
  assert.equal(run(['export', folder, 'a.pdf', dest], PW).status, 0);
  assert.ok(fs.readFileSync(dest).equals(bytes));

  const v2 = crypto.randomBytes(1000);
  const src = path.join(out, 'n.pdf');
  fs.writeFileSync(src, v2);
  assert.equal(run(['import', folder, src, 'a.pdf'], PW).status, 0);
  const log = JSON.parse(run(['log', folder, 'a.pdf'], PW).stdout);
  assert.equal(log.length, 2);
  assert.equal(run(['restore', folder, 'a.pdf', '1'], PW).status, 0);
  const dest2 = path.join(out, 'back2.pdf');
  assert.equal(run(['export', folder, 'a.pdf', dest2], PW).status, 0);
  assert.ok(fs.readFileSync(dest2).equals(bytes), 'restored the first version');

  const wrong = run(['status', folder], 'not the password\n');
  assert.equal(wrong.status, 3);
  assert.match(wrong.stderr, /password/);

  const changed = run(['change-password'], `${PW}new password\n`);
  assert.equal(changed.status, 0, changed.stderr);
  assert.equal(run(['status', folder], PW).status, 3, 'the old password stops working');
  assert.equal(run(['status', folder], 'new password\n').status, 0);
  // The new password opens the folder's own key file too.
  const keyFile = JSON.parse(fs.readFileSync(path.join(folder, '.rutbalocker-keys'), 'utf8'));
  assert.equal(keyFile.unlockers.length, 2);

  const unlock = run(['unlock', folder, 'a.pdf'], 'new password\n');
  assert.equal(unlock.status, 0, unlock.stderr);
  assert.ok(fs.existsSync(path.join(folder, 'a.pdf')));
  assert.ok(!fs.existsSync(path.join(folder, 'a.pdf.rlock')));
});

test('the password is never accepted as an argument', () => {
  const dir = tmp('rl-cli-');
  globalThis.__lockerDir = dir;
  run(['setup'], PW);
  const r = run(['status', '--password', 'cli password'], 'wrong\n');
  assert.equal(r.status, 3, 'the flag is ignored and standard input is what was read');
});
