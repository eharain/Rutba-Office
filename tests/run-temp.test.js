// tools/run-temp.js: npm test, test:perf and the gate run with a temp folder
// of their own that goes when they end, so the folders the tests make in the
// temp folder and never remove do not pile up there.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const wrapper = path.join(root, 'tools', 'run-temp.js');

// A command that makes a temp folder the way the tests do, says where, and ends with the code given.
const leaky = (code) => `
  const fs = require('node:fs'), os = require('node:os'), path = require('node:path');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-leak-'));
  fs.writeFileSync(path.join(dir, 'left.txt'), 'left behind');
  console.log(JSON.stringify({ dir, tmp: os.tmpdir() }));
  process.exit(${code});`;

// The wrapper run as if the machine's temp folder were `base`.
function run(base, code = 0) {
  const env = { ...process.env, TEMP: base, TMP: base, TMPDIR: base };
  const r = spawnSync(process.execPath, [wrapper, 'node', '-e', leaky(code)], { encoding: 'utf8', env });
  return { status: r.status, ...JSON.parse(r.stdout.trim().split('\n').pop()) };
}

const plain = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-run-temp-test-'));

test('what the command leaves in the temp folder is gone when it ends, and its exit code comes through', () => {
  const base = plain();
  try {
    const r = run(base, 3);
    assert.equal(r.status, 3);
    assert.equal(path.dirname(r.tmp), base, 'the run had a folder of its own inside the temp folder');
    assert.match(path.basename(r.tmp), /^rutba-office-run-/);
    assert.equal(path.dirname(r.dir), r.tmp, 'the command made its folder there');
    assert.ok(!fs.existsSync(r.dir) && !fs.existsSync(r.tmp), 'both are gone');
    assert.deepEqual(fs.readdirSync(base), []);
  } finally {
    fs.rmSync(base, { recursive: true, force: true });
  }
});

test('inside a run that has a folder already (npm test under the gate) the command uses it and leaves it to that run', () => {
  const outer = fs.mkdtempSync(path.join(plain(), 'rutba-office-run-'));
  try {
    const r = run(outer);
    assert.equal(r.status, 0);
    assert.equal(r.tmp, outer);
    assert.ok(fs.existsSync(r.dir), 'the outer run removes it when it ends');
  } finally {
    fs.rmSync(path.dirname(outer), { recursive: true, force: true });
  }
});

test('a folder a killed run left is removed by the next run once nothing in it is a day old, a newer one is kept', () => {
  const base = plain();
  try {
    const old = path.join(base, 'rutba-office-run-killed');
    const recent = path.join(base, 'rutba-office-run-going');
    for (const dir of [old, recent]) {
      fs.mkdirSync(dir);
      fs.writeFileSync(path.join(dir, 'file.txt'), 'x');
    }
    const twoDaysAgo = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000);
    fs.utimesSync(path.join(old, 'file.txt'), twoDaysAgo, twoDaysAgo);
    fs.utimesSync(old, twoDaysAgo, twoDaysAgo);
    assert.equal(run(base).status, 0);
    assert.ok(!fs.existsSync(old), 'the day-old folder is gone');
    assert.ok(fs.existsSync(recent), 'a folder a run may still be using is kept');
  } finally {
    fs.rmSync(base, { recursive: true, force: true });
  }
});
