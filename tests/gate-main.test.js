// The gate's main(), run as a person runs it, on the quickest pass it has
// (the signed-mail fuzzer), with its state kept in a folder of the test's own
// so a real gate beside it is not touched.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync, execFileSync } from 'node:child_process';
import { CLEAN_TREE, passKept } from '../tools/gate-carry.js';
import { treeKey } from '../tools/gate.js';

const ROOT = path.join(import.meta.dirname, '..');
const head = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: ROOT, encoding: 'utf8' }).trim();
const kept = ['README.md', 'tools/gate-floors.json'].map((f) => [f, fs.existsSync(path.join(ROOT, f)) ? fs.readFileSync(path.join(ROOT, f)) : null]);

const gate = (dir, args, env = {}) => {
  const r = spawnSync(process.execPath, [path.join(ROOT, 'tools', 'gate.js'), ...args], { cwd: ROOT, encoding: 'utf8', env: { ...process.env, RUTBA_GATE_DIR: dir, ...env }, timeout: 180000 });
  return { code: r.status, out: `${r.stdout}${r.stderr}`.replace(/\u001b\[[0-9;]*m/g, '') };
};
const stateOf = (dir) => JSON.parse(fs.readFileSync(path.join(dir, 'state.json'), 'utf8'));

test('one pass named: it runs, the gate opens for it only, and nothing goes on the record', { timeout: 420000 }, (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-gate-main-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const r = gate(dir, ['fuzz:smime', '--record', '--quiet'], { RUTBA_VERIFY_ONLY: 'recorder-app' });
  assert.equal(r.code, 0, r.out);
  assert.match(r.out, /RUTBA_VERIFY_ONLY was set in this shell and is ignored/);
  assert.match(r.out, /The gate is open\. 1 passes, for fuzz:smime only\./);
  assert.match(r.out, /Nothing recorded/);
  for (const [f, bytes] of kept) assert.deepEqual(fs.existsSync(path.join(ROOT, f)) ? fs.readFileSync(path.join(ROOT, f)) : null, bytes, `${f} is not written by a run of one pass`);
  const s = stateOf(dir);
  assert.equal(s.head, head);
  assert.equal(s.passes['fuzz:smime'].ok, true);
  assert.equal(s.passes['fuzz:smime'].counts.of, 2000, 'the fuzzer ran its full count');
  assert.equal(fs.existsSync(path.join(dir, 'running')), false, 'the lock is let go at the end');
  assert.ok(fs.statSync(path.join(dir, 'logs', 'fuzz-smime.log')).size > 0, 'the pass logged to the run\'s own folder');

  // --resume on the same commit and tree: the pass stands.
  const again = gate(dir, ['fuzz:smime', '--resume', '--quiet']);
  assert.equal(again.code, 0, again.out);
  assert.match(again.out, /ok {3}fuzz:smime .* 0\.0s \(before\)/);

  // A state from another tree is not resumed: the pass runs again.
  fs.writeFileSync(path.join(dir, 'state.json'), JSON.stringify({ ...stateOf(dir), tree: 'another tree' }));
  const other = gate(dir, ['fuzz:smime', '--resume', '--quiet']);
  assert.equal(other.code, 0, other.out);
  assert.doesNotMatch(other.out, /\(before\)/);
  assert.doesNotMatch(`${r.out}${again.out}`, /^warning: /m, 'what git says to itself is not in the gate\'s output');

  // A state from a commit this one does not come after: nothing is carried, the run says why, and the pass runs again.
  fs.writeFileSync(path.join(dir, 'state.json'), JSON.stringify({ ...stateOf(dir), head: 'f'.repeat(40), tree: CLEAN_TREE }));
  const strange = gate(dir, ['fuzz:smime', '--resume']);
  assert.equal(strange.code, 0, strange.out);
  assert.match(strange.out, /Nothing is carried from the run on ffffffff: (this commit does not come after it|this tree has changes beyond its commit)\./);
  assert.doesNotMatch(strange.out, /\(before\)/);
  assert.equal(stateOf(dir).head, head);

  // A state from the commit before this one: the pass is kept only if this checkout holds nothing beyond its
  // commit and no file that changed between the two is one the pass runs; the gate says which, and why.
  let parent = null;
  try { parent = execFileSync('git', ['rev-parse', '--verify', '--quiet', 'HEAD^'], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim(); } catch { /* a checkout of one commit */ }
  if (parent) {
    const { head: _, ...rest } = stateOf(dir);
    fs.writeFileSync(path.join(dir, 'state.json'), JSON.stringify({ ...rest, head: parent, tree: CLEAN_TREE }));
    const files = execFileSync('git', ['diff', '--name-only', '--no-renames', '-z', parent, head], { cwd: ROOT, encoding: 'utf8' }).split('\0').filter(Boolean);
    const kept = treeKey(ROOT) === CLEAN_TREE && passKept('fuzz:smime', files);
    const next = gate(dir, ['fuzz:smime', '--resume']);
    assert.equal(next.code, 0, next.out);
    if (kept) {
      assert.match(next.out, new RegExp(`Carried from the run on ${parent.slice(0, 8)} .*: fuzz:smime\\.`));
      assert.match(next.out, new RegExp(`fuzz:smime: passed on ${parent.slice(0, 8)}, and no file it runs has changed since\\.`));
      assert.match(next.out, /ok {3}fuzz:smime .* 0\.0s \(before\)/);
      assert.equal(stateOf(dir).passes['fuzz:smime'].at, parent);
    } else {
      assert.match(next.out, new RegExp(`Nothing is carried from the run on ${parent.slice(0, 8)}: `));
      assert.doesNotMatch(next.out, /\(before\)/);
      assert.equal(stateOf(dir).passes['fuzz:smime'].at, undefined);
    }
    assert.equal(stateOf(dir).head, head);
  }
});

test('a pass name the gate does not have is refused, not run as nothing and passed', { timeout: 60000 }, (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-gate-main-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const r = gate(dir, ['fuzz:smoke', '--quiet']);
  assert.equal(r.code, 2, r.out);
  assert.match(r.out, /The gate has no pass called fuzz:smoke\. Its passes: test, test:perf, .*fuzz:smime.*corpus\./);
  assert.doesNotMatch(r.out, /The gate is open/);
  assert.equal(fs.existsSync(path.join(dir, 'state.json')), false, 'nothing ran');
});

test('a second gate does not start while one is running', { timeout: 60000 }, (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-gate-main-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  fs.writeFileSync(path.join(dir, 'running'), String(process.pid));
  const r = gate(dir, ['fuzz:smime', '--quiet']);
  assert.equal(r.code, 1, r.out);
  assert.match(r.out, new RegExp(`Another gate is running \\(process ${process.pid}\\)`));
  assert.equal(fs.readFileSync(path.join(dir, 'running'), 'utf8'), String(process.pid), 'the other gate\'s lock is left alone');
});
