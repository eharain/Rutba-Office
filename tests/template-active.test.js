// The templates in use (D-2026-10-21-24): the gallery and Data Files read
// the set Office was built with, or, when Office names a newer signed set it
// downloaded (RUTBA_TEMPLATES_DIR, set by template-set-choice.js before
// anything reads templates), that set's own, its pictures and files beside it.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.join(root, 'packages', 'templates');

/** What the active entry points say, read in a process of its own with this environment. */
function read(env) {
  const code = `const g = await import('@rutba/templates/active'); const d = await import('@rutba/templates/data-files/active');
    console.log(JSON.stringify({ dir: g.ACTIVE_DIR, files: g.FILES_DIR, count: g.TEMPLATES.length, ids: d.TEMPLATE_IDS }));`;
  const out = execFileSync(process.execPath, ['--input-type=module', '-e', code], { cwd: root, env: { ...process.env, ...env }, encoding: 'utf8' });
  return JSON.parse(out.trim().split('\n').pop());
}

test('with no set named, the set Office was built with', () => {
  const r = read({ RUTBA_TEMPLATES_DIR: '' });
  assert.equal(r.dir, null);
  assert.ok(r.files.startsWith(path.join(pkg, 'files')));
  assert.ok(r.count > 10);
  assert.ok(r.ids.length >= 5);
});

test('with a downloaded set named, its own gallery and Data Files, its files beside it', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'templates-set-'));
  try {
    for (const part of ['src', 'data-files']) fs.cpSync(path.join(pkg, part), path.join(dir, part), { recursive: true });
    fs.copyFileSync(path.join(pkg, 'manifest.json'), path.join(dir, 'manifest.json'));
    // Marked, so what is read can be told from the bundled set.
    const index = path.join(dir, 'data-files', 'index.js');
    const text = fs.readFileSync(index, 'utf8');
    const line = /export \{ TEMPLATE_IDS, ([^}]*)\} from '\.\/simple\.js';/;
    assert.match(text, line);
    fs.writeFileSync(index, text.replace(line, "export { $1} from './simple.js';\nexport const TEMPLATE_IDS = ['from-the-set'];"));
    const r = read({ RUTBA_TEMPLATES_DIR: dir });
    assert.equal(r.dir, dir);
    assert.ok(r.files.startsWith(dir), r.files);
    assert.deepEqual(r.ids, ['from-the-set']);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
