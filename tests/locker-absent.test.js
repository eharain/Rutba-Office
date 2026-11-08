// A build made without RUTBA_LOCKER=1 carries no Locker at all: no window, no
// strings, no source file, no file association, no command-line entry. The
// check is on what such a build actually contains, by making one, and by a
// build made with the switch showing the same scan does find it.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const here = path.dirname(fileURLToPath(import.meta.url));
const desktop = path.resolve(here, '..', 'apps', 'desktop');
const WORDS = /locker|rlock/i;

function bundle(env) {
  const out = fs.mkdtempSync(path.join(os.tmpdir(), 'rl-bundle-'));
  execFileSync(process.execPath, ['build/bundle.js'], { cwd: desktop, stdio: 'pipe', env: { ...process.env, RUTBA_LOCKER: '', ...env, RUTBA_BUNDLE_OUT: path.join(out, 'out') } });
  return path.join(out, 'out');
}

function listAll(dir, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) listAll(p, out);
    else out.push(p);
  }
  return out;
}

const textual = (f) => /\.(js|cjs|css|html|json)$/.test(f);

test('a build made without the switch has no Locker in any file it outputs; the same build with the switch does', () => {
  const off = bundle({});
  const offFiles = listAll(off).filter(textual);
  assert.ok(offFiles.length > 10, 'a real build came out');
  const mentions = offFiles.filter((f) => WORDS.test(fs.readFileSync(f, 'utf8')));
  assert.deepEqual(mentions.map((f) => path.relative(off, f)), [], 'nothing in the build mentions the Locker');
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(off, 'extras.json'), 'utf8')).extras, []);
  fs.rmSync(path.dirname(off), { recursive: true, force: true });

  const on = bundle({ RUTBA_LOCKER: '1' });
  const onMentions = listAll(on).filter(textual).filter((f) => WORDS.test(fs.readFileSync(f, 'utf8')));
  assert.ok(onMentions.length >= 1, 'the switch puts it in, so the scan above means something');
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(on, 'extras.json'), 'utf8')).extras, ['locker']);
  fs.rmSync(path.dirname(on), { recursive: true, force: true });
});

test('the installer\'s file list leaves the Locker\'s code out unless the switch is set', async () => {
  const require = createRequire(import.meta.url);
  const configPath = path.join(desktop, 'electron-builder.config.cjs');
  const before = process.env.RUTBA_LOCKER;
  try {
    delete process.env.RUTBA_LOCKER;
    delete require.cache[require.resolve(configPath)];
    const offConfig = await require(configPath)();
    assert.ok(offConfig.files.includes('!main/extras/**'), 'the code folder is excluded');
    // Names and extensions only: the icons' paths carry this clone's folder name.
    const assoc = (cfg) => JSON.stringify([...(cfg.fileAssociations || []), ...(cfg.mac?.fileAssociations || []), ...(cfg.linux?.fileAssociations || [])].map((a) => [a.ext, a.name]));
    assert.ok(!WORDS.test(assoc(offConfig)), 'no file association names it');
    assert.ok(!WORDS.test(JSON.stringify(offConfig.extraResources || [])), 'no extra resource carries it');

    process.env.RUTBA_LOCKER = '1';
    delete require.cache[require.resolve(configPath)];
    const onConfig = await require(configPath)();
    assert.ok(!onConfig.files.includes('!main/extras/**'), 'with the switch the code is carried');
  } finally {
    if (before === undefined) delete process.env.RUTBA_LOCKER;
    else process.env.RUTBA_LOCKER = before;
    delete require.cache[require.resolve(configPath)];
  }
});

test('nothing an ordinary build ships as source names it: not the main process, not the shared registry, not the bridge', () => {
  const skip = /[\\/](extras|tests?)[\\/]|[\\/](verify-|screen-)[^\\/]*$/;
  const files = listAll(path.join(desktop, 'main')).filter((f) => f.endsWith('.js') && !skip.test(f));
  assert.ok(files.includes(path.join(desktop, 'main', 'main.js')));
  const named = files.filter((f) => WORDS.test(fs.readFileSync(f, 'utf8')));
  assert.deepEqual(named.map((f) => path.relative(desktop, f)), [], 'the main process has no word of it outside the extras folder');
  for (const f of ['packages/office-formats/src/registry.js', 'packages/office-shell/src/contract.js', 'packages/office-shell/src/client.js', 'packages/office-shell/src/electron/main.js']) {
    assert.ok(!WORDS.test(fs.readFileSync(path.resolve(here, '..', f), 'utf8')), `${f} has no word of it`);
  }
});

test('no catalogue carries the Locker\'s words, because no shipped window source names them', () => {
  const list = JSON.parse(fs.readFileSync(path.resolve(here, '..', 'packages/office-ui/src/catalogues/messages.json'), 'utf8'));
  const named = list.messages.filter((m) => WORDS.test(m.message) || m.where.some((w) => /locker/i.test(w)));
  assert.deepEqual(named, []);
});

test('no release, README or site page mentions it', () => {
  const root = path.resolve(here, '..');
  const hits = [];
  for (const f of ['README.md', ...fs.readdirSync(path.join(root, 'docs', 'releases')).map((n) => path.join('docs', 'releases', n))]) {
    if (WORDS.test(fs.readFileSync(path.join(root, f), 'utf8'))) hits.push(f);
  }
  assert.deepEqual(hits, []);
});
