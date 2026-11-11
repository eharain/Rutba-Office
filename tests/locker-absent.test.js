// The Locker is in every build, and a build made with RUTBA_LOCKER=0 carries no
// trace of it: no window, no strings, no source file, no file association. The
// check is on what each build actually contains, by making one of each. A
// release that left it out would strand every file people have locked.

import { test } from 'node:test';
import { tempDir } from './temp-dirs.js';
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
  const out = tempDir('rl-bundle-');
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

test('an ordinary build has the Locker, and a build made with the switch off has no trace of it in any file', () => {
  const off = bundle({ RUTBA_LOCKER: '0' });
  const offFiles = listAll(off).filter(textual);
  assert.ok(offFiles.length > 10, 'a real build came out');
  // The release notes are shown in What's new whatever the build carries.
  const mentions = offFiles.filter((f) => path.basename(f) !== 'whatsnew.json' && WORDS.test(fs.readFileSync(f, 'utf8')));
  assert.deepEqual(mentions.map((f) => path.relative(off, f)), [], 'nothing in the build mentions the Locker');
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(off, 'extras.json'), 'utf8')).extras, []);
  fs.rmSync(path.dirname(off), { recursive: true, force: true });

  const on = bundle({});
  const onMentions = listAll(on).filter(textual).filter((f) => WORDS.test(fs.readFileSync(f, 'utf8')));
  assert.ok(onMentions.length >= 1, 'the switch puts it in, so the scan above means something');
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(on, 'extras.json'), 'utf8')).extras, ['locker']);
  fs.rmSync(path.dirname(on), { recursive: true, force: true });
});

test('the installer\'s file list carries the Locker\'s code unless the switch is off', async () => {
  const require = createRequire(import.meta.url);
  const configPath = path.join(desktop, 'electron-builder.config.cjs');
  const before = process.env.RUTBA_LOCKER;
  try {
    process.env.RUTBA_LOCKER = '0';
    delete require.cache[require.resolve(configPath)];
    const offConfig = await require(configPath)();
    assert.ok(offConfig.files.includes('!main/extras/**'), 'the code folder is excluded');
    // Names and extensions only: the icons' paths carry this clone's folder name.
    const assoc = (cfg) => JSON.stringify([...(cfg.fileAssociations || []), ...(cfg.mac?.fileAssociations || []), ...(cfg.linux?.fileAssociations || [])].map((a) => [a.ext, a.name]));
    assert.ok(!WORDS.test(assoc(offConfig)), 'no file association names it');
    assert.ok(!WORDS.test(JSON.stringify(offConfig.extraResources || [])), 'no extra resource carries it');

    delete process.env.RUTBA_LOCKER;
    delete require.cache[require.resolve(configPath)];
    const onConfig = await require(configPath)();
    assert.ok(!onConfig.files.includes('!main/extras/**'), 'an ordinary build carries the code');
  } finally {
    if (before === undefined) delete process.env.RUTBA_LOCKER;
    else process.env.RUTBA_LOCKER = before;
    delete require.cache[require.resolve(configPath)];
  }
});
