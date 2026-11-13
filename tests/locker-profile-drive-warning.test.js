// The Locker warns when Office's profile is on a network share or a drive that
// keeps no permissions (D-2026-11-13-20, the owner: "Warn (Recommended)"). The
// service asks where the profile sits once, in the background, and hands the
// answer to the window in its state; the window puts one line under the offer
// to reopen files. Nothing is refused.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { transformSync } from 'esbuild';
import { tempDir } from './temp-dirs.js';
import { register } from '../apps/desktop/main/extras/locker/index.js';

// No Windows Hello prompt from a test.
process.env.RUTBA_LOCKER_FAKE_HELLO = '1';

/** The Locker's service on an empty profile, with Electron's parts handed in and `profileDriveOf` as the seam. */
async function service(profileDriveOf) {
  const dir = tempDir('rl-pd-');
  const settings = new Map();
  const sent = [];
  const { locker } = await register({
    stores: { dir, vault: null, settings: { get: (k, d) => (settings.has(k) ? settings.get(k) : d), set: (k, v) => settings.set(k, v) } },
    broadcast: (name, payload) => sent.push({ name, payload }),
    electron: { isPackaged: false, on() {}, quit() {}, getPath: () => dir },
    system: {
      powerMonitor: { on() {} },
      safeStorage: { isEncryptionAvailable: () => false, encryptString: () => Buffer.alloc(0), decryptString: () => '' },
      systemPreferences: {},
    },
    profileDriveOf,
  });
  return { locker, sent, dir };
}

const until = async (fn, ms = 5000) => {
  const end = Date.now() + ms;
  while (!fn() && Date.now() < end) await new Promise((r) => setTimeout(r, 10));
};

test('the state carries where the profile sits, and the window is told once the answer comes', async () => {
  for (const said of ['share', 'portable']) {
    let release;
    const asked = [];
    const { locker, sent, dir } = await service((d) => {
      asked.push(d);
      return new Promise((r) => { release = () => r(said); });
    });
    assert.deepEqual(asked, [dir], `${said}: the profile's own folder is asked about, once`);
    assert.equal((await locker.state()).profileDrive, null, `${said}: nothing is claimed before the answer`);
    release();
    await until(() => sent.some((s) => s.name === 'locker:state' && s.payload.profileDrive === said));
    assert.ok(sent.some((s) => s.name === 'locker:state' && s.payload.profileDrive === said), `${said}: the window is told`);
    assert.equal((await locker.state()).profileDrive, said);
    assert.equal(asked.length, 1, `${said}: asked once only`);
  }
});

test('a drive that keeps permissions, an answer that fails, or one not known says nothing, and the window is not told', async () => {
  for (const answer of [() => null, () => Promise.reject(new Error('no PowerShell')), () => { throw new Error('at once'); }, () => 'elsewhere']) {
    let done = false;
    const { locker, sent } = await service(async () => {
      try {
        return await answer();
      } finally {
        done = true;
      }
    });
    await until(() => done);
    await new Promise((r) => setTimeout(r, 20));
    assert.equal((await locker.state()).profileDrive, null);
    assert.equal(sent.filter((s) => s.name === 'locker:state').length, 0, 'nothing published for it');
  }
});

test('nothing waits on the asking: the Locker is ready, and its state answers, while PowerShell has not', async () => {
  const started = Date.now();
  const { locker } = await service(() => new Promise(() => {}));
  const state = await locker.state();
  assert.equal(state.profileDrive, null);
  assert.equal(state.open, false);
  assert.ok(Date.now() - started < 4000, `ready without the answer (${Date.now() - started} ms)`);
});

// The check runs main.js counts as such.
const CHECK_RUNS = ['RUTBA_OFFICE_VERIFY_EDIT', 'RUTBA_OFFICE_VERIFY_APPS', 'RUTBA_OFFICE_VERIFY_CORPUS', 'RUTBA_OFFICE_SMOKE', 'RUTBA_SCREENS_FEATURES', 'RUTBA_PROFILE_PEER'];

test('a check run of any kind main.js knows does not ask, so no window it checks waits on PowerShell', async () => {
  const main = fs.readFileSync(new URL('../apps/desktop/main/main.js', import.meta.url), 'utf8');
  const counted = /const CHECK_RUN = Boolean\(([^;]*)\);/.exec(main);
  assert.ok(counted, 'main.js has its list of check runs');
  for (const name of CHECK_RUNS) assert.ok(counted[1].includes(name.replace('RUTBA_PROFILE_PEER', 'PROFILE_PEER')), `main.js counts ${name}`);
  for (const name of CHECK_RUNS) {
    const before = process.env[name];
    process.env[name] = '1';
    try {
      let asked = 0;
      const { locker } = await service(() => { asked += 1; return 'share'; });
      await new Promise((r) => setTimeout(r, 20));
      assert.equal(asked, 0, `${name}: not asked`);
      assert.equal((await locker.state()).profileDrive, null);
    } finally {
      if (before === undefined) delete process.env[name];
      else process.env[name] = before;
    }
  }
});

// The window is JSX and is not loaded whole here: the line's own component is taken from its source, built
// with esbuild, and rendered.
const UI = fs.readFileSync(new URL('../apps/desktop/renderer/apps/locker.js', import.meta.url), 'utf8').replace(/\r\n/g, '\n');

function noteComponent() {
  const m = /^function ProfileDriveNote\([\s\S]*?\n\}\n/m.exec(UI);
  assert.ok(m, 'the window has the line\'s component');
  const { code } = transformSync(m[0], { loader: 'jsx', jsx: 'transform', jsxFactory: 'React.createElement' });
  const t = (s) => s;
  const Icon = ({ name }) => React.createElement('i', { 'data-icon': name });
  return new Function('React', 't', 'Icon', `${code}\nreturn ProfileDriveNote;`)(React, t, Icon);
}

test('the window shows one warning line for each case, saying what it means for the copies, and none otherwise', () => {
  const Note = noteComponent();
  const html = (drive) => renderToStaticMarkup(React.createElement(Note, { drive }));
  const share = html('share');
  const portable = html('portable');
  for (const [name, out] of [['share', share], ['portable', portable]]) {
    assert.match(out, /^<div class="lk-warn lk-profile-drive" role="note"/, `${name}: a note of its own`);
    assert.match(out, /data-icon="warn"/, `${name}: with the warning icon`);
    assert.match(out, /The locked files stay locked\./, `${name}: says the files stay locked`);
    // Only a file that opens from a temporary plain copy has a copy there: Office's own kinds open in memory.
    assert.match(out, /A file that opens from a temporary plain copy \(the Locker says when one does\) has that copy there while it is open/, `${name}: names the temporary plain copy`);
    assert.doesNotMatch(out, /copies of files you open/, `${name}: does not say every file leaves a copy`);
    assert.ok(!out.includes(String.fromCharCode(8212)), `${name}: no em dash`);
  }
  assert.match(share, /network share/);
  assert.match(share, /the server&#x27;s permissions, not Windows, decide who can read it/);
  assert.match(portable, /exFAT or FAT, which keeps no permissions/);
  assert.match(portable, /anyone with this drive can read it/);
  assert.notEqual(share, portable);
  for (const none of [null, undefined, '', 'elsewhere']) assert.equal(html(none), '', `nothing for ${none}`);
});

test('the line sits in the open view, straight under the offer to reopen files', () => {
  assert.match(UI, /\n( +)<ReopenOffer [^\n]*\/>\n\1<ProfileDriveNote drive=\{state\.profileDrive\} \/>\n/);
  assert.equal((UI.match(/<ProfileDriveNote /g) || []).length, 1, 'in one place');
  // ReopenOffer keeps its own description straight above it.
  const doc = /\/\*\*\n \* The files that were open when the Locker locked[\s\S]*?\*\/\nfunction ReopenOffer\(/.exec(UI);
  assert.ok(doc && !doc[0].includes('ProfileDriveNote'), 'ReopenOffer\'s description is straight above it');
});
