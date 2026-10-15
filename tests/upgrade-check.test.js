/**
 * The upgrade test's judgement (tools/upgrade-check.js): which release is the
 * previous one, what the files are called, the profile an older version is
 * started on, when it has run, and whether the newer one kept what it found.
 * The running itself (tools/upgrade-test.js) needs the executables and is
 * run by hand or by .github/workflows/upgrade.yml.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseVersion, compareVersions, previousRelease, newestRelease, assetName, seedProfile, olderHasRun, judgeUpgrade, judgeInstallRace } from '../tools/upgrade-check.js';

test('versions compare as numbers, a pre-release before its release', () => {
  assert.equal(parseVersion('v2.0.0').text, '2.0.0');
  assert.equal(parseVersion('nonsense'), null);
  assert.equal(compareVersions('1.51.0', '2.0.0'), -1);
  assert.equal(compareVersions('1.10.0', '1.9.3'), 1, 'not as text');
  assert.equal(compareVersions('2.0.0-beta.1', '2.0.0'), -1);
  assert.equal(compareVersions('v2.0.0', '2.0.0'), 0);
});

test('the previous release is the newest published one older than the target, drafts left out', () => {
  const releases = [
    { tag_name: 'v2.1.0', draft: true },
    { tag_name: 'v2.0.0', draft: false },
    { tag_name: 'v1.51.0', draft: false },
    { tag_name: 'v1.9.0', draft: false },
    { tag_name: 'not-a-version', draft: false },
  ];
  assert.equal(previousRelease(releases, '2.1.0').tag_name, 'v2.0.0');
  assert.equal(previousRelease(releases, '2.0.0').tag_name, 'v1.51.0');
  assert.equal(previousRelease(releases, '1.9.0'), null);
  assert.equal(newestRelease(releases).tag_name, 'v2.0.0', 'a draft is nobody\'s newest');
});

test('the files are named as electron-builder.config.cjs names them', () => {
  assert.equal(assetName('v2.0.0', 'portable'), 'Rutba-Office-2.0.0-portable.exe');
  assert.equal(assetName('2.0.0', 'installer'), 'Rutba-Office-2.0.0-win-x64.exe');
  assert.throws(() => assetName('2.0.0', 'dmg'));
});

const profile = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-upgrade-test-'));

test('the seeded profile has settings, a marker and a recent document, with updates and the notice board off', () => {
  const dir = profile();
  try {
    const seeded = seedProfile(dir, { document: 'C:\\docs\\a.docx', token: 'abc' });
    const settings = JSON.parse(fs.readFileSync(path.join(dir, 'settings.json'), 'utf8'));
    assert.equal(settings['upgradeTest.marker'], 'abc');
    assert.equal(settings['updates.automatic'], false, 'the older version must not fetch an update while under test');
    assert.equal(settings['announcements.enabled'], false);
    assert.deepEqual(seeded.settings, settings);
    assert.equal(JSON.parse(fs.readFileSync(path.join(dir, 'recent.json'), 'utf8'))[0].path, 'C:\\docs\\a.docx');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('the older version has run once it records itself as the version that ran last, and opened its document when asked to', () => {
  const dir = profile();
  try {
    seedProfile(dir, { document: path.join(dir, 'a.docx'), token: 't' });
    assert.equal(olderHasRun(dir, '1.51.0'), false);
    const settings = JSON.parse(fs.readFileSync(path.join(dir, 'settings.json'), 'utf8'));
    fs.writeFileSync(path.join(dir, 'settings.json'), JSON.stringify({ ...settings, 'updates.lastRun': '1.51.0' }));
    assert.equal(olderHasRun(dir, 'v1.51.0'), true);
    assert.equal(olderHasRun(dir, '1.51.0', { document: path.join(dir, 'b.docx') }), false);
    assert.equal(olderHasRun(dir, '1.51.0', { document: path.join(dir, 'a.docx') }), true);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

const seeded = { settings: { theme: 'dark', 'upgradeTest.marker': 'abc' }, document: path.resolve('docs/a.docx') };
const good = () => ({
  version: '2.0.0',
  settings: { theme: 'dark', 'upgradeTest.marker': 'abc', 'updates.lastRun': '2.0.0', 'updates.arrived': { from: '1.51.0', to: '2.0.0', seen: false } },
  recent: [{ path: path.resolve('docs/a.docx') }],
  feed: { ok: true, version: '2.0.0', files: [{ url: 'Rutba-Office-2.0.0-win-x64.exe', sha512: 'x', size: 1 }] },
});

test('an upgrade that kept everything is proven, with the evidence said', () => {
  const v = judgeUpgrade({ from: 'v1.51.0', to: '2.0.0', seeded, report: good() });
  assert.deepEqual(v.failures, []);
  assert.equal(v.ok, true);
  assert.ok(v.evidence.some((e) => /arrival is recorded, 1\.51\.0 to 2\.0\.0/.test(e)));
  assert.ok(v.evidence.some((e) => /updater read the feed: 2\.0\.0/.test(e)));
});

test('each thing an upgrade can lose is a failure of its own', () => {
  const lost = (change, pattern, opts = {}) => {
    const report = good();
    change(report);
    const v = judgeUpgrade({ from: '1.51.0', to: '2.0.0', seeded, report, ...opts });
    assert.equal(v.ok, false);
    assert.ok(v.failures.some((f) => pattern.test(f)), v.failures.join(' | '));
  };
  lost((r) => { r.version = '1.51.0'; }, /says it is 1\.51\.0/);
  lost((r) => { delete r.settings.theme; }, /setting theme was "dark"/);
  lost((r) => { delete r.settings['upgradeTest.marker']; }, /upgradeTest\.marker/);
  lost((r) => { r.settings['updates.arrived'] = null; }, /arrival is null/);
  lost((r) => { r.recent = []; }, /recent list lost/);
  lost((r) => { r.feed = { ok: false, error: 'net::ERR_INTERNET_DISCONNECTED' }; }, /did not read the feed: net::ERR/);
  assert.equal(judgeUpgrade({ from: '1.51.0', to: '2.0.0', seeded, report: { ...good(), feed: null }, feed: false }).ok, true, '--no-feed leaves the feed out');
  assert.deepEqual(judgeUpgrade({ from: '1.51.0', to: '2.0.0', seeded, report: null }).failures, ['the newer version wrote no report']);
});

test('an update that meets the app still closing must install once the app has closed', () => {
  const base = { race: 'close', limitSeconds: 300, from: '1.51.0', to: '2.0.0', closedAfter: 3.2 };
  const ok = judgeInstallRace({ ...base, exit: 0, seconds: 9, copies: ['2.0.0 user'] });
  assert.equal(ok.ok, true, ok.failures.join(' | '));
  assert.ok(ok.evidence.some((e) => /closed 3\.2 s after it started/.test(e)));
  const error2 = judgeInstallRace({ ...base, exit: 2, seconds: 4, copies: ['1.51.0 user'] });
  assert.deepEqual(error2.failures, ['the installer ended with 2 while the app was closing', 'the registry has 1.51.0 user, not one copy of 2.0.0']);
  assert.match(judgeInstallRace({ ...base, exit: 'timeout', seconds: 300, copies: [] }).failures[0], /still running after 300 s/);
  assert.match(judgeInstallRace({ ...base, exit: 0, seconds: 9, copies: ['2.0.0 user', '1.51.0 machine'] }).failures[0], /not one copy of 2\.0\.0/);
});

test('an update that meets the app left open closes it itself and installs, as the installer has since 2.1.0', () => {
  const base = { race: 'never-close', limitSeconds: 180, from: '2.0.0', to: '2.1.0' };
  // As the first run on a runner found it: 13 s, the copy ended, one 2.1.0.
  const clean = judgeInstallRace({ ...base, exit: 0, seconds: 13, copies: ['2.1.0 user'], appEnded: true });
  assert.equal(clean.ok, true, clean.failures.join(' | '));
  assert.match(judgeInstallRace({ ...base, exit: 'timeout', seconds: 180, copies: ['2.0.0 user'] }).failures[0], /hung: still running after 180 s/);
  assert.match(judgeInstallRace({ ...base, exit: 2, seconds: 17, copies: ['2.0.0 user'] }).failures.join(' | '), /ended with 2: the copy left open was not closed for it/, 'the 1.51 to 2.0 failure');
  assert.match(judgeInstallRace({ ...base, exit: 0, seconds: 1, copies: ['2.1.0 user'], appEnded: true }).failures[0], /without waiting 5 s/);
  assert.match(judgeInstallRace({ ...base, exit: 0, seconds: 13, copies: ['2.1.0 user'], appEnded: false }).failures[0], /still running when the installer finished/);
  assert.match(judgeInstallRace({ ...base, exit: 0, seconds: 13, copies: ['2.1.0 user', '2.0.0 machine'], appEnded: true }).failures[0], /not one copy of 2\.1\.0/);
  assert.throws(() => judgeInstallRace({ ...base, race: 'sometimes', exit: 0, seconds: 1, copies: [] }));
});
