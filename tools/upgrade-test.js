// The upgrade test: a person's profile, made by an older release, opened by
// the newer one, with nothing lost; and the newer one's updater reading the
// release feed.
//
//   node tools/upgrade-test.js                         the previous GitHub release to this build
//   node tools/upgrade-test.js --from v1.51.0 --to v2.0.0
//   node tools/upgrade-test.js --from D:\old\Rutba-Office-1.51.0-portable.exe --to apps\desktop\release\win-unpacked\Rutba Office.exe
//   node tools/upgrade-test.js --install --open-file   on a CI runner: the installers, one over the other
//
// --from and --to each take a tag (its file is fetched from the GitHub
// release into a cache in the temporary folder, kept for the next run) or a
// path to an executable; --to newest is the newest published release. --to defaults to this checkout's own build
// (apps/desktop/release, the portable executable of package.json's version,
// else win-unpacked); --from defaults to the newest release older than --to.
//
// By default nothing is installed: both versions run as the portable
// executable (or an unpacked build), each with --user-data-dir set to one
// profile in the temporary folder, so this machine's own copy, its profile,
// its registry and its Start menu are not touched. With --install (CI only:
// it installs over whatever copy the machine has) the older installer is run
// silently, then the newer over it, and the copy under test is the installed
// one; the registry then has to show one copy, of the newer version.
//
// The profile is seeded as a person's would be (tools/upgrade-check.js
// seedProfile), the older version started on it until it has recorded itself
// as the version that ran last, and closed; with --open-file it also opens a
// document, so it writes the recent list itself (which puts the document in
// Windows' recent items, so this is for a runner, not a desk). The newer one
// is started on the same profile with RUTBA_OFFICE_UPGRADE_REPORT, writes what
// it read (apps/desktop/main/upgrade-report.js) and quits, and
// judgeUpgrade says whether settings, the recent list and the arrival record
// survived, and whether the updater read the feed (--no-feed leaves that out).
//
// --race close and --race never-close (with --install) reproduce an update
// that meets the app still closing (1.51 to 2.0 stopped on "could not be
// removed ... Uninstaller error 2"): the older copy is started again, the
// newer installer run silently while it is open (/S --updated, as the
// updater runs it), and the copy closed 2 to 5 s later, as quitAndInstall
// closes it. The install must succeed and the report follow as above. With
// never-close the copy stays open: since 2.1.0 the installer closes it
// itself (its windows asked after 5 s, what is left ended after 10), so the
// install must succeed by itself, not hang, leave one copy of the newer
// version and the report follow (tools/upgrade-check.js judgeInstallRace).
//
// Exit codes: 0 proven; 1 failed (what failed is printed); 2 could not run.
// The verdict, with the evidence, is also written to --out (default
// <temp>/rutba-upgrade-result.json).

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawn, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { parseVersion, previousRelease, newestRelease, assetName, seedProfile, olderHasRun, judgeUpgrade, judgeInstallRace } from './upgrade-check.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OWNER = 'eharain';
const REPO = 'Rutba-Office';

const args = process.argv.slice(2);
const flag = (name) => args.includes(`--${name}`);
const option = (name) => {
  const at = args.indexOf(`--${name}`);
  const value = at >= 0 ? args[at + 1] : null;
  return value && !value.startsWith('--') ? value : null;
};
const install = flag('install');
const openFile = flag('open-file');
const race = option('race');
const feed = !flag('no-feed');
const keep = flag('keep');
const out = path.resolve(option('out') || path.join(os.tmpdir(), 'rutba-upgrade-result.json'));
const cache = path.join(os.tmpdir(), 'rutba-upgrade-cache');
const kind = install ? 'installer' : 'portable';
const say = (s) => console.log(`[upgrade] ${s}`);
const wait = (ms) => new Promise((r) => setTimeout(r, ms));

async function github(url, { raw = false } = {}) {
  const token = process.env.GH_TOKEN || process.env.GITHUB_TOKEN;
  const response = await fetch(url, {
    headers: { 'user-agent': 'rutba-office-upgrade-test', accept: raw ? 'application/octet-stream' : 'application/vnd.github+json', ...(token ? { authorization: `Bearer ${token}` } : {}) },
    redirect: 'follow',
  });
  if (!response.ok) throw new Error(`${url} answered ${response.status}`);
  return raw ? Buffer.from(await response.arrayBuffer()) : response.json();
}

let releasesCache = null;
const releases = async () => (releasesCache ??= await github(`https://api.github.com/repos/${OWNER}/${REPO}/releases?per_page=50`));

/** A release's file, from the cache or fetched into it. */
async function fetchAsset(tag, name) {
  const release = (await releases()).find((r) => r.tag_name === tag);
  if (!release) throw new Error(`there is no release ${tag}`);
  const asset = (release.assets || []).find((a) => a.name === name);
  if (!asset) throw new Error(`release ${tag} has no ${name} (it has ${(release.assets || []).map((a) => a.name).join(', ')})`);
  const file = path.join(cache, tag, name);
  if (fs.existsSync(file) && fs.statSync(file).size === asset.size) {
    say(`${name}: in the cache`);
    return file;
  }
  say(`${name}: fetching ${(asset.size / 1048576).toFixed(1)} MB from the ${tag} release`);
  const bytes = await github(asset.browser_download_url, { raw: true });
  if (bytes.length !== asset.size) throw new Error(`${name} came down as ${bytes.length} bytes, not ${asset.size}`);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, bytes);
  return file;
}

/** The version a path's name says (Rutba-Office-2.0.0-portable.exe), or null. */
const versionInName = (file) => /Rutba-Office-(\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)-/.exec(path.basename(file))?.[1] || null;

/** --to: a tag, a path, or this checkout's build. */
async function resolveTo() {
  let given = option('to');
  if (given === 'newest') given = newestRelease(await releases())?.tag_name || null;
  if (given && fs.existsSync(given)) return { file: path.resolve(given), version: option('to-version') || versionInName(given) || localVersion() };
  if (given) return { file: await fetchAsset(given, assetName(given, kind)), version: parseVersion(given).text };
  const version = localVersion();
  const release = path.join(root, 'apps', 'desktop', 'release');
  for (const candidate of [path.join(release, assetName(version, kind)), ...(install ? [] : [path.join(release, 'win-unpacked', 'Rutba Office.exe')])]) {
    if (fs.existsSync(candidate)) return { file: candidate, version };
  }
  throw new Error(`no build of ${version} in ${release}: run npm run dist:win, or give --to a tag or a path`);
}

const localVersion = () => JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8')).version;

/** --from: a tag, a path, or the newest release older than --to. */
async function resolveFrom(to) {
  const given = option('from');
  if (given && fs.existsSync(given)) {
    const version = option('from-version') || versionInName(given);
    if (!version) throw new Error(`cannot tell the version of ${given} from its name; give --from-version`);
    return { file: path.resolve(given), version };
  }
  const tag = given || previousRelease(await releases(), to.version)?.tag_name;
  if (!tag) throw new Error(`there is no release older than ${to.version}`);
  return { file: await fetchAsset(tag, assetName(tag, kind)), version: parseVersion(tag).text };
}

/** Stop a process and everything it started: a portable executable's stub starts the application. */
function stopTree(pid, force) {
  if (process.platform === 'win32') spawnSync('taskkill', ['/PID', String(pid), '/T', ...(force ? ['/F'] : [])], { stdio: 'ignore' });
  else {
    try {
      process.kill(-pid, force ? 'SIGKILL' : 'SIGTERM');
    } catch {
      /* gone */
    }
  }
}

function launch(exe, argv, env = {}) {
  const child = spawn(exe, argv, {
    stdio: ['ignore', 'pipe', 'pipe'],
    detached: process.platform !== 'win32',
    env: { ...process.env, RUTBA_WINDOW_DISPLAY: 'offscreen', ELECTRON_DISABLE_SECURITY_WARNINGS: '1', ...env },
  });
  child.stdout.on('data', () => {});
  child.stderr.on('data', () => {});
  child.done = new Promise((resolve) => child.on('exit', (code) => resolve(code)));
  return child;
}

/** Close a copy as gently as Windows allows (its windows asked to close), then for certain. */
async function close(child) {
  stopTree(child.pid, false);
  const code = await Promise.race([child.done, wait(15000).then(() => 'still running')]);
  if (code === 'still running') {
    stopTree(child.pid, true);
    await Promise.race([child.done, wait(10000)]);
  }
}

function powershell(command) {
  const r = spawnSync('powershell', ['-NoProfile', '-Command', command], { encoding: 'utf8' });
  return (r.stdout || '').trim();
}

/** The copies of Rutba Office Windows has registered, as "<version> <scope>". */
function registered() {
  const keys = "'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*','HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*','HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'";
  const text = powershell(`Get-ItemProperty ${keys} -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match 'Rutba' } | ForEach-Object { '{0} {1}' -f $_.DisplayVersion, $(if ($_.PSPath -match 'HKEY_CURRENT_USER') { 'user' } else { 'machine' }) }`);
  return text ? text.split(/\r?\n/) : [];
}

const INSTALLED = () => path.join(process.env.LOCALAPPDATA || '', 'Programs', 'Rutba Office', 'Rutba Office.exe');

/** Whether the copy at `exe` has a window up, within `ms`. */
async function windowUp(exe, ms) {
  const started = Date.now();
  while (Date.now() - started < ms) {
    const n = powershell(`@(Get-Process -Name 'Rutba Office' -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq '${exe.replace(/'/g, "''")}' -and $_.MainWindowHandle -ne 0 }).Count`);
    if (Number(n) > 0) return true;
    await wait(1000);
  }
  return false;
}

/**
 * The older copy started again on the profile; the newer installer run
 * silently while it is open, as the updater runs it (/S --updated); with
 * race 'close' the copy is closed 2 to 5 s later, as quitAndInstall closes
 * it, and with 'never-close' it is left open. Answers whether the newer
 * version is now installed for the report to follow.
 */
async function installWhileRunning({ race: how, from, to, profile, verdict }) {
  say(`starting ${from.version} again, to be open when the ${to.version} installer starts`);
  const running = launch(INSTALLED(), [`--user-data-dir=${profile}`, '--app=home']);
  if (!(await windowUp(INSTALLED(), 120000))) {
    await close(running);
    throw new Error(`${from.version} did not open a window within two minutes`);
  }
  say(`running ${path.basename(to.file)} /S --updated while ${from.version} is open`);
  const startedAt = Date.now();
  const installer = spawn(to.file, ['/S', '--updated'], { stdio: 'ignore' });
  const exited = new Promise((resolve) => installer.on('exit', (code) => resolve(code)));
  let closedAfter = null;
  if (how === 'close') {
    const delay = 2000 + Math.floor(Math.random() * 3000);
    await wait(delay);
    closedAfter = Math.round(delay / 100) / 10;
    say(`closing ${from.version} ${closedAfter} s after the installer started`);
    await close(running);
  }
  // The retries take 15 s; a hang is anything past three minutes (five for an install).
  const limitSeconds = how === 'close' ? 300 : 180;
  const exit = await Promise.race([exited, wait(limitSeconds * 1000).then(() => 'timeout')]);
  const seconds = (Date.now() - startedAt) / 1000;
  if (exit === 'timeout') stopTree(installer.pid, true);
  say(`the installer ${exit === 'timeout' ? 'was stopped' : `ended with ${exit}`} after ${Math.round(seconds)} s`);
  // Whether the installer ended the copy left open, asked before this closes it anyway.
  const appEnded = how === 'never-close' ? await Promise.race([running.done.then(() => true), wait(50).then(() => false)]) : null;
  if (how === 'never-close') await close(running);
  const judged = judgeInstallRace({ race: how, exit, seconds, limitSeconds, copies: registered(), from: from.version, to: to.version, closedAfter, appEnded });
  verdict.failures.push(...judged.failures);
  verdict.evidence.push(...judged.evidence);
  verdict.installer = { exit, seconds: Math.round(seconds), closedAfter };
  return exit === 0;
}

function runInstaller(file) {
  say(`installing ${path.basename(file)}`);
  const r = spawnSync(file, ['/S'], { stdio: 'ignore', timeout: 300000 });
  if (r.status !== 0) throw new Error(`${path.basename(file)} ended with ${r.status}`);
  if (!fs.existsSync(INSTALLED())) throw new Error(`no installed copy at ${INSTALLED()}`);
}

async function main() {
  if (process.platform !== 'win32') throw new Error('the upgrade test runs the Windows executables; run it on Windows');
  if (race && (!install || !['close', 'never-close'].includes(race))) throw new Error('--race is close or never-close, with --install');
  if (install && !process.env.CI && !flag('yes-install-here')) {
    throw new Error('--install installs over this machine\'s copy; it is for a CI runner (CI is set there). Leave it out to run the portable executables in a profile of their own.');
  }
  const to = await resolveTo();
  const from = await resolveFrom(to);
  say(`from ${from.version} (${from.file})`);
  say(`to   ${to.version} (${to.file})`);

  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-upgrade-'));
  const profile = path.join(work, 'profile');
  const document = path.join(work, 'documents', 'Upgrade letter.docx');
  fs.mkdirSync(path.dirname(document), { recursive: true });
  fs.copyFileSync(path.join(root, 'tests', 'corpus-public', 'docx', 'letter.docx'), document);
  const token = crypto.randomBytes(6).toString('hex');
  const seeded = seedProfile(profile, { document, token });
  const verdict = { from: from.version, to: to.version, mode: race ? `installed, race ${race}` : install ? 'installed' : 'portable', profile, ok: false, failures: [], evidence: [] };

  try {
    // The older version, on the profile, until it has recorded itself.
    let older = from.file;
    if (install) {
      runInstaller(from.file);
      older = INSTALLED();
    }
    if (openFile) {
      // With a document of its own, so the older version writes the recent list itself.
      const second = path.join(path.dirname(document), 'Opened by the older version.docx');
      fs.copyFileSync(document, second);
      seeded.opened = second;
    }
    say(`starting ${from.version} on ${profile}`);
    const old = launch(older, [`--user-data-dir=${profile}`, ...(openFile ? ['--app=word', seeded.opened] : ['--app=home'])]);
    const started = Date.now();
    let ran = false;
    while (Date.now() - started < 180000) {
      if (olderHasRun(profile, from.version, { document: seeded.opened || null })) {
        ran = true;
        break;
      }
      if (old.exitCode != null) break;
      await wait(1000);
    }
    // A beat more for anything it writes after its first window.
    if (ran) await wait(3000);
    await close(old);
    if (!ran) throw new Error(`${from.version} did not record itself in the profile within three minutes${old.exitCode != null ? ` (it exited with ${old.exitCode})` : ''}`);
    say(`${from.version} ran over the profile in ${Math.round((Date.now() - started) / 1000)} s and was closed`);

    // The update that races the app's closing: the newer installer started
    // while the older copy is open, as the updater starts it.
    if (race) {
      const done = await installWhileRunning({ race, from, to, profile, verdict });
      if (!done) return finish(verdict, from, to);
    }

    // The newer version, on the same profile, writing what it read.
    let newer = to.file;
    if (race) newer = INSTALLED();
    else if (install) {
      runInstaller(to.file);
      newer = INSTALLED();
      const copies = registered();
      const one = copies.length === 1 && copies[0].startsWith(`${to.version} `);
      (one ? verdict.evidence : verdict.failures).push(one ? `one copy registered: ${copies[0]}` : `the registry has ${copies.length ? copies.join('; ') : 'no copy'}, not one copy of ${to.version}`);
    }
    const reportFile = path.join(work, 'report.json');
    say(`starting ${to.version} on the same profile`);
    const now = launch(newer, [`--user-data-dir=${profile}`, '--app=home'], { RUTBA_OFFICE_UPGRADE_REPORT: reportFile, RUTBA_UPGRADE_FEED: feed ? '1' : '0' });
    const code = await Promise.race([now.done, wait(180000).then(() => 'timeout')]);
    if (code === 'timeout') await close(now);
    // A portable stub can end before the copy it started has written: give the file a moment.
    for (let i = 0; i < 20 && !fs.existsSync(reportFile); i++) await wait(500);
    const report = fs.existsSync(reportFile) ? JSON.parse(fs.readFileSync(reportFile, 'utf8')) : null;
    if (seeded.opened) {
      const kept = (report?.recent || []).some((r) => path.resolve(r.path).toLowerCase() === path.resolve(seeded.opened).toLowerCase());
      (kept ? verdict.evidence : verdict.failures).push(kept ? `the document ${from.version} opened is in ${to.version}'s recent list` : `the document ${from.version} opened is not in ${to.version}'s recent list`);
    }
    const judged = judgeUpgrade({ from: from.version, to: to.version, seeded, report, feed });
    verdict.failures.push(...judged.failures);
    verdict.evidence.push(...judged.evidence);
    verdict.report = report;
  } finally {
    verdict.ok = verdict.failures.length === 0;
    fs.writeFileSync(out, JSON.stringify(verdict, null, 2));
    if (!keep) fs.rmSync(work, { recursive: true, force: true });
  }
  return finish(verdict, from, to);
}

function finish(verdict, from, to) {
  verdict.ok = verdict.failures.length === 0;
  fs.writeFileSync(out, JSON.stringify(verdict, null, 2));
  for (const line of verdict.evidence) say(`ok    ${line}`);
  for (const line of verdict.failures) say(`FAIL  ${line}`);
  say(verdict.ok ? `proven: ${from.version} to ${to.version}, ${verdict.mode}` : `failed: ${verdict.failures.length} of ${verdict.failures.length + verdict.evidence.length}`);
  say(`result: ${out}`);
  return verdict.ok ? 0 : 1;
}

main().then(
  (code) => process.exit(code),
  (err) => {
    console.error(`[upgrade] could not run: ${err?.message || err}`);
    process.exit(2);
  }
);
