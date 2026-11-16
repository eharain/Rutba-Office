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
// After every --install run the uninstall entry is judged too
// (tools/upgrade-check.js judgeRegistration): one in the scope installed to,
// saying the version installed and, from 2.26.0, naming a folder that holds
// the copy; none left in the other scope.
//
// --layout <name> (with --install, not with --race) puts the computer in
// one of the states an install can meet, and proves the installer keeps the
// kind of copy there is and its folder (owner, 16 Nov, D-2026-11-16-03):
//   user              per-user over per-user
//   machine           all users over all users (also --all-users)
//   custom-user       per-user, in a folder of its own (/D=), kept there
//   custom-machine    all users, in a folder of its own (/D=), kept there
//   both              a whole per-user copy, registered, beside the all-users
//                     one: all users kept, the per-user copy removed
//   husk-machine      the all-users copy cut down to its program, with no
//                     entry left (what an earlier removal left on the
//                     owner's PC): made whole again, for all users
//   husk-beside-user  a per-user copy, and the program alone in Program
//                     Files with no entry: per-user kept, the husk removed
//                     (the runner is an administrator)
//   stale-user-entry  the owner's 2.26.0 fault (also --stale-user-entry):
//                     an all-users copy beside the state 2.18.0 left, a
//                     per-user InstallLocation under HKCUSoftware<guid>
//                     and a per-user entry whose uninstaller is gone,
//                     InstallLocation empty: all users kept
//   stale-user-entry-held  the same, with a helper of the older copy still
//                     running from its program (run as Node) through the
//                     update, as on the owner's PC: all users kept, the
//                     helper ended, no husk
//   husk-machine-unelevated  the all-users copy cut down to its program,
//                     and the update run without elevation (a basic-user
//                     token): no prompt, so this person's copy, the husk
//                     left (and said), the pin repointed by the app
// The older installer is run for the layout's kind (and folder), and its
// copy run over the profile; then an app pin is made from it as a taskbar
// pin is made (a .lnk in User PinnedTaskBar aimed at the installed program
// with --app=word), the layout arranged, and the newer installer run as an
// update runs it (/S --updated: no /allusers, /currentuser or /D=). The newer
// copy is started once as the installer's finish page starts it (no check
// variables, so its pin repair runs) and closed, and the pin clicked cold.
// judgeLayout: one entry, in the kind kept, naming the folder kept; a whole
// copy there; no husk (the program without resourcesapp.asar) and no second
// copy anywhere; the pin's target whole, and its command opening Documents.
//
// --fetch-only (with --install) fetches the two installers, writes where they
// are to --out and stops: the upgrade workflow's install and update proofs
// start from them.
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
import { parseVersion, previousRelease, newestRelease, assetName, seedProfile, olderHasRun, judgeUpgrade, profileFiles, goneFromProfile, judgeInstallRace, judgeRegistration, judgeLayout } from './upgrade-check.js';

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

function launch(exe, argv, env = {}, base = process.env) {
  const child = spawn(exe, argv, {
    stdio: ['ignore', 'pipe', 'pipe'],
    detached: process.platform !== 'win32',
    env: { ...base, RUTBA_WINDOW_DISPLAY: 'offscreen', ELECTRON_DISABLE_SECURITY_WARNINGS: '1', ...env },
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

const MACHINE_DIR = () => path.join(process.env.ProgramFiles || 'C:\\Program Files', 'Rutba Office');
const USER_DIR = () => path.join(process.env.LOCALAPPDATA || '', 'Programs', 'Rutba Office');

/**
 * The layouts --layout makes before the newer installer runs, and the kind
 * and folder the install must keep (owner, 16 Nov, D-2026-11-16-03: "the
 * installer should detect these changes and variations"). `older` is what
 * the older installer is run with; `kind` and `dir` what must be there
 * after; `arrange` what is done to the computer once the older copy has run.
 */
const LAYOUTS = {
  user: { older: [], kind: 'user', dir: USER_DIR },
  machine: { older: ['/allusers'], kind: 'machine', dir: MACHINE_DIR },
  'custom-user': { older: ['/currentuser'], kind: 'user', dir: () => path.join(process.env.LOCALAPPDATA || '', 'Rutba Office Test Folder', 'Rutba Office') },
  'custom-machine': { older: ['/allusers'], kind: 'machine', dir: () => path.join(process.env.ProgramFiles || 'C:\\Program Files', 'Rutba Office Test Folder', 'Rutba Office') },
  // A whole per-user copy beside the all-users one: all users is kept.
  both: { older: ['/allusers'], kind: 'machine', dir: MACHINE_DIR, arrange: () => arrangeBoth() },
  // The all-users copy left as its program alone, its entries gone: made whole again, for all users.
  'husk-machine': { older: ['/allusers'], kind: 'machine', dir: MACHINE_DIR, arrange: () => arrangeMachineHusk() },
  // A per-user copy beside an all-users husk with no entries: per-user kept, the husk gone (this runner is elevated).
  'husk-beside-user': { older: [], kind: 'user', dir: USER_DIR, arrange: () => arrangeHuskBesideUser() },
  // The owner's 2.26.0 fault: 2.25.0 for all users, beside what 2.18.0 left per-user.
  'stale-user-entry': { older: ['/allusers'], kind: 'machine', dir: MACHINE_DIR, arrange: () => makeStaleUserEntry() },
  // The owner's fault as it happened: the same, and a helper of the older copy still running from its program
  // (the program run as Node, as the AI tools' helpers run it) while the update runs, holding the file open.
  'stale-user-entry-held': { older: ['/allusers'], kind: 'machine', dir: MACHINE_DIR, arrange: () => { const made = makeStaleUserEntry(); holdOlderProgram(MACHINE_DIR()); return made; } },
  // An update that is not elevated, with only a husk left for all users: it must raise no prompt, so it installs
  // for this person and leaves the husk (said in its log); the pin made from the all-users copy is repointed by
  // the app's pin repair.
  'husk-machine-unelevated': { older: ['/allusers'], kind: 'user', dir: USER_DIR, unelevated: true, husks: () => [MACHINE_DIR()], arrange: () => arrangeMachineHusk() },
};
const layoutName = option('layout') || (flag('stale-user-entry') ? 'stale-user-entry' : flag('all-users') ? 'machine' : 'user');
const LAYOUT = LAYOUTS[layoutName];
// A pin made, the copy started once as the finish page starts it, and the pin clicked cold: for any layout asked for.
const pinTest = Boolean(option('layout')) || flag('stale-user-entry');
const allUsers = LAYOUT?.kind === 'machine';
const INSTALLED = () => path.join(LAYOUT.dir(), 'Rutba Office.exe');
/** What the older installer is run with: its kind, and /D= (last, as NSIS reads it) for a custom folder. */
const olderArgs = () => [...LAYOUT.older, ...(layoutName.startsWith('custom-') ? [`/D=${LAYOUT.dir()}`] : [])];

/** Whether `dir` holds the program, and resources\app.asar beside it. */
const folderFacts = (dir, scope) => ({ dir, scope, exe: fs.existsSync(path.join(dir, 'Rutba Office.exe')), asar: fs.existsSync(path.join(dir, 'resources', 'app.asar')) });

/** Every Rutba Office uninstall entry, as judgeRegistration reads them, with its key's name. */
function entries() {
  const keys = "'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*','HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*','HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'";
  const text = powershell(`@(Get-ItemProperty ${keys} -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match 'Rutba' } | ForEach-Object { [pscustomobject]@{ key = $_.PSChildName; scope = $(if ($_.PSPath -match 'HKEY_CURRENT_USER') { 'user' } else { 'machine' }); version = $_.DisplayVersion; installLocation = $_.InstallLocation; uninstallString = $_.UninstallString } }) | ConvertTo-Json -Compress`);
  let list = [];
  try {
    list = text ? [].concat(JSON.parse(text)) : [];
  } catch {
    list = [];
  }
  return list.map((e) => ({ ...e, installLocation: e.installLocation || '', works: Boolean(e.installLocation) && folderFacts(e.installLocation).exe && folderFacts(e.installLocation).asar }));
}

const regAdd = (key, name, value) => spawnSync('reg', ['add', key, '/v', name, '/t', 'REG_SZ', '/d', value, '/f'], { stdio: 'ignore' }).status === 0;
const regDelete = (key, value = null) => spawnSync('reg', ['delete', key, ...(value ? ['/v', value] : []), '/f', '/reg:64'], { stdio: 'ignore' });

/** The GUID the all-users entry is filed by, which every scope's keys share. */
function machineGuid() {
  const machine = entries().find((e) => e.scope === 'machine');
  if (!machine) throw new Error(`the older installer, run with /allusers, registered no copy for all users (${registered().join('; ') || 'none'})`);
  return machine.key;
}

/** The User Pinned\TaskBar folder, where Windows keeps a taskbar pin's .lnk. */
const TASKBAR = () => path.join(process.env.APPDATA || '', 'Microsoft', 'Internet Explorer', 'Quick Launch', 'User Pinned', 'TaskBar');
const PIN = () => path.join(TASKBAR(), 'Rutba Office upgrade test.lnk');

/** A pin made from the installed copy as a taskbar pin is made: a .lnk aimed at the program, for Documents. */
function makePin() {
  fs.mkdirSync(TASKBAR(), { recursive: true });
  const q = (s) => `'${String(s).replace(/'/g, "''")}'`;
  powershell(`$l = (New-Object -ComObject WScript.Shell).CreateShortcut(${q(PIN())}); $l.TargetPath = ${q(INSTALLED())}; $l.Arguments = '--app=word'; $l.WorkingDirectory = ${q(path.dirname(INSTALLED()))}; $l.Save()`);
  if (!fs.existsSync(PIN())) throw new Error(`could not make the pin at ${PIN()}`);
  say(`made a pin at ${PIN()}, aimed at ${INSTALLED()}`);
}

/**
 * What 2.18.0 left beside an all-users copy: the per-user InstallLocation and
 * a per-user entry whose uninstaller is gone, InstallLocation empty, under
 * the GUID the all-users entry is filed by. Answers what was made, for the
 * clean-up.
 */
function makeStaleUserEntry() {
  const guid = machineGuid();
  const userDir = USER_DIR();
  const uninstallKey = `HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${guid}`;
  const installKey = `HKCU\\Software\\${guid}`;
  const uninstaller = `"${path.join(userDir, 'Uninstall Rutba Office.exe')}" /currentuser`;
  const ok = regAdd(installKey, 'InstallLocation', userDir)
    && regAdd(uninstallKey, 'DisplayName', 'Rutba Office')
    && regAdd(uninstallKey, 'DisplayVersion', '2.18.0')
    && regAdd(uninstallKey, 'UninstallString', uninstaller)
    && regAdd(uninstallKey, 'QuietUninstallString', `${uninstaller} /S`)
    && regAdd(uninstallKey, 'InstallLocation', '');
  if (!ok) throw new Error('could not write the stale per-user entry');
  say(`made the stale per-user entry (${guid}, 2.18.0, uninstaller gone)`);
  return { guid, userDir, uninstallKey, installKey, marker: '2.18.0' };
}

/** A whole per-user copy beside the all-users one, registered as a per-user install registers it. */
function arrangeBoth() {
  const guid = machineGuid();
  const userDir = USER_DIR();
  fs.cpSync(MACHINE_DIR(), userDir, { recursive: true });
  const uninstallKey = `HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${guid}`;
  const installKey = `HKCU\\Software\\${guid}`;
  const uninstaller = `"${path.join(userDir, 'Uninstall Rutba Office.exe')}" /currentuser`;
  const ok = regAdd(installKey, 'InstallLocation', userDir)
    && regAdd(uninstallKey, 'DisplayName', 'Rutba Office')
    && regAdd(uninstallKey, 'DisplayVersion', 'both-test')
    && regAdd(uninstallKey, 'UninstallString', uninstaller)
    && regAdd(uninstallKey, 'QuietUninstallString', `${uninstaller} /S`);
  if (!ok) throw new Error('could not register the per-user copy');
  say(`made a whole per-user copy in ${userDir} beside the one for all users`);
  return { guid, userDir, uninstallKey, installKey, marker: 'both-test' };
}

/** The all-users copy left as an earlier cut-short removal left the owner's: the program alone, no entries. */
function arrangeMachineHusk() {
  const guid = machineGuid();
  const dir = MACHINE_DIR();
  for (const e of fs.readdirSync(dir)) if (e !== 'Rutba Office.exe') fs.rmSync(path.join(dir, e), { recursive: true, force: true });
  regDelete(`HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${guid}`);
  regDelete(`HKLM\\Software\\${guid}`);
  say(`left only the program in ${dir}, and no entry for it`);
  return null;
}

/** An all-users husk with no entries beside the per-user copy. */
function arrangeHuskBesideUser() {
  fs.mkdirSync(MACHINE_DIR(), { recursive: true });
  fs.copyFileSync(path.join(USER_DIR(), 'Rutba Office.exe'), path.join(MACHINE_DIR(), 'Rutba Office.exe'));
  say(`put the program alone, unregistered, in ${MACHINE_DIR()}`);
  return null;
}

/** A helper of the older copy: its program run as Node, with a script that waits, holding the program open. */
let holder = null;
function holdOlderProgram(dir) {
  const script = path.join(os.tmpdir(), `rutba-upgrade-hold-${process.pid}.js`);
  fs.writeFileSync(script, 'setTimeout(() => {}, 30 * 60 * 1000);\n');
  holder = spawn(path.join(dir, 'Rutba Office.exe'), [script], { env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }, stdio: 'ignore' });
  holder.ended = new Promise((resolve) => holder.on('exit', () => resolve(true)));
  say(`started a helper of the older copy from ${dir} (pid ${holder.pid})`);
}

/**
 * The newer installer run as a person who is not an administrator runs it: with a basic-user token (runas
 * /trustlevel), which has no administrators group, as an app's update runs it on a computer where the person
 * signs in without elevation. runas does not wait: the command writes its exit code to a file, read here.
 */
async function runUnelevated(file, more, work) {
  if (/\s/.test(work) || /\s/.test(file)) throw new Error(`runas needs paths without spaces: ${work}, ${file}`);
  const codeFile = path.join(work, 'unelevated-exit.txt');
  const cmdFile = path.join(work, 'unelevated.cmd');
  fs.rmSync(codeFile, { force: true });
  fs.writeFileSync(cmdFile, `@"${file}" /S ${more.join(' ')}\r\n@echo %errorlevel%> "${codeFile}"\r\n`);
  say(`installing ${path.basename(file)} /S ${more.join(' ')}, not elevated`);
  spawnSync('runas', ['/trustlevel:0x20000', `cmd /d /c ${cmdFile}`], { stdio: 'ignore', windowsVerbatimArguments: false });
  for (let i = 0; i < 300 && !fs.existsSync(codeFile); i++) await wait(1000);
  const code = fs.existsSync(codeFile) ? Number(fs.readFileSync(codeFile, 'utf8').trim()) : null;
  if (code !== 0) throw new Error(`${path.basename(file)}, not elevated, ended with ${code ?? 'nothing after five minutes'}`);
}

/** What a layout made in the registry, taken away again if the run left it: only what it made. */
function dropArranged(made) {
  fs.rmSync(PIN(), { force: true });
  if (holder && holder.exitCode == null) stopTree(holder.pid, true);
  if (!made) return;
  const version = powershell(`(Get-ItemProperty 'Registry::HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${made.guid}' -ErrorAction SilentlyContinue).DisplayVersion`);
  if (version === made.marker) spawnSync('reg', ['delete', made.uninstallKey, '/f'], { stdio: 'ignore' });
  const facts = folderFacts(made.userDir, 'user');
  if (!(facts.exe && facts.asar)) spawnSync('reg', ['delete', made.installKey, '/v', 'InstallLocation', '/f'], { stdio: 'ignore' });
}

/** A pin's target and arguments, as Windows reads the .lnk. */
function readLink(file) {
  const q = `'${file.replace(/'/g, "''")}'`;
  const text = powershell(`$l = (New-Object -ComObject WScript.Shell).CreateShortcut(${q}); [pscustomobject]@{ target = $l.TargetPath; arguments = $l.Arguments } | ConvertTo-Json -Compress`);
  try {
    return JSON.parse(text);
  } catch {
    return { target: null, arguments: null };
  }
}

/** The titles of the windows of every process started from `exe`. */
const windowTitles = (exe) => powershell(`@(Get-Process -Name 'Rutba Office' -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq '${exe.replace(/'/g, "''")}' -and $_.MainWindowTitle } | ForEach-Object { $_.MainWindowTitle }) -join [char]10`).split(/\r?\n/).filter(Boolean);

/** The environment without the check variables, as a person's start has it: with them, Office runs no pin repair. */
const personEnv = () => Object.fromEntries(Object.entries(process.env).filter(([k]) => !/^RUTBA_OFFICE_(VERIFY_|SMOKE)/.test(k)));

/**
 * After the newer installer: its copy (where its entry says, else where the
 * layout keeps it) started once as the finish page starts it, given time for
 * its pin repair (10 s after start) and closed; then the pin clicked cold,
 * and the whole judged (judgeLayout).
 */
async function judgeLayoutRun({ to, work, verdict }) {
  const found = entries();
  const mine = found.find((e) => e.scope === LAYOUT.kind && e.installLocation);
  const installed = path.join(mine?.installLocation || LAYOUT.dir(), 'Rutba Office.exe');
  if (fs.existsSync(installed)) {
    say(`starting ${to.version} once, as the installer's finish page would`);
    const first = launch(installed, [`--user-data-dir=${path.join(work, 'finish-profile')}`], {}, personEnv());
    await wait(25000);
    await close(first);
  }
  const link = readLink(PIN());
  const target = link.target || '';
  const facts = target ? folderFacts(path.dirname(target)) : { exe: false, asar: false };
  const dirs = new Map();
  for (const [dir, scope] of [[MACHINE_DIR(), 'machine'], [USER_DIR(), 'user'], [LAYOUT.dir(), LAYOUT.kind], ...found.filter((e) => e.installLocation).map((e) => [e.installLocation, e.scope]), ...(target ? [[path.dirname(target), 'pin']] : [])]) {
    if (!dirs.has(dir.toLowerCase())) dirs.set(dir.toLowerCase(), folderFacts(dir, scope));
  }
  let titles = [];
  if (target && facts.exe) {
    say(`clicking the pin cold: ${target} ${link.arguments}`);
    const clicked = launch(target, [...String(link.arguments || '').split(/\s+/).filter(Boolean), `--user-data-dir=${path.join(work, 'pin-profile')}`], {}, personEnv());
    for (let i = 0; i < 30 && !titles.some((t) => /Rutba (Documents|Word)/.test(t)); i++) {
      await wait(1000);
      titles = windowTitles(target);
    }
    await close(clicked);
  }
  // A helper of the older copy left running from its program: ended by the install, or the program could not go.
  const held = holder ? { ended: holder.exitCode != null || (await Promise.race([holder.ended, wait(50).then(() => false)])) } : null;
  const husksLeft = LAYOUT.husks ? LAYOUT.husks() : [];
  const judged = judgeLayout({ to: to.version, scope: LAYOUT.kind, dir: LAYOUT.dir(), entries: found, folders: [...dirs.values()], link: { ...link, exe: facts.exe, asar: facts.asar }, titles, husksLeft, held });
  verdict.failures.push(...judged.failures);
  verdict.evidence.push(...judged.evidence);
  verdict.layout = { name: layoutName, entries: found, folders: [...dirs.values()], link, titles };
}

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

/** The installer run silently; with `expect`, the copy must be where the layout keeps it, or the run cannot go on. */
function runInstaller(file, more = [], expect = true) {
  say(`installing ${path.basename(file)} ${['/S', ...more].join(' ')}`);
  const r = spawnSync(file, ['/S', ...more], { stdio: 'ignore', timeout: 300000 });
  if (r.status !== 0) throw new Error(`${path.basename(file)} ended with ${r.status}`);
  if (expect && !fs.existsSync(INSTALLED())) throw new Error(`no installed copy at ${INSTALLED()}`);
}

async function main() {
  if (process.platform !== 'win32') throw new Error('the upgrade test runs the Windows executables; run it on Windows');
  if (race && (!install || !['close', 'never-close'].includes(race))) throw new Error('--race is close or never-close, with --install');
  if (!LAYOUT) throw new Error(`--layout is one of ${Object.keys(LAYOUTS).join(', ')}`);
  if (pinTest && (!install || race)) throw new Error('--layout and --stale-user-entry are with --install, and not with --race');
  if (install && !process.env.CI && !flag('yes-install-here')) {
    throw new Error('--install installs over this machine\'s copy; it is for a CI runner (CI is set there). Leave it out to run the portable executables in a profile of their own.');
  }
  const to = await resolveTo();
  const from = await resolveFrom(to);
  say(`from ${from.version} (${from.file})`);
  say(`to   ${to.version} (${to.file})`);
  // --fetch-only: the two files fetched and named in --out, for the installer and update proofs (tools/install-test.ps1, update-test.ps1).
  if (flag('fetch-only')) {
    fs.writeFileSync(out, JSON.stringify({ from, to }, null, 2));
    return 0;
  }

  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-upgrade-'));
  const profile = path.join(work, 'profile');
  const document = path.join(work, 'documents', 'Upgrade letter.docx');
  fs.mkdirSync(path.dirname(document), { recursive: true });
  fs.copyFileSync(path.join(root, 'tests', 'corpus-public', 'docx', 'letter.docx'), document);
  const token = crypto.randomBytes(6).toString('hex');
  const seeded = seedProfile(profile, { document, token });
  const verdict = { from: from.version, to: to.version, mode: pinTest ? `installed over the ${layoutName} layout, kept for ${allUsers ? 'all users' : 'this person'} in ${LAYOUT.dir()}` : allUsers ? `installed for all users, race ${race}` : race ? `installed, race ${race}` : install ? 'installed' : 'portable', profile, ok: false, failures: [], evidence: [] };
  let made = null;
  const judgeEntry = () => {
    const judged = judgeRegistration({ to: to.version, scope: allUsers ? 'machine' : 'user', entries: entries() });
    verdict.failures.push(...judged.failures);
    verdict.evidence.push(...judged.evidence);
  };

  try {
    // The older version, on the profile, until it has recorded itself.
    let older = from.file;
    if (install) {
      runInstaller(from.file, olderArgs());
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
      judgeEntry();
    }

    // The newer version, on the same profile, writing what it read.
    let newer = to.file;
    if (race) newer = INSTALLED();
    else if (install) {
      // A layout: a pin made from the older copy, the computer arranged as
      // the layout says, and the newer installer run as the updater runs it:
      // silent, --updated, no /allusers or /currentuser, no /D=.
      if (pinTest) {
        makePin();
        made = LAYOUT.arrange ? LAYOUT.arrange() : null;
      }
      if (LAYOUT.unelevated) await runUnelevated(to.file, ['--updated'], work);
      else runInstaller(to.file, pinTest ? ['--updated'] : [], !pinTest);
      const copies = registered();
      const one = copies.length === 1 && copies[0].startsWith(`${to.version} `);
      (one ? verdict.evidence : verdict.failures).push(one ? `one copy registered: ${copies[0]}` : `the registry has ${copies.length ? copies.join('; ') : 'no copy'}, not one copy of ${to.version}`);
      if (pinTest) await judgeLayoutRun({ to, work, verdict });
      else judgeEntry();
      // Where the newer copy went, if not where it was kept: the profile is still opened by it.
      const went = entries().find((e) => e.version === parseVersion(to.version)?.text && e.works);
      newer = fs.existsSync(INSTALLED()) ? INSTALLED() : went ? path.join(went.installLocation, 'Rutba Office.exe') : null;
      if (!newer) {
        verdict.failures.push(`no working copy of ${to.version} anywhere after the install`);
        return finish(verdict, from, to);
      }
    }
    const reportFile = path.join(work, 'report.json');
    const leftByOlder = profileFiles(profile);
    say(`starting ${to.version} on the same profile`);
    const now = launch(newer, [`--user-data-dir=${profile}`, '--app=home'], { RUTBA_OFFICE_UPGRADE_REPORT: reportFile, RUTBA_UPGRADE_FEED: feed ? '1' : '0' });
    const code = await Promise.race([now.done, wait(180000).then(() => 'timeout')]);
    if (code === 'timeout') await close(now);
    (code === 0 ? verdict.evidence : verdict.failures).push(code === 0 ? `${to.version} wrote its report and ended cleanly` : `${to.version} did not end cleanly: ${code}`);
    const gone = goneFromProfile(leftByOlder, profileFiles(profile));
    (gone.length ? verdict.failures : verdict.evidence).push(gone.length ? `${to.version} removed what ${from.version} left: ${gone.slice(0, 8).join(', ')}` : `none of the ${leftByOlder.length} files ${from.version} left is gone`);
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
    if (pinTest) dropArranged(made);
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
