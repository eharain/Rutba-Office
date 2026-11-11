// Everything before a release, in lanes so that what can run side by side does.
//
//   npm run gate
//   npm run gate -- --record     and write what it counted into the release
//                                note and the README (a full run that passed)
//   npm run gate -- --apps word,mail   the window checks of those chunks only
//   npm run gate -- --changed    the window checks of the chunks changed since
//                                the last release tag (all of them when what
//                                every app shares has changed); with --record
//                                on a patch release (2.17.1), what it checked
//                                is put on the record (D-2026-10-23-01)
//   npm run gate -- --resume     carry on from the last run on this commit:
//                                what passed is not run again, and --record
//                                adds the runs up
//   npm run gate -- --jobs 3     how many check processes side by side
//                                (otherwise from the processors and memory)
//   npm run gate -- --serial     one process, every check in turn, as before
//
// The order (owner, 2026-10-21: "splitting this long running gate into
// smaller chunks is very important so that lighter or non locking parts can
// be run in parallel"):
//
//   build         the windows bundled once, for every check process below.
//   test:perf     the performance budgets, alone, while the machine is still:
//                 a budget timed beside other work measures the machine.
//   side by side  the engine tests, the editing checks and a window check
//                 process a chunk (an app, or a part of what every app shares:
//                 shell, ribbon, look, theme, access, languages), as many at once as the
//                 processors and memory allow, the longest chunk first. Each
//                 process has a clipboard and global shortcuts of its own
//                 (machine-fakes.js), so none touches another's or the
//                 person's; each logs to .gate/logs/<chunk>.log.
//   machine       alone: the parts that need the computer to themselves (its
//                 own clipboard and keys, checked through the system adapters,
//                 timing budgets, drags, native copies).
//   again         alone, the parts whose checks failed above, in a process
//                 started just for them (owner, 2026-10-21: "Queue the failed
//                 ones to run at the end"). The gate passes only if every
//                 check that failed is seen passing here, by its name; one
//                 that fails again closes it. The record names each one.
//   smoke         every window paints without logging anything.
//   then          the app packs added and removed; the rich fixtures saved,
//                 edited and saved as Strict, opened in Word, Excel and
//                 PowerPoint themselves (office:reopen, so this machine needs
//                 Microsoft Office); the four fuzzers; the corpus.
//
// Lanes side by side: as many as the processors allow (one to each two and a
// half, four at most), one more started only while 1.5 GB of memory is free,
// looked at again every ten seconds, so a run that starts short of memory
// takes more lanes as memory comes free, and says why it has fewer.
//
// Nothing here is optional before a release. With --record the gate puts its
// own counts on the record: a Gate section in docs/releases/v<version>.md
// saying where and when it ran and what each pass counted, and the README's
// counts brought to the same numbers.

import { spawn } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/*
 * The window checks by app (owner, 2026-10-20: "lets do that first then").
 * verify:apps names each part's app; 'shell' is what every app shares. A run
 * can be limited to some apps, and picks up after the apps that passed.
 */
export const VERIFY_APPS = ['shell', 'ribbon', 'look', 'theme', 'access', 'languages', 'word', 'sheets', 'slides', 'pictures', 'image', 'video', 'recorder', 'broadcast', 'talk', 'base', 'pdf', 'mail', 'calendar', 'contacts', 'locker'];

/** The apps a changed file touches; 'shell' means every app. */
export function appsForPath(file) {
  const f = String(file).replace(/\\/g, '/');
  // What makes or publishes the shipped build touches every app.
  if (/^tools\/(?:make-|sbom|check-packs|publish-release)|^\.github\/workflows\/release\.yml$/.test(f)) return ['shell'];
  if (/^(docs|tools|\.github)\/|^[^/]+\.md$|^tests\//.test(f)) return [];
  const app = /^apps\/desktop\/renderer\/apps\/([a-z]+)/.exec(f)?.[1];
  if (app) {
    const key = { home: 'shell', edit: 'shell' }[app] || app;
    return VERIFY_APPS.includes(key) ? [key] : ['shell'];
  }
  if (/^apps\/desktop\/main\/extras\/locker\//.test(f)) return ['locker'];
  // The studio both the Recorder and Live Broadcast are.
  if (/^apps\/desktop\/renderer\/studio\//.test(f)) return ['recorder', 'broadcast'];
  const check = /^apps\/desktop\/main\/verify-([a-z]+)/.exec(f)?.[1];
  if (check) return [{ locker: 'locker', word: 'word', sheet: 'sheets', deck: 'slides', slide: 'slides', mail: 'mail', video: 'video', recorder: 'recorder', broadcast: 'broadcast', talk: 'talk', base: 'base', pdf: 'pdf', picture: 'pictures', pictures: 'pictures', image: 'image', calendar: 'calendar', contacts: 'contacts', ribbon: 'ribbon', polish: 'theme', dark: 'theme', gallery: 'look', soon: 'look', panels: 'look', icons: 'look', a11y: 'access', keyboard: 'access', high: 'access', screen: 'access', dialog: 'access', lost: 'access', accessibility: 'access', urdu: 'languages', spelling: 'languages' }[check] || 'shell'];
  const service = /^apps\/desktop\/main\/([a-z]+)/.exec(f)?.[1];
  if (service) return [{ mail: 'mail', live: 'broadcast', talk: 'talk', base: 'base', pdf: 'pdf', calendar: 'calendar', contacts: 'contacts', dav: 'calendar' }[service] || 'shell'];
  const pkg = /^packages\/([a-z-]+)\//.exec(f)?.[1];
  if (pkg) {
    return {
      // tests/gate-apps.test.js holds this to what each app really imports (tools/app-map.mjs).
      ooxml: ['word', 'sheets', 'slides', 'base', 'mail'], 'doc-view': ['word'], 'sheet-view': ['sheets', 'base'], formula: ['sheets', 'base'], presentation: ['slides'],
      mailbox: ['mail'], calendar: ['calendar', 'mail', 'contacts'], contacts: ['contacts', 'mail'], pdf: ['pdf', 'word', 'sheets', 'slides'],
      imaging: ['image', 'pictures', 'broadcast', 'slides', 'word'], media: ['video', 'recorder', 'broadcast', 'pictures'], stream: ['broadcast'], studio: ['video', 'recorder', 'broadcast', 'image', 'slides'], database: ['base'], ocr: ['pdf', 'pictures', 'sheets', 'image', 'mail'],
    }[pkg] || ['shell'];
  }
  return ['shell'];
}

/** The apps to check for these changed files: every app when the shared part changed. */
export function appsForChanges(files) {
  const set = new Set(files.flatMap(appsForPath));
  return set.has('shell') ? VERIFY_APPS.slice() : VERIFY_APPS.filter((a) => set.has(a));
}

const git = (...args) => { try { return execFileSync('git', args, { cwd: ROOT, encoding: 'utf8', maxBuffer: 2 ** 30 }).trim(); } catch { return ''; } };

/** `git status --porcelain` as git says it: the first column is a space when a file is changed but not staged, so it is not trimmed. */
export function porcelain(cwd = ROOT) {
  try { return execFileSync('git', ['status', '--porcelain'], { cwd, encoding: 'utf8' }); } catch { return ''; }
}

/**
 * What the working tree holds beyond HEAD, as one key for --resume: the
 * status, every change to a tracked file byte for byte (binary ones too), and
 * the contents of each file git does not track yet, so an old pass never
 * stands for a file edited since.
 */
export function treeKey(cwd = ROOT) {
  const run = (args) => { try { return execFileSync('git', args, { cwd, maxBuffer: 2 ** 31 - 1, stdio: ['ignore', 'pipe', 'ignore'] }); } catch { return null; } };
  const hash = crypto.createHash('sha256');
  hash.update(porcelain(cwd));
  hash.update(run(['diff', 'HEAD', '--binary']) ?? 'no diff');
  const untracked = String(run(['ls-files', '--others', '--exclude-standard', '-z']) || '').split('\0').filter(Boolean).sort();
  for (const f of untracked) {
    hash.update(`\0${f}\0`);
    try { hash.update(fs.readFileSync(path.join(cwd, f))); } catch { hash.update('unreadable'); }
  }
  return hash.digest('hex');
}

/** The apps a --changed run checks. With no tag to compare with, nothing can be said to be unchanged: all of them. */
export function wantedForChanged(since, files) {
  return since ? appsForChanges(files) : VERIFY_APPS.slice();
}

/** Whether a machine lane that ran is good for these apps: it ran for each of them, and not as part of a --serial run. */
export function machineCovers(machine, apps) {
  return Boolean(machine?.ok) && !machine.serial && apps.every((a) => machine.scope?.includes(a));
}
// RUTBA_GATE_DIR moves the run's own state, logs and lock elsewhere: the gate's tests run it beside a real one.
const STATE_DIR = process.env.RUTBA_GATE_DIR ? path.resolve(process.env.RUTBA_GATE_DIR) : path.join(ROOT, '.gate');
const STATE = path.join(STATE_DIR, 'state.json');
const FLOORS = path.join(ROOT, 'tools', 'gate-floors.json');
/** What the last run on this commit, with this working tree, left to do. */
function readState(head, tree) {
  try {
    const s = JSON.parse(fs.readFileSync(STATE, 'utf8'));
    return s.head === head && s.tree === tree ? s : null;
  } catch {
    return null;
  }
}


/** What each pass counted, read from the lines it ends with. */
export function countsOf(name, output) {
  const text = String(output || '').replace(/\u001b\[[0-9;]*m/g, '').replace(/\r/g, '');
  const all = (re) => [...text.matchAll(re)].map((m) => Number(m[1]));
  const sum = (re) => all(re).reduce((n, v) => n + v, 0);
  const ratio = (re) => { const m = [...text.matchAll(re)].pop(); return m ? { passed: Number(m[1]), of: Number(m[2]) } : null; };
  if (name === 'test' || name === 'test:perf') {
    // The spec reporter ends with "ℹ tests 2209", the TAP reporter (output
    // that is not a terminal, as a gate run into a log is) with "# tests 2209".
    const tests = sum(/^(?:ℹ|#) tests (\d+)$/gm);
    return tests ? { tests, passed: sum(/^(?:ℹ|#) pass (\d+)$/gm), failed: sum(/^(?:ℹ|#) fail (\d+)$/gm), skipped: sum(/^(?:ℹ|#) skipped (\d+)$/gm) } : null;
  }
  if (name === 'verify:edit') return ratio(/^(\d+)\/(\d+) editing checks passed/gm);
  if (name === 'verify:apps') return ratio(/^(\d+)\/(\d+) application checks passed/gm);
  if (name === 'smoke') return ratio(/^(\d+)\/(\d+) windows rendered clean/gm);
  if (name.startsWith('fuzz:')) {
    const m = [...text.matchAll(/^(\d+) (?:damaged files|damaged PDFs|messages): .*faults (\d+|undefined)\s*$/gm)].pop();
    if (!m) return null;
    const faults = Number(m[2]) || 0;
    return { passed: Number(m[1]) - faults, of: Number(m[1]) };
  }
  if (name === 'verify:packs' || name === 'office:reopen') {
    const lines = (re) => text.match(re)?.length || 0;
    const [ok, bad] = [lines(/^ok\b/gm), lines(/^FAIL\b/gm)];
    return ok + bad ? { passed: ok, of: ok + bad } : null;
  }
  if (name === 'corpus') {
    const m = /^# Corpus run: (\d+) files/m.exec(text);
    if (!m) return null;
    // One that faulted or hung did not pass, whatever the count of files: columns five and six of the table.
    const bad = [...text.matchAll(/^\| \S+ \| \d+ \| \d+ \| \d+ \| \d+ \| (\d+) \| (\d+) \|/gm)].reduce((t, r) => t + Number(r[1]) + Number(r[2]), 0);
    return { passed: Number(m[1]) - bad, of: Number(m[1]) };
  }
  return null;
}

const n = (v) => Number(v).toLocaleString('en-GB');
const PLATFORM = { win32: 'Windows', darwin: 'macOS', linux: 'Linux' };

/** The damaged files the fuzzers went through. */
const fuzzed = (counts) => FUZZERS.reduce((total, [name]) => total + (counts[name]?.of || 0), 0);

/** The fuzzers: each damages good files with the seed it is given, and fails on an answer that is not a sentence. */
export const FUZZERS = [['fuzz:open', 'tools/fuzz-open.js', 'damaged Office files', 300], ['fuzz:pdf', 'tools/fuzz-pdf.js', 'damaged PDFs', 300], ['fuzz:smime', 'tools/fuzz-smime.js', 'damaged signed mail', 2000], ['fuzz:grammar', 'tools/fuzz-grammar.js', 'files made to do harm', 600]];

/** The release note's Gate section, in the note's own plain style. */
export function gateSection(counts, { version, date, fewer = [], platform = process.platform, arch = process.arch, reran = [], apps = null, commit = null }) {
  const t = engineCounts(counts);
  const parts = [
    t ? `${n(t.tests)} engine tests, ${n(t.passed)} passing${t.skipped ? ` and ${n(t.skipped)} skipped where a package they need is not installed` : ''}` : null,
    counts['verify:edit'] ? `${n(counts['verify:edit'].passed)} of ${n(counts['verify:edit'].of)} editing checks` : null,
    counts['verify:apps'] && apps?.length !== 0 ? `${n(counts['verify:apps'].passed)} of ${n(counts['verify:apps'].of)} application checks${apps ? `, of ${listed(apps)}, the ${apps.length === 1 ? 'app' : 'apps'} it changed` : ''}` : null,
    counts.smoke ? `${n(counts.smoke.passed)} of ${n(counts.smoke.of)} windows painted clean` : null,
    counts['office:reopen'] ? `${n(counts['office:reopen'].passed)} of ${n(counts['office:reopen'].of)} saves reopened in Word, Excel and PowerPoint with what the original holds` : null,
    fuzzed(counts) ? `${n(fuzzed(counts))} damaged files read or refused in a sentence` : null,
    counts.corpus ? `${n(counts.corpus.of)} files from outside the project opened` : null,
  ].filter(Boolean);
  const list = parts.length > 1 ? `${parts.slice(0, -1).join('; ')}; and ${parts.at(-1)}` : parts[0];
  const again = reran.length
    ? `\n\n${wrap(`${reran.length === 1 ? 'One check' : `${n(reran.length)} checks`} failed beside the others and passed when run again alone at the end of the same gate:`)}\n\n${reran.map((r) => `- ${r}`).join('\n')}`
    : '';
  const less = fewer.length ? `\n\n${wrap(`Fewer checks than the gate before, allowed by hand: ${fewer.join('; ')}.`)}` : '';
  const at = `${less}${commit ? `\n\nCommit checked: ${commit}` : ''}`;
  return `## Gate\n\n${wrap(`Run on ${PLATFORM[platform] || platform} (${arch}) on ${date}, before ${version} was released: ${list}.`)}${at}${again}\n`;
}

/**
 * What a run must not inherit. Each of these makes a check process run less
 * than the whole, and a gate started from a shell that has one set would
 * count what little ran as the gate.
 */
export const SHRINKING_ENV = ['RUTBA_VERIFY_ONLY', 'RUTBA_VERIFY_APPS', 'RUTBA_VERIFY_PART_IDS', 'RUTBA_VERIFY_LANE', 'RUTBA_VERIFY_PARTS', 'RUTBA_LOCKER', 'RUTBA_CORPUS_LIMIT', 'RUTBA_A11Y_APPS', 'RUTBA_DARK_APPS', 'RUTBA_POLISH_APPS', 'RUTBA_RIBBON_APPS', 'RUTBA_SCREENS_ONLY'];

/** More checks than this failing beside the others closes the gate: that is a fault, not a busy machine. */
export const RERUN_LIMIT = 5;

/** The counts the last full gate wrote into the README, the least this one may run. */
export function floorsFrom(readme) {
  const text = String(readme || '');
  const num = (re) => { const m = re.exec(text); return m ? Number(m[1].replace(/,/g, '')) : null; };
  return {
    tests: num(/what runs before a release: ([\d,]+) engine tests/),
    edit: num(/engine tests, ([\d,]+) editing checks/),
    apps: num(/editing checks,\s+([\d,]+) application checks/),
    smoke: num(/and ([\d,]+) window captures/),
  };
}

/** What this full run counted fewer of than the last: [] when nothing fell. */
export function belowFloor(counts, floors) {
  const all = engineCounts(counts);
  return [
    ['engine tests', all?.tests, floors.tests],
    ['editing checks', counts['verify:edit']?.of, floors.edit],
    ['application checks', counts['verify:apps']?.of, floors.apps],
    ['window captures', counts.smoke?.of, floors.smoke],
  ].filter(([, now, was]) => was != null && (now == null || now < was)).map(([what, now, was]) => `${what}: ${now ?? 'none'}, down from ${was}`);
}

/** The changes in a tree that stop a run going on the record: all but the note and README a record writes. */
export function dirtyForRecord(porcelain, version) {
  const own = new Set(['README.md', 'tools/gate-floors.json', `docs/releases/v${version}.md`]);
  // A floors file taken away is not a file the record writes: it would leave nothing to be held to.
  return String(porcelain || '').split('\n').filter(Boolean).filter((l) => /^.?D/.test(l.slice(0, 2)) || !own.has(l.slice(3).trim().replace(/\\/g, '/').replace(/^"|"$/g, '')));
}

/**
 * What is installed here that is not what the lock file says, by package: the
 * release installs from the lock, so a gate run against anything else checked
 * another program. An optional package for another platform may be absent.
 */
export function lockDrift(lock, versionAt) {
  const out = [];
  for (const [at, entry] of Object.entries(lock.packages || {})) {
    if (!at.includes('node_modules/') || entry.link) continue;
    const have = versionAt(at);
    if (have == null) { if (!entry.optional) out.push(`${at} is missing`); } else if (have !== entry.version) out.push(`${at} is ${have}, and the lock file says ${entry.version}`);
  }
  return out;
}

/** What this run counted, one by one, for the next run to be held to. */
export function floorsNow(counts, { apps = {}, machine = null } = {}) {
  const all = engineCounts(counts);
  return {
    ran: all ? all.tests - all.skipped : null,
    apps: Object.fromEntries(Object.entries(apps).map(([a, v]) => [a, v.of])),
    machine: machine?.of ?? null,
    // A --serial run has no machine lane of its own: those checks ran inside each app's count.
    ...(machine?.serial ? { serial: true } : {}),
    ...Object.fromEntries(['fuzz:open', 'fuzz:pdf', 'fuzz:grammar', 'fuzz:smime', 'verify:packs', 'office:reopen', 'corpus'].map((k) => [k, counts[k]?.of ?? null])),
  };
}

/** Where this full run did fewer than the last on the record, part by part: an aggregate hides one app losing what another gained. */
export function belowRecorded(now, was) {
  if (!was) return [];
  const fell = [];
  const held = (what, a, b) => { if (b != null && (a == null || a < b)) fell.push(`${what}: ${a ?? 'none'}, down from ${b}`); };
  held('engine tests that ran rather than skipped', now.ran, was.ran);
  if (now.serial) {
    // Its machine checks are inside the apps' counts, so the whole is held to the whole.
    const windows = (f) => Object.values(f.apps || {}).reduce((t, v) => t + (v || 0), 0) + (f.machine || 0);
    held('window checks, the machine lane\'s among them', windows(now), windows(was));
  } else held('machine lane', now.machine, was.machine);
  for (const [a, of] of Object.entries(was.apps || {})) held(`${a} window checks`, now.apps?.[a], of);
  for (const k of ['fuzz:open', 'fuzz:pdf', 'fuzz:grammar', 'fuzz:smime', 'verify:packs', 'office:reopen', 'corpus']) held(k, now[k], was[k]);
  return fell;
}

/** a, b and c */
const listed = (xs) => (xs.length > 1 ? `${xs.slice(0, -1).join(', ')} and ${xs.at(-1)}` : xs[0]);

/**
 * A patch release (2.17.1, not 2.17.0) may go on the record with the window
 * checks of only the apps it changed (owner, D-2026-10-23-01: "has no way to
 * just test one single app when the changes is in single app only").
 */
export function isPatch(version) {
  return /^\d+\.\d+\.([1-9]\d*)$/.test(String(version));
}

/** The engine tests in all: the suite, and the budgets run alone after it. */
export function engineCounts(counts) {
  const [t, p] = [counts.test, counts['test:perf']];
  if (!t || !p) return t || p || null;
  return { tests: t.tests + p.tests, passed: t.passed + p.passed, failed: t.failed + p.failed, skipped: t.skipped + p.skipped };
}

/** A paragraph wrapped at about 75 characters, as the notes are. */
export function wrap(text, width = 75) {
  const lines = [];
  let line = '';
  for (const word of String(text).split(/\s+/).filter(Boolean)) {
    if (line && line.length + 1 + word.length > width) { lines.push(line); line = word; } else line = line ? `${line} ${word}` : word;
  }
  if (line) lines.push(line);
  return lines.join('\n');
}

/** Today as this computer's calendar has it, not the UTC one: a gate run after midnight is that day's. */
export function localDate(d = new Date()) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/** The note with its Gate section put in, or put in again. */
export function withGateSection(note, section) {
  const body = String(note || '').replace(/\r\n/g, '\n');
  const at = body.search(/^## Gate\n/m);
  if (at < 0) return `${body.replace(/\n*$/, '\n\n')}${section}`;
  const next = body.slice(at + 1).search(/^## /m);
  return body.slice(0, at) + section + (next < 0 ? '' : `\n${body.slice(at + 1 + next)}`);
}

/** The README's counts brought to what the gate counted. */
export function withReadmeCounts(readme, counts) {
  let out = String(readme || '');
  const t = counts.test;
  const all = engineCounts(counts);
  if (all) out = out.replace(/(what runs before a release: )[\d,]+( engine tests)/, (m, a, b) => a + n(all.tests) + b);
  if (t) {
    out = out.replace(/(npm test\s+# )[\d,]+( tests, no network needed)/, (m, a, b) => a + n(t.tests) + b);
    out = out.replace(/(the engine suite: )[\d,]+( checks, no windows)/, (m, a, b) => a + n(t.tests) + b);
  }
  if (counts['verify:edit']) out = out.replace(/(engine tests, )[\d,]+( editing checks)/, (m, a, b) => a + n(counts['verify:edit'].of) + b);
  if (counts['verify:apps']) out = out.replace(/(editing checks,\s+)[\d,]+( application checks)/, (m, a, b) => a + n(counts['verify:apps'].of) + b);
  if (counts.smoke) out = out.replace(/(and )[\d,]+( window captures)/, (m, a, b) => a + n(counts.smoke.of) + b);
  return out;
}

/**
 * Whether the screen is locked: on Windows the lock screen is its own
 * process, LogonUI. With it up, the checks of the computer's own clipboard,
 * real input and the screen's capture fail after an hour of waiting (three
 * gates for 2.17.1), so the machine lane waits for the screen instead.
 * `list` answers what tasklist printed, for tests.
 */
export function screenLocked({ platform = process.platform, list = () => execFileSync('tasklist', ['/FI', 'IMAGENAME eq LogonUI.exe', '/NH'], { encoding: 'utf8' }) } = {}) {
  if (platform !== 'win32') return false;
  try { return /^\s*LogonUI\.exe\s/im.test(list()); } catch { return false; }
}

/** Wait up to `limitMs` for the screen to be unlocked, saying so once; whether it was. */
async function unlockedWithin(limitMs, { tickMs = 30_000 } = {}) {
  if (!screenLocked()) return true;
  console.log(`The screen is locked. The machine lane needs it unlocked: waiting up to ${Math.round(limitMs / 60_000)} minutes.`);
  const started = Date.now();
  while (Date.now() - started < limitMs) {
    await new Promise((r) => setTimeout(r, tickMs));
    if (!screenLocked()) { console.log('The screen is unlocked: the machine lane starts.'); return true; }
  }
  return false;
}

const argv = process.argv.slice(2);
const valueOf = (flag) => { const i = argv.indexOf(flag); return i >= 0 && argv[i + 1] && !argv[i + 1].startsWith('-') ? argv[i + 1] : null; };
const appsArg = valueOf('--apps');
const jobsArg = valueOf('--jobs');
const only = argv.filter((a, i) => !a.startsWith('-') && argv[i - 1] !== '--apps' && argv[i - 1] !== '--jobs');
const quiet = argv.includes('--quiet');
const record = argv.includes('--record');
const resume = argv.includes('--resume');
const changed = argv.includes('--changed');
const serial = argv.includes('--serial');
const fewer = argv.includes('--fewer-checks');

/** Run a command; its output shown as it comes (unless quiet, or sent to `log`), and the end of it kept. */
function run(command, args, env = {}, { log = null, cwd = ROOT } = {}) {
  return new Promise((resolve) => {
    const child = spawn(command, args, { cwd, stdio: ['ignore', 'pipe', 'pipe'], shell: process.platform === 'win32', env: { ...process.env, ...env } });
    const file = log ? fs.createWriteStream(log) : null;
    let tail = '';
    const keep = (to) => (chunk) => {
      if (file) file.write(chunk);
      else if (!quiet) to.write(chunk);
      tail = (tail + chunk).slice(-200000);
    };
    child.stdout.on('data', keep(process.stdout));
    child.stderr.on('data', keep(process.stderr));
    child.on('exit', (code) => { file?.end(); resolve({ code: code ?? 1, tail }); });
  });
}

/** The share of the processors busy over `ms`, from the operating system's own counters. */
async function busyShare(ms) {
  const sample = () => os.cpus().reduce((t, c) => {
    const all = Object.values(c.times).reduce((a, b) => a + b, 0);
    return { idle: t.idle + c.times.idle, all: t.all + all };
  }, { idle: 0, all: 0 });
  const a = sample();
  await new Promise((r) => setTimeout(r, ms));
  const b = sample();
  const all = b.all - a.all;
  return all > 0 ? 1 - (b.idle - a.idle) / all : 0;
}

/**
 * The budgets are run alone, and alone means the machine has settled: files
 * written just before are still being read by the virus scanner for a while,
 * which held the main thread past its budget in two of three gates for 2.2.0.
 * So the gate waits until the processors have been under a quarter busy for
 * three seconds running. It used to go on after two minutes whatever the load
 * (a run again began at 57% busy and failed its budgets): now it waits up to
 * fifteen, and a machine that never settles is said to be one, so the budgets
 * are not measured on it (review of 11 Nov). Answers { calm, waited, busy }.
 */
export async function settle(limitMs = 15 * 60_000, { sample = busyShare, log = (line) => { if (!quiet) console.log(line); } } = {}) {
  const started = Date.now();
  let calm = 0;
  let share = 1;
  let said = false;
  while (Date.now() - started < limitMs) {
    share = await sample(1000);
    calm = share < 0.25 ? calm + 1 : 0;
    if (calm >= 3) break;
    if (!said && Date.now() - started > 60_000) { said = true; log(`The machine is ${Math.round(share * 100)}% busy: the budgets wait for it to settle, up to ${Math.round(limitMs / 60_000)} minutes.`); }
  }
  const waited = Math.round((Date.now() - started) / 1000);
  const busy = Math.round(share * 100);
  if (calm >= 3) log(`The machine settled for the budgets: waited ${waited} s, the processors ${busy}% busy.`);
  else log(`The machine did not settle: ${busy}% busy after ${waited} s.`);
  return { calm: calm >= 3, waited, busy };
}

/**
 * Whether this computer runs on its battery: the budgets are not measured
 * then (a laptop holds its processors and graphics back). Windows answers
 * through Win32_Battery: BatteryStatus 1 is discharging; no battery is mains.
 */
export function onBattery({ platform = process.platform, query = () => execFileSync('powershell', ['-NoProfile', '-Command', '(Get-CimInstance Win32_Battery | Select-Object -First 1).BatteryStatus'], { encoding: 'utf8', timeout: 30000 }) } = {}) {
  if (platform !== 'win32') return false;
  // 1 discharging, 4 low, 5 critical: each is the battery running the machine (review of 11 Nov).
  try { return ['1', '4', '5'].includes(String(query()).trim()); } catch { return false; }
}

/* ── the lanes ─────────────────────────────────────────────────────── */

const TIMING = path.join(STATE_DIR, 'timing.json');
const LOGS = path.join(STATE_DIR, 'logs');
const PARTS_DIR = path.join(STATE_DIR, 'parts');
/** A chunk's length the first time, before a run has measured it, in seconds. */
const FIRST_GUESS = { shell: 900, ribbon: 550, look: 400, theme: 400, word: 420, sheets: 400, slides: 350, access: 300, mail: 300, video: 300, languages: 200, pdf: 150, image: 150, recorder: 90, broadcast: 90, pictures: 60, talk: 40, base: 30, calendar: 40, contacts: 20, locker: 300 };

/** The chunks in the order they start: the longest first, so the last to finish is a short one. */
export function longestFirst(chunks, timing = {}) {
  const len = (c) => timing[c] ?? FIRST_GUESS[c] ?? 120;
  return [...chunks].sort((a, b) => len(b) - len(a) || chunks.indexOf(a) - chunks.indexOf(b));
}

/**
 * How many check processes run side by side: one to each two and a half
 * processors (a check process mostly waits on its windows) and to each 1.5 GB
 * of free memory, at least one and at most four.
 */
export function jobsFor({ cpus = os.cpus().length, freeGb = os.freemem() / 2 ** 30 } = {}) {
  return Math.max(1, Math.min(4, Math.floor(cpus / 2.5), Math.floor(freeGb / 1.5)));
}

/** The lanes the processors allow: one to each two and a half, at least one and at most four. */
export function laneCap(cpus = os.cpus().length) {
  return Math.max(1, Math.min(4, Math.floor(cpus / 2.5)));
}

/**
 * Run `tasks` in their order, at most `cap` at once, a task started beside
 * others only while `freeGb()` leaves `perLane` for it after what the lanes
 * started in the last `settleMs` will still take. Looked at again every
 * `tickMs`, so lanes are added as memory comes free. `onShort(free, running)`
 * hears when a lane waits for memory, once each time it starts to.
 */
export async function poolByMemory(tasks, { cap, freeGb = () => os.freemem() / 2 ** 30, perLane = 1.5, settleMs = 30_000, tickMs = 10_000, now = Date.now, onShort = () => {} }) {
  const results = new Array(tasks.length);
  if (!tasks.length) return results;
  let next = 0;
  let running = 0;
  let waiting = false;
  const started = [];
  return new Promise((resolve, reject) => {
    let timer = null;
    const finish = () => { clearInterval(timer); resolve(results); };
    const fill = () => {
      while (next < tasks.length && running < cap) {
        const recent = started.filter((t) => now() - t < settleMs).length;
        const free = freeGb();
        if (running > 0 && free - recent * perLane < perLane) {
          if (!waiting) onShort(free, running);
          waiting = true;
          return;
        }
        waiting = false;
        const i = next++;
        running++;
        started.push(now());
        Promise.resolve().then(() => tasks[i]()).then((r) => { results[i] = r; }, (err) => { clearInterval(timer); reject(err); }).finally(() => {
          running--;
          if (next >= tasks.length && running === 0) finish();
          else fill();
        });
      }
    };
    timer = setInterval(fill, tickMs);
    if (timer.unref) timer.unref();
    fill();
  });
}

/**
 * The parts to run again at the end, from the chunks' and the machine
 * lane's counts: { parallel: { apps, ids, names, whole }, machine: { ... } }.
 * `ids` are the failed parts' ids (verify-select.js); a failure outside any
 * part ('*') asks for its app whole, as `<app>#*`, and only that app.
 */
export function partsToRerun(apps, machine) {
  const pick = (entries) => {
    const out = { apps: [], ids: [], names: [], whole: [] };
    for (const [app, v] of entries) {
      if (!v || v.ok || !v.failed?.length) continue;
      out.apps.push(app);
      out.names.push(...v.failed);
      const ids = v.failedParts || ['*'];
      if (ids.includes('*')) {
        out.whole.push(app);
        out.ids.push(`${app}#*`);
      } else out.ids.push(...ids);
    }
    return out;
  };
  // A machine lane that wrote no parts file (its process ended, or it was not measured) has no part to name: it is not run again, and says why.
  if (machine && !machine.ok && !machine.byApp) return { parallel: pick(Object.entries(apps)), machine: { apps: [], ids: [], names: [], whole: [], unnamed: machine.notMeasured ? 'not measured' : 'ended without its parts file' } };
  const machineApps = machine && !machine.ok ? Object.entries(machine.byApp) : [];
  return { parallel: pick(Object.entries(apps)), machine: pick(machineApps) };
}

/** Each part's count by its id, from a parts file's apps: { [id]: { app, of, passed, failed, threw } }; `<app>#*` for each app whole. */
export function partsById(apps) {
  const out = {};
  for (const [app, v] of Object.entries(apps || {})) {
    if (!v) continue;
    // '*' is a check outside any part; the app whole stands for it.
    for (const [id, p] of Object.entries(v.parts || {})) if (id !== '*') out[id] = { app, ...p };
    out[`${app}#*`] = { app, of: v.of || 0, passed: v.passed || 0, failed: v.failed || [], realFailed: v.realFailed || [], passedNames: v.passedNames || [], threw: Object.values(v.parts || {}).some((p) => p.threw) };
  }
  return out;
}

/**
 * Whether the run again cleared the parts that failed, part by part (review
 * of 11 Nov: a name matched to a name left about 25 catch-alls that no clean
 * run again could clear). Each failed id must have run again with checks, none
 * failing, without throwing, and with at least as many checks as passed in it
 * the first time; nothing else in the run again may fail. Answers { ok,
 * missing, failing, short, threw }: `missing` (asked for, and none of its
 * checks ran) is the gate's own fault, not Office's.
 */
export function rerunCleared(ids, first, again) {
  const before = partsById(first);
  const now = partsById(again);
  const out = { ok: false, missing: [], failing: [], short: [], threw: [], unseen: [], crashed: false };
  // No parts file at all: the run again's own process ended, which is not the gate missing a part.
  if (!again) { out.crashed = true; return out; }
  if (!ids?.length) return out;
  for (const id of ids) {
    const p = now[id];
    if (!p || !p.of) { out.missing.push(id); continue; }
    if (p.threw) out.threw.push(id);
    const was = before[id]?.passed || 0;
    if (p.of < was) out.short.push(`${id} (${p.of} checks, ${was} passed before)`);
    // A count alone could clear a check that took another path and never ran
    // again (review of 11 Nov): each of the product's own failed checks must
    // pass again by name. The run's catch-alls (a throw, a console line) are
    // written only when something went wrong, so a clean run has none to show.
    const passedAgain = new Set(p.passedNames || []);
    const unseen = (before[id]?.realFailed || []).filter((n) => !passedAgain.has(n));
    if (unseen.length) out.unseen.push(`${id}: ${unseen.slice(0, 3).join('; ')}${unseen.length > 3 ? ` and ${unseen.length - 3} more` : ''}`);
  }
  out.failing = Object.values(again).flatMap((v) => v?.failed || []);
  out.ok = !out.missing.length && !out.failing.length && !out.short.length && !out.threw.length && !out.unseen.length;
  return out;
}

/** What the run again found, in a line. */
export function againSaid(r) {
  if (r.ok) return 'each passed alone';
  if (r.crashed) return 'the run again ended without writing its parts file: its window process crashed or was stopped';
  const said = [];
  if (r.missing.length) said.push(`the gate's run again found none of these parts: ${r.missing.join(', ')} (a fault of the gate, not of Office)`);
  if (r.failing.length) said.push(`failed again: ${r.failing.slice(0, 5).join('; ')}${r.failing.length > 5 ? ` and ${r.failing.length - 5} more` : ''}`);
  if (r.threw.length) said.push(`threw again: ${r.threw.join(', ')}`);
  if (r.short.length) said.push(`ran fewer checks than passed the first time: ${r.short.join(', ')}`);
  if (r.unseen?.length) said.push(`did not pass again by name: ${r.unseen.join(', ')}`);
  return said.join('; ');
}

/**
 * The parts marked for the machine lane that no lane ran: the chunks skip
 * them for the machine lane, and a part inside a section (`if (runs(app))`)
 * is never reached there, since the machine lane enters no section. `skipped`
 * are the ids the chunks left to the machine lane; `machineSeen` the machine
 * lane's own `seen` list.
 */
export function machinePartsUnrun(skipped, machineSeen) {
  const reached = new Set((machineSeen || []).filter((s) => s.because !== 'lane').map((s) => s.id));
  return [...new Set(skipped)].filter((id) => !reached.has(id));
}

/** The machine-lane parts a chunk's run passed by, from its `seen` list. */
export const leftToMachine = (seen) => (seen || []).filter((s) => s.machine && s.because === 'lane').map((s) => s.id);

/** Run `tasks` (functions answering promises), at most `limit` at once, in their order. */
export async function pool(tasks, limit) {
  const results = new Array(tasks.length);
  let next = 0;
  const worker = async () => {
    for (;;) {
      const i = next++;
      if (i >= tasks.length) return;
      results[i] = await tasks[i]();
    }
  };
  await Promise.all(Array.from({ length: Math.max(1, Math.min(limit, tasks.length)) }, worker));
  return results;
}

const readJson = (file) => { try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return null; } };
const secs = (ms) => (ms / 1000).toFixed(1);
const say = (line) => { if (!quiet) console.log(line); };

async function main() {
  const results = [];
  const counts = {};
  let failed = false;
  // Faults of the gate itself, not of Office: a run again that found none of
  // what it was asked for, a machine part no lane ran, budgets on a machine
  // that never settled. Said apart at the end, and the gate stays closed.
  const harness = [];
  const gateFault = (line) => { harness.push(line); failed = true; say(`  the gate's own fault: ${line}`); };
  /** Whether the budgets can be measured now: on mains power, the processors settled. */
  const budgetsReady = async (what) => {
    if (onBattery()) { gateFault(`${what} not measured: this computer is on its battery`); return false; }
    const calm = await settle();
    if (!calm.calm) { gateFault(`${what} not measured: the machine stayed ${calm.busy}% busy for ${Math.round(calm.waited / 60)} minutes`); return false; }
    return true;
  };

  for (const name of SHRINKING_ENV) {
    if (process.env[name] !== undefined) {
      console.log(`${name} was set in this shell and is ignored: the gate runs everything.`);
      delete process.env[name];
    }
  }

  // A pass named that the gate does not have would run nothing and still open the gate.
  const known = ['test', 'test:perf', 'verify:edit', 'verify:apps', 'smoke', 'verify:packs', 'office:reopen', ...FUZZERS.map(([name]) => name), 'corpus'];
  const unknown = only.filter((name) => !known.includes(name));
  if (unknown.length) {
    console.log(`The gate has no pass called ${unknown.join(', ')}. Its passes: ${known.join(', ')}.`);
    process.exit(2);
  }

  const drifted = lockDrift(readJson(path.join(ROOT, 'package-lock.json')) || {}, (at) => readJson(path.join(ROOT, at, 'package.json'))?.version ?? null);
  if (drifted.length) {
    console.log(`The gate is closed: node_modules is not what the lock file says (${drifted.length}): ${drifted.slice(0, 5).join('; ')}. Run npm ci.`);
    process.exit(1);
  }

  // Which apps' window checks this run is for: those named, those changed, or all.
  let wanted = VERIFY_APPS.slice();
  if (appsArg) wanted = appsArg.split(',').map((a) => a.trim()).filter((a) => VERIFY_APPS.includes(a));
  else if (changed) {
    // The tag before this commit: one already on HEAD would make every change look like none.
    const since = git('describe', '--tags', '--abbrev=0', 'HEAD^');
    const files = [...git('diff', '--name-only', `${since}..HEAD`).split('\n'), ...git('diff', '--name-only').split('\n')].filter(Boolean);
    wanted = wantedForChanged(since, files);
    say(since ? `Changed since ${since}: ${files.length} files; the window checks of ${wanted.join(', ') || 'no app'}.` : 'No tag to compare with: the window checks of every app.');
  }
  const partial = wanted.length < VERIFY_APPS.length || only.length > 0;
  const want = (name) => !only.length || only.includes(name);

  // The record of runs on this commit: started afresh unless resuming.
  const head = git('rev-parse', 'HEAD');
  // What the tree holds, not only which files differ: an old pass must not stand for a file edited since.
  const tree = treeKey();
  for (const d of [STATE_DIR, LOGS, PARTS_DIR]) fs.mkdirSync(d, { recursive: true });
  // One gate at a time: two would write over each other's parts and counts.
  const running = path.join(STATE_DIR, 'running');
  try {
    const other = Number(fs.readFileSync(running, 'utf8'));
    if (other && other !== process.pid) {
      process.kill(other, 0);
      console.log(`Another gate is running (process ${other}). Wait for it to end.`);
      process.exit(1);
    }
  } catch { /* none, or it has ended */ }
  fs.writeFileSync(running, String(process.pid));
  process.on('exit', () => { try { fs.rmSync(running, { force: true }); } catch { /* gone */ } });
  const state = (resume && readState(head, tree)) || { head, tree, passes: {}, apps: {}, machine: null };
  state.apps ||= {};
  const save = () => fs.writeFileSync(STATE, JSON.stringify(state, null, 2));
  if (resume && !state.passes.test && !Object.keys(state.apps).length) say('Nothing to resume on this commit: a run from the start.');
  const timing = readJson(TIMING) || {};
  const jobs = Math.max(1, Number(jobsArg) || laneCap());
  if (want('verify:apps') && wanted.length && screenLocked()) {
    say('The screen is locked: everything that does not need it runs now, and the machine lane (the computer\'s own clipboard, real input, the screen\'s capture) waits for it to be unlocked.');
  }

  /** A pass of its own (test, test:perf, verify:edit, smoke), its counts kept on the state. */
  const pass = async (name, what, command, args, { log = null, env = {} } = {}) => {
    if (resume && state.passes[name]?.ok) {
      say(`${name}: passed in the run before on this commit.`);
      results.push({ name, what, ok: true, seconds: '0.0', resumed: true });
      if (state.passes[name].counts) counts[name] = state.passes[name].counts;
      return true;
    }
    const started = Date.now();
    if (!log) say(`\n\u001b[1m── ${name}: ${what} ─────────────────────────────\u001b[0m\n`);
    else say(`  [${name}] started`);
    const out = await run(command, args, env, { log });
    const counted = countsOf(name, out.tail);
    // A pass that ended without a count, or counted nothing, tested nothing, whatever it exited with.
    const code = out.code === 0 && (!counted || counted.of === 0) ? 1 : out.code;
    if (code !== out.code) say(`${name} ended without saying what it counted: it did not finish.`);
    results.push({ name, what, ok: code === 0, seconds: secs(Date.now() - started), log });
    if (counted) counts[name] = counted;
    state.passes[name] = { ok: code === 0, counts: counted || null };
    save();
    if (log) say(`  [${name}] ${code === 0 ? 'ok  ' : 'FAIL'} in ${secs(Date.now() - started)} s${counted?.of ? `, ${counted.passed}/${counted.of}` : counted?.tests ? `, ${counted.passed}/${counted.tests}` : ''}${code === 0 ? '' : ` (${path.relative(ROOT, log)})`}`);
    if (code !== 0) failed = true;
    return code === 0;
  };

  /** One run of window checks: a chunk in the parallel lane, or the machine lane. Its counts from its parts file. */
  const windowRun = async (label, env, { mayBeEmpty = false } = {}) => {
    const partsFile = path.join(PARTS_DIR, `${label}.json`);
    try { fs.rmSync(partsFile, { force: true }); } catch { /* none */ }
    const log = path.join(LOGS, `${label}.log`);
    const started = Date.now();
    say(`  [${label}] started`);
    const out = await run('node', [path.join('apps', 'desktop', 'build', 'verify-apps-run.js')], { ...env, RUTBA_VERIFY_PARTS: partsFile }, { log });
    const seconds = (Date.now() - started) / 1000;
    const file = readJson(partsFile);
    const parts = file?.apps || null;
    const sum = parts ? Object.values(parts).reduce((t, p) => ({ passed: t.passed + p.passed, of: t.of + p.of, failed: [...t.failed, ...p.failed] }), { passed: 0, of: 0, failed: [] }) : null;
    // A chunk that ran no check tested nothing: it is not a pass.
    const ok = out.code === 0 && Boolean(sum) && sum.failed.length === 0 && (mayBeEmpty || sum.of > 0);
    say(`  [${label}] ${ok ? 'ok  ' : 'FAIL'} ${sum ? `${sum.passed}/${sum.of}` : 'ended without its count'} in ${seconds.toFixed(0)} s${ok ? '' : ` (${path.relative(ROOT, log)})`}`);
    if (sum && !ok) for (const f of sum.failed.slice(0, 8)) say(`         FAIL ${f}`);
    return { ok, seconds, sum: sum || { passed: 0, of: 0, failed: ['the run ended before it counted'] }, parts, seen: file?.seen || [] };
  };

  // 1. The windows' bundle, once: every check process below loads the same one.
  if (want('verify:edit') || want('verify:apps') || want('smoke')) {
    say('\n\u001b[1m── build: the windows, bundled once ─────────────────\u001b[0m');
    const b = await run('npm', ['run', 'build'], {}, { log: path.join(LOGS, 'build.log') });
    if (b.code !== 0) {
      console.log(b.tail.slice(-4000));
      console.log('\n\u001b[31mThe gate is closed.\u001b[0m The windows did not build.');
      process.exit(1);
    }
  }

  // 2. The budgets, alone, while the machine is otherwise still.
  if (want('test:perf')) {
    if (resume && state.passes['test:perf']?.ok) await pass('test:perf', 'the performance budgets, run alone', 'npm', ['run', 'test:perf']);
    else if (await budgetsReady('test:perf')) await pass('test:perf', 'the performance budgets, run alone', 'npm', ['run', 'test:perf']);
    else { state.passes['test:perf'] = { ok: false, counts: null }; save(); results.push({ name: 'test:perf', what: 'the performance budgets, run alone', ok: false, seconds: '0.0' }); }
  }

  const lanes = serial ? null : { parallel: [], machine: null };
  if (serial) {
    // The run as it was: one process, every check in turn.
    if (want('test')) await pass('test', 'the engines', 'npm', ['test']);
    if (want('verify:edit') && !failed) await pass('verify:edit', 'typing reaches the document', 'node', [path.join('apps', 'desktop', 'build', 'verify-edit-run.js')]);
    if (want('verify:apps') && !failed) {
      const left = wanted.filter((a) => !(resume && state.apps[a]?.ok));
      const r = await windowRun('serial', left.length < VERIFY_APPS.length ? { RUTBA_VERIFY_APPS: left.join(',') } : {});
      for (const [a, p] of Object.entries(r.parts || {})) state.apps[a] = { ok: p.failed.length === 0, passed: p.passed, of: p.of, failed: p.failed, seconds: p.seconds };
      state.machine = { ok: true, passed: 0, of: 0, failed: [], serial: true, scope: [] };
      results.push({ name: 'verify:apps', what: 'each app opens, changes and saves', ok: r.ok, seconds: r.seconds.toFixed(1) });
      if (!r.ok) failed = true;
      save();
    }
  } else {
    // 3. Side by side: the engine tests, the editing checks and a check
    // process a chunk, each with a clipboard and shortcuts of its own.
    const chunks = longestFirst(wanted.filter((a) => !(resume && state.apps[a]?.ok)), timing);
    const freeNow = os.freemem() / 2 ** 30;
    say(`\n\u001b[1m── side by side: up to ${jobs} at once ─────────────────────\u001b[0m`);
    say(jobsArg ? `  ${jobs} at once, as asked` : `  ${os.cpus().length} processors allow ${jobs}; ${freeNow.toFixed(1)} GB of memory free now, and a lane is added only while 1.5 GB is free for it`);
    say(`  ${[want('test') ? 'the engine tests' : null, want('verify:edit') ? 'the editing checks' : null, want('verify:apps') ? `the window checks of ${chunks.length ? chunks.join(', ') : 'no chunk left'}` : null].filter(Boolean).join(', ')}; the logs are in ${path.relative(ROOT, LOGS)}`);
    const started = Date.now();
    const tasks = [];
    if (want('test')) tasks.push(() => pass('test', 'the engines', 'npm', ['test'], { log: path.join(LOGS, 'test.log') }));
    if (want('verify:apps')) {
      for (const c of chunks) {
        tasks.push(async () => {
          const r = await windowRun(c, { RUTBA_VERIFY_LANE: 'parallel', RUTBA_VERIFY_APPS: c });
          const p = r.parts?.[c] || { passed: r.sum.passed, of: r.sum.of, failed: r.sum.failed, seconds: r.seconds };
          state.apps[c] = { ok: r.ok, passed: p.passed, of: p.of, failed: p.failed, failedParts: p.failedParts, parts: p.parts || {}, toMachine: leftToMachine(r.seen), seconds: Math.round(r.seconds) };
          if (r.ok) timing[c] = Math.round(r.seconds);
          save();
          return r.ok;
        });
      }
    }
    if (want('verify:edit')) tasks.push(() => pass('verify:edit', 'typing reaches the document', 'node', [path.join('apps', 'desktop', 'build', 'verify-edit-run.js')], { log: path.join(LOGS, 'verify-edit.log') }));
    if (jobsArg) await pool(tasks, jobs);
    else await poolByMemory(tasks, { cap: jobs, onShort: (free, running) => say(`  ${running} at once for now: ${free.toFixed(1)} GB free, and another lane needs 1.5 GB (close programs to run more)`) });
    say(`  side by side: ${secs(Date.now() - started)} s`);

    // 4. The machine lane, alone: the computer's own clipboard and shortcuts,
    // timing budgets, drags and native copies, with nothing else running.
    if (want('verify:apps')) {
      if (resume && machineCovers(state.machine, wanted)) say('  [machine] passed in the run before on this commit, for these apps');
      else if (!(await unlockedWithin(60 * 60_000))) {
        // Not run, so not failed: what it checks waits, and --resume runs it once the screen is unlocked.
        say('  [machine] not run: the screen stayed locked. Unlock it and run the gate again with --resume.');
        failed = true;
      } else {
        say(`\n\u001b[1m── machine: the computer's own clipboard and keys, alone ──\u001b[0m`);
        // Its budgets are measured on a settled machine, as test:perf's are: it used to start with no settle at all.
        if (await budgetsReady('the machine lane')) {
          const r = await windowRun('machine', { RUTBA_VERIFY_LANE: 'machine', ...(wanted.length < VERIFY_APPS.length ? { RUTBA_VERIFY_APPS: wanted.join(',') } : {}) }, { mayBeEmpty: wanted.length < VERIFY_APPS.length });
          state.machine = { ok: r.ok, passed: r.sum.passed, of: r.sum.of, failed: r.sum.failed, seconds: Math.round(r.seconds), byApp: r.parts || null, scope: wanted.slice(), seen: r.seen };
          // A part marked for the machine lane that the chunks passed by and the machine lane never reached ran in no lane at all.
          const unrun = machinePartsUnrun(wanted.flatMap((a) => state.apps[a]?.toMachine || []), r.seen);
          if (unrun.length) gateFault(`parts marked for the machine lane that no lane ran (a part inside a section is never reached there): ${unrun.join(', ')}`);
        } else state.machine = { ok: false, passed: 0, of: 0, failed: [], scope: [], notMeasured: true };
        save();
      }

      // 5. Again, alone: the parts whose checks failed, in a process of their
      // own after the rest, the machine perhaps quieter by now. The gate
      // passes only if each check that failed passes here.
      const again = partsToRerun(Object.fromEntries(wanted.map((a) => [a, state.apps[a]])), state.machine);
      if (again.machine.unnamed) say(`
  the machine lane ${again.machine.unnamed}, so it is not run again and the gate stays closed.`);
      for (const [laneName, set] of [['parallel', again.parallel], ['machine', again.machine]]) {
        if (!set.names.length) continue;
        // Counted by part, not by check: one export that stalled failed six checks in two parts in gate 9.
        if (set.ids.length > RERUN_LIMIT) {
          say(`\n  ${set.ids.length} parts failed ${laneName === 'machine' ? 'in the machine lane' : 'beside the others'}: more than ${RERUN_LIMIT}, so they are not run again and the gate stays closed.`);
          continue;
        }
        say(`\n\u001b[1m── again, alone: ${set.ids.length} part${set.ids.length === 1 ? '' : 's'} that failed ${laneName === 'machine' ? 'in the machine lane' : 'beside the others'} ──\u001b[0m`);
        for (const id of set.ids) say(`         ${id}`);
        for (const n of set.names) say(`           ${n}`);
        // Only the parts that failed, and an app whole (`<app>#*`) only where a check outside its parts failed.
        const env = { RUTBA_VERIFY_LANE: laneName, RUTBA_VERIFY_APPS: set.apps.join(','), RUTBA_VERIFY_PART_IDS: set.ids.join(',') };
        if (laneName === 'machine') {
          if (!(await budgetsReady('the machine lane\'s run again'))) continue;
        } else await settle();
        const r = await windowRun(laneName === 'machine' ? 'again-machine' : 'again', env, { mayBeEmpty: true });
        const first = laneName === 'machine' ? state.machine.byApp : Object.fromEntries(set.apps.map((a) => [a, state.apps[a]]));
        const verdict = rerunCleared(set.ids, first, r.parts);
        const cleared = r.ok && verdict.ok;
        say(`  again: ${cleared ? 'each passed alone' : againSaid(verdict) || 'still failing'}`);
        if (verdict.missing.length) gateFault(`the run again found none of these parts: ${verdict.missing.join(', ')}`);
        if (!cleared) continue;
        if (laneName === 'machine') {
          state.machine = { ...state.machine, ok: true, reran: set.names, failed: [] };
        } else {
          for (const a of set.apps) {
            const v = state.apps[a];
            state.apps[a] = { ...v, ok: true, reran: v.failed, failed: [] };
          }
        }
        save();
      }
      const appsOk = wanted.every((a) => state.apps[a]?.ok) && state.machine?.ok;
      if (!appsOk) failed = true;
      const all = [...Object.values(state.apps), state.machine].filter(Boolean);
      results.push({ name: 'verify:apps', what: 'each app opens, changes and saves', ok: appsOk, seconds: secs(Date.now() - started) });
      say(VERIFY_APPS.filter((a) => state.apps[a]).map((a) => `  ${state.apps[a].ok ? 'ok  ' : 'FAIL'} ${a.padEnd(10)} ${state.apps[a].passed}/${state.apps[a].of}${state.apps[a].seconds ? `  ${state.apps[a].seconds} s` : ''}`).join('\n'));
      if (state.machine) say(`  ${state.machine.ok ? 'ok  ' : 'FAIL'} ${'machine'.padEnd(10)} ${state.machine.passed}/${state.machine.of}${state.machine.seconds ? `  ${state.machine.seconds} s` : ''}`);
      if (all.length) counts['verify:apps'] = { passed: all.reduce((n, v) => n + v.passed, 0), of: all.reduce((n, v) => n + v.of, 0) };
    }
    fs.writeFileSync(TIMING, JSON.stringify(timing, null, 2));
  }

  // 6. Every window painted, once the rest has passed.
  if (want('smoke') && !failed) await pass('smoke', 'every window paints', 'npm', ['run', 'smoke']);

  // The app packs, added and removed across two starts.
  if (want('verify:packs') && !failed) await pass('verify:packs', 'app packs add and remove', 'node', [path.join('apps', 'desktop', 'build', 'verify-packs-run.js')], { log: path.join(LOGS, 'verify-packs.log') });

  // What the suite saves, opened in Word, Excel and PowerPoint themselves: they must find what the original holds, and the edit.
  if (want('office:reopen') && !failed) await pass('office:reopen', 'Word, Excel and PowerPoint reopen our saves', 'powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join('tools', 'check-strict-office.ps1'), '-Dir', path.join(STATE_DIR, 'office-saves')], { log: path.join(LOGS, 'office-reopen.log') });

  // 7. Files nobody made on purpose: damaged ones, then real ones from outside the project.
  for (const [name, tool, what, rounds] of FUZZERS) {
    if (want(name) && !failed) await pass(name, what, 'node', [tool, String(rounds), String(parseInt(head.slice(0, 8), 16))], { log: path.join(LOGS, `${name.replace(':', '-')}.log`) });
  }
  if (want('corpus') && !failed) {
    const corpus = path.join(ROOT, '.corpus');
    if (fs.existsSync(corpus)) await pass('corpus', 'real files, opened and saved', 'node', ['tools/corpus.js'], { log: path.join(LOGS, 'corpus.log'), env: { RUTBA_CORPUS_RATES: path.join(ROOT, 'tools', 'corpus-rates.json'), RUTBA_CORPUS: corpus, RUTBA_CORPUS_OUT: path.join(STATE_DIR, 'corpus') } });
    else { console.log('corpus: there is no .corpus folder here (tools/downloads.js fetches it), so this gate cannot pass.'); failed = true; }
  }

  // A full run must have run at least what the last one did.
  if (!partial && !failed) {
    const fell = [
      ...belowFloor(counts, floorsFrom(fs.readFileSync(path.join(ROOT, 'README.md'), 'utf8'))),
      ...belowRecorded(floorsNow(counts, state), readJson(FLOORS)),
      ...(fs.existsSync(FLOORS) && !readJson(FLOORS) ? ['tools/gate-floors.json cannot be read, so there is nothing to compare this run with'] : []),
    ];
    state.fewer = fell;
    save();
    if (fell.length) {
      for (const f of fell) console.log(`Fewer checks than the last gate: ${f}`);
      if (fewer) console.log('Allowed by --fewer-checks, and said so in the release note.');
      else { console.log('Say why in the commit that removed them and run again with --fewer-checks.'); failed = true; }
    }
  }

  console.log('\n────────────────────────────────────────────────');
  for (const h of harness) console.log(`\u001b[31mGATE FAULT\u001b[0m ${h}`);
  state.harness = harness;
  save();
  for (const r of results) {
    console.log(`${r.ok ? '\u001b[32mok  \u001b[0m' : '\u001b[31mFAIL\u001b[0m'} ${r.name.padEnd(13)} ${r.what.padEnd(38)} ${r.seconds}s${r.resumed ? ' (before)' : ''}`);
  }
  console.log(
    failed
      ? `\n\u001b[31mThe gate is closed.\u001b[0m Fix what failed and run it again, or with --resume to check again only what failed.`
      : `\n\u001b[32mThe gate is open.\u001b[0m ${results.length} passes${partial ? `, for ${only.length ? only.join(', ') : wanted.join(', ')} only` : ''}.`
  );

  // On the record only when every pass and every chunk passed on this commit, in one run or resumed;
  // a patch release with --changed, every pass and the chunks it changed.
  const version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
  const scoped = changed && !only.length && !appsArg && isPatch(version) && wanted.length < VERIFY_APPS.length;
  const whole = (!partial || scoped) && ['test', 'test:perf', 'verify:edit', 'smoke', 'verify:packs', 'office:reopen', ...FUZZERS.map(([name]) => name), 'corpus'].every((name) => state.passes[name]?.ok)
    && (scoped ? wanted : VERIFY_APPS).every((a) => state.apps[a]?.ok) && (!wanted.length || machineCovers(state.machine, scoped ? wanted : VERIFY_APPS));
  if (record) {
    const dirty = dirtyForRecord(porcelain(), version);
    if (!failed && whole && (git('rev-parse', 'HEAD') !== head || dirty.length)) {
      console.log(`Nothing recorded: the commit moved during the run, or the tree has changes a commit would not hold (${dirty.slice(0, 5).join('; ') || 'HEAD changed'}).`);
    } else if (failed || !whole) {
      console.log(changed && !isPatch(version)
        ? `Nothing recorded: ${version} is not a patch release, and only a full run that passed is put on its record.`
        : 'Nothing recorded: only a run that passed, in full or for the apps a patch release changed, is put on the record.');
    } else {
      const date = localDate();
      const notePath = path.join(ROOT, 'docs', 'releases', `v${version}.md`);
      if (fs.existsSync(notePath)) {
        const reran = [...VERIFY_APPS.flatMap((a) => state.apps[a]?.reran || []), ...(state.machine?.reran || [])];
        fs.writeFileSync(notePath, withGateSection(fs.readFileSync(notePath, 'utf8'), gateSection(counts, { version, date, reran, apps: scoped ? wanted : null, commit: head, fewer: state.fewer || [] })));
        console.log(`Recorded in ${path.relative(ROOT, notePath)}.`);
      } else {
        console.log(`No note at ${path.relative(ROOT, notePath)} to record in; write the note first.`);
      }
      const readmePath = path.join(ROOT, 'README.md');
      fs.writeFileSync(readmePath, withReadmeCounts(fs.readFileSync(readmePath, 'utf8'), counts));
      console.log("The README's counts brought up to date.");
      if (!scoped) fs.writeFileSync(FLOORS, `${JSON.stringify(floorsNow(counts, state), null, 2)}\n`);
    }
  }

  process.exit(failed ? 1 : 0);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await main();
