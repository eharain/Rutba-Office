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
//   npm run gate -- --no-settle  the budgets and the machine lane measured
//                                without waiting for the machine to settle,
//                                every limit as it is, and said in the record
//                                with how busy it was (owner, 13 Nov 2026)
//   npm run gate -- verify:apps --parts mail#reading,video#verifyVideoMp4
//                                those parts and sections alone, each by its id
//   npm run gate -- verify:apps --failed
//                                the parts of every open entry in the failure list
//   npm run gate -- ledger       the failure list (.gate/ledger.json, and
//                                .gate/failures.md to read): each check that
//                                failed, its number, the line that judged it,
//                                its part and the wait that gave up before it
//   npm run gate -- ledger close 18.0142 --by <commit>
//                                close an entry once a run after the fix passed
//                                it (or --reason "why", in a sentence); --record
//                                is refused while any entry is open (owner, 11
//                                Nov 2026: find all the failures, fix them, then
//                                one full run)
//   npm run gate -- ledger move 00.0001 --to tests/locker-large.test.js
//                                an engine test moved to another file to fix
//                                it: its entry and number go with it, and only
//                                passes in that file prove it
//   npm run gate -- test --parts test#tests/calendar.test.js,fuzz:pdf#*
//                                engine test files and fuzz passes by id too
//   RUTBA_GATE_VM=<runner> RUTBA_GATE_VM_HOST=<name> npm run gate
//                                the engine tests (all but those that touch
//                                Windows itself, tools/engine-hosts.json) and
//                                the fuzz passes on a Linux machine, from the
//                                start and beside the rest (tools/gate-hosts.js);
//                                one that cannot be reached closes the gate
//   npm run gate -- --local      with the runner set, all of it here instead
//
// The order (owner, 2026-10-21: "splitting this long running gate into
// smaller chunks is very important so that lighter or non locking parts can
// be run in parallel"):
//
//   build         the windows bundled once, for every check process below.
//   test:perf     the performance budgets, alone, while the machine is still:
//                 a budget timed beside other work measures the machine.
//   test:removal  the folder-removal suites on Node and on Office's own
//                 Electron as Node, here, where a junction is real.
//   side by side  the engine tests, the editing checks and a window check
//                 process a chunk (an app, or a part of what every app shares:
//                 shell, ribbon, look, theme, access, languages), as many at once as the
//                 processors and memory allow, the longest chunk first, by
//                 the seconds each took before (a chunk never measured goes
//                 first); an app longer than a lane's share runs as pieces of
//                 its parts, each a process of its own (tools/gate-plan.js). Each
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
import { execFile, execFileSync } from 'node:child_process';
import { readLedger, writeLedger, runFindings, recordRun, openEntries, openParts, ledgerLines, ledgerMarkdown, closeEntry, moveEntry, reasonClosures, markRecorded } from './gate-ledger.js';
import { readRegistry, writeRegistry, assign as assignIds } from '../apps/desktop/main/verify-ids.js';
import { readSeconds, writeSeconds, noteSeconds, planChunks, mergePieces, sourceIds } from './gate-plan.js';
import { timeLost, timeLostLines } from './gate-time-lost.js';
import { FUZZERS, vmConfig, vmCommand, engineFiles, engineGlobs, readHosts, splitEngine, osReach, vmFault, snapshotOf, worktreeHash, readRun, testArgs, engineFindings, mergeFindings, engineParts, isEngineId, hostsSentence } from './gate-hosts.js';

export { FUZZERS };

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
/** The one file a run writes that its tree key leaves out. */
export const UNKEYED = 'tools/check-ids.json';
export function treeKey(cwd = ROOT) {
  const run = (args) => { try { return execFileSync('git', args, { cwd, maxBuffer: 2 ** 31 - 1, stdio: ['ignore', 'pipe', 'ignore'] }); } catch { return null; } };
  const hash = crypto.createHash('sha256');
  // The check numbers a run writes change no check: a run that numbered a new one is still resumed (review of 11 Nov).
  hash.update(String(porcelain(cwd)).split('\n').filter((l) => l.slice(3).trim().replace(/^"|"$/g, '') !== UNKEYED).join('\n'));
  hash.update(run(['diff', 'HEAD', '--binary', '--', '.', `:(exclude)${UNKEYED}`]) ?? 'no diff');
  const untracked = String(run(['ls-files', '--others', '--exclude-standard', '-z']) || '').split('\0').filter((f) => f && f !== UNKEYED).sort();
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
/** The failure list (tools/gate-ledger.js) and its page. */
const LEDGER = path.join(STATE_DIR, 'ledger.json');
const FAILURES_PAGE = path.join(STATE_DIR, 'failures.md');
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
  if (name === 'test' || name === 'test:perf' || name === 'test:removal') {
    // The spec reporter ends with "ℹ tests 2209", the TAP reporter (output
    // that is not a terminal, as a gate run into a log is) with "# tests 2209".
    const tests = sum(/^(?:ℹ|#) tests (\d+)$/gm);
    if (!tests) return null;
    const out = { tests, passed: sum(/^(?:ℹ|#) pass (\d+)$/gm), failed: sum(/^(?:ℹ|#) fail (\d+)$/gm), skipped: sum(/^(?:ℹ|#) skipped (\d+)$/gm) };
    // The folder-removal suites run once on each runtime, each run ending with its own totals: how many runs, for the record.
    if (name === 'test:removal') out.runs = all(/^(?:ℹ|#) tests (\d+)$/gm).length;
    return out;
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
/**
 * What the record says of a run with `--no-settle`: what was measured without
 * the wait, and how busy the processors were at each, so a budget met under
 * load is not read as one met on a quiet machine. `unsettled` is
 * { what: percent busy }; empty or null, nothing is said.
 */
export function unsettledNote(unsettled) {
  const steps = Object.entries(unsettled || {});
  if (!steps.length) return '';
  return `The budgets were measured without waiting for the computer to settle (--no-settle): just before each, the processors were ${listed(steps.map(([what, busy]) => `${busy}% busy (${what})`))}. Every budget and limit is as before. The owner, 13 Nov 2026: "Run them under load (Recommended)".`;
}

/**
 * What the record says of the gate's priority (raiseOwn), as --no-settle is
 * said: it changes what the budgets were measured under, never the budgets.
 * `priority` holds 'raised' and/or 'normal', one for each way a resumed run went.
 */
export function priorityNote(priority) {
  const ways = new Set(priority || []);
  if (!ways.has('raised')) return '';
  const part = ways.has('normal') ? 'For part of the run (a resumed one), the' : 'The';
  return `${part} gate raised its own processes on this computer to above normal priority (those that lived past its ten-second walk), beside whatever else the computer was doing; every budget and limit is as before.`;
}

/** The folder-removal suites' runs, on Node and again on Office's own Electron, for the Gate section: their skips said apart. */
const removalPart = (r) => (r ? `the folder-removal suites on Node and again on Office's own Electron: ${n(r.tests)} tests in ${r.runs ? `${n(r.runs)} ${r.runs === 1 ? 'run' : 'runs'}` : 'its runs'}, ${n(r.passed)} passing${r.skipped ? ` and ${n(r.skipped)} skipped` : ''}` : null);

/** The release note's Gate section, in the note's own plain style. */
export function gateSection(counts, { version, date, fewer = [], platform = process.platform, arch = process.arch, reran = [], apps = null, commit = null, closedByReason = [], hosts = null, unsettled = null, priority = null }) {
  const t = engineCounts(counts);
  const parts = [
    t ? `${n(t.tests)} engine tests, ${n(t.passed)} passing${t.skipped ? ` and ${n(t.skipped)} skipped where a package they need is not installed` : ''}` : null,
    removalPart(counts['test:removal']),
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
  // Failures closed with a reason and no fixing commit: each said here, so the list is never got round quietly (PM, 11 Nov 2026).
  const byReason = closedByReason.length
    ? `\n\n${wrap(`${closedByReason.length === 1 ? 'One failed check was' : `${n(closedByReason.length)} failed checks were`} closed with a reason rather than a fixing commit:`)}\n\n${closedByReason.map((c) => `- ${c.id} ${c.name}: ${c.reason}`).join('\n')}`
    : '';
  // Where each pass ran, when some ran on another machine: the host and its system, as that machine said (PM, 11 Nov 2026).
  const where = hosts ? `\n\n${wrap(hosts)}` : '';
  const underLoad = unsettledNote(unsettled);
  const raisedNote = priorityNote(priority);
  const loaded = `${underLoad ? `\n\n${wrap(underLoad)}` : ''}${raisedNote ? `\n\n${wrap(raisedNote)}` : ''}`;
  const at = `${less}${commit ? `\n\nCommit checked: ${commit}` : ''}`;
  return `## Gate\n\n${wrap(`Run on ${PLATFORM[platform] || platform} (${arch}) on ${date}, before ${version} was released: ${list}.`)}${where}${loaded}${at}${again}${byReason}\n`;
}

/**
 * What a run must not inherit. Each of these makes a check process run less
 * than the whole, and a gate started from a shell that has one set would
 * count what little ran as the gate.
 */
export const SHRINKING_ENV = ['RUTBA_VERIFY_ONLY', 'RUTBA_VERIFY_APPS', 'RUTBA_VERIFY_PART_IDS', 'RUTBA_VERIFY_SKIP_IDS', 'RUTBA_VERIFY_LANE', 'RUTBA_VERIFY_PARTS', 'RUTBA_LOCKER', 'RUTBA_CORPUS_LIMIT', 'RUTBA_A11Y_APPS', 'RUTBA_DARK_APPS', 'RUTBA_POLISH_APPS', 'RUTBA_RIBBON_APPS', 'RUTBA_SCREENS_ONLY'];

/** More checks than this failing beside the others closes the gate: that is a fault, not a busy machine. */
export const RERUN_LIMIT = 5;
/** The longest a run on the lab machine may take, its wait in the machine's queue included, before the gate ends it and stops. */
const VM_MINUTES = Math.max(5, Number(process.env.RUTBA_GATE_VM_MINUTES) || 90);
/** The app a check belongs to by its number in the registry, when that app is one of the window checks' apps; null when it has no number yet. */
export function appOfCheck(registry, name) {
  const id = registry?.ids?.[name];
  const num = id ? Number(String(id).split('.')[0]) : NaN;
  const app = Number.isFinite(num) ? Object.entries(registry.apps || {}).find(([, v]) => v === num)?.[0] : null;
  return app && VERIFY_APPS.includes(app) ? app : null;
}

/** The longest a window lane may go without a check result before it is ended as hung. */
export const CHECK_MINUTES = Math.max(2, Number(process.env.RUTBA_GATE_CHECK_MINUTES) || 15);
/** A check's result in a window lane's log: `ok   [07.0221] word: ... : detail`; the name is what is kept. */
export const CHECK_LINE = /^(?:ok|FAIL) +\[[^\]]+\] (.*?)(?: : .*)?$/;

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
  // tools/check-ids.json: the numbers the run gave the checks it saw for the first time.
  const own = new Set(['README.md', 'tools/gate-floors.json', 'tools/check-ids.json', `docs/releases/v${version}.md`]);
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
    // Counted in tests, as the engine tests are: a runtime's half or a suite left out shows as fewer.
    'test:removal': counts['test:removal']?.tests ?? null,
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
  held('folder-removal tests', now['test:removal'], was['test:removal']);
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
// `--parts mail#verifyMailSignIn,video#*`: those parts alone; `--failed`: the parts of every open entry in the failure list.
const partsArg = valueOf('--parts');
const failedOnly = argv.includes('--failed');
const VALUED = new Set(['--apps', '--jobs', '--parts', '--by', '--reason']);
const only = argv[0] === 'ledger' ? [] : argv.filter((a, i) => !a.startsWith('-') && !VALUED.has(argv[i - 1]));
const quiet = argv.includes('--quiet');
const record = argv.includes('--record');
const resume = argv.includes('--resume');
const changed = argv.includes('--changed');
const serial = argv.includes('--serial');
const fewer = argv.includes('--fewer-checks');
// `--no-settle`: the budgets and the machine lane are measured without waiting for the machine to
// settle, beside whatever else it is doing, every budget and limit as it is. Off unless asked, and
// said in the record when used. Owner, 13 Nov 2026, of "Skip the wait for a settled PC, run the
// budgets at above-normal priority, and keep every limit as it is": "Run them under load (Recommended)".
const noSettle = argv.includes('--no-settle');

/** Run a command; its output shown as it comes (unless quiet, or sent to `log`), and the end of it kept. */
/** The processes the gate has running, so a stop can end them. */
const live = new Set();

/**
 * End a process and everything it started. On Windows a child whose parent
 * has already gone is out of taskkill /T's reach, but it still names that
 * parent: so those are found by their parent's id and ended too (11 Nov: the
 * lab runner's bash died, its ssh lived on holding the gate's pipe, and the
 * time limit could not end it).
 */
export function endTree(pid, child = null, { startedAt = child?.startedAt ?? 0, endedAt = child?.endedAt ?? Infinity, list = listProcesses } = {}) {
  if (process.platform !== 'win32') { try { child?.kill('SIGKILL'); } catch { /* ended already */ } return; }
  const procs = list();
  // No listing: only the child itself, and only while it has not exited, so its id is still its own.
  if (!procs.length) {
    if (child && child.exitCode === null && child.signalCode === null) { try { execFileSync('taskkill', ['/pid', String(pid), '/T', '/F'], { stdio: 'ignore' }); } catch { /* ended already */ } }
    return;
  }
  // No /T: it goes by parent id alone, and would take a stranger's tree that once named one of these ids (review of 11 Nov).
  // Root first: the list is fixed already, and a parent ended first starts no new child in the gap.
  // With no start time (no child given) the root is never known as ours, and nothing is ended.
  for (const id of treeOf(pid, startedAt, procs, endedAt)) {
    try { execFileSync('taskkill', ['/pid', String(id), '/F'], { stdio: 'ignore' }); } catch { /* ended already */ }
  }
}

const LIST_PROCESSES = 'Get-CimInstance Win32_Process | ForEach-Object { "{0} {1} {2}" -f $_.ProcessId, $_.ParentProcessId, ([DateTimeOffset]$_.CreationDate).ToUnixTimeMilliseconds() }';
const parseProcesses = (out) => String(out || '').split(/\r?\n/).map((l) => l.trim().split(' ').map(Number)).filter((v) => v.length === 3 && v.every(Number.isFinite)).map(([p, pp, created]) => ({ pid: p, ppid: pp, created }));

/**
 * listProcesses without holding the gate still while PowerShell lists them: for the walk that raises priority.
 * Answers [] when the listing fails or passes its 30 s; a busy computer can take that long (the unit run of
 * 14 Nov 2026 05:12), and the walk then raises nothing that round and walks again at the next.
 */
export function listProcessesAsync() {
  return new Promise((resolve) => {
    execFile('powershell', ['-NoProfile', '-NonInteractive', '-Command', LIST_PROCESSES], { encoding: 'utf8', timeout: 30_000, maxBuffer: 2 ** 24, windowsHide: true }, (err, out) => resolve(err ? [] : parseProcesses(out)));
  });
}

/** Every process now: [{ pid, ppid, created }] (created in ms), or [] when they cannot be listed. */
function listProcesses() {
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', 'Get-CimInstance Win32_Process | ForEach-Object { "{0} {1} {2}" -f $_.ProcessId, $_.ParentProcessId, ([DateTimeOffset]$_.CreationDate).ToUnixTimeMilliseconds() }'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'], timeout: 30_000, maxBuffer: 2 ** 24 });
    return out.split(/\r?\n/).map((l) => l.trim().split(' ').map(Number)).filter((v) => v.length === 3 && v.every(Number.isFinite)).map(([p, pp, created]) => ({ pid: p, ppid: pp, created }));
  } catch {
    return [];
  }
}

/**
 * The ids to end for a process and its descendants, from a list of every
 * process, root first. Windows gives a dead process's id to another later,
 * so an id is listed only while the process holding it is the one we mean
 * (made within a few seconds of when we know it started), and a child counts
 * as a node's only when it was made after that node started and before the
 * node ended: its exit (`endedAt`, known for the root), or the start of a
 * later process that took its id. A dead node is not listed, but its children
 * are still walked (review of 11 Nov: the root id was ended whoever held it).
 */
export function treeOf(pid, startedAt, procs, endedAt = Infinity) {
  const out = [];
  const seen = new Set();
  const walk = (id, since, ended) => {
    if (seen.has(id)) return;
    seen.add(id);
    const now = procs.find((p) => p.pid === id);
    const ours = now && Math.abs(now.created - since) <= 5_000;
    if (ours) out.push(id);
    // The id taken over by a later process: only children made before it are this node's.
    const until = Math.min(ended, now && !ours && now.created > since ? now.created : Infinity);
    for (const k of procs) if (k.ppid === id && k.pid !== id && k.created >= since - 1_000 && k.created < until) walk(k.pid, k.created, Infinity);
  };
  walk(pid, startedAt, endedAt);
  return out;
}

/**
 * The gate and everything it starts run above normal priority, so other work
 * on the computer beside it (another session's build, a batch job) does not
 * starve its windows and timed checks (the owner, 13 Nov 2026: "you should
 * learn to co exisit"). Its own processes only, and no setting of the
 * computer is changed. One already higher is left as it is. Answers how many
 * were raised.
 */
export function raiseOwn(pids, { get = os.getPriority, set = (pid) => os.setPriority(pid, os.constants.priority.PRIORITY_ABOVE_NORMAL) } = {}) {
  let n = 0;
  for (const pid of pids) {
    try {
      // Lower numbers are higher priority: above normal is -7, high -14.
      if (get(pid) > os.constants.priority.PRIORITY_ABOVE_NORMAL) { set(pid); n += 1; }
    } catch { /* ended already, or not ours to raise */ }
  }
  return n;
}

// Windows does not pass a raised priority on to the processes a process starts (an Electron window started by
// npm through cmd.exe is three down), so every so often the gate's own trees are walked and what is new is raised.
const RAISE_EVERY_MS = 10_000;
// The processes walked already, as `pid:created` (raiseTrees).
const raisedIds = new Set();
// Set by main(): the gate run as a command raises; its parts imported by the tests do not.
let raiseOn = false;
let raising = null;
let raisingNow = false;
let raisedCount = 0;
function keepRaised() {
  if (!raiseOn || raising || process.env.RUTBA_GATE_PRIORITY === 'normal') return;
  raisedCount += raiseOwn([process.pid]);
  if (process.platform !== 'win32') return;
  raising = setInterval(async () => {
    if (raisingNow || !live.size) return;
    raisingNow = true;
    try {
      raisedCount += await raiseTrees(live, { skip: raisedIds });
    } finally {
      raisingNow = false;
    }
  }, RAISE_EVERY_MS);
  raising.unref?.();
}

/**
 * Every process in the trees of `children` (each { pid, startedAt, endedAt? },
 * as the gate keeps them), raised as raiseOwn does: the walk endTree makes, so
 * only processes that are these children's own. Processes in `skip` are passed
 * over and those walked are added to it, each kept by its id and when it was
 * made: Windows gives an ended process's id to a new one, which is walked as
 * new (review of 13 Nov). Answers how many were raised.
 */
export async function raiseTrees(children, { skip = new Set(), list = listProcessesAsync, raise = raiseOwn } = {}) {
  const procs = await list();
  const made = new Map(procs.map((p) => [p.pid, p.created]));
  const ids = new Map();
  for (const c of children) {
    for (const id of treeOf(c.pid, c.startedAt, procs, c.endedAt ?? Infinity)) {
      const key = `${id}:${made.get(id)}`;
      if (!skip.has(key)) ids.set(id, key);
    }
  }
  for (const key of ids.values()) skip.add(key);
  return raise([...ids.keys()]);
}
/** A child the gate has just started: raised at once (its own children follow with the next walk, which finds it raised). */
function raiseChild(child) {
  if (!raiseOn || process.env.RUTBA_GATE_PRIORITY === 'normal' || !child?.pid) return;
  raisedCount += raiseOwn([child.pid]);
  keepRaised();
}
/** How many of the gate's processes were raised above normal, for the run's last lines. */
export const raisedSoFar = () => raisedCount;

/** End every process the gate has running, and what each started. */
function stopLive() {
  for (const child of live) endTree(child.pid, child);
}

// Whatever way the gate ends, nothing it started is left running: the windows, or a run on the lab machine (review of 11 Nov).
process.on('exit', () => { if (live.size) stopLive(); });

export function run(command, args, env = {}, { log = null, cwd = ROOT, whole = false, shell = process.platform === 'win32', limitMs = 0, stallMs = 0 } = {}) {
  // `whole`: wait for the output to be all written before answering, for a log read in full after (the TAP of the engine tests).
  // `shell: false` for a program named by a path that may hold a space (RUTBA_GATE_VM_BASH); `limitMs`: ended past it, answering timedOut.
  if (whole) {
    return new Promise((resolve) => {
      const child = spawn(command, args, { cwd, stdio: ['ignore', 'pipe', 'pipe'], shell, env: { ...process.env, ...env } });
      child.startedAt = Date.now();
      live.add(child);
      raiseChild(child);
      const file = log ? fs.createWriteStream(log) : null;
      let tail = '';
      let timedOut = false;
      const keep = (chunk) => { if (file) file.write(chunk); tail = (tail + chunk).slice(-200000); };
      const timer = limitMs > 0 ? setTimeout(() => {
        timedOut = true;
        endTree(child.pid, child);
        // Its pipes held open by something it left: answered without them.
        setTimeout(() => finish(child.exitCode ?? 1), 15_000).unref?.();
      }, limitMs) : null;
      child.stdout.on('data', keep);
      child.stderr.on('data', keep);
      child.on('error', (err) => keep(`\n${err.message}\n`));
      let answered = false;
      const finish = (code) => {
        if (answered) return;
        answered = true;
        if (timer) clearTimeout(timer);
        live.delete(child);
        child.stdout.destroy();
        child.stderr.destroy();
        const done = { code: code ?? 1, tail, timedOut };
        if (file) file.end(() => resolve(done)); else resolve(done);
      };
      child.on('close', (code) => finish(code));
      // Ended, but its pipes held open by a process it left behind (an ssh whose bash died): what it started is ended, and the run answered.
      child.on('exit', (code) => {
        // From here its id may be given to another: only what it made before now is its.
        child.endedAt = Date.now();
        setTimeout(() => { if (!answered) { endTree(child.pid, child); finish(code); } }, 15_000).unref?.();
      });
    });
  }
  return new Promise((resolve) => {
    const child = spawn(command, args, { cwd, stdio: ['ignore', 'pipe', 'pipe'], shell, env: { ...process.env, ...env } });
    child.startedAt = Date.now();
    live.add(child);
    raiseChild(child);
    child.on('exit', () => { child.endedAt = Date.now(); live.delete(child); });
    const file = log ? fs.createWriteStream(log) : null;
    let tail = '';
    // `stallMs`: a window lane that gives no check result for this long is ended, and the last check it gave is kept (PM, 11 Nov 2026:
    // a lane hung over an hour with only clock lines, holding the lock).
    const watch = stallMs > 0 ? { last: Date.now(), after: null, line: '', failed: [] } : null;
    let hung = null;
    let exited = false;
    const onLines = (chunk) => {
      watch.line += chunk;
      const lines = watch.line.split(/\r?\n/);
      watch.line = lines.pop();
      for (const l of lines) {
        const m = CHECK_LINE.exec(l);
        if (!m) continue;
        watch.last = Date.now();
        watch.after = m[1];
        // The checks that failed before a hang: a hung lane writes no counts, and they would go unnamed (review of 11 Nov).
        if (l.startsWith('FAIL') && !watch.failed.includes(m[1])) watch.failed.push(m[1]);
      }
    };
    const timer = watch ? setInterval(() => {
      if (!hung && Date.now() - watch.last < stallMs) return;
      hung ||= { minutes: Math.round(stallMs / 60_000), after: watch.after, failed: watch.failed.slice() };
      // Tried again on every tick until the lane has gone: one kill that failed must not leave it hanging.
      if (exited) return;
      endTree(child.pid, child);
    }, Math.min(15_000, stallMs)) : null;
    if (timer?.unref) timer.unref();
    const keep = (to) => (chunk) => {
      if (file) file.write(chunk);
      else if (!quiet) to.write(chunk);
      tail = (tail + chunk).slice(-200000);
      if (watch) onLines(String(chunk));
    };
    child.stdout.on('data', keep(process.stdout));
    child.stderr.on('data', keep(process.stderr));
    child.on('exit', (code) => { exited = true; if (timer) clearInterval(timer); file?.end(); resolve({ code: code ?? 1, tail, hung }); });
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
const SECONDS = path.join(STATE_DIR, 'part-seconds.json');
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
      // No part named for what failed: the app whole.
      const ids = v.failedParts?.length ? v.failedParts : ['*'];
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
    // again (review of 11 Nov): a failed check not seen passing again is let
    // go only when the run again ran at least as many checks as the first run
    // did in all. A part's own catch-all ("the pages check ran", 416 of them)
    // is written only when something threw, so a clean run never shows it by
    // name; the steps the throw cut short show up as the extra checks. A run
    // again with fewer checks than that must pass each one by name.
    const passedAgain = new Set(p.passedNames || []);
    const ranBefore = before[id]?.of || 0;
    const unseen = p.of >= ranBefore ? [] : (before[id]?.realFailed || []).filter((n) => !passedAgain.has(n));
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
  // The gate itself and what it starts, above normal priority (raiseOwn). RUTBA_GATE_PRIORITY=normal leaves them as they are.
  raiseOn = true;
  keepRaised();
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
    // A reading kept from an earlier run (--resume) is for that run: the lane measured now replaces it, and its run again with it.
    const unsettled = { ...(state.unsettled || {}) };
    delete unsettled[what];
    if (what === 'the machine lane') delete unsettled['the machine lane\'s run again'];
    if (noSettle) {
      const busy = Math.round((await busyShare(1000)) * 100);
      say(`${what}: measured without waiting for the machine to settle (--no-settle), the processors ${busy}% busy.`);
      state.unsettled = { ...unsettled, [what]: busy };
      save();
      return true;
    }
    const calm = await settle();
    if (!calm.calm) { gateFault(`${what} not measured: the machine stayed ${calm.busy}% busy for ${Math.round(calm.waited / 60)} minutes`); return false; }
    state.unsettled = unsettled;
    save();
    return true;
  };

  for (const name of SHRINKING_ENV) {
    if (process.env[name] !== undefined) {
      console.log(`${name} was set in this shell and is ignored: the gate runs everything.`);
      delete process.env[name];
    }
  }

  // `npm run gate -- ledger`: the failure list; `ledger close <number> --by <commit>` or `--reason "..."` closes an entry.
  if (argv[0] === 'ledger') {
    let ledger;
    try { ledger = readLedger(LEDGER); } catch (err) { console.log(err.message); process.exit(2); }
    if (argv[1] === 'close' && argv[2]) {
      const by = valueOf('--by');
      const full = by ? git('rev-parse', '--verify', '--quiet', `${by}^{commit}`) : null;
      if (by && !full) { console.log(`No commit ${by} here.`); process.exit(2); }
      const descends = (head, commit) => { try { execFileSync('git', ['merge-base', '--is-ancestor', commit, head], { cwd: ROOT, stdio: 'ignore' }); return true; } catch { return false; } };
      const r = closeEntry(ledger, argv[2], { by: full, reason: valueOf('--reason'), descends });
      if (r.ok) { writeLedger(ledger, LEDGER); fs.writeFileSync(FAILURES_PAGE, ledgerMarkdown(ledger, { head: git('rev-parse', 'HEAD') })); }
      console.log(r.said);
      process.exit(r.ok ? 0 : 1);
    }
    // `ledger move <number> --to tests/<file>.test.js`: an engine test moved to another file to fix it; its entry and number go with it.
    if (argv[1] === 'move' && argv[2]) {
      const titles = (file) => {
        try {
          const text = fs.readFileSync(path.join(ROOT, file), 'utf8');
          return [...text.matchAll(/\btest\(\s*(['"`])((?:\\.|(?!\1)[^\\])*)\1/g)].map((m) => m[2].replace(/\\(.)/g, '$1'));
        } catch { return []; }
      };
      const r = moveEntry(ledger, argv[2], valueOf('--to'), { has: (file, title) => titles(file).includes(title) });
      if (r.ok) {
        const registry = readRegistry();
        if (registry.ids[r.from] && !registry.ids[r.name]) { registry.ids[r.name] = registry.ids[r.from]; delete registry.ids[r.from]; writeRegistry(registry); }
        writeLedger(ledger, LEDGER);
        fs.writeFileSync(FAILURES_PAGE, ledgerMarkdown(ledger, { head: git('rev-parse', 'HEAD') }));
      }
      console.log(r.said);
      process.exit(r.ok ? 0 : 1);
    }
    const open = openEntries(ledger);
    console.log(`The failure list: ${open.length} open, ${Object.keys(ledger.entries).length - open.length} closed (${path.relative(ROOT, LEDGER)}).`);
    for (const line of ledgerLines(open, { max: 400 })) console.log(line);
    if (open.length) console.log(`\nRun their parts: npm run gate -- verify:apps --failed. Close one once fixed and passed: npm run gate -- ledger close <number> --by <commit>.`);
    process.exit(0);
  }

  // A pass named that the gate does not have would run nothing and still open the gate.
  const known = ['test', 'test:perf', 'test:removal', 'verify:edit', 'verify:apps', 'smoke', 'verify:packs', 'office:reopen', ...FUZZERS.map(([name]) => name), 'corpus'];
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

  // A full run on the record waits for the failure list to be worked through:
  // each failure fixed and proved on its own app first, then one full run
  // (owner, 11 Nov 2026). Open entries are listed, and the run does not start.
  // A list that is there and does not read stops the gate: read as empty, it would let --record past it.
  let ledgerAtStart;
  try {
    ledgerAtStart = readLedger(LEDGER);
  } catch (err) {
    console.log(`The gate will not start: ${err.message}.`);
    process.exit(1);
  }
  if (record) {
    const open = openEntries(ledgerAtStart);
    if (open.length) {
      console.log(`The gate will not go on the record: the failure list has ${open.length} open ${open.length === 1 ? 'entry' : 'entries'}. Fix each at its cause, run its parts (npm run gate -- verify:apps --failed), and close it (npm run gate -- ledger close <number> --by <commit>).`);
      for (const line of ledgerLines(open)) console.log(line);
      process.exit(1);
    }
  }

  // The parts asked for: by `--parts`, or every open entry's with `--failed`.
  if (argv.includes('--parts') && !partsArg) { console.log('--parts needs the ids: --parts mail#reading,video#verifyVideoMp4'); process.exit(2); }
  const askedIds = partsArg ? partsArg.split(',').map((x) => x.trim()).filter(Boolean) : failedOnly ? openParts(ledgerAtStart) : [];
  if (failedOnly && !askedIds.length) { console.log('The failure list has no open entry: nothing to run.'); process.exit(0); }
  // The engine's test files, and which of them touch Windows itself and stay on this computer.
  let engineAll;
  let hostsList;
  try {
    engineAll = engineFiles(ROOT);
    hostsList = readHosts(ROOT);
  } catch (err) {
    console.log(`The gate will not start: ${err.message}.`);
    process.exit(1);
  }
  // Ids of the engine's (test#<file>, fuzz:pdf#*) and of the windows' (<app>#<part>).
  const engineAsked = engineParts(askedIds.filter(isEngineId), engineAll);
  const partIds = askedIds.filter((id) => !isEngineId(id));
  if (askedIds.length) {
    // A part's id names its app: one that names none would run nothing and look like a pass.
    const strange = [...partIds.filter((id) => !VERIFY_APPS.includes(id.split('#')[0]) || !id.split('#')[1]), ...engineAsked.strange];
    if (strange.length) { console.log(`No such part: ${strange.join(', ')}. A window part is <app>#<part>, the apps ${VERIFY_APPS.join(', ')}; an engine one is test#<file>, test#* or fuzz:<name>#*.`); process.exit(2); }
    // A run of some parts says nothing a resume or the record may take for a chunk whole (review of 11 Nov).
    if (record || resume) { console.log(`${record ? '--record' : '--resume'} is for whole chunks: a run of ${failedOnly ? 'the failed parts' : 'named parts'} goes on neither.`); process.exit(2); }
  }

  // Which apps' window checks this run is for: those named, those changed, those of the parts asked for, or all.
  let wanted = VERIFY_APPS.slice();
  if (askedIds.length) wanted = VERIFY_APPS.filter((a) => partIds.some((id) => id.startsWith(`${a}#`)));
  else if (appsArg) wanted = appsArg.split(',').map((a) => a.trim()).filter((a) => VERIFY_APPS.includes(a));
  else if (changed) {
    // The tag before this commit: one already on HEAD would make every change look like none.
    const since = git('describe', '--tags', '--abbrev=0', 'HEAD^');
    const files = [...git('diff', '--name-only', `${since}..HEAD`).split('\n'), ...git('diff', '--name-only').split('\n')].filter(Boolean);
    wanted = wantedForChanged(since, files);
    say(since ? `Changed since ${since}: ${files.length} files; the window checks of ${wanted.join(', ') || 'no app'}.` : 'No tag to compare with: the window checks of every app.');
  }
  const partial = wanted.length < VERIFY_APPS.length || only.length > 0 || askedIds.length > 0;
  // A run of some parts runs those parts: the window checks, engine test files or fuzz passes they name, nothing else (review of 11 Nov).
  const want = (name) => (askedIds.length
    ? (name === 'verify:apps' && partIds.length > 0) || (name === 'test' && engineAsked.tests.length > 0) || engineAsked.fuzz.includes(name)
    : !only.length || only.includes(name));
  if (askedIds.length && serial) { console.log('--serial runs every check in one process: it cannot run parts by id. Leave it out.'); process.exit(2); }

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
  // How the gate's processes ran, each run that a resume adds up: the record says it (priorityNote).
  // A state kept from before the gate raised anything ran at normal.
  state.priority = [...new Set([...(state.priority || (Object.keys(state.passes || {}).length || Object.keys(state.apps).length ? ['normal'] : [])), process.platform === 'win32' && process.env.RUTBA_GATE_PRIORITY !== 'normal' ? 'raised' : 'normal'])];
  // A run of some parts keeps no state: a resume would take its chunk for passed whole (review of 11 Nov).
  const save = () => { if (!askedIds.length) fs.writeFileSync(STATE, JSON.stringify(state, null, 2)); };
  // What ran in this run, not before it: the failure list adds only these (a resumed chunk's failures were counted when they happened).
  const ranNow = new Set();
  let machineRanNow = false;
  if (resume && !state.passes.test && !Object.keys(state.apps).length) say('Nothing to resume on this commit: a run from the start.');
  const timing = readJson(TIMING) || {};
  // Each part's seconds, measured in earlier runs: the chunks are cut by them (tools/gate-plan.js).
  const measured = readSeconds(SECONDS);
  const jobs = Math.max(1, Number(jobsArg) || laneCap());
  if (want('verify:apps') && wanted.length && screenLocked()) {
    say('The screen is locked: everything that does not need it runs now, and the machine lane (the computer\'s own clipboard, real input, the screen\'s capture) waits for it to be unlocked.');
  }

  /** A pass of its own (test, test:perf, verify:edit, smoke), its counts kept on the state. */
  const pass = async (name, what, command, args, { log = null, env = {}, whole = false } = {}) => {
    if (resume && state.passes[name]?.ok) {
      say(`${name}: passed in the run before on this commit.`);
      results.push({ name, what, ok: true, seconds: '0.0', resumed: true });
      if (state.passes[name].counts) counts[name] = state.passes[name].counts;
      return true;
    }
    const started = Date.now();
    if (!log) say(`\n\u001b[1m── ${name}: ${what} ─────────────────────────────\u001b[0m\n`);
    else say(`  [${name}] started`);
    // `whole`: every line written before the count is read, and read from the log, not the last 200,000 characters kept.
    const out = await run(command, args, env, { log, whole });
    let text = out.tail;
    if (whole && log) try { text = fs.readFileSync(log, 'utf8'); } catch { /* the tail, then */ }
    const counted = countsOf(name, text);
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
  // Each window lane's log this run, for the table of what the waits that gave up cost.
  const laneLogs = new Set();
  const windowRun = async (label, env, { mayBeEmpty = false } = {}) => {
    const partsFile = path.join(PARTS_DIR, `${label}.json`);
    try { fs.rmSync(partsFile, { force: true }); } catch { /* none */ }
    const log = path.join(LOGS, `${label}.log`);
    laneLogs.add(log);
    const started = Date.now();
    say(`  [${label}] started`);
    const out = await run('node', [path.join('apps', 'desktop', 'build', 'verify-apps-run.js')], { ...env, RUTBA_VERIFY_PARTS: partsFile }, { log, stallMs: CHECK_MINUTES * 60_000 });
    if (out.hung) say(`  [${label}] HUNG: no check result for ${out.hung.minutes} minutes${out.hung.after ? ` after "${out.hung.after}"` : ', none at all'}; ended (RUTBA_GATE_CHECK_MINUTES)`);
    const seconds = (Date.now() - started) / 1000;
    // A lane ended as hung is read as one that crashed, whatever parts file it left: its app runs again whole, and the checks
    // it said failed before the hang keep their names (review of 11 Nov: a parts file written, then a window that never quit, failed unnamed).
    const file = out.hung ? null : readJson(partsFile);
    const parts = file?.apps || null;
    const sum = parts ? Object.values(parts).reduce((t, p) => ({ passed: t.passed + p.passed, of: t.of + p.of, failed: [...t.failed, ...p.failed] }), { passed: 0, of: 0, failed: [] }) : null;
    // A chunk that ran no check tested nothing: it is not a pass.
    const ok = out.code === 0 && Boolean(sum) && sum.failed.length === 0 && (mayBeEmpty || sum.of > 0);
    say(`  [${label}] ${ok ? 'ok  ' : 'FAIL'} ${sum ? `${sum.passed}/${sum.of}` : 'ended without its count'} in ${seconds.toFixed(0)} s${ok ? '' : ` (${path.relative(ROOT, log)})`}`);
    if (sum && !ok) for (const f of sum.failed.slice(0, 8)) say(`         FAIL ${f}`);
    return { ok, seconds, sum: sum || { passed: 0, of: 0, failed: ['the run ended before it counted', ...(out.hung?.failed || [])] }, parts, seen: file?.seen || [], hung: out.hung || null };
  };

  // 0. The engine tests and the fuzz passes: here, or on the lab machine
  // from the very start, beside everything else (PM, 11 Nov 2026). The tests
  // that touch Windows itself stay here (tools/engine-hosts.json).
  const vm = vmConfig(process.env, argv);
  if (vm.said) say(`  ${vm.said}`);
  const seed = String(parseInt(head.slice(0, 8), 16));
  // Listed in engine-hosts.json, or reaching through its imports a module that touches Windows: here.
  const reachCache = new Map();
  const split = splitEngine(engineAll, hostsList, vm.runner ? (f) => osReach(ROOT, f, reachCache) : () => null);
  /** The host a test file belongs to now: with a lab machine set up, the one the split gives it, even in a --local run. */
  const homeOf = (file) => (vm.runner && !split.local.includes(file) ? vm.host : 'local');
  const someFiles = engineAsked.tests.length > 0 && engineAsked.tests[0] !== '*';
  const testFiles = someFiles ? engineAsked.tests : engineAll;
  const testGroups = !want('test') ? [] : vm.on
    ? [{ key: 'vm', host: vm.host, files: testFiles.filter((f) => split.vm.includes(f)), what: `the engines, on ${vm.host}` },
      { key: 'local-windows', host: 'local', files: testFiles.filter((f) => split.local.includes(f)), what: 'the engines that touch Windows, here' }].filter((g) => g.files.length)
    : [{ key: 'local-all', host: 'local', files: testFiles, what: 'the engines' }];
  const groupDone = (key) => resume && state.passes.test?.groups?.[key]?.ok;
  /** What each test group and fuzz pass found in this run, for the failure list. */
  const engineRuns = [];
  const fuzzRuns = [];
  const testOk = (r) => r.code === 0 && Boolean(r.tap?.totals) && r.tap.totals.tests > 0 && !r.tap.failures.length;
  /** A test group's files that failed, when a run of them again alone can say anything: each failure placed in a file, few enough. */
  const testsToRerun = (r) => {
    if (!r.tap?.totals || !r.tap.failures.length || r.tap.failures.some((f) => !f.file)) return null;
    const files = [...new Set(r.tap.failures.map((f) => f.file))];
    return files.length <= RERUN_LIMIT ? files : null;
  };
  /** Cleared by the run again: it passed, and passed each test that failed by its name. */
  const clearedBy = (first, again) => again && again.code === 0 && again.tap?.totals && !again.tap.failures.length && first.tap.failures.every((f) => again.tap.passed.includes(f.name));

  /** Engine test files run here, into one TAP log read whole. */
  const localTests = async (files, { label, alone = false } = {}) => {
    const log = path.join(LOGS, `${label}.log`);
    const started = Date.now();
    say(`  [${label}] started: ${files === engineAll ? 'every engine test file' : `${files.length} engine test file${files.length === 1 ? '' : 's'}`}`);
    // The whole suite by the globs npm test names; some files by their names.
    // Each test's file and line go to a file of their own beside the log (tools/gate-reporter.js): TAP names no file for a pass.
    const events = path.relative(ROOT, path.join(LOGS, `${label}-events.jsonl`)).split(path.sep).join('/');
    try { fs.rmSync(path.join(ROOT, events), { force: true }); } catch { /* none */ }
    const args = [path.join('tools', 'run-temp.js'), 'node', ...testArgs(files === engineAll ? engineGlobs(ROOT) : files, { events, alone })];
    const out = await run('node', args, {}, { log, whole: true });
    let text = '';
    try { text = fs.readFileSync(log, 'utf8'); } catch { text = out.tail; }
    let lines = '';
    try { lines = fs.readFileSync(path.join(ROOT, events), 'utf8'); } catch { /* the reporter wrote none */ }
    const tap = readRun(text, lines, engineAll);
    const r = { code: out.code, tap, seconds: (Date.now() - started) / 1000, log };
    say(`  [${label}] ${testOk(r) ? 'ok  ' : 'FAIL'} in ${secs(Date.now() - started)} s${tap.totals ? `, ${tap.totals.passed}/${tap.totals.tests}` : ', ended without its count'}${testOk(r) ? '' : ` (${path.relative(ROOT, log)})`}`);
    for (const f of tap.failures.slice(0, 8)) say(`         FAIL ${f.name}${f.at ? ` [where] ${f.at}` : ''}`);
    return r;
  };

  /** One run on the lab machine of some passes; its logs come back beside the gate's own. */
  let startTree = null;
  const vmBatch = async (label, passes, { files = null, alone = false } = {}) => {
    const cmd = ['node', 'tools/run-temp.js', 'node', 'tools/gate-vm.js', '--passes', passes.join(','), '--seed', seed, ...(files ? ['--files', files.join(',')] : []), ...(alone ? ['--alone'] : [])];
    const [prog, args] = vmCommand(vm, ROOT, label, cmd);
    const log = path.join(LOGS, `${label}.log`);
    const started = Date.now();
    say(`  [${label}] started on ${vm.host}: ${passes.join(', ')}${files ? `, ${files.length} test file${files.length === 1 ? '' : 's'}` : ''}${alone ? ', one at a time' : ''}`);
    const back = path.join(ROOT, '.gate', 'vm', 'logs');
    try { fs.rmSync(path.join(back, 'vm-result.json'), { force: true }); } catch { /* none */ }
    // Without a shell: a bash named by a path with a space in it (RUTBA_GATE_VM_BASH) is one program, not two words. Ended past its time limit.
    const out = await run(prog, args, {}, { log, whole: true, shell: false, limitMs: VM_MINUTES * 60_000 });
    const result = readJson(path.join(back, 'vm-result.json'));
    let fault = out.timedOut ? `it did not end within ${VM_MINUTES} minutes (RUTBA_GATE_VM_MINUTES), waiting in the machine's queue or stuck: ended` : null;
    fault ||= vmFault(out.code, out.tail) || (result ? null : 'it ran and brought back no word of what it ran (.gate/vm/logs/vm-result.json)');
    // The machine ran the tree this gate started on, or its passes stand for nothing here.
    const snap = snapshotOf(out.tail);
    const changedSince = () => { try { return git('diff-tree', '-r', '--name-only', startTree, `${snap}^{tree}`).split('\n').filter(Boolean).slice(0, 8).join(', '); } catch { return ''; } };
    if (!fault && !snap) fault = 'the runner did not say which snapshot of the tree it ran';
    else if (!fault && startTree && git('rev-parse', `${snap}^{tree}`) !== startTree) fault = `it ran another tree than this gate started on: a file changed after the start (${changedSince() || 'which, git could not say'})`;
    // What the machine says it ran, not only what was sent: another run from this clone checks out its own snapshot in the same folder there.
    else if (!fault && startTree && result.tree !== startTree) fault = `it ran the tree ${String(result.tree).slice(0, 12)}, not the ${startTree.slice(0, 12)} this gate started on: another run from this clone took its folder there`;
    const got = {};
    for (const name of passes) {
      const own = path.join(LOGS, `${label}-${name.replace(':', '-')}.log`);
      let text = '';
      try { text = fs.readFileSync(path.join(back, `${name.replace(':', '-')}.log`), 'utf8'); fs.writeFileSync(own, text); } catch { /* none came back */ }
      const p = result?.passes?.[name];
      let lines = '';
      if (name === 'test') try { lines = fs.readFileSync(path.join(back, 'test-events.jsonl'), 'utf8'); fs.writeFileSync(path.join(LOGS, `${label}-test-events.jsonl`), lines); } catch { /* the reporter wrote none */ }
      got[name] = { code: p ? p.code : 1, seconds: p?.seconds ?? 0, text, log: own, tap: name === 'test' ? readRun(text, lines, engineAll) : null, counted: name === 'test' ? null : countsOf(name, text) };
    }
    say(`  [${label}] ${fault ? 'FAULT' : out.code === 0 ? 'ok  ' : 'FAIL'} in ${secs(Date.now() - started)} s${fault ? `: ${fault}` : ''} (${path.relative(ROOT, log)})`);
    return { fault, result, passes: got };
  };

  /** The lab machine could not be reached or ran the wrong tree: the gate stops and says so, and nothing runs here in its place. */
  const vmStop = (fault) => {
    gateFault(`${vm.host}: ${fault}`);
    console.log(`\n\u001b[31mThe gate stopped.\u001b[0m The engine tests and fuzz passes did not run on ${vm.host}: ${fault}. Run it again once the machine answers, or with --local to run them on this computer.`);
    stopLive();
    state.harness = harness;
    save();
    process.exit(1);
  };

  // The lab machine's passes start now: the first run, then each that failed again alone there (PM, 11 Nov 2026).
  const vmTests = testGroups.find((g) => g.key === 'vm' && !groupDone('vm'));
  const vmFuzz = vm.on ? FUZZERS.filter(([name]) => want(name) && !(resume && state.passes[name]?.ok)).map(([name]) => name) : [];
  let vmSettled = false;
  const vmJob = vm.on && (vmTests || vmFuzz.length) ? (async () => {
    startTree = worktreeHash(ROOT);
    const first = await vmBatch(`vm-${head.slice(0, 8)}`, [...(vmTests ? ['test'] : []), ...vmFuzz], { files: someFiles ? vmTests?.files : null });
    if (first.fault) vmStop(first.fault);
    // The machine ran the files this gate gave it, no more and no fewer: else a file would run on both machines, or on neither.
    const ranThere = first.result?.passes?.test?.list;
    if (vmTests && Array.isArray(ranThere)) {
      const want = new Set(vmTests.files);
      const extra = ranThere.filter((f) => !want.has(f));
      const missing = vmTests.files.filter((f) => !ranThere.includes(f));
      if (extra.length || missing.length) vmStop(`it split the engine tests another way than this gate did: ${extra.length} more (${extra.slice(0, 4).join(', ')}), ${missing.length} fewer (${missing.slice(0, 4).join(', ')})`);
    } else if (vmTests) vmStop('it did not say which test files it ran (vm-result.json passes.test.list)');
    const t = first.passes.test;
    const files = t && !testOk(t) ? testsToRerun(t) : null;
    if (t && !testOk(t) && !files) say(`  [vm] ${t.tap?.totals ? `the engine tests failed in more than ${RERUN_LIMIT} files, or outside any` : 'the engine tests ended without their count'}: not run again, and the gate stays closed`);
    const fuzzAgain = vmFuzz.filter((name) => first.passes[name].code !== 0);
    if (!files && !fuzzAgain.length) return { first, again: null };
    const again = await vmBatch(`vm-again-${head.slice(0, 8)}`, [...(files ? ['test'] : []), ...fuzzAgain], { files, alone: true });
    if (again.fault) vmStop(again.fault);
    return { first, again };
  })().catch((err) => vmStop(`the gate could not run it: ${err?.message || err}`)).finally(() => { vmSettled = true; }) : null;

  /** The test groups' verdict, once all have run: into the state, the counts and the results. */
  /** A test group's record in the state; a group cleared by its run again counts its tests as passed, with those that needed it named. */
  const groupRecord = (g, d) => {
    const cleared = !testOk(d.first) && clearedBy(d.first, d.again);
    const t = d.first.tap?.totals || null;
    return { cleared, record: { ok: testOk(d.first) || Boolean(cleared), host: g.host, os: d.os || null, osName: d.osName || null, cpus: d.cpus || null, files: g.files.length, counts: t && cleared ? { ...t, passed: t.passed + t.failed, failed: 0 } : t, reran: cleared ? d.first.tap.failures.map((f) => f.name) : [] } };
  };
  const engineVerdict = (done) => {
    const groups = { ...(state.passes.test?.groups || {}) };
    for (const g of testGroups) {
      const d = done[g.key];
      if (!d) {
        if (groups[g.key]?.ok) results.push({ name: 'test', what: g.what, ok: true, seconds: '0.0', resumed: true });
        continue;
      }
      const { cleared, record: rec } = groupRecord(g, d);
      const ok = rec.ok;
      groups[g.key] = rec;
      engineRuns.push({ host: g.host === 'local' ? 'local' : vm.host, code: d.first.code, tap: d.first.tap, again: d.again?.tap || null, againCode: d.again?.code ?? null });
      if (cleared) say(`  [${g.key}] again: each passed alone`);
      else if (d.again) say(`  [${g.key}] again: still failing`);
      results.push({ name: 'test', what: g.what, ok, seconds: secs(d.first.seconds * 1000), log: d.first.log });
      if (!ok) failed = true;
    }
    const needed = testGroups.map((g) => groups[g.key]);
    const sum = needed.filter((g) => g?.counts).reduce((t, g) => ({ tests: t.tests + g.counts.tests, passed: t.passed + g.counts.passed, failed: t.failed + g.counts.failed, skipped: t.skipped + g.counts.skipped }), { tests: 0, passed: 0, failed: 0, skipped: 0 });
    if (needed.length) {
      // Only this run's groups: one kept from a run split another way (the lab machine then, --local now) would be named in the record.
      const own = Object.fromEntries(testGroups.map((g) => [g.key, groups[g.key]]).filter(([, v]) => v));
      state.passes.test = { ok: needed.every((g) => g?.ok), counts: sum.tests ? sum : null, groups: own };
      if (sum.tests) counts.test = sum;
      save();
    }
  };
  /** A test group run here, then again alone for what failed. */
  const localGroup = async (g) => {
    const first = await localTests(g.files, { label: g.key === 'local-all' ? 'test' : 'test-windows' });
    const files = !testOk(first) ? testsToRerun(first) : null;
    const again = files ? await localTests(files, { label: 'again-test', alone: true }) : null;
    return { first, again, os: `${os.type()} ${os.release()}`, osName: PLATFORM[process.platform] || process.platform };
  };
  const localDone = {};
  const localTask = testGroups.find((g) => g.host === 'local' && !groupDone(g.key));
  // A group cleared only by its run again is kept as not passed until the failure list has its flaky entry: --resume runs it again.
  const early = ({ cleared, record }) => (cleared ? { ...record, ok: false } : record);
  const runLocalTests = async () => {
    if (!localTask) return true;
    localDone[localTask.key] = await localGroup(localTask);
    // Kept at once, not after the lanes: a run stopped later does not run these again on --resume (review of 11 Nov). The pass's own verdict waits for every group.
    state.passes.test = { counts: null, ...(state.passes.test || {}), ok: false, groups: { ...(state.passes.test?.groups || {}), [localTask.key]: early(groupRecord(localTask, localDone[localTask.key])) } };
    save();
    return true;
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

  // 2b. The folder-removal suites again, on Office's own Electron as Node as well as on Node (D-2026-11-13-37):
  // Electron 44's Node follows a Windows junction in a recursive fs.rm, and the engine suite above runs on Node
  // only. Here, on this computer, where junctions are real; never on the lab machine.
  // Logged and read whole: three runtimes' TAP, each run's totals counted (review of dfd0566b).
  if (want('test:removal')) await pass('test:removal', 'folder removal on Node and Electron', 'npm', ['run', 'test:removal'], { log: path.join(LOGS, 'test-removal.log'), whole: true });

  const lanes = serial ? null : { parallel: [], machine: null };
  // What the window checks found, and the names they passed (for their numbers): the failure list takes them at the end.
  let windowFindings = null;
  const windowSeen = [];
  if (serial) {
    // The run as it was: one process, every check in turn.
    say('  --serial keeps no failure list and gives no check numbers: the lanes do (npm run gate).');
    if (want('test')) await runLocalTests();
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
    // Chunks by the seconds measured before (tools/gate-plan.js): a long app in pieces of its parts; a run of named parts is not cut.
    const left = wanted.filter((a) => !(resume && state.apps[a]?.ok));
    let current = null;
    try { current = sourceIds(fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'main', 'verify-apps.js'), 'utf8')); } catch { /* no piece is named, every app runs whole */ }
    const plan = partIds.length ? left.map((a) => ({ label: a, app: a, ids: null, skip: null, seconds: timing[a] ?? null })) : planChunks(left, { timing, seconds: measured, jobs, current: current || new Set() });
    const chunks = plan.map((c) => c.label);
    const unmeasured = plan.filter((c) => c.seconds == null).map((c) => c.label);
    if (unmeasured.length) say(`  not measured before, so run first: ${unmeasured.join(', ')}`);
    const freeNow = os.freemem() / 2 ** 30;
    say(`\n\u001b[1m── side by side: up to ${jobs} at once ─────────────────────\u001b[0m`);
    say(jobsArg ? `  ${jobs} at once, as asked` : `  ${os.cpus().length} processors allow ${jobs}; ${freeNow.toFixed(1)} GB of memory free now, and a lane is added only while 1.5 GB is free for it`);
    say(`  ${[want('test') ? 'the engine tests' : null, want('verify:edit') ? 'the editing checks' : null, want('verify:apps') ? `the window checks of ${chunks.length ? chunks.join(', ') : 'no chunk left'}` : null].filter(Boolean).join(', ')}; the logs are in ${path.relative(ROOT, LOGS)}`);
    const started = Date.now();
    const tasks = [];
    if (localTask) tasks.push(runLocalTests);
    if (want('verify:apps')) {
      // The pieces of an app run as chunks of their own; the app's record is written once the last of them ends.
      const pieces = {};
      for (const c of plan) {
        tasks.push(async () => {
          // Asked for a machine part alone, an app's chunk may run nothing: what each id ran is checked after the lanes.
          const ids = partIds.length ? partIds : c.ids;
          const r = await windowRun(c.label, { RUTBA_VERIFY_LANE: 'parallel', RUTBA_VERIFY_APPS: c.app, ...(ids ? { RUTBA_VERIFY_PART_IDS: ids.join(',') } : {}), ...(c.skip ? { RUTBA_VERIFY_SKIP_IDS: c.skip.join(',') } : {}) }, { mayBeEmpty: partIds.length > 0 });
          (pieces[c.app] ||= []).push({ ...r, planIds: c.ids });
          const all = plan.filter((x) => x.app === c.app).length;
          if (pieces[c.app].length < all) return r.ok;
          const a = c.app;
          ranNow.add(a);
          // A part a piece was cut to run that ran nothing: the plan's fault, said by its id, and forgotten so the next plan does not name it.
          // A piece that wrote its parts file but nothing for the app ran none of its parts either (review of 11 Nov).
          const stale = pieces[a].filter((p) => p.parts).flatMap((p) => (p.planIds || []).filter((id) => !(p.parts[a]?.parts?.[id]?.of > 0)));
          if (stale.length) {
            gateFault(`${a}: a piece was cut to run ${stale.join(', ')}, and ${stale.length === 1 ? 'it' : 'they'} ran nothing (renamed or removed since measured); forgotten for the next run`);
            for (const id of stale) delete measured.parts[id];
          }
          // realFailed and passedNames too: the run again of an app whole (`<app>#*`) is held to them (review of 11 Nov).
          state.apps[a] = mergePieces(pieces[a], a, leftToMachine);
          // A run of some parts is no measure of the chunk whole, which orders the next full run.
          if (state.apps[a].ok && !partIds.length) {
            timing[a] = state.apps[a].seconds;
            noteSeconds(measured, { [a]: state.apps[a] }, { whole: [a] });
          } else noteSeconds(measured, { [a]: state.apps[a] });
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
      // Asked for parts that are none of them the machine lane's: it is not waited for or settled for (review of 11 Nov).
      const machineAsked = !partIds.length || partIds.some((id) => id.endsWith('#*') || wanted.some((a) => (state.apps[a]?.toMachine || []).includes(id)));
      if (!machineAsked) {
        say('  [machine] not needed: none of the parts asked for is the machine lane\'s');
        state.machine = { ok: true, passed: 0, of: 0, failed: [], seconds: 0, byApp: null, scope: wanted.slice(), skipped: true };
      } else if (resume && machineCovers(state.machine, wanted)) say('  [machine] passed in the run before on this commit, for these apps');
      else if (!(await unlockedWithin(60 * 60_000))) {
        // Not run, so not failed: what it checks waits, and --resume runs it once the screen is unlocked.
        say('  [machine] not run: the screen stayed locked. Unlock it and run the gate again with --resume.');
        failed = true;
      } else {
        say(`\n\u001b[1m── machine: the computer's own clipboard and keys, alone ──\u001b[0m`);
        // Its budgets are measured on a settled machine, as test:perf's are: it used to start with no settle at all.
        if (await budgetsReady('the machine lane')) {
          const r = await windowRun('machine', { RUTBA_VERIFY_LANE: 'machine', ...(wanted.length < VERIFY_APPS.length ? { RUTBA_VERIFY_APPS: wanted.join(',') } : {}), ...(partIds.length ? { RUTBA_VERIFY_PART_IDS: partIds.join(',') } : {}) }, { mayBeEmpty: wanted.length < VERIFY_APPS.length || partIds.length > 0 });
          state.machine = { ok: r.ok, passed: r.sum.passed, of: r.sum.of, failed: r.sum.failed, seconds: Math.round(r.seconds), byApp: r.parts || null, scope: wanted.slice(), seen: r.seen, hung: r.hung || null };
          machineRanNow = true;
          // A part marked for the machine lane that the chunks passed by and the machine lane never reached ran in no lane at all.
          // A machine lane that wrote no parts file crashed: that is said as itself, not as every machine part unrun.
          const unrun = r.parts ? machinePartsUnrun(wanted.flatMap((a) => state.apps[a]?.toMachine || []), r.seen) : [];
          if (!r.parts) say('  the machine lane ended without writing its parts file: its window process crashed or was stopped');
          if (unrun.length) gateFault(`parts marked for the machine lane that no lane ran (a part inside a section is never reached there): ${unrun.join(', ')}`);
        } else state.machine = { ok: false, passed: 0, of: 0, failed: [], scope: [], notMeasured: true };
        save();
      }

      // 5. Again, alone: the parts whose checks failed, in a process of their
      // own after the rest, the machine perhaps quieter by now. The gate
      // passes only if each check that failed passes here.
      // What the lanes found before any run again clears it: the failure list's first sight.
      const firstApps = JSON.parse(JSON.stringify(Object.fromEntries([...ranNow].map((a) => [a, state.apps[a]]).filter(([, v]) => v))));
      const firstMachine = machineRanNow && state.machine?.byApp ? JSON.parse(JSON.stringify(state.machine.byApp)) : null;
      const machineCrashed = machineRanNow && !state.machine?.byApp;
      // What the machine lane said failed before it hung, each by the app its number says (tools/check-ids.json), so --failed and a later pass find it.
      const machineHung = machineCrashed && state.machine?.hung ? { ...state.machine.hung, failed: (state.machine.hung.failed || []).map((name) => ({ name, app: appOfCheck(readRegistry(), name) })) } : null;
      // Each part asked for by id must have run in some lane: an id that names no part ran nothing and would look like a pass.
      if (partIds.length) {
        const ranIds = new Set();
        for (const [a, v] of [...Object.entries(firstApps), ...Object.entries(firstMachine || {})]) {
          for (const [pid, p] of Object.entries(v?.parts || {})) if (p.of > 0) { ranIds.add(pid); ranIds.add(`${a}#*`); }
        }
        const unrun = partIds.filter((id) => !ranIds.has(id));
        if (unrun.length) gateFault(`asked for by id and run in no lane (no such part now, or its lane did not run): ${unrun.join(', ')}`);
      }
      const clearedApps = new Set();
      const again = partsToRerun(Object.fromEntries(wanted.map((a) => [a, state.apps[a]])), state.machine);
      const againFiles = [];
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
        if (r.parts) againFiles.push(r.parts);
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
            // Only an app the run again ran some part of: the verdict says nothing of another.
            if (!set.ids.some((id) => id.startsWith(`${a}#`))) continue;
            const v = state.apps[a];
            state.apps[a] = { ...v, ok: true, reran: v.failed, failed: [] };
            clearedApps.add(a);
          }
        }
        save();
      }
      // What the window checks found, for the failure list written once the engine passes are in too.
      // Apps run only in part: their parts prove nothing of the app whole.
      const narrowed = new Set(partIds.length ? wanted.filter((a) => !partIds.includes(`${a}#*`)) : []);
      windowFindings = runFindings({ apps: firstApps, machine: firstMachine, machineCrashed, machineHung, machineScope: wanted, again: againFiles, cleared: clearedApps, narrowed });
      for (const [app, v] of [...Object.entries(firstApps), ...Object.entries(firstMachine || {})]) {
        for (const p of Object.values(v?.parts || {})) for (const n of p.passedNames || []) windowSeen.push({ name: n, app });
        for (const n of v?.passedNames || []) windowSeen.push({ name: n, app });
      }

      const appsOk = wanted.every((a) => state.apps[a]?.ok) && state.machine?.ok;
      if (!appsOk) failed = true;
      const all = [...Object.values(state.apps), state.machine].filter(Boolean);
      results.push({ name: 'verify:apps', what: 'each app opens, changes and saves', ok: appsOk, seconds: secs(Date.now() - started) });
      say(VERIFY_APPS.filter((a) => state.apps[a]).map((a) => `  ${state.apps[a].ok ? 'ok  ' : 'FAIL'} ${a.padEnd(10)} ${state.apps[a].passed}/${state.apps[a].of}${state.apps[a].seconds ? `  ${state.apps[a].seconds} s` : ''}`).join('\n'));
      if (state.machine) say(`  ${state.machine.ok ? 'ok  ' : 'FAIL'} ${'machine'.padEnd(10)} ${state.machine.passed}/${state.machine.of}${state.machine.seconds ? `  ${state.machine.seconds} s` : ''}`);
      if (all.length) counts['verify:apps'] = { passed: all.reduce((n, v) => n + v.passed, 0), of: all.reduce((n, v) => n + v.of, 0) };
      // What the waits that gave up cost this run, by check file (verify-waits.js says each one in its lane's log).
      const logsNow = [...laneLogs].map((f) => { try { return fs.readFileSync(f, 'utf8'); } catch { return ''; } });
      say(timeLostLines(timeLost(logsNow)).join('\n'));
    }
    fs.writeFileSync(TIMING, JSON.stringify(timing, null, 2));
    if (!serial) writeSeconds(SECONDS, measured);
  }

  // The engine passes' verdict: the tests here and on the lab machine, and its fuzz passes, each with the run again alone.
  if (vmJob && !vmSettled) {
    // Said, not waited for in silence: the runner's last line says whether it waits in the machine's queue.
    const last = (() => { try { return fs.readdirSync(LOGS).filter((f) => /^vm-.*\.log$/.test(f) && !/-(test|fuzz-[\w-]+)\.log$/.test(f)).map((f) => path.join(LOGS, f)).sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs)[0]; } catch { return null; } })();
    const line = last ? String(fs.readFileSync(last, 'utf8')).trim().split(/\r?\n/).pop() : '';
    say(`  waiting for ${vm.host}${line ? `: ${line.slice(0, 160)}` : ''} (each run there is ended after ${VM_MINUTES} minutes)`);
  }
  const vmDone = vmJob ? await vmJob : null;
  const testDone = { ...localDone };
  if (vmDone?.first.passes.test) testDone.vm = { first: vmDone.first.passes.test, again: vmDone.again?.passes.test || null, os: vmDone.first.result?.os, osName: vmDone.first.result?.osName, cpus: vmDone.first.result?.cpus };
  if (testGroups.length) engineVerdict(testDone);
  for (const name of vmFuzz) {
    const [, , what] = FUZZERS.find(([n]) => n === name);
    const first = vmDone.first.passes[name];
    const again = vmDone.again?.passes[name] || null;
    const ranOk = (p) => p && p.code === 0 && p.counted?.of > 0;
    const ok = ranOk(first) || ranOk(again);
    if (first.code === 0 && !ranOk(first)) say(`  [${name}] ended without saying what it counted on ${vm.host}: it did not finish`);
    if (!ranOk(first)) say(`  [${name}] again: ${ok ? 'passed alone' : 'still failing'}`);
    state.passes[name] = { ok, counts: first.counted || null, host: vm.host, os: vmDone.first.result?.os || null, osName: vmDone.first.result?.osName || null, cpus: vmDone.first.result?.cpus || null, reran: !ranOk(first) && ok ? [name] : [] };
    if (first.counted) counts[name] = first.counted;
    fuzzRuns.push({ name, host: vm.host, ok: ranOk(first), counted: first.counted, again: again ? { ok: ranOk(again) } : null });
    results.push({ name, what: `${what}, on ${vm.host}`, ok, seconds: String(first.seconds), log: first.log });
    if (!ok) failed = true;
  }
  // A fuzz pass that passed on the lab machine in the run before on this commit: its counts and its line, as pass() keeps them (review of 11 Nov).
  if (vm.on) {
    for (const [name, , what] of FUZZERS) {
      if (!want(name) || vmFuzz.includes(name) || !(resume && state.passes[name]?.ok)) continue;
      say(`${name}: passed in the run before on this commit, on ${state.passes[name].host || vm.host}.`);
      results.push({ name, what: `${what}, on ${state.passes[name].host || vm.host}`, ok: true, seconds: '0.0', resumed: true });
      if (state.passes[name].counts) counts[name] = state.passes[name].counts;
    }
  }
  save();

  // 6. Every window painted, once the rest has passed.
  if (want('smoke') && !failed) await pass('smoke', 'every window paints', 'npm', ['run', 'smoke']);

  // The app packs, added and removed across two starts.
  if (want('verify:packs') && !failed) await pass('verify:packs', 'app packs add and remove', 'node', [path.join('apps', 'desktop', 'build', 'verify-packs-run.js')], { log: path.join(LOGS, 'verify-packs.log') });

  // What the suite saves, opened in Word, Excel and PowerPoint themselves: they must find what the original holds, and the edit.
  if (want('office:reopen') && !failed) await pass('office:reopen', 'Word, Excel and PowerPoint reopen our saves', 'powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join('tools', 'check-strict-office.ps1'), '-Dir', path.join(STATE_DIR, 'office-saves')], { log: path.join(LOGS, 'office-reopen.log') });

  // 7. Files nobody made on purpose: damaged ones, then real ones from outside the project.
  // With the lab machine they ran there, from the start; here they run here.
  for (const [name, tool, what, rounds] of FUZZERS) {
    if (vm.on || !want(name) || failed) continue;
    const before = resume && state.passes[name]?.ok;
    await pass(name, what, 'node', [tool, String(rounds), seed], { log: path.join(LOGS, `${name.replace(':', '-')}.log`) });
    if (!before) fuzzRuns.push({ name, host: 'local', ok: Boolean(state.passes[name]?.ok), counted: state.passes[name]?.counts || null, again: null });
  }
  if (want('corpus') && !failed) {
    const corpus = path.join(ROOT, '.corpus');
    if (fs.existsSync(corpus)) await pass('corpus', 'real files, opened and saved', 'node', ['tools/corpus.js'], { log: path.join(LOGS, 'corpus.log'), env: { RUTBA_CORPUS_RATES: path.join(ROOT, 'tools', 'corpus-rates.json'), RUTBA_CORPUS: corpus, RUTBA_CORPUS_OUT: path.join(STATE_DIR, 'corpus') } });
    else { console.log('corpus: there is no .corpus folder here (tools/downloads.js fetches it), so this gate cannot pass.'); failed = true; }
  }

  // The failure list: every check that failed in this run, window or engine,
  // with its number, where it was judged and the wait behind it, kept until
  // closed. --serial keeps none.
  if (!serial && (windowFindings || engineRuns.length || fuzzRuns.length)) {
    try {
      const ledger = readLedger(LEDGER);
      const findings = mergeFindings(windowFindings || { failures: [], passed: new Set(), clean: new Set() }, engineFindings({ tests: engineRuns, fuzz: fuzzRuns, wholeSuite: !someFiles, home: homeOf, fuzzHome: vm.runner ? vm.host : 'local' }));
      // Numbers for the checks seen for the first time, so the next run says them: real checks only, not those the run says of itself.
      // The engine's are numbered only when they fail (app 00): its thousands of passing tests need none.
      const registry = readRegistry();
      if (registry.apps.engine == null) registry.apps.engine = 0;
      const seenNames = [...windowSeen];
      for (const f of findings.failures) if (!f.synthetic && (VERIFY_APPS.includes(f.app) || f.app === 'engine')) seenNames.push({ name: f.name, app: f.app });
      const added = assignIds(registry, seenNames, VERIFY_APPS);
      // Numbers are given on the Windows computer only (PM, 11 Nov 2026): a copy of the
      // registry on the lab machine would hand out numbers this one never sees, so there a
      // check with no number stops the run, loudly, and is numbered by a run here.
      if (added.length && process.platform !== 'win32') {
        gateFault(`${added.length} check${added.length === 1 ? ' has' : 's have'} no number, and numbers are given only on the Windows computer: run ${added.length === 1 ? 'it' : 'them'} there first (${added.slice(0, 5).join('; ')})`);
        for (const name of added) delete registry.ids[name];
      } else if (added.length) {
        writeRegistry(registry);
        say(`  ${added.length} check${added.length === 1 ? '' : 's'} seen for the first time ${added.length === 1 ? 'was' : 'were'} given ${added.length === 1 ? 'a number' : 'numbers'} (tools/check-ids.json).`);
      }
      for (const f of findings.failures) if (!f.synthetic && registry.ids[f.name]) f.id = registry.ids[f.name];
      // A pass on a tree with changes beyond the commit is no proof for closing an entry by that commit.
      const version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
      const cleanTree = dirtyForRecord(porcelain(), version).length === 0;
      recordRun(ledger, { head, clean: cleanTree, label: record ? 'record' : partial ? `partial: ${askedIds.length ? askedIds.join(',') : only.length ? only.join(',') : wanted.join(',')}` : 'full', scope: wanted }, findings);
      writeLedger(ledger, LEDGER);
      fs.writeFileSync(FAILURES_PAGE, ledgerMarkdown(ledger, { head }));
      const open = openEntries(ledger);
      if (findings.failures.length || open.length) {
        say(`\n\u001b[1m── the failure list: ${findings.failures.length} failed in this run, ${open.length} open in all ──\u001b[0m`);
        for (const line of ledgerLines(open)) say(line);
        say(`  The list: ${path.relative(ROOT, FAILURES_PAGE)}. Work through all of it before a full run; --record is refused while any is open.`);
      }
    } catch (err) {
      gateFault(`the failure list was not written: ${err.message}`);
    }
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
  // Said with --quiet too, and whatever passes ran: the conditions the budgets were measured under (review of 13 Nov).
  if (raisedSoFar()) console.log(`\n${raisedSoFar()} of the gate's own processes (itself, what it started, and what those started) were raised to above normal priority, beside whatever else the computer was doing (RUTBA_GATE_PRIORITY=normal for none).`);
  console.log(
    failed
      ? `\n\u001b[31mThe gate is closed.\u001b[0m Fix what failed and run it again, or with --resume to check again only what failed.`
      : `\n\u001b[32mThe gate is open.\u001b[0m ${results.length} passes${partial ? `, for ${only.length ? only.join(', ') : wanted.join(', ')} only` : ''}.`
  );

  // On the record only when every pass and every chunk passed on this commit, in one run or resumed;
  // a patch release with --changed, every pass and the chunks it changed.
  const version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
  const scoped = changed && !only.length && !appsArg && isPatch(version) && wanted.length < VERIFY_APPS.length;
  const whole = (!partial || scoped) && ['test', 'test:perf', 'test:removal', 'verify:edit', 'smoke', 'verify:packs', 'office:reopen', ...FUZZERS.map(([name]) => name), 'corpus'].every((name) => state.passes[name]?.ok)
    && (scoped ? wanted : VERIFY_APPS).every((a) => state.apps[a]?.ok) && (!wanted.length || machineCovers(state.machine, scoped ? wanted : VERIFY_APPS));
  if (record) {
    const dirty = dirtyForRecord(porcelain(), version);
    // What this run itself added to the failure list (a check that failed and then passed alone) keeps it off the record too.
    let openNow = [];
    try { openNow = openEntries(readLedger(LEDGER)); } catch (err) { openNow = [{ id: 'the list', name: err.message }]; }
    if (!failed && whole && openNow.length) {
      console.log(`Nothing recorded: the failure list has ${openNow.length} open ${openNow.length === 1 ? 'entry' : 'entries'} from this run (${openNow.slice(0, 5).map((e) => e.id).join(', ')}). Close each at its cause first.`);
    } else if (!failed && whole && (git('rev-parse', 'HEAD') !== head || dirty.length)) {
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
        const ledgerNow = readLedger(LEDGER);
        // Each engine group and fuzz pass by the machine that ran it, as the state kept it.
        const hostGroups = [
          ...Object.values(state.passes.test?.groups || {}).map((g) => ({ pass: 'test', host: g.host, osName: g.osName, cpus: g.cpus, tests: g.counts?.tests })),
          ...FUZZERS.map(([name]) => ({ pass: name, host: state.passes[name]?.host || 'local', osName: state.passes[name]?.osName, cpus: state.passes[name]?.cpus })),
        ];
        reran.push(...Object.values(state.passes.test?.groups || {}).flatMap((g) => g.reran || []), ...FUZZERS.flatMap(([name]) => state.passes[name]?.reran || []));
        fs.writeFileSync(notePath, withGateSection(fs.readFileSync(notePath, 'utf8'), gateSection(counts, { version, date, reran, apps: scoped ? wanted : null, commit: head, fewer: state.fewer || [], closedByReason: reasonClosures(ledgerNow), hosts: hostsSentence(hostGroups), unsettled: state.unsettled || null, priority: state.priority || null })));
        writeLedger(markRecorded(ledgerNow, version), LEDGER);
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
