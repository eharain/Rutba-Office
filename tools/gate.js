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
export const VERIFY_APPS = ['shell', 'ribbon', 'look', 'theme', 'access', 'languages', 'word', 'sheets', 'slides', 'pictures', 'image', 'video', 'recorder', 'talk', 'base', 'pdf', 'mail', 'calendar', 'contacts'];

/** The apps a changed file touches; 'shell' means every app. */
export function appsForPath(file) {
  const f = String(file).replace(/\\/g, '/');
  if (/^(docs|tools|\.github)\/|^[^/]+\.md$|^tests\//.test(f)) return [];
  const app = /^apps\/desktop\/renderer\/apps\/([a-z]+)/.exec(f)?.[1];
  if (app) {
    const key = { home: 'shell', edit: 'shell' }[app] || app;
    return VERIFY_APPS.includes(key) ? [key] : ['shell'];
  }
  const check = /^apps\/desktop\/main\/verify-([a-z]+)/.exec(f)?.[1];
  if (check) return [{ word: 'word', sheet: 'sheets', deck: 'slides', slide: 'slides', mail: 'mail', video: 'video', recorder: 'recorder', talk: 'talk', base: 'base', pdf: 'pdf', picture: 'pictures', pictures: 'pictures', image: 'image', calendar: 'calendar', contacts: 'contacts', ribbon: 'ribbon', polish: 'theme', dark: 'theme', gallery: 'look', soon: 'look', panels: 'look', icons: 'look', a11y: 'access', keyboard: 'access', high: 'access', screen: 'access', dialog: 'access', lost: 'access', accessibility: 'access', urdu: 'languages', spelling: 'languages' }[check] || 'shell'];
  const service = /^apps\/desktop\/main\/([a-z]+)/.exec(f)?.[1];
  if (service) return [{ mail: 'mail', live: 'recorder', talk: 'talk', base: 'base', pdf: 'pdf', calendar: 'calendar', contacts: 'contacts', dav: 'calendar' }[service] || 'shell'];
  const pkg = /^packages\/([a-z-]+)\//.exec(f)?.[1];
  if (pkg) {
    return {
      ooxml: ['word', 'sheets', 'slides'], 'doc-view': ['word'], 'sheet-view': ['sheets'], formula: ['sheets'], presentation: ['slides'],
      mailbox: ['mail'], calendar: ['calendar', 'mail'], contacts: ['contacts', 'mail'], pdf: ['pdf', 'word', 'sheets', 'slides'],
      imaging: ['image', 'pictures'], media: ['video', 'recorder'], stream: ['recorder'], studio: ['video', 'recorder'], database: ['base'], ocr: ['pdf', 'pictures', 'sheets'],
    }[pkg] || ['shell'];
  }
  return ['shell'];
}

/** The apps to check for these changed files: every app when the shared part changed. */
export function appsForChanges(files) {
  const set = new Set(files.flatMap(appsForPath));
  return set.has('shell') ? VERIFY_APPS.slice() : VERIFY_APPS.filter((a) => set.has(a));
}

const git = (...args) => { try { return execFileSync('git', args, { cwd: ROOT, encoding: 'utf8' }).trim(); } catch { return ''; } };
const STATE_DIR = path.join(ROOT, '.gate');
const STATE = path.join(STATE_DIR, 'state.json');
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
  return null;
}

const n = (v) => Number(v).toLocaleString('en-GB');
const PLATFORM = { win32: 'Windows', darwin: 'macOS', linux: 'Linux' };

/** The release note's Gate section, in the note's own plain style. */
export function gateSection(counts, { version, date, platform = process.platform, arch = process.arch, reran = [], apps = null }) {
  const t = engineCounts(counts);
  const parts = [
    t ? `${n(t.tests)} engine tests, ${n(t.passed)} passing${t.skipped ? ` and ${n(t.skipped)} skipped where a package they need is not installed` : ''}` : null,
    counts['verify:edit'] ? `${n(counts['verify:edit'].passed)} of ${n(counts['verify:edit'].of)} editing checks` : null,
    counts['verify:apps'] ? `${n(counts['verify:apps'].passed)} of ${n(counts['verify:apps'].of)} application checks${apps ? `, of ${listed(apps)}, the ${apps.length === 1 ? 'app' : 'apps'} it changed` : ''}` : null,
    counts.smoke ? `${n(counts.smoke.passed)} of ${n(counts.smoke.of)} windows painted clean` : null,
  ].filter(Boolean);
  const list = parts.length > 1 ? `${parts.slice(0, -1).join('; ')}; and ${parts.at(-1)}` : parts[0];
  const again = reran.length
    ? `\n\n${wrap(`${reran.length === 1 ? 'One check' : `${n(reran.length)} checks`} failed beside the others and passed when run again alone at the end of the same gate:`)}\n\n${reran.map((r) => `- ${r}`).join('\n')}`
    : '';
  return `## Gate\n\n${wrap(`Run on ${PLATFORM[platform] || platform} (${arch}) on ${date}, before ${version} was released: ${list}.`)}${again}\n`;
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
 * So the gate waits, up to two minutes, until the processors have been under
 * a quarter busy for three seconds running.
 */
async function settle(limitMs = 120_000) {
  const started = Date.now();
  let calm = 0;
  let share = 1;
  while (Date.now() - started < limitMs) {
    share = await busyShare(1000);
    calm = share < 0.25 ? calm + 1 : 0;
    if (calm >= 3) break;
  }
  const waited = ((Date.now() - started) / 1000).toFixed(0);
  if (!quiet) console.log(`The machine settled for the budgets: waited ${waited} s, the processors ${Math.round(share * 100)}% busy.`);
}

/* ── the lanes ─────────────────────────────────────────────────────── */

const TIMING = path.join(STATE_DIR, 'timing.json');
const LOGS = path.join(STATE_DIR, 'logs');
const PARTS_DIR = path.join(STATE_DIR, 'parts');
/** A chunk's length the first time, before a run has measured it, in seconds. */
const FIRST_GUESS = { shell: 900, ribbon: 550, look: 400, theme: 400, word: 420, sheets: 400, slides: 350, access: 300, mail: 300, video: 300, languages: 200, pdf: 150, image: 150, recorder: 90, pictures: 60, talk: 40, base: 30, calendar: 40, contacts: 20 };

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
 * lane's counts: { parallel: { apps, ids, names }, machine: { ... } }. A
 * failure outside any part ('*') runs its app's chunk whole.
 */
export function partsToRerun(apps, machine) {
  const pick = (entries) => {
    const out = { apps: [], ids: [], names: [], whole: [] };
    for (const [app, v] of entries) {
      if (!v || v.ok || !v.failed?.length) continue;
      out.apps.push(app);
      out.names.push(...v.failed);
      const ids = v.failedParts || ['*'];
      if (ids.includes('*')) out.whole.push(app);
      else out.ids.push(...ids);
    }
    return out;
  };
  const machineApps = machine && !machine.ok ? Object.entries(machine.byApp || { machine: machine }) : [];
  return { parallel: pick(Object.entries(apps)), machine: pick(machineApps) };
}

/** Whether the run again cleared `names`: each one ran and passed, and nothing in it failed. */
export function rerunCleared(names, parts) {
  const all = Object.values(parts || {});
  const passed = new Set(all.flatMap((p) => p.passedNames || []));
  return all.every((p) => !p.failed?.length) && names.every((n) => passed.has(n));
}

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

  // Which apps' window checks this run is for: those named, those changed, or all.
  let wanted = VERIFY_APPS.slice();
  if (appsArg) wanted = appsArg.split(',').map((a) => a.trim()).filter((a) => VERIFY_APPS.includes(a));
  else if (changed) {
    const since = git('describe', '--tags', '--abbrev=0');
    const files = [...git('diff', '--name-only', `${since}..HEAD`).split('\n'), ...git('diff', '--name-only').split('\n')].filter(Boolean);
    wanted = appsForChanges(files);
    say(`Changed since ${since}: ${files.length} files; the window checks of ${wanted.join(', ') || 'no app'}.`);
  }
  const partial = wanted.length < VERIFY_APPS.length || only.length > 0;
  const want = (name) => !only.length || only.includes(name);

  // The record of runs on this commit: started afresh unless resuming.
  const head = git('rev-parse', 'HEAD');
  const tree = git('status', '--porcelain', '--untracked-files=no');
  for (const d of [STATE_DIR, LOGS, PARTS_DIR]) fs.mkdirSync(d, { recursive: true });
  const state = (resume && readState(head, tree)) || { head, tree, passes: {}, apps: {}, machine: null };
  state.apps ||= {};
  const save = () => fs.writeFileSync(STATE, JSON.stringify(state, null, 2));
  if (resume && !state.passes.test && !Object.keys(state.apps).length) say('Nothing to resume on this commit: a run from the start.');
  const timing = readJson(TIMING) || {};
  const jobs = Math.max(1, Number(jobsArg) || laneCap());

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
    // A pass that counts and ended without its count did not finish, whatever it exited with.
    const code = out.code === 0 && name !== 'test' && name !== 'test:perf' && !counted ? 1 : out.code;
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
  const windowRun = async (label, env) => {
    const partsFile = path.join(PARTS_DIR, `${label}.json`);
    try { fs.rmSync(partsFile, { force: true }); } catch { /* none */ }
    const log = path.join(LOGS, `${label}.log`);
    const started = Date.now();
    say(`  [${label}] started`);
    const out = await run('node', [path.join('apps', 'desktop', 'build', 'verify-apps-run.js')], { ...env, RUTBA_VERIFY_PARTS: partsFile }, { log });
    const seconds = (Date.now() - started) / 1000;
    const parts = readJson(partsFile)?.apps || null;
    const sum = parts ? Object.values(parts).reduce((t, p) => ({ passed: t.passed + p.passed, of: t.of + p.of, failed: [...t.failed, ...p.failed] }), { passed: 0, of: 0, failed: [] }) : null;
    const ok = out.code === 0 && Boolean(sum) && sum.failed.length === 0;
    say(`  [${label}] ${ok ? 'ok  ' : 'FAIL'} ${sum ? `${sum.passed}/${sum.of}` : 'ended without its count'} in ${seconds.toFixed(0)} s${ok ? '' : ` (${path.relative(ROOT, log)})`}`);
    if (sum && !ok) for (const f of sum.failed.slice(0, 8)) say(`         FAIL ${f}`);
    return { ok, seconds, sum: sum || { passed: 0, of: 0, failed: ['the run ended before it counted'] }, parts };
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
    if (!(resume && state.passes['test:perf']?.ok)) await settle();
    await pass('test:perf', 'the performance budgets, run alone', 'npm', ['run', 'test:perf']);
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
      state.machine = { ok: true, passed: 0, of: 0, failed: [] };
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
          state.apps[c] = { ok: r.ok, passed: p.passed, of: p.of, failed: p.failed, seconds: Math.round(r.seconds) };
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
      if (resume && state.machine?.ok) say('  [machine] passed in the run before on this commit');
      else {
        say(`\n\u001b[1m── machine: the computer's own clipboard and keys, alone ──\u001b[0m`);
        const r = await windowRun('machine', { RUTBA_VERIFY_LANE: 'machine', ...(wanted.length < VERIFY_APPS.length ? { RUTBA_VERIFY_APPS: wanted.join(',') } : {}) });
        state.machine = { ok: r.ok, passed: r.sum.passed, of: r.sum.of, failed: r.sum.failed, seconds: Math.round(r.seconds), byApp: r.parts || null };
        save();
      }

      // 5. Again, alone: the parts whose checks failed, in a process of their
      // own after the rest, the machine perhaps quieter by now. The gate
      // passes only if each check that failed passes here.
      const again = partsToRerun(Object.fromEntries(wanted.map((a) => [a, state.apps[a]])), state.machine);
      for (const [laneName, set] of [['parallel', again.parallel], ['machine', again.machine]]) {
        if (!set.names.length) continue;
        say(`\n\u001b[1m── again, alone: ${set.names.length} check${set.names.length === 1 ? '' : 's'} that failed ${laneName === 'machine' ? 'in the machine lane' : 'beside the others'} ──\u001b[0m`);
        for (const n of set.names) say(`         ${n}`);
        await settle();
        const env = laneName === 'machine'
          ? { RUTBA_VERIFY_LANE: 'machine' }
          : { RUTBA_VERIFY_LANE: 'parallel', RUTBA_VERIFY_APPS: set.apps.join(',') };
        // A chunk with a failure outside its parts runs whole; the rest only the parts that failed.
        if (!set.whole.length) env.RUTBA_VERIFY_PART_IDS = set.ids.join(',');
        const r = await windowRun(laneName === 'machine' ? 'again-machine' : 'again', env);
        const cleared = rerunCleared(set.names, r.parts);
        say(`  again: ${cleared ? 'each passed alone' : 'still failing'}`);
        if (!cleared) continue;
        if (laneName === 'machine') {
          state.machine = { ...state.machine, ok: true, passed: state.machine.of, reran: set.names, failed: [] };
        } else {
          for (const a of set.apps) {
            const v = state.apps[a];
            state.apps[a] = { ...v, ok: true, passed: v.of, reran: v.failed, failed: [] };
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

  console.log('\n────────────────────────────────────────────────');
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
  const whole = (!partial || scoped) && ['test', 'test:perf', 'verify:edit', 'smoke'].every((name) => state.passes[name]?.ok)
    && (scoped ? wanted : VERIFY_APPS).every((a) => state.apps[a]?.ok) && (!wanted.length || state.machine?.ok);
  if (record) {
    if (failed || !whole) {
      console.log(changed && !isPatch(version)
        ? `Nothing recorded: ${version} is not a patch release, and only a full run that passed is put on its record.`
        : 'Nothing recorded: only a run that passed, in full or for the apps a patch release changed, is put on the record.');
    } else {
      const date = localDate();
      const notePath = path.join(ROOT, 'docs', 'releases', `v${version}.md`);
      if (fs.existsSync(notePath)) {
        const reran = [...VERIFY_APPS.flatMap((a) => state.apps[a]?.reran || []), ...(state.machine?.reran || [])];
        fs.writeFileSync(notePath, withGateSection(fs.readFileSync(notePath, 'utf8'), gateSection(counts, { version, date, reran, apps: scoped ? wanted : null })));
        console.log(`Recorded in ${path.relative(ROOT, notePath)}.`);
      } else {
        console.log(`No note at ${path.relative(ROOT, notePath)} to record in; write the note first.`);
      }
      const readmePath = path.join(ROOT, 'README.md');
      fs.writeFileSync(readmePath, withReadmeCounts(fs.readFileSync(readmePath, 'utf8'), counts));
      console.log("The README's counts brought up to date.");
    }
  }

  process.exit(failed ? 1 : 0);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await main();
