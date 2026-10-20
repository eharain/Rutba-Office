// Everything, in the order that finds problems soonest.
//
//   npm run gate
//   npm run gate -- --record     and write what it counted into the release
//                                note and the README (a full run that passed)
//   npm run gate -- --apps word,mail   the window checks of those apps only
//   npm run gate -- --changed    the window checks of the apps changed since
//                                the last release tag (all of them when what
//                                every app shares has changed)
//   npm run gate -- --resume     carry on from the last run on this commit:
//                                the passes and the apps that passed are not
//                                run again, and --record adds the runs up
//
// Five passes, each answering a question the one before it cannot:
//
//   test          do the engines do what they say? no windows.
//   test:perf     does a large file still open, recalculate and save in
//                 seconds? run alone, after the suite: a budget timed while
//                 the rest of the suite shares the cores measures the machine
//                 (on a four-core laptop a recalculation of 13 s of its own
//                 CPU took 20 to 23 s beside the suite).
//   verify:edit   do keystrokes reach the document? two real windows.
//   verify:apps   does each app open, change and save a real file? every window.
//   smoke         does every window paint without logging anything?
//
// Unit tests first because they take seconds and catch most of it. The window
// runs are slower, need a display, and are the only ones that can prove the
// thing a person actually came for.
//
// Nothing here is optional before a release. A suite that cannot save is not a
// suite, and the only way to know it still saves is to have watched it. With
// --record the gate puts its own counts on the record: a Gate section in
// docs/releases/v<version>.md saying where and when it ran and what each
// pass counted, and the README's counts brought to the same numbers, so a
// release says the gate ran rather than leaving it to be believed.

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
export const VERIFY_APPS = ['shell', 'word', 'sheets', 'slides', 'pictures', 'image', 'video', 'recorder', 'pdf', 'mail', 'calendar', 'contacts'];

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
  if (check) return [{ word: 'word', sheet: 'sheets', deck: 'slides', slide: 'slides', mail: 'mail', video: 'video', recorder: 'recorder', pdf: 'pdf', picture: 'pictures', pictures: 'pictures', image: 'image', calendar: 'calendar', contacts: 'contacts' }[check] || 'shell'];
  const service = /^apps\/desktop\/main\/([a-z]+)/.exec(f)?.[1];
  if (service) return [{ mail: 'mail', live: 'recorder', pdf: 'pdf', calendar: 'calendar', contacts: 'contacts', dav: 'calendar' }[service] || 'shell'];
  const pkg = /^packages\/([a-z-]+)\//.exec(f)?.[1];
  if (pkg) {
    return {
      ooxml: ['word', 'sheets', 'slides'], 'doc-view': ['word'], 'sheet-view': ['sheets'], formula: ['sheets'], presentation: ['slides'],
      mailbox: ['mail'], calendar: ['calendar', 'mail'], contacts: ['contacts', 'mail'], pdf: ['pdf', 'word', 'sheets', 'slides'],
      imaging: ['image', 'pictures'], media: ['video', 'recorder'], stream: ['recorder'], studio: ['video', 'recorder'], ocr: ['pdf', 'pictures', 'sheets'],
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
const PARTS = path.join(STATE_DIR, 'parts.json');
/** What the last run on this commit, with this working tree, left to do. */
function readState(head, tree) {
  try {
    const s = JSON.parse(fs.readFileSync(STATE, 'utf8'));
    return s.head === head && s.tree === tree ? s : null;
  } catch {
    return null;
  }
}

const PASSES = [
  ['test', 'the engines', 'npm', ['test']],
  ['test:perf', 'the performance budgets, run alone', 'npm', ['run', 'test:perf']],
  ['verify:edit', 'typing reaches the document', 'npm', ['run', 'verify:edit']],
  ['verify:apps', 'each app opens, changes and saves', 'npm', ['run', 'verify:apps']],
  ['smoke', 'every window paints', 'npm', ['run', 'smoke']],
];

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
export function gateSection(counts, { version, date, platform = process.platform, arch = process.arch }) {
  const t = engineCounts(counts);
  const parts = [
    t ? `${n(t.tests)} engine tests, ${n(t.passed)} passing${t.skipped ? ` and ${n(t.skipped)} skipped where a package they need is not installed` : ''}` : null,
    counts['verify:edit'] ? `${n(counts['verify:edit'].passed)} of ${n(counts['verify:edit'].of)} editing checks` : null,
    counts['verify:apps'] ? `${n(counts['verify:apps'].passed)} of ${n(counts['verify:apps'].of)} application checks` : null,
    counts.smoke ? `${n(counts.smoke.passed)} of ${n(counts.smoke.of)} windows painted clean` : null,
  ].filter(Boolean);
  const list = parts.length > 1 ? `${parts.slice(0, -1).join('; ')}; and ${parts.at(-1)}` : parts[0];
  return `## Gate\n\n${wrap(`Run on ${PLATFORM[platform] || platform} (${arch}) on ${date}, before ${version} was released: ${list}.`)}\n`;
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
const only = argv.filter((a, i) => !a.startsWith('-') && argv[i - 1] !== '--apps');
const quiet = argv.includes('--quiet');
const record = argv.includes('--record');
const resume = argv.includes('--resume');
const changed = argv.includes('--changed');

function run(command, args, env = {}) {
  return new Promise((resolve) => {
    const child = spawn(command, args, { stdio: ['ignore', 'pipe', 'pipe'], shell: process.platform === 'win32', env: { ...process.env, ...env } });
    // Shown as it comes (unless quiet), and the end of it kept: the counts
    // are in the last lines, and a quiet run shows the tail of a failure.
    let tail = '';
    const keep = (to) => (chunk) => {
      if (!quiet) to.write(chunk);
      tail = (tail + chunk).slice(-200000);
    };
    child.stdout.on('data', keep(process.stdout));
    child.stderr.on('data', keep(process.stderr));
    child.on('exit', (code) => resolve({ code: code ?? 1, tail }));
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
 * The budgets are run alone, and alone means the machine has settled: the
 * engine tests before them write thousands of files that the virus scanner
 * goes on reading for a while, which held the main thread past its budget in
 * two of three gates for 2.2.0, where the same budgets passed on their own
 * every time. So before them the gate waits, up to two minutes, until the
 * processors have been under a quarter busy for three seconds running, and
 * says how long it waited and how busy the machine still was.
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
    if (!quiet) console.log(`Changed since ${since}: ${files.length} files; the window checks of ${wanted.join(', ') || 'no app'}.`);
  }
  const partial = wanted.length < VERIFY_APPS.length || only.length > 0;

  // The record of runs on this commit: started afresh unless resuming.
  const head = git('rev-parse', 'HEAD');
  const tree = git('status', '--porcelain', '--untracked-files=no');
  fs.mkdirSync(STATE_DIR, { recursive: true });
  const state = (resume && readState(head, tree)) || { head, tree, passes: {}, apps: {} };
  const save = () => fs.writeFileSync(STATE, JSON.stringify(state, null, 2));
  if (resume && !state.passes.test && !Object.keys(state.apps).length && !quiet) console.log('Nothing to resume on this commit: a run from the start.');

  for (const [name, what, command, args] of PASSES) {
    if (only.length && !only.includes(name)) continue;
    let env = {};
    if (name === 'verify:apps') {
      const left = wanted.filter((a) => !(resume && state.apps[a]?.ok));
      if (!left.length) {
        if (!quiet) console.log(`\nverify:apps: every app asked for passed in the run before (${wanted.join(', ')}).`);
        results.push({ name, what, ok: true, seconds: '0.0', resumed: true });
        continue;
      }
      try { fs.rmSync(PARTS, { force: true }); } catch { /* none */ }
      env = { RUTBA_VERIFY_PARTS: PARTS, ...(left.length < VERIFY_APPS.length ? { RUTBA_VERIFY_APPS: left.join(',') } : {}) };
      if (!quiet && left.length < VERIFY_APPS.length) console.log(`\nverify:apps for ${left.join(', ')}${resume ? ' (the rest passed in the run before)' : ''}.`);
    } else if (resume && state.passes[name]?.ok) {
      if (!quiet) console.log(`\n${name}: passed in the run before on this commit.`);
      results.push({ name, what, ok: true, seconds: '0.0', resumed: true });
      if (state.passes[name].counts) counts[name] = state.passes[name].counts;
      continue;
    }
    if (name === 'test:perf') await settle();

    const started = Date.now();
    if (!quiet) console.log(`\n\u001b[1m── ${name} — ${what} ─────────────────────────────\u001b[0m\n`);
    const out = await run(command, args, env);
    const { tail } = out;
    const counted = countsOf(name, tail);
    // A pass that counts and ended without its count did not finish, whatever it exited with.
    const code = out.code === 0 && name !== 'test' && name !== 'test:perf' && !counted ? 1 : out.code;
    if (code !== out.code && !quiet) console.log(`\n${name} ended without saying what it counted: it did not finish.`);
    const seconds = ((Date.now() - started) / 1000).toFixed(1);
    results.push({ name, what, ok: code === 0, seconds });
    if (counted) counts[name] = counted;
    if (name === 'verify:apps') {
      // Each app's own counts; an app with a check that failed is run again on --resume.
      let parts = null;
      try { parts = JSON.parse(fs.readFileSync(PARTS, 'utf8')).apps; } catch { /* the run ended before writing them */ }
      const asked = env.RUTBA_VERIFY_APPS ? env.RUTBA_VERIFY_APPS.split(',') : VERIFY_APPS;
      if (parts) for (const a of asked) {
        const p = parts[a] || { passed: 0, of: 0, failed: [] };
        state.apps[a] = { ok: p.failed.length === 0, passed: p.passed, of: p.of, failed: p.failed };
      }
      const apps = Object.entries(state.apps);
      if (apps.length) counts[name] = { passed: apps.reduce((n, [, v]) => n + v.passed, 0), of: apps.reduce((n, [, v]) => n + v.of, 0) };
      if (!quiet && parts) console.log(VERIFY_APPS.filter((a) => state.apps[a]).map((a) => `  ${state.apps[a].ok ? 'ok  ' : 'FAIL'} ${a.padEnd(9)} ${state.apps[a].passed}/${state.apps[a].of}`).join('\n'));
    } else {
      state.passes[name] = { ok: code === 0, counts: counted || null };
    }
    save();

    if (code !== 0) {
      failed = true;
      if (quiet && tail) console.log(tail.slice(-4000));
      // The later passes are slower and mostly re-report the same fault, so a
      // failure stops here rather than spending four minutes confirming it.
      break;
    }
  }

  console.log('\n────────────────────────────────────────────────');
  for (const r of results) {
    console.log(`${r.ok ? '\u001b[32mok  \u001b[0m' : '\u001b[31mFAIL\u001b[0m'} ${r.name.padEnd(13)} ${r.what.padEnd(38)} ${r.seconds}s`);
  }
  const total = results.reduce((s, r) => s + Number(r.seconds), 0).toFixed(1);
  console.log(
    failed
      ? `\n\u001b[31mThe gate is closed.\u001b[0m Fix the pass above and run it again${Object.values(state.apps).some((a) => a.ok) ? ', or with --resume to check again only what failed' : ''}.`
      : `\n\u001b[32mThe gate is open.\u001b[0m ${results.length} passes in ${total}s${partial ? `, for ${only.length ? only.join(', ') : wanted.join(', ')} only` : ''}.`
  );

  // On the record only when every pass and every app passed on this commit, in one run or resumed.
  const whole = !partial && PASSES.every(([name]) => name === 'verify:apps' || state.passes[name]?.ok) && VERIFY_APPS.every((a) => state.apps[a]?.ok);
  if (record) {
    if (failed || !whole) {
      console.log('Nothing recorded: only a full run that passed is put on the record.');
    } else {
      const version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
      const date = localDate();
      const notePath = path.join(ROOT, 'docs', 'releases', `v${version}.md`);
      if (fs.existsSync(notePath)) {
        fs.writeFileSync(notePath, withGateSection(fs.readFileSync(notePath, 'utf8'), gateSection(counts, { version, date })));
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
