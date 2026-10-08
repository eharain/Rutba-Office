// Everything, in the order that finds problems soonest.
//
//   npm run gate
//   npm run gate -- --record     and write what it counted into the release
//                                note and the README (a full run that passed)
//
// Four passes, each answering a question the one before it cannot:
//
//   test          do the engines do what they say? no windows.
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
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const PASSES = [
  ['test', 'the engines', 'npm', ['test']],
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
  if (name === 'test') {
    const tests = sum(/^ℹ tests (\d+)$/gm);
    return tests ? { tests, passed: sum(/^ℹ pass (\d+)$/gm), failed: sum(/^ℹ fail (\d+)$/gm), skipped: sum(/^ℹ skipped (\d+)$/gm) } : null;
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
  const t = counts.test;
  const parts = [
    t ? `${n(t.tests)} engine tests, ${n(t.passed)} passing${t.skipped ? ` and ${n(t.skipped)} skipped where a package they need is not installed` : ''}` : null,
    counts['verify:edit'] ? `${n(counts['verify:edit'].passed)} of ${n(counts['verify:edit'].of)} editing checks` : null,
    counts['verify:apps'] ? `${n(counts['verify:apps'].passed)} of ${n(counts['verify:apps'].of)} application checks` : null,
    counts.smoke ? `${n(counts.smoke.passed)} of ${n(counts.smoke.of)} windows painted clean` : null,
  ].filter(Boolean);
  const list = parts.length > 1 ? `${parts.slice(0, -1).join('; ')}; and ${parts.at(-1)}` : parts[0];
  return `## Gate\n\nRun on ${PLATFORM[platform] || platform} (${arch}) on ${date}, before ${version} was released: ${list}.\n`;
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
  if (t) {
    out = out.replace(/(npm test\s+# )[\d,]+( tests, no network needed)/, (m, a, b) => a + n(t.tests) + b);
    out = out.replace(/(the engine suite — )[\d,]+( checks, no windows)/, (m, a, b) => a + n(t.tests) + b);
    out = out.replace(/(what runs before a release: )[\d,]+( engine tests)/, (m, a, b) => a + n(t.tests) + b);
  }
  if (counts['verify:edit']) out = out.replace(/(engine tests, )[\d,]+( editing checks)/, (m, a, b) => a + n(counts['verify:edit'].of) + b);
  if (counts['verify:apps']) out = out.replace(/(editing checks,\s+)[\d,]+( application checks)/, (m, a, b) => a + n(counts['verify:apps'].of) + b);
  if (counts.smoke) out = out.replace(/(and )[\d,]+( window captures)/, (m, a, b) => a + n(counts.smoke.of) + b);
  return out;
}

const only = process.argv.slice(2).filter((a) => !a.startsWith('-'));
const quiet = process.argv.includes('--quiet');
const record = process.argv.includes('--record');

function run(command, args) {
  return new Promise((resolve) => {
    const child = spawn(command, args, { stdio: ['ignore', 'pipe', 'pipe'], shell: process.platform === 'win32' });
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

async function main() {
  const results = [];
  const counts = {};
  let failed = false;

  for (const [name, what, command, args] of PASSES) {
    if (only.length && !only.includes(name)) continue;

    const started = Date.now();
    if (!quiet) console.log(`\n\u001b[1m── ${name} — ${what} ─────────────────────────────\u001b[0m\n`);
    const { code, tail } = await run(command, args);
    const seconds = ((Date.now() - started) / 1000).toFixed(1);
    results.push({ name, what, ok: code === 0, seconds });
    const counted = countsOf(name, tail);
    if (counted) counts[name] = counted;

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
      ? `\n\u001b[31mThe gate is closed.\u001b[0m Fix the pass above and run it again.`
      : `\n\u001b[32mThe gate is open.\u001b[0m ${results.length} passes in ${total}s.`
  );

  if (record) {
    if (failed || only.length) {
      console.log('Nothing recorded: only a full run that passed is put on the record.');
    } else {
      const version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
      const date = new Date().toISOString().slice(0, 10);
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
