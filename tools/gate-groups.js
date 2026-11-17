// The gate run unit by unit (the owner, 16 and 17 Nov 2026: "split the big
// run to smaller parallel or sequential runs, run them by modules and by
// apps not combined one long run"; "split the long running tests to
// reasonably sized smaller ones by grouping them logically").
//
//   npm run gate:groups                       every unit, one after another
//   npm run gate:groups -- --record           and, when every unit passed, the
//                                             gate's own record (--resume --record)
//   npm run gate:groups -- --only word,mail   those units alone
//   npm run gate:groups -- --from slides      that unit and those after it
//   npm run gate:groups -- --jobs 3           check processes side by side in a unit
//   npm run gate:groups -- --list             the units, and what each runs
//
// A unit is a run of the gate itself on this commit (tools/gate.js with
// --resume), limited to the passes or the apps that belong together: the
// gate adds the runs up, a unit that passed is not run again, and one that
// failed is run again alone. Nothing is checked here that the gate does not
// check, and nothing less: tests/gate-groups.test.js holds every app and
// every pass to exactly one unit. The record is the gate's, written only by
// its own --record once everything passed on this commit.
//
// It does not start while the screen is locked.
//
// Each unit logs to .gate/groups/<unit>.log, and a line for each goes to
// .gate/groups/summary.txt. Every unit runs even after one fails, so one
// round finds every failure.

import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { screenLocked } from './gate.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/**
 * The units, in the order they run: what needs a still machine first, then
 * the engines, then the windows by the code they share, then what is read
 * from outside. `passes` are the gate's passes by name; `apps` are window
 * checks (verify:apps) of those apps.
 */
export const UNITS = [
  { key: 'budgets', what: 'the performance budgets, alone', passes: ['test:perf'] },
  { key: 'removal', what: 'the folder-removal suites', passes: ['test:removal'] },
  { key: 'engine', what: 'the engine tests and the editing checks', passes: ['test', 'verify:edit'] },
  { key: 'shell', what: 'the windows every app shares: the shell', apps: ['shell'] },
  { key: 'look', what: 'the ribbon, the look and the themes', apps: ['ribbon', 'look', 'theme'] },
  { key: 'access', what: 'accessibility and the languages', apps: ['access', 'languages'] },
  { key: 'word', what: 'Word', apps: ['word'] },
  { key: 'sheets', what: 'Sheets and Data Files', apps: ['sheets', 'base'] },
  { key: 'slides', what: 'Slides', apps: ['slides'] },
  { key: 'mail', what: 'Mail', apps: ['mail'] },
  { key: 'people', what: 'Calendar, Contacts and Talk', apps: ['calendar', 'contacts', 'talk'] },
  { key: 'image', what: 'Image and Pictures', apps: ['image', 'pictures'] },
  { key: 'video', what: 'Video, the Recorder and Live Broadcast', apps: ['video', 'recorder', 'broadcast'] },
  { key: 'files', what: 'PDF and the Locker', apps: ['pdf', 'locker'] },
  { key: 'build', what: 'every window painting clean, and the app packs added and removed', passes: ['smoke', 'verify:packs'] },
  { key: 'outside', what: 'saves reopened in Microsoft Office, damaged Office files and signed mail, the corpus', passes: ['office:reopen', 'fuzz:open', 'fuzz:smime', 'corpus'] },
  { key: 'damaged-pdf', what: 'damaged PDFs', passes: ['fuzz:pdf'] },
  { key: 'harm', what: 'files made to do harm', passes: ['fuzz:grammar'] },
];

/** The gate's arguments for a unit: its passes by name, or the window checks of its apps. */
export function unitArgs(unit, { jobs = 2 } = {}) {
  const what = unit.apps ? ['verify:apps', '--apps', unit.apps.join(',')] : unit.passes.slice();
  return ['--resume', '--jobs', String(jobs), ...what];
}

/** The units a run is for: all, those named (`only`), or one and those after it (`from`). Throws on a name that is no unit's. */
export function chosen(units, { only = null, from = null } = {}) {
  const keys = units.map((u) => u.key);
  const strange = [...(only || []), ...(from ? [from] : [])].filter((k) => !keys.includes(k));
  if (strange.length) throw new Error(`No such unit: ${strange.join(', ')}. The units: ${keys.join(', ')}.`);
  let list = units;
  if (from) list = list.slice(keys.indexOf(from));
  if (only) list = list.filter((u) => only.includes(u.key));
  return list;
}

/** What a unit's log says of its run: the gate's verdict, each pass's count, and how many checks failed. */
export function readLog(text) {
  const t = String(text).replace(/\r/g, '');
  const verdict = (t.match(/The gate is (?:open|closed)/g) || []).pop() || 'no verdict';
  const counted = t.match(/\[[a-z:]+\] (?:ok {2}|FAIL) in [\d.]+ s, \d+\/\d+/g) || [];
  const fails = (t.match(/^\s*FAIL /gm) || []).length;
  const listed = (t.match(/the failure list: \d+ failed in this run, \d+ open in all/g) || []).pop() || '';
  return { verdict, counted, fails, listed };
}

/** A unit's line for the summary. */
export function summaryLine(unit, { code, seconds, log }) {
  const r = readLog(log);
  return `${unit.key.padEnd(12)} exit ${code} ${String(Math.round(seconds)).padStart(4)} s | ${r.verdict} | ${r.counted.join('; ')}${r.fails ? ` | FAIL lines ${r.fails}` : ''}${r.listed ? ` | ${r.listed}` : ''}`;
}

const gate = (args, logFile) => new Promise((resolve) => {
  const out = fs.openSync(logFile, 'w');
  const child = spawn(process.execPath, [path.join('tools', 'run-temp.js'), 'node', path.join('tools', 'gate.js'), ...args], { cwd: ROOT, stdio: ['ignore', out, out], env: process.env });
  child.on('exit', (code) => { fs.closeSync(out); resolve(code ?? 1); });
  child.on('error', () => { try { fs.closeSync(out); } catch { /* closed */ } resolve(1); });
});

async function main() {
  const argv = process.argv.slice(2);
  const valueOf = (flag) => { const i = argv.indexOf(flag); return i >= 0 && argv[i + 1] && !argv[i + 1].startsWith('-') ? argv[i + 1] : null; };
  const jobs = Number(valueOf('--jobs')) || 2;
  if (argv.includes('--list')) {
    for (const u of UNITS) console.log(`${u.key.padEnd(12)} ${u.what}\n${' '.repeat(13)}npm run gate -- ${unitArgs(u, { jobs }).join(' ')}`);
    return 0;
  }
  let units;
  try {
    units = chosen(UNITS, { only: valueOf('--only')?.split(',').map((k) => k.trim()).filter(Boolean) || null, from: valueOf('--from') });
  } catch (err) {
    console.log(err.message);
    return 2;
  }
  // Never while the screen is locked (the owner's standing rule for a gate): the windows cannot be judged then, and
  // a run that starts must be one whose every unit can be believed.
  if (screenLocked()) { console.log('The screen is locked: the gate is not run while it is. Unlock it and run this again.'); return 2; }
  const record = argv.includes('--record');
  if (record && units.length < UNITS.length) { console.log('--record is for every unit: leave out --only and --from. A unit that passed on this commit is not run again.'); return 2; }
  const dir = path.join(process.env.RUTBA_GATE_DIR ? path.resolve(process.env.RUTBA_GATE_DIR) : path.join(ROOT, '.gate'), 'groups');
  fs.mkdirSync(dir, { recursive: true });
  const summary = path.join(dir, 'summary.txt');
  const say = (line) => { console.log(line); fs.appendFileSync(summary, `${line}\n`); };
  say(`== ${new Date().toISOString()} ${units.map((u) => u.key).join(' ')}`);
  let bad = 0;
  for (const unit of units) {
    const logFile = path.join(dir, `${unit.key}.log`);
    const started = Date.now();
    const code = await gate(unitArgs(unit, { jobs }), logFile);
    say(summaryLine(unit, { code, seconds: (Date.now() - started) / 1000, log: fs.readFileSync(logFile, 'utf8') }));
    if (code !== 0) bad += 1;
  }
  if (bad) {
    say(`== ${bad} unit${bad === 1 ? '' : 's'} did not pass. Their logs are in ${path.relative(ROOT, dir)}; fix, then run those units again.`);
    return 1;
  }
  if (!record) {
    say('== every unit asked for passed on this commit.');
    return 0;
  }
  // The gate adds the units up, runs alone what needs the computer to itself, and writes its own record.
  const logFile = path.join(dir, 'record.log');
  const started = Date.now();
  const code = await gate(['--resume', '--record', '--jobs', String(jobs)], logFile);
  const text = fs.readFileSync(logFile, 'utf8');
  say(`${'record'.padEnd(12)} exit ${code} ${String(Math.round((Date.now() - started) / 1000)).padStart(4)} s | ${(text.match(/The gate is (?:open|closed)|Recorded in [^\s]+|Nothing recorded[^.]*/g) || []).join('; ')}`);
  return code === 0 ? 0 : 1;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) process.exit(await main());
