// The gate's passes on the lab machine: run there by the gate through its
// runner (tools/gate-hosts.js), never by hand on this computer.
//
//   node tools/run-temp.js node tools/gate-vm.js --passes test,fuzz:open --seed 1234
//   node tools/run-temp.js node tools/gate-vm.js --passes test --files tests/a.test.js --alone
//
// Each pass writes its output to .gate/logs/<pass>.log, which the runner
// brings back, and the machine says what it is in .gate/logs/vm-result.json.
// The engine tests here are those that touch nothing of Windows itself
// (tools/engine-hosts.json); --files names some of them, for a run again.
// --alone runs one file and one pass at a time.

import { execFileSync, spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { FUZZERS, VM_PASSES, engineFiles, readHosts, splitEngine, osReach, testArgs } from './gate-hosts.js';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const LOGS = path.join(ROOT, '.gate', 'logs');

const arg = (name) => { const i = process.argv.indexOf(name); return i > 0 ? process.argv[i + 1] : null; };
const passes = String(arg('--passes') || '').split(',').filter(Boolean);
const files = String(arg('--files') || '').split(',').filter(Boolean);
const seed = String(arg('--seed') || '1');
const alone = process.argv.includes('--alone');

const strange = passes.filter((p) => !VM_PASSES.includes(p));
if (!passes.length || strange.length) {
  console.log(`gate-vm: --passes names ${strange.join(', ') || 'nothing'}; the passes here are ${VM_PASSES.join(', ')}`);
  process.exit(2);
}

/** One pass into its log: its exit code and seconds. */
function run(name, command, args) {
  return new Promise((resolve) => {
    const started = Date.now();
    const log = fs.openSync(path.join(LOGS, `${name.replace(':', '-')}.log`), 'w');
    const child = spawn(command, args, { cwd: ROOT, stdio: ['ignore', log, log] });
    child.on('error', (err) => { fs.writeSync(log, `\n${err.message}\n`); });
    child.on('close', (code) => {
      fs.closeSync(log);
      const seconds = Math.round((Date.now() - started) / 1000);
      console.log(`[gate-vm] ${name} exit ${code} in ${seconds} s`);
      resolve({ code: code ?? 1, seconds });
    });
  });
}

fs.mkdirSync(LOGS, { recursive: true });
const result = { os: `${os.type()} ${os.release()}`, osName: osName(), platform: process.platform, arch: process.arch, node: process.version, cpus: os.cpus().length, tree: treeHere(), passes: {} };

/**
 * The tree this machine runs, as git names it, said before anything runs:
 * another run from the same clone checks out its own snapshot in the same
 * folder, so the gate compares this with the tree it started on.
 */
function treeHere() {
  try {
    const dirty = execFileSync('git', ['status', '--porcelain'], { cwd: ROOT, encoding: 'utf8' }).trim();
    return `${execFileSync('git', ['rev-parse', 'HEAD^{tree}'], { cwd: ROOT, encoding: 'utf8' }).trim()}${dirty ? '+changed' : ''}`;
  } catch {
    return null;
  }
}

/** The distribution's own name where it says one: Rocky Linux 9.5, not only Linux. */
function osName() {
  try {
    const rel = fs.readFileSync('/etc/os-release', 'utf8');
    const m = /^PRETTY_NAME="?([^"\n]+)"?$/m.exec(rel);
    if (m) return m[1];
  } catch { /* not Linux, or not said */ }
  return os.type();
}

if (passes.includes('test')) {
  const all = engineFiles(ROOT);
  // Split as the gate splits, imports followed: a file it keeps on Windows is not run here as well (review of 11 Nov).
  const reachCache = new Map();
  const mine = splitEngine(all, readHosts(ROOT), (f) => osReach(ROOT, f, reachCache)).vm;
  const asked = files.length ? files : mine;
  const wrong = asked.filter((f) => !mine.includes(f));
  if (wrong.length) {
    console.log(`gate-vm: not test files this machine runs: ${wrong.join(', ')}`);
    process.exit(2);
  }
  result.passes.test = { ...(await run('test', process.execPath, testArgs(asked, { events: '.gate/logs/test-events.jsonl', alone }))), files: asked.length, list: asked };
}
const fuzzers = FUZZERS.filter(([name]) => passes.includes(name));
const fuzz = (z) => run(z[0], process.execPath, [z[1], String(z[3]), seed]).then((r) => { result.passes[z[0]] = r; });
if (alone) for (const z of fuzzers) await fuzz(z);
else await Promise.all(fuzzers.map(fuzz));

fs.writeFileSync(path.join(LOGS, 'vm-result.json'), `${JSON.stringify(result, null, 2)}\n`);
process.exit(Object.values(result.passes).every((p) => p.code === 0) ? 0 : 1);
