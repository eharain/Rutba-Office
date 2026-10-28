// Runs a command with a temp folder of its own and removes that folder when
// the command ends, so what the tests and checks write to the temp folder
// (os.tmpdir(), %TEMP%) does not stay behind on the machine.
//
//   node tools/run-temp.js node --test "tests/*.test.js"
//
// npm test, npm run test:perf and npm run gate go through it, and whatever
// they start inherits the folder: a test that makes a temp folder and never
// removes it, or one stopped halfway, leaves nothing once the run is over.
// Inside a run that already has one (npm test under the gate) the command
// uses that folder, and the outer run removes it. A folder left by a run that
// was killed is removed by the next run once nothing in it is a day old.

import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const PREFIX = 'rutba-office-run-';
const DAY = 24 * 60 * 60 * 1000;

/** The newest write in a folder and everything in it. */
function newest(dir) {
  let latest = fs.statSync(dir).mtimeMs;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    try {
      latest = Math.max(latest, entry.isDirectory() ? newest(full) : fs.lstatSync(full).mtimeMs);
    } catch { /* gone while looking */ }
  }
  return latest;
}

/** Removes the run folders that killed runs left, once nothing in them is a day old. */
function sweep(dir) {
  for (const name of fs.readdirSync(dir)) {
    if (!name.startsWith(PREFIX)) continue;
    const full = path.join(dir, name);
    try {
      if (Date.now() - newest(full) > DAY) fs.rmSync(full, { recursive: true, force: true, maxRetries: 3 });
    } catch { /* still in use, or already gone */ }
  }
}

const [command, ...args] = process.argv.slice(2);
if (!command) {
  console.error('usage: node tools/run-temp.js <command> [args...]');
  process.exit(2);
}

const outer = os.tmpdir();
const own = !path.basename(outer).startsWith(PREFIX);
let temp = outer;
if (own) {
  sweep(outer);
  temp = fs.mkdtempSync(path.join(outer, PREFIX));
}

const child = spawn(command === 'node' ? process.execPath : command, args, {
  stdio: 'inherit',
  env: { ...process.env, TEMP: temp, TMP: temp, TMPDIR: temp },
  shell: command !== 'node' && process.platform === 'win32',
});

// Ctrl+C reaches the command as well; this waits for it to end and then
// removes the folder. A termination sent to this process alone is passed on.
process.on('SIGINT', () => {});
process.on('SIGTERM', () => child.kill('SIGTERM'));

function done(code) {
  if (own) {
    try {
      fs.rmSync(temp, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
    } catch (err) {
      console.error(`run-temp: ${temp} is left behind (${err.code}); the next run removes it once it is a day old`);
    }
  }
  process.exit(code);
}

child.on('error', (err) => {
  console.error(`run-temp: ${command} did not start: ${err.message}`);
  done(1);
});
child.on('exit', (code) => done(code ?? 1));
