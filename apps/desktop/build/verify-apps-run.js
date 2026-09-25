// Launch Electron with the application checks enabled.
//
// In a profile of its own. A verification run that shares the developer's
// profile inherits every account imported while working on the thing being
// verified, and then passes — or fails — for reasons that have nothing to do
// with the code. The directory is fresh each run and removed afterwards, so
// what the checks see is exactly what the fixtures put there.

import { spawn, execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';

const app = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-'));

const child = spawn(electron, [app, `--user-data-dir=${profile}`], {
  stdio: ['inherit', 'pipe', 'inherit'],
  env: { ...process.env, RUTBA_OFFICE_VERIFY_APPS: '1', RUTBA_WINDOW_DISPLAY: process.env.RUTBA_WINDOW_DISPLAY ?? 'offscreen', RUTBA_SMOKE_SEED: '1' },
});

// The run's own verdict is the code its main process ends on, and it says so
// on its last line. Read it on the way through.
let verdict = null;
let tail = '';
child.stdout.on('data', (chunk) => {
  process.stdout.write(chunk);
  tail = (tail + chunk.toString()).slice(-2000);
  const m = /\[exit\] the process ended with (\d+)/.exec(tail);
  if (m) verdict = Number(m[1]);
});

/**
 * Every process still holding this run's profile. On a machine busy with
 * other work Electron's teardown can overrun, end on a code of its own
 * (7014 was seen) and leave its renderers running with nothing to close them;
 * they are stopped here, by the profile only this run used.
 */
function stopLeftovers() {
  if (process.platform !== 'win32') return 0;
  try {
    const out = execFileSync('powershell.exe', [
      '-NoProfile', '-Command',
      `$p = @(Get-CimInstance Win32_Process -Filter "Name='electron.exe'" | Where-Object { $_.CommandLine -like '*${profile.replace(/'/g, "''")}*' }); $p | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }; $p.Count`,
    ], { encoding: 'utf8', timeout: 30000 });
    return Number(out.trim()) || 0;
  } catch {
    return 0;
  }
}

child.on('exit', (code, signal) => {
  let result = code ?? 1;
  if (code !== 0) console.log(`     [exit] Electron ended with code ${code}${signal ? `, signal ${signal}` : ''}`);
  // A code the run never chose, after the run had already said its verdict:
  // the teardown's, not the checks'. The verdict is what the gate reads.
  if (verdict !== null && code !== verdict && code !== 0 && code !== 1) {
    console.log(`     [exit] the checks' own verdict was ${verdict}; Electron's teardown ended on ${code} after it — the verdict stands`);
    result = verdict;
  }
  const stopped = stopLeftovers();
  if (stopped) console.log(`     [exit] stopped ${stopped} Electron process(es) the run left behind`);
  try {
    fs.rmSync(profile, { recursive: true, force: true });
  } catch {
    // Held a moment longer by a process just stopped; the temp folder is the system's to clear.
  }
  process.exit(result);
});
