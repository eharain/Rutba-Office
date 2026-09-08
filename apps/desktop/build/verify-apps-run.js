// Launch Electron with the application checks enabled.
//
// In a profile of its own. A verification run that shares the developer's
// profile inherits every account imported while working on the thing being
// verified, and then passes — or fails — for reasons that have nothing to do
// with the code. The directory is fresh each run and removed afterwards, so
// what the checks see is exactly what the fixtures put there.

import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';

const app = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-'));

const child = spawn(electron, [app, `--user-data-dir=${profile}`], {
  stdio: 'inherit',
  env: { ...process.env, RUTBA_OFFICE_VERIFY_APPS: '1', RUTBA_WINDOW_DISPLAY: process.env.RUTBA_WINDOW_DISPLAY ?? 'offscreen', RUTBA_SMOKE_SEED: '1' },
});

child.on('exit', (code) => {
  fs.rmSync(profile, { recursive: true, force: true });
  process.exit(code ?? 1);
});
