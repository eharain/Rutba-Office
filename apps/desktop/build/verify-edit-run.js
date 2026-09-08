// Launch Electron with the editing checks enabled, and pass through the code.
//
// In a profile of its own, for the reason the other two runs have one: a check
// that shares the developer's profile inherits every setting and every account
// left behind by working on the thing being checked, and then passes — or
// fails — for reasons that have nothing to do with the code. This was the one
// runner still using the real profile.
//
// Off the desktop by default too, like the application checks: a run must
// never take the keyboard from whoever is at the machine. Set
// RUTBA_WINDOW_DISPLAY=secondary to watch it on a second display.

import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';

const app = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-edit-'));

const child = spawn(electron, [app, `--user-data-dir=${profile}`], {
  stdio: 'inherit',
  env: { ...process.env, RUTBA_OFFICE_VERIFY_EDIT: '1', RUTBA_WINDOW_DISPLAY: process.env.RUTBA_WINDOW_DISPLAY ?? 'offscreen', ELECTRON_ENABLE_LOGGING: '1' },
});

child.on('exit', (code) => {
  fs.rmSync(profile, { recursive: true, force: true });
  process.exit(code ?? 1);
});
