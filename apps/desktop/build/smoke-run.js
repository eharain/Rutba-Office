// Launch Electron with the smoke flag set, and pass through its exit code.
//
// In a profile of its own, for the same reason the verification run has one: a
// screenshot taken against the developer's profile shows the developer's
// accounts, seventeen copies of a fixture imported by seventeen previous runs,
// and whatever was left selected last time. The point of a smoke capture is to
// show what a person sees, and a person's profile is empty until they use it.

import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';

const app = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

// `--profile keep` when you want to photograph your own mailbox rather than the
// fixture — useful once, never in a build.
const keep = process.argv.includes('--keep-profile');
const profile = keep ? null : fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-smoke-'));

const child = spawn(electron, [app, ...(profile ? [`--user-data-dir=${profile}`] : []), ...process.argv.slice(2).filter((a) => a !== '--keep-profile')], {
  stdio: 'inherit',
  env: { ...process.env, RUTBA_OFFICE_SMOKE: '1', RUTBA_SMOKE_SEED: '1', ELECTRON_DISABLE_SECURITY_WARNINGS: '1' },
});

child.on('exit', (code) => {
  if (profile) fs.rmSync(profile, { recursive: true, force: true });
  process.exit(code ?? 1);
});
