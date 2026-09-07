// Launch Electron with the application checks enabled.
import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';

const app = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const child = spawn(electron, [app], {
  stdio: 'inherit',
  env: { ...process.env, RUTBA_OFFICE_VERIFY_APPS: '1', RUTBA_SMOKE_SEED: '1' },
});
child.on('exit', (code) => process.exit(code ?? 1));
