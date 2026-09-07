// Launch Electron with the smoke flag set, and pass through its exit code.
import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';

const app = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const child = spawn(electron, [app, ...process.argv.slice(2)], {
  stdio: 'inherit',
  env: { ...process.env, RUTBA_OFFICE_SMOKE: '1', ELECTRON_DISABLE_SECURITY_WARNINGS: '1' },
});
child.on('exit', (code) => process.exit(code ?? 1));
