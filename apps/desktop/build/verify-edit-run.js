// Launch Electron with the editing checks enabled, and pass through the code.
import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';

const app = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const child = spawn(electron, [app], {
  stdio: 'inherit',
  env: { ...process.env, RUTBA_OFFICE_VERIFY_EDIT: '1', ELECTRON_ENABLE_LOGGING: '1' },
});
child.on('exit', (code) => process.exit(code ?? 1));
