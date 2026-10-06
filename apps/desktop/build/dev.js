// A run from source that keeps up with the editor.
//
// The renderer and the preload are built once, then rebuilt on every save
// while the application runs from them; closing the application ends the
// watch. A change to the renderer shows after Reload (Ctrl+Shift+R, which a
// copy run from source has in its View menu); a change to the main process
// needs the application started again.
//
//   npm run dev
//   npm run dev -- --user-data-dir=<folder>    any argument goes to Electron
//
// `npm run dev` pointed at this file from the first commit, and the file was
// never written: the command failed with a missing module.

import { spawn, spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';

const here = path.dirname(fileURLToPath(import.meta.url));
const app = path.resolve(here, '..');
const bundle = path.join(here, 'bundle.js');

// Built once first, so the window never opens on a bundle still being written.
const built = spawnSync(process.execPath, [bundle], { cwd: app, stdio: 'inherit' });
if (built.status !== 0) process.exit(built.status ?? 1);

const watcher = spawn(process.execPath, [bundle, '--watch'], { cwd: app, stdio: 'inherit' });
const suite = spawn(electron, ['.', ...process.argv.slice(2)], { cwd: app, stdio: 'inherit' });

suite.on('exit', (code) => {
  watcher.kill();
  process.exit(code ?? 0);
});
watcher.on('exit', (code) => {
  if (code) console.error(`the renderer watch stopped (${code}); the application is still running on the last build`);
});
