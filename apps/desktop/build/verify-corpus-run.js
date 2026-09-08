// Launch Electron with the corpus check enabled, in a profile of its own.
//
// The profile matters twice over here: a run over a thousand files must not
// write a thousand entries into the developer's recent list (which is what
// the first check runs did, and pushed every real entry out of it), and what
// the check sees must be the fixtures' state, not the developer's.
//
//   npm run verify:corpus
//   RUTBA_CORPUS_DIRS="D:\folder;E:\other" npm run verify:corpus
//   RUTBA_CORPUS_LIMIT=25 npm run verify:corpus

import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';

const app = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-corpus-'));
const out = process.env.RUTBA_CORPUS_OUT || path.join(app, 'build', 'corpus');

const child = spawn(electron, [app, `--user-data-dir=${profile}`], {
  stdio: 'inherit',
  env: {
    ...process.env,
    RUTBA_OFFICE_VERIFY_CORPUS: '1',
    RUTBA_CORPUS_OUT: out,
    RUTBA_WINDOW_DISPLAY: process.env.RUTBA_WINDOW_DISPLAY ?? 'offscreen',
    ELECTRON_DISABLE_SECURITY_WARNINGS: '1',
  },
});

child.on('exit', (code) => {
  fs.rmSync(profile, { recursive: true, force: true });
  console.log(`report: ${path.join(out, 'report.txt')}`);
  process.exit(code ?? 1);
});
