// The app packs' window checks (main/verify-packs.js): two starts of Rutba
// Office from the source over one profile and one packs folder.
//
// The packs are made for the run and signed with a key made for it and
// thrown away after it; the core is handed the public half
// (RUTBA_PACK_PUBLIC_KEY), as a run from the source may be. The packs
// folder starts with Documents, Mail, AI Tools and Talk, which carries its
// own code (build/make-app-packs.js, CARRY_CODE); the release's file for
// every pack is in a folder of its own, which Add or remove apps adds from
// (RUTBA_PACKS_SOURCE). The first start adds Calendar and removes Mail; the
// second sees the change. Everything is in a temporary folder, removed
// afterwards.
//
//   node build/bundle.js && node build/verify-packs-run.js

import { spawn, execFileSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';
import { makeKeyPair } from '../../../tools/make-pack-key.mjs';
import { makeAppPacks } from './make-app-packs.js';
import { packFiles } from '@rutba/office-shell/electron/pack-install';

const app = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-packs-'));
const profile = path.join(tmp, 'profile');
const installed = path.join(tmp, 'installed');
const release = path.join(tmp, 'release');
const { version } = JSON.parse(fs.readFileSync(path.join(app, 'package.json'), 'utf8'));

const pair = makeKeyPair();
const publicKey = path.join(tmp, 'pack-key.pub.pem');
fs.writeFileSync(publicKey, pair.publicPem);
const built = await makeAppPacks({ privateKey: crypto.createPrivateKey(pair.privatePem), version, release, root: path.join(tmp, 'packs'), keysFile: path.join(tmp, 'pack-keys.json'), installerFile: path.join(tmp, 'installer-packs.nsh') });
fs.mkdirSync(installed, { recursive: true });
for (const id of ['word', 'mail', 'ai', 'talk']) {
  const f = packFiles(id);
  for (const name of [f.archive, f.unpacked, f.manifest]) {
    if (fs.existsSync(path.join(built.out, name))) fs.cpSync(path.join(built.out, name), path.join(installed, name), { recursive: true });
  }
}

/** Every process still holding this run's profile, stopped. */
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

/** One start, its verdict the code its main process ends on. */
function start(phase) {
  return new Promise((resolve) => {
    const child = spawn(electron, [app, `--user-data-dir=${profile}`], {
      stdio: ['inherit', 'pipe', 'inherit'],
      env: {
        ...process.env,
        RUTBA_OFFICE_VERIFY_APPS: '1',
        RUTBA_VERIFY_ONLY: 'packs',
        RUTBA_PACKS_PHASE: phase,
        RUTBA_PACKS_DIR: installed,
        RUTBA_PACK_PUBLIC_KEY: publicKey,
        RUTBA_PACKS_SOURCE: release,
        RUTBA_WINDOW_DISPLAY: process.env.RUTBA_WINDOW_DISPLAY ?? 'offscreen',
        RUTBA_SCREEN_SCALE: process.env.RUTBA_SCREEN_SCALE ?? '1',
      },
    });
    let verdict = null;
    let tail = '';
    child.stdout.on('data', (chunk) => {
      process.stdout.write(chunk);
      tail = (tail + chunk.toString()).slice(-2000);
      const m = /\[exit\] the process ended with (\d+)/.exec(tail);
      if (m) verdict = Number(m[1]);
    });
    child.on('exit', (code) => {
      const stopped = stopLeftovers();
      if (stopped) console.log(`     [exit] stopped ${stopped} Electron process(es) the ${phase} start left behind`);
      resolve(verdict ?? code ?? 1);
    });
  });
}

let result = 1;
try {
  const first = await start('first');
  const next = first === 0 ? await start('next') : 1;
  result = first || next;
  console.log(`     [packs] the first start ended with ${first}, the next with ${next}`);
} finally {
  try {
    fs.rmSync(tmp, { recursive: true, force: true });
  } catch {
    // Held a moment longer by a process just stopped; the temp folder is the system's to clear.
  }
}
process.exit(result);
