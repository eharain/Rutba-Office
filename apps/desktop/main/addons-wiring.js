// The add-ons as main.js uses them: the contract this run decides with, the
// add-ons of this start (addons-start.js), the service (addons-service.js),
// and the run an installer or an administrator's prompt makes with
// `--addons-apply-staged`, which moves staged add-ons in and quits without a
// window. Kept apart from main.js so it can be read and tried on its own.

import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { contractFor } from '@rutba/office-shell/electron/addons-compat';
import { addonsFolder, lockFolder, markRemoval, applyRemovals, canWrite } from './addons-folder.js';
import { startAddons, applyStaged } from './addons-start.js';
import { createAddonsService } from './addons-service.js';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');
const CONTRACT_FILE = require('@rutba/office-shell/addons-contract.json');

/** The contract this run decides with: the file as built, this Office's version, the core's release pack keys for the engine. */
export const runContract = ({ office, releaseKeys }) => contractFor(CONTRACT_FILE, { office, releaseKeys });

/** Where an install for everyone stages add-ons as the person. */
export const personStaging = (env = process.env) => path.join(env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local'), 'Rutba Office', 'addons-staging');

const flag = (argv, name) => {
  const a = argv.find((x) => x === name || x.startsWith(`${name}=`));
  if (!a) return null;
  return a.includes('=') ? a.slice(name.length + 1).replace(/^"|"$/g, '') : '';
};

/**
 * `--addons-apply-staged[=<folder>]` and `--addons-remove=<id,id>`: run by
 * the installer's elevated step during an update for everyone, or by the
 * administrator's prompt at Restart. Checks the staged add-ons with this
 * version's own rules, moves them in, does the removals, and answers the
 * exit code: 0, or 1 when something could not be done (it is tried again
 * at the next start). Answers null when this run is not that.
 */
export function applyStagedRun({ argv = process.argv, execPath = process.execPath, isPackaged, office, releaseKeys, env = process.env, log = () => {} }) {
  const from = flag(argv, '--addons-apply-staged');
  const remove = flag(argv, '--addons-remove');
  if (from === null && remove === null) return null;
  const folder = addonsFolder({ execPath, isPackaged, env });
  if (!folder) return 0;
  try {
    fs.mkdirSync(folder.dir, { recursive: true });
    if (!canWrite(folder.dir)) {
      log('the add-ons folder cannot be written by this run');
      return 1;
    }
    const release = lockFolder(folder.dir);
    if (!release) return 1;
    try {
      for (const id of String(remove || '').split(',').map((s) => s.trim()).filter(Boolean)) markRemoval(folder.dir, id);
      const removed = applyRemovals(folder.dir);
      const at = from ? path.join(from, office) : null;
      const r = applyStaged({ dir: folder.dir, from: at && fs.existsSync(at) ? at : null, contract: runContract({ office, releaseKeys }), office });
      if (at && fs.existsSync(at)) fs.rmSync(at, { recursive: true, force: true });
      log(`moved in ${r.moved.join(', ') || 'nothing'}; removed ${removed.removed.join(', ') || 'nothing'}${r.refused.length ? `; refused ${r.refused.map((x) => `${x.id} (${x.code})`).join(', ')}` : ''}`);
      return r.refused.length || removed.held.length ? 1 : 0;
    } finally {
      release();
    }
  } catch (err) {
    log(err?.message || String(err));
    return 1;
  }
}

/**
 * Runs this program once as administrator with the arguments given, Windows
 * asking first, and waits for it. No script file: the arguments are passed
 * on the command line, and the run checks what it moves with its own rules.
 */
export function elevateWith(execPath) {
  const q = (s) => `'${String(s).replace(/['‘’‚‛]/g, (m) => m + m)}'`;
  return (args) => new Promise((resolve) => {
    const list = args.map((a) => q(`"${a}"`)).join(',');
    const command = `try { $p = Start-Process -FilePath ${q(execPath)} -ArgumentList ${list} -Verb RunAs -Wait -PassThru -WindowStyle Hidden; exit $p.ExitCode } catch { exit 5 }`;
    const child = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], { windowsHide: true, stdio: 'ignore' });
    child.on('error', () => resolve({ ok: false, code: -1 }));
    child.on('exit', (code) => resolve({ ok: code === 0, code }));
  });
}

/** The feed built into this release, for the list when the network cannot be reached, or null. */
export function bundledFeed(dir) {
  try {
    return { feed: JSON.parse(fs.readFileSync(path.join(dir, 'feed.json'), 'utf8')), sig: JSON.parse(fs.readFileSync(path.join(dir, 'feed.json.sig'), 'utf8')) };
  } catch {
    return null;
  }
}

/**
 * This run's add-ons and their service.
 * @returns {{ start, folder, contract, service: (deps) => object }}
 */
export function addonsForRun({ execPath = process.execPath, isPackaged, office, releaseKeys, userData, argv = process.argv, env = process.env, stores = null, checkRun = false }) {
  const contract = runContract({ office, releaseKeys });
  // A check run uses no add-ons unless it names a folder of its own.
  const folder = checkRun && !env.RUTBA_ADDONS_DIR ? null : addonsFolder({ execPath, isPackaged, env });
  const off = stores?.settings?.get('addons.off', []) || [];
  const start = startAddons({ folder, contract, office, memoryFile: path.join(userData, 'addons-starts.json'), argv, off });
  for (const r of start.registry.refused) console.error(`[addons] ${r.id}: ${r.reason}`);
  if (start.said) console.error(`[addons] ${start.said}`);
  return {
    start,
    folder,
    contract,
    service: ({ stores: s, broadcast, app, bundledDir, updatesOn, engineState, languages = null }) => createAddonsService({
      folder, start, contract, office, stores: s, broadcast, app,
      bundled: bundledDir ? bundledFeed(bundledDir) : null,
      personStaging: folder?.scope === 'machine' ? personStaging(env) : null,
      elevate: process.platform === 'win32' ? elevateWith(execPath) : null,
      updatesOn,
      engineState,
      languages,
    }),
  };
}
