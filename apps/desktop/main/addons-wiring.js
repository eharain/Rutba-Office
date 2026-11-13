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
import { addonsFolder, lockFolder, markRemoval, applyRemovals, letGo, stagingFor, canWrite, secureMachineFolder, PARTS } from './addons-folder.js';
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
/** Waits, without a thread of its own, until the process `pid` has gone or `ms` has passed. */
export function waitForExit(pid, ms = 60_000, { alive = (p) => { try { process.kill(p, 0); return true; } catch (err) { return err?.code === 'EPERM'; } }, now = Date.now } = {}) {
  const until = now() + ms;
  const pause = new Int32Array(new SharedArrayBuffer(4));
  while (alive(pid)) {
    if (now() >= until) return false;
    Atomics.wait(pause, 0, 0, 200);
  }
  return true;
}

/**
 * `--addons-apply-staged[=<folder>]` and `--addons-remove=<id,id>`: run by
 * the installer's elevated step during an update for everyone, or by the
 * administrator's prompt at Restart. Checks the staged add-ons with this
 * version's own rules, moves them in, does the removals, and answers the
 * exit code: 0, or 1 when something could not be done (it is tried again
 * at the next start). Answers null when this run is not that.
 *
 * At Restart it is given `--addons-after=<pid>`: Office's own process, which
 * quits as this run starts. Nothing is moved until it has gone, since a file
 * it holds open cannot be moved; then `--addons-relaunch` starts Office
 * again, through the shell, so that it runs as the person and not as the
 * administrator. A staged add-on that still could not be moved in stays
 * staged for the next start.
 */
export function applyStagedRun({ argv = process.argv, execPath = process.execPath, isPackaged, office, releaseKeys, env = process.env, log = () => {}, wait = waitForExit, relaunch = relaunchAsPerson, secure = secureMachineFolder }) {
  const from = flag(argv, '--addons-apply-staged');
  const remove = flag(argv, '--addons-remove');
  if (from === null && remove === null) return null;
  const after = Number(flag(argv, '--addons-after'));
  const again = flag(argv, '--addons-relaunch') !== null;
  if (Number.isInteger(after) && after > 0 && !wait(after)) log(`Rutba Office (${after}) had not quit after a minute; moving what can be moved`);
  try {
    return applyStagedNow({ from, remove, execPath, isPackaged, office, releaseKeys, env, log, secure });
  } finally {
    if (again) {
      try { relaunch(execPath, env); } catch (err) { log(`could not start Rutba Office again: ${err?.message || err}`); }
    }
  }
}

/** Starts Office through the shell, so it runs as the person signed in, not as the administrator. */
export function relaunchAsPerson(execPath, env = process.env) {
  const explorer = path.join(env.WINDIR || 'C:\\Windows', 'explorer.exe');
  spawn(explorer, [execPath], { detached: true, stdio: 'ignore', windowsHide: false }).unref();
}

function applyStagedNow({ from, remove, execPath, isPackaged, office, releaseKeys, env, log, secure }) {
  const folder = addonsFolder({ execPath, isPackaged, env });
  if (!folder) return 0;
  try {
    fs.mkdirSync(folder.dir, { recursive: true });
    // For everyone: only administrators may write here, before anything is moved in.
    if (folder.scope === 'machine' && !secure(folder.dir)) {
      log('the permissions of the add-ons folder could not be set; nothing was moved');
      return 1;
    }
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
      // The person's staging emptied but for what is kept for the next start (and the downloads in progress).
      if (at && fs.existsSync(at)) {
        const keep = new Set([...(r.kept || []).flatMap((id) => [id, `${id}.pack.json`]), '.parts']);
        for (const name of fs.readdirSync(at)) if (!keep.has(name)) fs.rmSync(path.join(at, name), { recursive: true, force: true });
        if (!fs.readdirSync(at).length) fs.rmSync(at, { recursive: true, force: true });
      }
      // What nothing will use: older versions kept, and downloads left for a week, the person's too.
      letGo(folder.dir, { office, partsIn: [path.join(stagingFor(folder.dir, office), PARTS), ...(at ? [path.join(at, PARTS)] : [])] });
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
 * asking first. No script file: the arguments are passed on the command
 * line, and the run checks what it moves with its own rules. It waits for
 * the run and answers its exit code; with `wait: false` it answers as soon
 * as Windows has started it (ok), or the person said no (code 5), so Office
 * can quit for the run to do its work.
 */
export function elevateWith(execPath) {
  const q = (s) => `'${String(s).replace(/['‘’‚‛]/g, (m) => m + m)}'`;
  return (args, { wait = true } = {}) => new Promise((resolve) => {
    const list = args.map((a) => q(`"${a}"`)).join(',');
    const command = wait
      ? `try { $p = Start-Process -FilePath ${q(execPath)} -ArgumentList ${list} -Verb RunAs -Wait -PassThru -WindowStyle Hidden; exit $p.ExitCode } catch { exit 5 }`
      : `try { Start-Process -FilePath ${q(execPath)} -ArgumentList ${list} -Verb RunAs -WindowStyle Hidden | Out-Null; exit 0 } catch { exit 5 }`;
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
    service: ({ stores: s, broadcast, app, bundledDir, updatesOn, engineState, languages = null, openWindows = null }) => createAddonsService({
      folder, start, contract, office, stores: s, broadcast, app, openWindows,
      bundled: bundledDir ? bundledFeed(bundledDir) : null,
      personStaging: folder?.scope === 'machine' ? personStaging(env) : null,
      elevate: process.platform === 'win32' ? elevateWith(execPath) : null,
      updatesOn,
      engineState,
      languages,
    }),
  };
}
