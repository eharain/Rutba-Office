// The add-ons as main.js uses them: the contract this run decides with, the
// add-ons of this start (addons-start.js) and the service (addons-service.js).
// Kept apart from main.js so it can be read and tried on its own.
//
// There is no administrator's step in 2.26.0: every copy keeps its add-ons
// in the person's own folder (addons-folder.js, D-2026-11-13-35), written as
// the person, so nothing is moved in with the administrator's permission.

import path from 'node:path';
import { createRequire } from 'node:module';
import { contractFor } from '@rutba/office-shell/electron/addons-compat';
import { addonsFolder } from './addons-folder.js';
import { startAddons } from './addons-start.js';
import { createAddonsService, KEEP_KEY } from './addons-service.js';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');
const CONTRACT_FILE = require('@rutba/office-shell/addons-contract.json');

/** The contract this run decides with: the file as built, this Office's version, the core's release pack keys for the engine. */
export const runContract = ({ office, releaseKeys }) => contractFor(CONTRACT_FILE, { office, releaseKeys });

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
  const keep = stores?.settings?.get(KEEP_KEY, []) || [];
  const start = startAddons({ folder, contract, office, memoryFile: path.join(userData, 'addons-starts.json'), argv, off, keep });
  // The take-backs the folder has now: the profile lets go of them.
  if (start.takenBack?.length) stores.settings.set(KEEP_KEY, keep.filter((x) => !start.takenBack.includes(x)));
  for (const r of start.registry.refused) console.error(`[addons] ${r.id}: ${r.reason}`);
  if (start.said) console.error(`[addons] ${start.said}`);
  return {
    start,
    folder,
    contract,
    service: ({ stores: s, broadcast, app, bundledDir, updatesOn, engineState, languages = null, openWindows = null }) => createAddonsService({
      folder, start, contract, office, stores: s, broadcast, app, openWindows,
      bundled: bundledDir ? bundledFeed(bundledDir) : null,
      updatesOn,
      engineState,
      languages,
    }),
  };
}
