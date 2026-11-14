// The language packs a gate's window checks need, added to each check run's
// own add-ons folder the way a person adds one from a file (Add from file on
// the add-ons page, addons-service.js addFromFile): the signed release pack,
// its envelope checked with the contract's packages key and each file with
// its hash, staged and put in place, used from the run's start. Nothing is
// copied past that path.
//
// Since D-2026-11-14-16 Office carries no Urdu, Arabic or Hindi words of its
// own. A check run has no add-ons folder unless one is named
// (addons-wiring.js), so with no pack installed setLanguage('ur') is English
// and every check of a window in those languages fails. The gate is given the
// release's signed packs by one setting, RUTBA_GATE_PACKS: the packages
// release folder of this version, dist/v<version> in the packages repository
// (lang-<tag>.pack.json and the lang-<tag> folder beside it). A gate run whose
// window checks include a part below and has no such folder is refused at its
// start (tools/gate.js); build/verify-apps-run.js adds the packs before each
// check run's Electron starts.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { writeZip, ZipEntry } from '@rutba/ooxml/zip';
import { runContract } from './addons-wiring.js';
import { startAddons } from './addons-start.js';
import { createAddonsService } from './addons-service.js';
import { removeTree } from './addons-folder.js';

/** The setting naming the release folder of signed packs. */
export const GATE_PACKS_ENV = 'RUTBA_GATE_PACKS';

/** The packs every check run gets: the languages the checks below open windows in. */
export const GATE_LANGUAGE_PACKS = Object.freeze(['lang-ur', 'lang-ar', 'lang-hi']);

/**
 * The window check parts (<app>#<part>, as verify-apps.js names them) that
 * open a window in Urdu, Arabic or Hindi and fail without its pack:
 * verify-urdu.js (all three), verify-ribbon.js (the ribbon in Urdu runs right
 * to left), verify-video-editor.js, verify-smime.js and verify-mail-accounts.js.
 */
export const PARTS_NEEDING_PACKS = Object.freeze([
  'languages#verifyUrdu',
  'languages#verifyArabic',
  'languages#verifyHindi',
  'ribbon#verifyRibbon',
  'video#verifyVideoEditor',
  'mail#verifySmime',
  'mail#verifyMailAccounts',
]);

/**
 * The parts above a run of window checks includes: of the parts asked for by
 * id when there are any, of the apps whose checks run otherwise.
 * @param {{ apps?: string[], partIds?: string[] }} run
 */
export function partsNeedingPacks({ apps = [], partIds = [] } = {}) {
  if (partIds.length) return PARTS_NEEDING_PACKS.filter((id) => partIds.includes(id));
  return PARTS_NEEDING_PACKS.filter((id) => apps.includes(id.split('#')[0]));
}

/**
 * What is wrong with the release folder named, in a sentence, or null: not
 * set, not a folder, another version's, or a pack missing.
 * @param {string|undefined|null} release
 * @param {{ office: string, ids?: string[] }} o
 */
export function releaseProblem(release, { office, ids = GATE_LANGUAGE_PACKS }) {
  if (!release) {
    return `${GATE_PACKS_ENV} is not set. The window checks open windows in Urdu, Arabic and Hindi, whose words come only from their packs: set it to the packages release folder of ${office} (dist/v${office} in the packages repository, the signed ${ids.join(', ')}).`;
  }
  const dir = path.resolve(release);
  let isDir = false;
  try { isDir = fs.statSync(dir).isDirectory(); } catch { /* not there */ }
  if (!isDir) return `${GATE_PACKS_ENV} names ${dir}, which is not a folder.`;
  if (path.basename(dir) !== `v${office}`) return `${GATE_PACKS_ENV} names ${dir}, a release folder for another version: this Office is ${office}, so the folder is v${office}.`;
  const missing = ids.filter((id) => !fs.existsSync(path.join(dir, `${id}.pack.json`)) || !fs.existsSync(path.join(dir, id)));
  if (missing.length) return `${GATE_PACKS_ENV} names ${dir}, which has no signed ${missing.join(', ')} (each is <id>.pack.json and its <id> folder).`;
  return null;
}

/**
 * Why a gate run may not start, in plain words, or null: it runs window
 * checks, they include a part that needs the packs, and the release folder
 * is not set or not usable. Said at the start, before any check runs.
 * @param {{ verifying: boolean, apps?: string[], partIds?: string[], release?: string|null, office: string }} o
 */
export function packsRefusal({ verifying, apps = [], partIds = [], release = null, office }) {
  if (!verifying) return null;
  const parts = partsNeedingPacks({ apps, partIds });
  if (!parts.length) return null;
  const problem = releaseProblem(release, { office });
  return problem ? `The gate will not start: its window checks include ${parts.join(', ')}, which open windows in Urdu, Arabic or Hindi. ${problem}` : null;
}

/**
 * A release pack as the file a person adds: a .zip of its envelope and of the
 * files its signed manifest lists, under its own folder, and nothing else.
 */
export function zipOfPack(release, id) {
  const envelopeBytes = fs.readFileSync(path.join(release, `${id}.pack.json`));
  const envelope = JSON.parse(envelopeBytes.toString('utf8'));
  const entry = (name, data) => {
    const e = new ZipEntry({ name, method: 8 });
    e.data = data;
    return e;
  };
  const entries = [entry(`${id}.pack.json`, envelopeBytes)];
  for (const rel of Object.keys(envelope.manifest?.files || {})) {
    entries.push(entry(`${id}/${rel}`, fs.readFileSync(path.join(release, id, ...rel.split('/')))));
  }
  return writeZip(entries);
}

/**
 * Adds the packs to an add-ons folder through the add-ons service, as Add
 * from file does, each refused for what the service refuses (a stranger's
 * signature, another Office's pack, a file changed). Throws with the reason;
 * the ids added otherwise.
 * @param {{ release: string, addonsDir: string, office: string, ids?: string[], contract?: object }} o
 */
export async function installGatePacks({ release, addonsDir, office, ids = GATE_LANGUAGE_PACKS, contract = runContract({ office, releaseKeys: [] }) }) {
  const problem = releaseProblem(release, { office, ids });
  if (problem) throw new Error(problem);
  const from = path.resolve(release);
  const folder = { dir: path.resolve(addonsDir), scope: 'named' };
  fs.mkdirSync(folder.dir, { recursive: true });
  // The service's own working files, never the run's profile: a start counted there would be the run's first.
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-gate-packs-'));
  try {
    const start = startAddons({ folder, contract, office, memoryFile: path.join(work, 'starts.json'), argv: [], off: [], keep: [] });
    const kept = new Map();
    const stores = { settings: { get: (k, d) => (kept.has(k) ? kept.get(k) : d), set: (k, v) => kept.set(k, v) } };
    const service = createAddonsService({
      folder, start, contract, office, stores,
      source: null,
      fetchImpl: async () => { throw new Error('the gate adds packs from the release folder only'); },
    });
    for (const id of ids) {
      const file = path.join(work, `${id}.zip`);
      fs.writeFileSync(file, zipOfPack(from, id));
      let r;
      try {
        r = await service.addFromFile({ file });
      } catch (err) {
        throw new Error(`${id} from ${from} was refused: ${err?.message || err}`);
      }
      if (r?.added !== id || !fs.existsSync(path.join(folder.dir, `${id}.pack.json`))) throw new Error(`${id} from ${from} was not put in place in ${folder.dir}.`);
    }
    return [...ids];
  } finally {
    removeTree(work);
  }
}
