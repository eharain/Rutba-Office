// The first module of a process an add-on starts through the utilityProcess
// it is handed (apps/desktop/main/engine-host.js, heldUtilityProcess). That
// fork always starts this file, never the add-on's own: it puts
// pack-resolve's hooks in, from the resolution of the run that started it,
// and only then imports the add-on's module, so every import of the child is
// held to the add-on's folder from its first one. Nothing rests on execArgv
// or NODE_OPTIONS reaching the child, and a process started for any other
// module, or with no resolution, does not run it.

import { pathToFileURL } from 'node:url';
import { installPackResolution, closeAddonsFolder, isHeld, WORKER_RESOLUTION, CHILD_MODULE } from './pack-resolve.js';

const target = process.env[CHILD_MODULE];
delete process.env[CHILD_MODULE];
const given = process.env[WORKER_RESOLUTION];
if (!target || !given) throw new Error('This process starts only an add-on module, with the resolution of the run that started it.');
const told = JSON.parse(given);
if (told.packsDir) installPackResolution(told);
for (const dir of told.closed || []) closeAddonsFolder(dir);
// The parent checked it; checked again here, by its real path, under the hooks now in.
if (!isHeld(target)) throw new Error(`${target} is not inside an add-on whose imports are held, so it was not started.`);
await import(pathToFileURL(target).href);
