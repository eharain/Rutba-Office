// The first thing a worker thread of a pack runs (--import, set by
// pack-resolve.js), so that its own imports are answered as its parent's are:
// module hooks belong to one thread, and a worker's imports are resolved
// before any line of it runs.

import { installPackResolution, WORKER_RESOLUTION } from './pack-resolve.js';

const given = process.env[WORKER_RESOLUTION];
if (given) installPackResolution(JSON.parse(given));
