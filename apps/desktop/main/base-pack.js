// Data Files' pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
//
// Each open database's SQLite file (base.js). What each database may read on
// this computer is kept in its settings, never in the file (A2.18), and a
// database renamed from the launcher reads what it read before: the core's
// rename tells this pack, through the hook it registers here.

import { createBaseService } from './base.js';
import { createTrust } from './base-trust.js';

/** This computer's list of what each database may read, in its settings (base-trust.js). */
const trustIn = (stores) => createTrust({ get: () => stores.settings.get('base.allowedSources', {}), set: (v) => stores.settings.set('base.allowedSources', v) });

export function register({ stores, onRecentRenamed }) {
  onRecentRenamed?.((from, to) => {
    if (/\.rbase$/i.test(from)) trustIn(stores).carry(from, to, { move: true });
  });
  return {
    base: createBaseService({ secrets: stores.secrets, trust: trustIn(stores) }),
  };
}
