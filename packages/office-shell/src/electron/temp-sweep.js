// The folders Rutba Office makes in the system's temp folder for one job: a
// scratch file, an embedded object opened in its app. Each goes when its job
// or the app does; one left by a crash, or held open by another program at
// quit, goes here on a later start once it is a day old.
import fs from 'node:fs';
import path from 'node:path';

export const TEMP_PREFIXES = ['rutba-office-', 'rutba-object-'];

/** Removes `dir`'s folders named by one of `prefixes` and older than `age` ms. Returns how many went. */
export function sweepTemps(dir, { prefixes = TEMP_PREFIXES, age = 24 * 60 * 60 * 1000, now = Date.now() } = {}) {
  let gone = 0;
  let names = [];
  try { names = fs.readdirSync(dir); } catch { return 0; }
  for (const name of names) {
    if (!prefixes.some((p) => name.startsWith(p))) continue;
    const p = path.join(dir, name);
    try {
      const st = fs.statSync(p);
      if (!st.isDirectory() || now - st.mtimeMs < age) continue;
      fs.rmSync(p, { recursive: true, force: true });
      gone += 1;
    } catch {
      /* still in use by another program: it goes on a later start */
    }
  }
  return gone;
}
