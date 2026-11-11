// Which of the window checks a run runs: by lane, by app and by part.
//
// Kept apart from verify-apps.js so the gate's tests can hold it to what the
// gate asks of it without opening a window (review of 11 Nov: gate 7 ran its
// one failed part again as "0/0", and nothing had proved the run again).
//
// A part's id is its app and the name of the function it runs,
// `mail#verifyMailIndexWorker`, not its place in the order: a place moved
// when a section before it was skipped or threw, so a number named another
// part in another lane. A section of checks written in line is named too,
// `mail#reading`, and runs alone when asked for (owner, 11 Nov 2026: smaller
// groups, each run on its own); `<app>#*` asks for the whole app.

/** The name of the function a part runs: `() => verifyMailLinks({...})` is verifyMailLinks; null for any other shape. */
export function partSlug(run) {
  const m = /^\s*(?:async\s*)?\(\s*\)\s*=>\s*(?:await\s+)?([A-Za-z_$][\w$]*)\s*\(/.exec(String(run));
  return m ? m[1] : null;
}

/**
 * The run's choices: `lane` ('parallel', 'machine' or 'all'), `apps` (none
 * for every app), `ids` (none for every part) and `skip` (parts and sections
 * that do not run here: another piece of the same app runs them). `section(app, name)`
 * registers a section of checks written in line and says whether it runs;
 * `part(app, run, { machine })` registers a part and says whether it runs.
 * Both answer { id, runs, because } where because is null when it runs, or
 * 'lane', 'app', 'id', 'twice' when it does not. `wants(app)` says whether
 * the run is for the app at all. Everything registered is in `seen`.
 *
 * A section no longer runs to reach a part asked for: no part sits inside
 * one (tests/gate-sections.test.js holds verify-apps.js to that), so a part
 * run again runs that part alone, not every section of its app beside it.
 */
export function createSelector({ lane = 'all', apps = [], ids = [], skip = [] } = {}) {
  const wantedApp = (app) => !apps.length || apps.includes(app);
  const whole = (app) => ids.includes(`${app}#*`);
  const seen = [];
  const taken = new Set();
  return {
    seen,
    wants(app) {
      if (lane === 'machine' || !wantedApp(app)) return false;
      return !ids.length || ids.some((id) => id.startsWith(`${app}#`));
    },
    section(app, name) {
      const id = `${app}#${name}`;
      if (!/^[A-Za-z][\w-]*$/.test(String(name)) || taken.has(id)) {
        const because = taken.has(id) ? 'twice' : 'unnamed';
        seen.push({ id, app, machine: false, section: true, runs: false, because });
        return { id, runs: false, because };
      }
      taken.add(id);
      let because = null;
      if (!wantedApp(app)) because = 'app';
      // A section is the parallel lane's: the machine lane runs only machine parts.
      else if (lane === 'machine') because = 'lane';
      else if (ids.length && !ids.includes(id) && !whole(app)) because = 'id';
      else if (skip.includes(id)) because = 'id';
      seen.push({ id, app, machine: false, section: true, runs: because === null, because });
      return { id, runs: because === null, because };
    },
    part(app, run, { machine = false } = {}) {
      const slug = partSlug(run);
      const id = `${app}#${slug || 'unnamed'}`;
      if (!slug || taken.has(id)) {
        // Two parts with one id cannot be told apart by a run again: neither runs, and the run says so.
        seen.push({ id, app, machine, runs: false, because: slug ? 'twice' : 'unnamed' });
        return { id, runs: false, because: slug ? 'twice' : 'unnamed' };
      }
      taken.add(id);
      let because = null;
      // The app first: a chunk leaves to the machine lane only its own app's machine parts.
      if (!wantedApp(app)) because = 'app';
      else if (machine ? lane === 'parallel' : lane === 'machine') because = 'lane';
      else if (ids.length && !ids.includes(id) && !whole(app)) because = 'id';
      else if (skip.includes(id)) because = 'id';
      seen.push({ id, app, machine, runs: because === null, because });
      return { id, runs: because === null, because };
    },
  };
}
