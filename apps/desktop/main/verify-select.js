// Which of the window checks a run runs: by lane, by app and by part.
//
// Kept apart from verify-apps.js so the gate's tests can hold it to what the
// gate asks of it without opening a window (review of 11 Nov: gate 7 ran its
// one failed part again as "0/0", and nothing had proved the run again).
//
// A part's id is its app and the name of the function it runs,
// `mail#verifyMailIndexWorker`, not its place in the order: a place moved
// when a section before it was skipped or threw, so a number named another
// part in another lane. The checks of a section outside any part are
// `<app>#main`; `<app>#*` asks for the whole app.

/** The name of the function a part runs: `() => verifyMailLinks({...})` is verifyMailLinks; null for any other shape. */
export function partSlug(run) {
  const m = /^\s*(?:async\s*)?\(\s*\)\s*=>\s*(?:await\s+)?([A-Za-z_$][\w$]*)\s*\(/.exec(String(run));
  return m ? m[1] : null;
}

/**
 * The run's choices: `lane` ('parallel', 'machine' or 'all'), `apps` (none
 * for every app) and `ids` (none for every part). `runs(app)` says whether
 * a section of the app's checks outside any part runs; `part(app, run,
 * { machine })` registers a part and says whether it runs: { id, runs,
 * because } where because is null when it runs, or 'lane', 'app', 'id',
 * 'twice' when it does not. Every part registered is in `seen`.
 */
export function createSelector({ lane = 'all', apps = [], ids = [] } = {}) {
  const wantedApp = (app) => !apps.length || apps.includes(app);
  const whole = (app) => ids.includes(`${app}#*`);
  const seen = [];
  const taken = new Set();
  return {
    seen,
    // A section also runs when a part of its app is asked for: parts sit
    // inside sections (Mail's did), and a skipped section never reaches them.
    runs(app) {
      if (lane === 'machine' || !wantedApp(app)) return false;
      if (ids.length && !ids.some((id) => id.startsWith(`${app}#`))) return false;
      return true;
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
      seen.push({ id, app, machine, runs: because === null, because });
      return { id, runs: because === null, because };
    },
  };
}
