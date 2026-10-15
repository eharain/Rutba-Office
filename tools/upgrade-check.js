// The upgrade test's judgement, apart from the running (tools/upgrade-test.js),
// so it can be proven in the engine tests: which release is the previous one,
// what each release's files are called, the profile an older version is
// started on, and whether what the newer one read from it is what was left
// there.

import fs from 'node:fs';
import path from 'node:path';

/** 2.0.0, v2.0.0, 2.0.0-beta.1 as numbers to compare; a pre-release tag sorts before its release. */
export function parseVersion(v) {
  const m = /^v?(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?$/.exec(String(v || '').trim());
  if (!m) return null;
  return { major: Number(m[1]), minor: Number(m[2]), patch: Number(m[3]), pre: m[4] || null, text: `${m[1]}.${m[2]}.${m[3]}${m[4] ? `-${m[4]}` : ''}` };
}

export function compareVersions(a, b) {
  const x = parseVersion(a);
  const y = parseVersion(b);
  if (!x || !y) throw new Error(`not a version: ${!x ? a : b}`);
  for (const k of ['major', 'minor', 'patch']) if (x[k] !== y[k]) return x[k] < y[k] ? -1 : 1;
  if (x.pre === y.pre) return 0;
  if (!x.pre) return 1;
  if (!y.pre) return -1;
  return x.pre < y.pre ? -1 : 1;
}

/**
 * The release to upgrade FROM: the newest published release older than `to`.
 * `releases` is GitHub's list (tag_name, draft); drafts are nobody's
 * installed copy, so they are left out.
 */
export function previousRelease(releases, to) {
  const older = (releases || [])
    .filter((r) => !r.draft && parseVersion(r.tag_name) && compareVersions(r.tag_name, to) < 0)
    .sort((a, b) => compareVersions(b.tag_name, a.tag_name));
  return older[0] || null;
}

/** The newest published release, or null. */
export function newestRelease(releases) {
  return (releases || []).filter((r) => !r.draft && parseVersion(r.tag_name)).sort((a, b) => compareVersions(b.tag_name, a.tag_name))[0] || null;
}

/** A release's Windows files, as electron-builder.config.cjs names them. */
export function assetName(version, kind) {
  const v = parseVersion(version)?.text;
  if (!v) throw new Error(`not a version: ${version}`);
  if (kind === 'portable') return `Rutba-Office-${v}-portable.exe`;
  if (kind === 'installer') return `Rutba-Office-${v}-win-x64.exe`;
  throw new Error(`no file of kind ${kind}`);
}

/**
 * The profile a person who used the older version would have: settings
 * they changed, one marker no version knows (a setting is kept whether the
 * version knows it or not), a recent document, and the updates and the
 * notice board off, so the older version contacts nothing while it runs.
 * Answers what was put there, for judgeUpgrade.
 */
export function seedProfile(dir, { document, token }) {
  fs.mkdirSync(dir, { recursive: true });
  const settings = {
    theme: 'dark',
    'updates.automatic': false,
    'announcements.enabled': false,
    'upgradeTest.marker': token,
  };
  const recent = [{ path: document, name: path.basename(document), app: 'word', at: Date.now() }];
  fs.writeFileSync(path.join(dir, 'settings.json'), JSON.stringify(settings, null, 2));
  fs.writeFileSync(path.join(dir, 'recent.json'), JSON.stringify(recent, null, 2));
  return { settings, recent, document };
}

/** What the older version left in the profile: has it run over it yet? */
export function olderHasRun(dir, version, { document = null } = {}) {
  let settings = {};
  let recent = [];
  try {
    settings = JSON.parse(fs.readFileSync(path.join(dir, 'settings.json'), 'utf8'));
    recent = JSON.parse(fs.readFileSync(path.join(dir, 'recent.json'), 'utf8'));
  } catch {
    return false;
  }
  // A version records itself as the one that ran last as it starts (updates.js, since 2026-09-20).
  if (settings['updates.lastRun'] !== parseVersion(version)?.text) return false;
  return !document || recent.some((r) => samePath(r.path, document));
}

const samePath = (a, b) => (process.platform === 'win32' ? path.resolve(a).toLowerCase() === path.resolve(b).toLowerCase() : path.resolve(a) === path.resolve(b));

/**
 * Whether the newer version kept what the older one left, from its report
 * (apps/desktop/main/upgrade-report.js). Answers { ok, failures, evidence }:
 * each a list of sentences.
 */
export function judgeUpgrade({ from, to, seeded, report, feed = true }) {
  const failures = [];
  const evidence = [];
  const want = (ok, yes, no) => (ok ? evidence : failures).push(ok ? yes : no);
  const oldV = parseVersion(from)?.text;
  const newV = parseVersion(to)?.text;
  if (!report) return { ok: false, failures: ['the newer version wrote no report'], evidence };
  const s = report.settings || {};

  want(report.version === newV, `the newer copy is ${report.version}`, `the newer copy says it is ${report.version}, not ${newV}`);
  for (const [key, value] of Object.entries(seeded.settings)) {
    want(JSON.stringify(s[key]) === JSON.stringify(value), `setting ${key} kept: ${JSON.stringify(value)}`, `setting ${key} was ${JSON.stringify(value)} and is ${JSON.stringify(s[key])}`);
  }
  want(s['updates.lastRun'] === newV, `the version that ran last is now ${newV}`, `updates.lastRun is ${JSON.stringify(s['updates.lastRun'])}, not ${newV}`);
  const arrived = s['updates.arrived'];
  want(
    arrived && arrived.from === oldV && arrived.to === newV,
    `the arrival is recorded, ${oldV} to ${newV}, for the card that says so`,
    `the arrival is ${JSON.stringify(arrived ?? null)}, not from ${oldV} to ${newV}`
  );
  const kept = (report.recent || []).some((r) => samePath(r.path, seeded.document));
  want(kept, `the recent list still has ${path.basename(seeded.document)}`, `the recent list lost ${seeded.document} (it holds ${(report.recent || []).length})`);

  if (feed) {
    const f = report.feed;
    const files = f?.files || [];
    want(
      Boolean(f?.ok && parseVersion(f.version) && files.length && files.every((x) => x.url && x.sha512)),
      `the updater read the feed: ${f?.version}, ${files.map((x) => x.url).join(', ')}`,
      `the updater did not read the feed: ${f?.error || JSON.stringify(f)}`
    );
    // What it offers is not judged: a copy built before its release is newer than the feed.
  }
  return { ok: failures.length === 0, failures, evidence };
}

/**
 * The update that races the app's own closing (the 1.51 to 2.0 failure:
 * "could not be removed ... Uninstaller error 2"). The updater starts the
 * new installer the moment the app begins to quit, so the installer meets a
 * copy still running. `race` is 'close' (the app is closed 2 to 5 s after
 * the installer starts, as quitAndInstall does: the install must succeed)
 * or 'never-close' (the app is left running: the silent install must give
 * up cleanly, with a code other than 0, after its 15 s of retries, and
 * leave the copy that was there). `exit` is the installer's exit code, or
 * 'timeout' when it was stopped at `limitSeconds`; `copies` is registered()
 * after it. Answers { ok, failures, evidence }.
 */
export function judgeInstallRace({ race, exit, seconds, limitSeconds, copies, from, to, closedAfter = null, appEnded = null }) {
  const failures = [];
  const evidence = [];
  const want = (ok, yes, no) => (ok ? evidence : failures).push(ok ? yes : no);
  const oldV = parseVersion(from)?.text;
  const newV = parseVersion(to)?.text;
  const only = (v) => copies.length === 1 && copies[0].startsWith(`${v} `);
  const s = Math.round(seconds);
  if (race === 'close') {
    want(exit === 0, `the installer succeeded in ${s} s, the app closed ${closedAfter} s after it started`, exit === 'timeout' ? `the installer was still running after ${limitSeconds} s` : `the installer ended with ${exit} while the app was closing`);
    want(only(newV), `one copy registered, ${newV}`, `the registry has ${copies.length ? copies.join('; ') : 'no copy'}, not one copy of ${newV}`);
  } else if (race === 'never-close') {
    // Since 2.1.0 the installer closes a copy left open itself: its windows
    // are asked to close after 5 s and what is left is ended after 10, so
    // the install succeeds rather than waiting on a window nobody closes.
    want(exit !== 'timeout', `the installer ended by itself in ${s} s`, `the installer hung: still running after ${limitSeconds} s, and was stopped`);
    if (exit !== 'timeout') {
      want(exit === 0, `it ended with 0, having closed the copy left open`, `it ended with ${exit}: the copy left open was not closed for it`);
      want(seconds >= 5, `it waited ${s} s for the copy to close before closing it`, `it ended after ${s} s, without waiting 5 s for the copy to close by itself`);
    }
    if (appEnded !== null) want(appEnded, `the copy left open had ended when the installer finished`, 'the copy left open was still running when the installer finished');
    want(only(newV), `one copy registered, ${newV}`, `the registry has ${copies.length ? copies.join('; ') : 'no copy'}, not one copy of ${newV}`);
  } else {
    throw new Error(`no race ${race}`);
  }
  return { ok: failures.length === 0, failures, evidence };
}
