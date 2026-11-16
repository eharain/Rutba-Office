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

/** Chromium's own folders in a profile: they come and go with no meaning for a person's work. */
const BROWSER_PARTS = /^(cache|code cache|gpucache|dawncache|dawngraphitecache|dawnwebgpucache|shadercache|grshadercache|blob_storage|session storage|local storage|network|shared dictionary|service worker|webstorage|crashpad|sentry|logs?)$/i;

/** Every file of Office's own in a profile, as a path from its folder. */
export function profileFiles(dir) {
  const out = [];
  const walk = (at, rel, depth) => {
    for (const e of fs.readdirSync(at, { withFileTypes: true })) {
      const name = rel ? `${rel}/${e.name}` : e.name;
      if (e.isDirectory()) {
        if (depth > 0 || !BROWSER_PARTS.test(e.name)) walk(path.join(at, e.name), name, depth + 1);
      } else if (depth > 0 || /\.(json|db|sqlite|key|bin)$/i.test(e.name)) {
        out.push(name);
      }
    }
  };
  walk(dir, '', 0);
  return out.sort();
}

/** What an upgrade took away: files the older version left that are gone after the newer one ran. */
export function goneFromProfile(before, after) {
  const now = new Set(after);
  // A lock, a temporary or a journal is meant to go.
  return before.filter((f) => !now.has(f) && !/(\.tmp|\.lock|-journal|-wal|-shm)$/i.test(f));
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
  // A candidate (2.26.0-rc.1) is held to what its release will be.
  const released = newV ? newV.split('-')[0] : newV;
  if (!report) return { ok: false, failures: ['the newer version wrote no report'], evidence };
  const s = report.settings || {};

  want(report.version === newV, `the newer copy is ${report.version}`, `the newer copy says it is ${report.version}, not ${newV}`);
  // Reports before 2.25.0 were written before any window opened, and cannot say.
  if (compareVersions(released, '2.25.0') >= 0) {
    want(report.window?.visible === true && report.window?.drawn === true, 'its first window came up and drew itself', `its first window did not come up: ${JSON.stringify(report.window ?? null)}`);
  }
  // From 2.26.0 a process the speech engine starts is held from its first import, by the core's bootstrap inside the archive.
  if (compareVersions(released, '2.26.0') >= 0) {
    const c = report.engineChild;
    const held = Boolean(c && !c.error && c.first === 'engine' && /only what it carries/.test(String(c.above || '')));
    const archived = report.packaged === false || c?.fromArchive === true;
    want(
      held && archived,
      `a process the speech engine starts is held to its folder, started by ${c?.bootstrap ? path.basename(c.bootstrap) : 'the bootstrap'} from the archive`,
      `a process the speech engine starts is not shown held: ${JSON.stringify(c ?? null)}`
    );
  }
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
 * Whether Windows' entry for the copy is true after an install. `entries`
 * is every Rutba Office uninstall entry found, each { scope: 'user' or
 * 'machine', version, installLocation, works } (works: the program and
 * resources\app.asar are in installLocation); `scope` is the scope the copy
 * was installed to. The entry there must say the version installed and, from
 * 2.26.0 (electron-builder writes no InstallLocation, the installer since
 * then does), name a folder that holds a working copy; the other scope must
 * keep none (a stale 2.18.0 per-user entry was left behind a per-user copy
 * of 2.26.0). Answers { ok, failures, evidence }.
 */
export function judgeRegistration({ to, scope, entries }) {
  const failures = [];
  const evidence = [];
  const want = (ok, yes, no) => (ok ? evidence : failures).push(ok ? yes : no);
  const newV = parseVersion(to)?.text;
  const released = newV ? newV.split('-')[0] : newV;
  const other = scope === 'machine' ? 'user' : 'machine';
  const here = (entries || []).filter((e) => e.scope === scope);
  const there = (entries || []).filter((e) => e.scope === other);
  const say = (e) => `${e.version || 'no version'} (${e.scope}, ${e.installLocation ? `in ${e.installLocation}` : 'no InstallLocation'})`;
  want(here.length === 1, `one entry for ${scope}`, here.length ? `${here.length} entries for ${scope}: ${here.map(say).join('; ')}` : `no entry for ${scope}, where the copy was installed`);
  const e = here[0];
  if (e) {
    want(e.version === newV, `the ${scope} entry says ${newV}`, `the ${scope} entry says ${e.version || 'no version'}, not ${newV}`);
    if (released && compareVersions(released, '2.26.0') >= 0) {
      want(Boolean(e.installLocation) && e.works === true, `it names its folder, ${e.installLocation}, which holds the copy`, e.installLocation ? `it names ${e.installLocation}, which holds no working copy` : 'it names no folder (InstallLocation is empty)');
    }
  }
  want(there.length === 0, `no entry is left for ${other}`, `an entry is left for ${other}: ${there.map(say).join('; ')}`);
  return { ok: failures.length === 0, failures, evidence };
}

/** Whether a link's command opened a Documents window: its title names the app, "Rutba Documents" since 1.29.0, "Rutba Word" before. */
export const documentsOpened = (titles) => (titles || []).some((t) => /Rutba (Documents|Word)/.test(String(t)));

/**
 * An install over a copy already there, judged (owner, 16 Nov,
 * D-2026-11-16-03: "the installer should detect these changes and
 * variations"): whatever kinds and folders were there before, the install
 * keeps the kind of copy there is (`scope`, 'machine' or 'user') and its
 * folder (`dir`, a custom one too). `folders` is each folder looked at,
 * { dir, scope, exe, asar } (exe: the program is there; asar:
 * resources\app.asar beside it); `link` is an app pin made from the copy
 * before the install, read after the newer copy has run once (its pin repair
 * may repoint it), { target, arguments, exe, asar }; `titles` the windows a
 * cold click on it opened. The entry must be the one judgeRegistration asks
 * for and, from 2.26.0, name `dir`; `dir` must hold a whole copy; no folder
 * may hold the program alone (a husk, which a pin starts and which ends at
 * once), and no other folder a second whole copy; and the pin must still
 * open Documents. `husksLeft` names folders where a husk is expected to stay
 * (an install that is not elevated leaves one for all users rather than
 * raise a prompt); `held`, when a helper of the older copy was left running
 * from its program, says whether the install ended it ({ ended }).
 */
export function judgeLayout({ to, scope, dir, entries, folders, link, titles, husksLeft = [], held = null }) {
  const judged = judgeRegistration({ to, scope, entries });
  const failures = [...judged.failures];
  const evidence = [...judged.evidence];
  const want = (ok, yes, no) => (ok ? evidence : failures).push(ok ? yes : no);
  const same = (a, b) => String(a || '').toLowerCase().replace(/\\+$/, '') === String(b || '').toLowerCase().replace(/\\+$/, '');
  const who = scope === 'machine' ? 'all users' : 'this person';
  const released = parseVersion(to)?.text?.split('-')[0];
  const entry = (entries || []).find((e) => e.scope === scope);
  if (entry && released && compareVersions(released, '2.26.0') >= 0) {
    want(same(entry.installLocation, dir), `its folder is the one the copy was in, ${dir}`, `its folder moved: ${entry.installLocation || 'none'}, not ${dir}`);
  }
  const here = (folders || []).find((f) => same(f.dir, dir));
  want(Boolean(here?.exe && here?.asar), `the copy is still for ${who}, in ${dir}`, `there is no working copy for ${who} in ${dir}`);
  const allowed = (f) => husksLeft.some((d) => same(d, f.dir));
  const husks = (folders || []).filter((f) => f.exe && !f.asar && !allowed(f));
  want(husks.length === 0, 'no folder holds the program without the rest', `a husk (the program, no resources\\app.asar): ${husks.map((f) => f.dir).join('; ')}`);
  for (const d of husksLeft) {
    const f = (folders || []).find((x) => same(x.dir, d));
    if (f?.exe && !f.asar) evidence.push(`the husk in ${d} is left, as an install that is not elevated must leave it`);
  }
  if (held) want(held.ended === true, 'the helper left running from the older program was ended by the install', 'the helper left running from the older program is still running after the install');
  const others = (folders || []).filter((f) => f.exe && f.asar && !same(f.dir, dir));
  want(others.length === 0, 'no second copy is left anywhere else', `a second copy is left: ${others.map((f) => `${f.dir} (${f.scope})`).join('; ')}`);
  want(Boolean(link?.exe && link?.asar), `the pin's target is a whole copy: ${link?.target}`, link?.exe ? `the pin starts a husk: ${link.target}` : `the pin's target is gone: ${link?.target}`);
  want(/--app=word\b/.test(String(link?.arguments || '')), 'the pin still opens Documents (--app=word)', `the pin lost its --app=word: ${JSON.stringify(link?.arguments ?? null)}`);
  want(documentsOpened(titles), `a cold click on the pin opened ${(titles || []).find((t) => /Rutba (Documents|Word)/.test(t))}`, `a cold click on the pin opened no Documents window (${(titles || []).length ? titles.join(' | ') : 'no window'})`);
  return { ok: failures.length === 0, failures, evidence };
}

/**
 * The owner's 2.26.0 fault, judged: an all-users copy, a stale per-user
 * entry left by 2.18.0, and an app pin made from the all-users copy; then
 * the newer installer run silently, as an update. The copy must still be
 * for all users, in `machineDir`, as judgeLayout asks of any layout.
 */
export function judgeStaleUserEntry({ to, entries, machineDir, folders, link, titles }) {
  return judgeLayout({ to, scope: 'machine', dir: machineDir, entries, folders, link, titles });
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
