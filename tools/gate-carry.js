// A part of the gate keeps its pass until what it runs changes (the owner,
// 16 Nov 2026, D-2026-11-16-02: "A gate pass is kept by what changed: each
// part of the gate keeps its pass until its own code, the shared code, its
// checks or the packs change; a change to shared code still runs everything
// again").
//
// The gate's state was for one commit: a fix to one app, once committed,
// voided the pass of every other. Here a state kept for an earlier commit is
// brought to this one (npm run gate -- --resume, which npm run gate:groups
// uses for every unit), and each part keeps its pass only when no file that
// changed between the two commits is one it runs:
//
//   the window checks of an app   by appsForPath (tools/gate.js), the map
//                                 --changed uses, held to what each app
//                                 really imports by tests/gate-apps.test.js;
//                                 a file every app shares voids them all
//   the machine lane              only when no app's files changed
//   the engine tests              never: they read the whole checkout
//   every other pass              past a change to words alone (docs, notes,
//                                 workflows, the check numbers); the passes
//                                 that open no window and import nothing of
//                                 the windows' own (ENGINE_ONLY, held by
//                                 tests/gate-carry.test.js) also past a
//                                 change to the windows' own files
//   anything at all               never past a change to the gate itself, or
//                                 to a release note (the build takes the
//                                 note into the app)
//
// Nothing is carried unless both commits were checked with nothing in the
// tree beyond them, the older is an ancestor of the newer, and git could
// list what changed. A part carried keeps the commit it ran on (`at`), the
// run says it, and the record says it in the release note.
//
// The packs the window checks add (RUTBA_GATE_PACKS) are not in the
// checkout: packsKey names them by their signed envelopes, and a state
// whose window checks ran with other packs loses those passes.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

/** treeKey (tools/gate.js) of a checkout that holds nothing beyond its commit. */
export const CLEAN_TREE = crypto.createHash('sha256').update('').digest('hex');

/**
 * The passes that open no window and import nothing of the windows' own
 * files (apps/desktop/renderer, apps/desktop/main/verify-*): a change there
 * is not theirs. tests/gate-carry.test.js walks what each one imports and
 * fails when one comes to import such a file.
 */
export const ENGINE_ONLY = Object.freeze(['test:perf', 'fuzz:open', 'fuzz:pdf', 'fuzz:smime', 'fuzz:grammar']);

/** Whether a file is one of the windows' own: the renderer, and the window checks. */
export const isWindowsOwn = (file) => /^apps\/desktop\/renderer\/|^apps\/desktop\/main\/verify-/.test(String(file).replace(/\\/g, '/'));

/**
 * What a changed file is to the gate's passes:
 *   'gate'     the gate itself, or a release note: nothing is carried past it
 *   'words'    read by the engine tests alone
 *   'windows'  the windows' own files
 *   'code'     anything else: every pass that is not a window check runs again
 */
export function kindOf(file) {
  const f = String(file).replace(/\\/g, '/');
  if (/^tools\/(?:gate|run-temp\.js$)/.test(f) || /^docs\/releases\//.test(f)) return 'gate';
  if (f === 'tools/check-ids.json' || /^docs\//.test(f) || /^[^/]+\.md$/.test(f) || (/^\.github\//.test(f) && f !== '.github/workflows/release.yml')) return 'words';
  if (isWindowsOwn(f)) return 'windows';
  return 'code';
}

/** Whether a pass that is not a window check keeps its pass over these changed files. */
export function passKept(name, files) {
  if (!files.length) return true;
  if (name === 'test') return false;
  return files.map(kindOf).every((k) => k === 'words' || (k === 'windows' && ENGINE_ONLY.includes(name)));
}

/**
 * A state kept for an earlier commit, brought to `head` over the files that
 * changed in between: each part that had passed keeps its pass, with the
 * commit it ran on, only if none of those files is one it runs. What had not
 * passed is left out, and runs again as it would have.
 * @param {object} old the state as kept: { head, passes, apps, machine, ... }
 * @param {{ head: string, files: string[], appsForPath: (file: string) => string[], apps: string[] }} o
 * @returns {{ state: object|null, kept: string[], again: string[] }} state null when nothing is kept
 */
export function carry(old, { head, files, appsForPath, apps }) {
  const ranOn = (entry) => entry.at || old.head;
  const state = { ...old, head, passes: {}, apps: {}, machine: null };
  const kept = [];
  const again = [];
  const voided = files.some((f) => kindOf(f) === 'gate');
  const touched = new Set(files.flatMap((f) => appsForPath(f)));
  const everyApp = voided || touched.has('shell');
  for (const [name, entry] of Object.entries(old.passes || {})) {
    if (!entry?.ok) continue;
    if (!voided && passKept(name, files)) { state.passes[name] = { ...entry, at: ranOn(entry) }; kept.push(name); } else again.push(name);
  }
  for (const app of apps) {
    const entry = old.apps?.[app];
    if (!entry?.ok) continue;
    if (!everyApp && !touched.has(app)) { state.apps[app] = { ...entry, at: ranOn(entry) }; kept.push(app); } else again.push(app);
  }
  // The machine lane ran the apps together: it stands only while no app's files changed.
  if (old.machine?.ok && !old.machine.serial) {
    if (!everyApp && touched.size === 0) { state.machine = { ...old.machine, at: ranOn(old.machine) }; kept.push('machine'); } else again.push('machine');
  }
  return { state: kept.length ? state : null, kept, again };
}

/**
 * The whole step, with git asked through `git`: why nothing is carried, in a
 * few words, or the state and what to say of it.
 * @param {object} old the state file as it is
 * @param {{ head: string, tree: string, git: { isAncestor: (a: string, b: string) => boolean, changed: (a: string, b: string) => string[]|null }, appsForPath: Function, apps: string[] }} o
 * @returns {{ state: object|null, why?: string, lines: string[] }}
 */
export function carryOver(old, { head, tree, git, appsForPath, apps }) {
  const none = (why) => ({ state: null, why, lines: [`Nothing is carried from the run on ${short(old?.head)}: ${why}.`] });
  if (!old?.head || typeof old.head !== 'string') return { state: null, why: 'no run before', lines: [] };
  if (old.head === head) return none('it was on this commit with other changes in the tree');
  if (old.tree !== CLEAN_TREE) return none('it ran on a tree with changes beyond its commit');
  if (tree !== CLEAN_TREE) return none('this tree has changes beyond its commit');
  if (!git.isAncestor(old.head, head)) return none('this commit does not come after it');
  const files = git.changed(old.head, head);
  if (!files) return none('git could not list what changed since');
  const r = carry(old, { head, files, appsForPath, apps });
  if (!r.state) return none(`${files.length === 1 ? 'the file' : `the ${files.length} files`} changed since ${files.length === 1 ? 'is' : 'are'} run by every part that had passed`);
  const lines = [`Carried from the run on ${short(old.head)} (${files.length} ${files.length === 1 ? 'file' : 'files'} changed since, none of them run by these): ${r.kept.join(', ')}.`];
  if (r.again.length) lines.push(`Passed there and run again here, since a file they run changed: ${r.again.join(', ')}.`);
  return { state: r.state, lines };
}

const short = (sha) => String(sha || 'no commit').slice(0, 8);

/** The passes, in the words of a release note. */
const PASS_WORDS = {
  test: 'the engine tests',
  'test:perf': 'the performance budgets',
  'test:removal': 'the folder-removal suites',
  'verify:edit': 'the editing checks',
  smoke: 'the windows painted clean',
  'verify:packs': 'the app packs added and removed',
  'office:reopen': 'the saves reopened in Word, Excel and PowerPoint',
  'fuzz:open': 'the damaged Office files',
  'fuzz:pdf': 'the damaged PDFs',
  'fuzz:smime': 'the damaged signed mail',
  'fuzz:grammar': 'the files made to do harm',
  corpus: 'the files from outside the project',
};

const listed = (items) => (items.length > 1 ? `${items.slice(0, -1).join(', ')} and ${items.at(-1)}` : items[0] || '');

/**
 * What the record says of the parts that were not run on the commit checked:
 * each earlier commit and what passed on it, or '' when everything ran here.
 */
export function carriedNote(state) {
  const by = new Map();
  const put = (at, kind, name) => {
    if (!at || at === state.head) return;
    if (!by.has(at)) by.set(at, { passes: [], apps: [], machine: false });
    if (kind === 'machine') by.get(at).machine = true; else by.get(at)[kind].push(name);
  };
  for (const [name, e] of Object.entries(state.passes || {})) if (e?.ok) put(e.at, 'passes', PASS_WORDS[name] || name);
  for (const [app, e] of Object.entries(state.apps || {})) if (e?.ok) put(e.at, 'apps', app);
  if (state.machine?.ok) put(state.machine.at, 'machine');
  if (!by.size) return '';
  const each = [...by].map(([at, p]) => {
    const parts = [
      ...(p.apps.length ? [`the window checks of ${listed(p.apps)}`] : []),
      ...(p.machine ? ["the checks that need the computer's own clipboard and keys"] : []),
      ...p.passes,
    ];
    return `on ${at}, ${listed(parts)}`;
  });
  return `Not every part was run on the commit checked. A part keeps its pass from an earlier commit while no file it runs has changed since, and these were kept: ${each.join('; ')}. Everything else was run on the commit checked.`;
}

/**
 * The packs a run's window checks add, named by their signed envelopes (each
 * lists its files with their digests): null with no release folder.
 */
export function packsKey(release, ids) {
  if (!release) return null;
  const hash = crypto.createHash('sha256');
  for (const id of ids.slice().sort()) {
    hash.update(`\0${id}\0`);
    try { hash.update(fs.readFileSync(path.join(release, `${id}.pack.json`))); } catch { hash.update('none'); }
  }
  return hash.digest('hex');
}
