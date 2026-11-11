// Each window check's own number and the place in the code it was made.
//
// Owner, 11 Nov 2026: a failure must say exactly where it failed, with a
// number beside its words, so it can be found, tracked and closed. A check's
// number is `<app>.<check>`, `18.0142` for Mail's check 142: the app's number
// is fixed in tools/check-ids.json, and each check's number is given once,
// the first time the gate sees the check, and never reused. The checks run
// with the registry as it was when they started; one the registry does not
// know yet says `new`, and the gate gives it a number at the end of the run.
//
// The place is the file and line of the check() call, read from the stack,
// so `verify-mail-signin.js:54` points at the line that judged it.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
/** The registry of numbers, in the repository: committed, so a number is the same on every machine. */
export const REGISTRY = path.join(HERE, '..', '..', '..', 'tools', 'check-ids.json');

/** The registry as it is on disk, or an empty one. */
export function readRegistry(file = REGISTRY) {
  try {
    const r = JSON.parse(fs.readFileSync(file, 'utf8'));
    return { apps: r.apps || {}, next: r.next || {}, ids: r.ids || {} };
  } catch {
    return { apps: {}, next: {}, ids: {} };
  }
}

/** A check's number, or 'new' when the registry does not know it yet. */
export function idOf(registry, name) {
  return registry.ids[name] || 'new';
}

/**
 * Give a number to each name the registry does not know, under `app`.
 * Answers the names given numbers. Apps are numbered in the order given
 * by `apps` the first time; an app added later takes the next number.
 */
export function assign(registry, entries, apps = []) {
  const added = [];
  for (const a of apps) {
    if (registry.apps[a] == null) registry.apps[a] = Math.max(0, ...Object.values(registry.apps)) + 1;
  }
  for (const { name, app } of entries) {
    if (!name || registry.ids[name]) continue;
    const a = app && registry.apps[app] != null ? app : 'shell';
    if (registry.apps[a] == null) registry.apps[a] = Math.max(0, ...Object.values(registry.apps)) + 1;
    const n = (registry.next[a] || 1);
    registry.next[a] = n + 1;
    registry.ids[name] = `${String(registry.apps[a]).padStart(2, '0')}.${String(n).padStart(4, '0')}`;
    added.push(name);
  }
  return added;
}

/** Write the registry with its keys in a steady order, so a diff shows only what was added. */
export function writeRegistry(registry, file = REGISTRY) {
  const sorted = (o) => Object.fromEntries(Object.entries(o).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)));
  const out = { about: 'Window check numbers: <app>.<check>, given once and never reused (tools/gate.js gives new ones at the end of a run).', apps: registry.apps, next: sorted(registry.next), ids: sorted(registry.ids) };
  fs.writeFileSync(file, `${escapeBanned(JSON.stringify(out, null, 1))}\n`);
}

/**
 * The characters the repository's commit hook refuses in an added line (the
 * long dash, zero-width and direction marks, word joiner, no-break space,
 * byte-order mark), written as JSON escapes: some checks' names hold them,
 * and the name read back is the same.
 */
export const escapeBanned = (text) => text.replace(/[\u2014\u200b\u200c\u200e\u200f\u2060\u00a0\ufeff]/g, (c) => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`);

/**
 * The file and line that called the function `skip` frames above this one:
 * `verify-mail-signin.js:54`, or null when the stack says nothing useful.
 * With `depth` 2, the caller's caller too, `verify-mail-signin.js:31 < :53`,
 * for a wait asked for through a helper of the check's own.
 */
export function callSite(stack, skip = 2, depth = 1) {
  const frames = String(stack || '').split('\n').slice(1);
  const out = [];
  for (const f of frames.slice(skip - 1)) {
    const m = /([\w.-]+\.(?:m?js)):(\d+):\d+\)?\s*$/.exec(f);
    if (!m || /^verify-(ids|waits)\.js$/.test(m[1])) continue;
    const here = `${m[1]}:${m[2]}`;
    if (out.includes(here)) continue;
    // The same file again is said by its line alone.
    out.push(out.length && out[0].startsWith(`${m[1]}:`) ? `:${m[2]}` : here);
    if (out.length >= depth) break;
  }
  return out.length ? out.join(' < ') : null;
}
