// Windows is told only about the kinds of file the chosen apps open
// (Choose apps, owner, 2026-10-20). The installer registers every type, as
// the classes electron-builder writes: `.ext` names a class, and
// `.ext\OpenWithProgids` lists it for Open with. For an app turned off this
// takes our class out of both, so its files go back to whatever opened them
// before; for an app turned on again it puts them back, without taking a
// type another program has since been given.
//
// Only a copy installed for this person (HKCU) is changed here: a copy
// installed for everyone keeps its types in HKLM, which needs the installer.

import { execFile } from 'node:child_process';
import { PACKS } from '@rutba/office-shell/contract';

export const CLASSES = 'HKCU\\Software\\Classes';

function reg(args) {
  return new Promise((resolve) => {
    execFile('reg', args, { windowsHide: true }, (err, stdout) => resolve({ ok: !err, out: String(stdout || '') }));
  });
}

/**
 * A key's default value, '' when it has none, or null when the key is not
 * there. An unset default is printed as "(value not set)", in the system's
 * language, and a class name is never in brackets: so a bracketed value is none.
 */
async function defaultOf(key) {
  const r = await reg(['query', key, '/ve']);
  if (!r.ok) return null;
  const v = /REG_(?:EXPAND_)?SZ\s+(.*?)\s*$/m.exec(r.out)?.[1] ?? '';
  return /^\(.*\)$/.test(v) ? '' : v;
}

/**
 * Each type's class, and whether a chosen app still opens it.
 * @param {{ ext: string, name: string }[]} all every type the suite registers
 * @param {{ ext: string }[]} kept the types the chosen apps open
 */
export function fileTypePlan(all, kept) {
  const keep = new Set(kept.map((a) => a.ext));
  return all.map((a) => ({ ext: a.ext, progId: a.name, keep: keep.has(a.ext) }));
}

/**
 * Brings the registry in line with the plan.
 * @returns {Promise<{ withdrawn: string[], restored: string[], notOurs: string[] }>}
 */
export async function applyFileTypes(plan, { root = CLASSES } = {}) {
  const ours = new Map();
  const isOurs = async (progId) => {
    if (!ours.has(progId)) ours.set(progId, (await reg(['query', `${root}\\${progId}\\shell\\open\\command`, '/ve'])).ok);
    return ours.get(progId);
  };
  const out = { withdrawn: [], restored: [], notOurs: [] };
  for (const { ext, progId, keep } of plan) {
    if (!(await isOurs(progId))) { out.notOurs.push(ext); continue; }
    const key = `${root}\\.${ext}`;
    const listed = (await reg(['query', `${key}\\OpenWithProgids`, '/v', progId])).ok;
    const current = await defaultOf(key);
    if (keep) {
      let changed = false;
      if (!listed) { await reg(['add', `${key}\\OpenWithProgids`, '/v', progId, '/t', 'REG_NONE', '/f']); changed = true; }
      // Ours again only where nothing else has taken the type meanwhile.
      if (!current) { await reg(['add', key, '/ve', '/d', progId, '/f']); changed = true; }
      if (changed) out.restored.push(ext);
    } else {
      let changed = false;
      if (listed) { await reg(['delete', `${key}\\OpenWithProgids`, '/v', progId, '/f']); changed = true; }
      if (current === progId) { await reg(['delete', key, '/ve', '/f']); changed = true; }
      if (changed) out.withdrawn.push(ext);
    }
  }
  return out;
}

/** Where the installer's Choose apps page and Office keep the apps turned off, for each other. */
export const CHOICE_KEY = 'HKCU\\Software\\Rutba Office';

/**
 * The apps the installer's page left unticked and those it left ticked, if it
 * was the last to choose: { off, on, addons } or null. `on` is null from an
 * installer older than AppsOn, which wrote the unticked ones only.
 */
export async function installerChoice({ key = CHOICE_KEY } = {}) {
  const by = await reg(['query', key, '/v', 'AppsChosenAt']);
  if (!by.ok || !/REG_SZ\s+installer\s*$/m.test(by.out)) return null;
  const list = async (name) => {
    const r = await reg(['query', key, '/v', name]);
    if (!r.ok) return null;
    const raw = new RegExp(`${name}\\s+REG_SZ\\s+(.*?)\\s*$`, 'm').exec(r.out)?.[1] ?? '';
    return raw.split(',').map((s) => s.trim()).filter(Boolean);
  };
  return { off: (await list('AppsOff')) || [], on: await list('AppsOn'), addons: (await list('AddOns')) || [] };
}

/**
 * The apps turned off and those turned on, written where the installer's page
 * reads them, as chosen in Office. Both, so an installer with an app this copy
 * has never heard of still shows exactly the apps that are on now.
 */
export async function recordChoice(off, { key = CHOICE_KEY, addons = [], on = null } = {}) {
  await reg(['add', key, '/v', 'AppsOff', '/t', 'REG_SZ', '/d', off.join(','), '/f']);
  if (on) await reg(['add', key, '/v', 'AppsOn', '/t', 'REG_SZ', '/d', on.join(','), '/f']);
  await reg(['add', key, '/v', 'AddOns', '/t', 'REG_SZ', '/d', addons.join(','), '/f']);
  await reg(['add', key, '/v', 'AppsChosenAt', '/t', 'REG_SZ', '/d', 'office', '/f']);
}

/**
 * The kinds of file an app pack opens, registered as the installer registers
 * them (electron-builder's APP_ASSOCIATE): `.ext` names the class and lists it
 * for Open with, and the class says how the file is opened and what it
 * looks like. For a pack added by Add or remove apps in a copy installed for
 * this person; one installed for everyone is the installer's to change.
 * @param {{ ext, name, description, icon }[]} types  icon: a path on disk
 * @param {{ exe: string, root?: string }} o
 */
export async function registerTypes(types, { exe, root = CLASSES } = {}) {
  const done = [];
  for (const t of types) {
    const progId = t.name;
    const key = `${root}\\.${t.ext}`;
    const current = await defaultOf(key);
    // A type another program has been given keeps it; Rutba Office is still offered under Open with.
    if (!current) await reg(['add', key, '/ve', '/d', progId, '/f']);
    await reg(['add', `${key}\\OpenWithProgids`, '/v', progId, '/t', 'REG_NONE', '/f']);
    await reg(['add', `${root}\\${progId}`, '/ve', '/d', t.description || t.name, '/f']);
    if (t.icon) await reg(['add', `${root}\\${progId}\\DefaultIcon`, '/ve', '/d', t.icon, '/f']);
    await reg(['add', `${root}\\${progId}\\shell`, '/ve', '/d', 'open', '/f']);
    await reg(['add', `${root}\\${progId}\\shell\\open`, '/ve', '/d', 'Open with Rutba Office', '/f']);
    await reg(['add', `${root}\\${progId}\\shell\\open\\command`, '/ve', '/d', `"${exe}" "%1"`, '/f']);
    done.push(t.ext);
  }
  return done;
}

/** The packs installed, where the next installer reads them (HKCU\\Software\\Rutba Office, Packs). */
export async function recordPacks(ids, { key = CHOICE_KEY, known = PACKS.map((p) => p.id) } = {}) {
  await reg(['add', key, '/v', 'Packs', '/t', 'REG_SZ', '/d', ids.join(','), '/f']);
  // Every pack this version knows, so the next installer adds only a pack new in its own (PacksKnown).
  await reg(['add', key, '/v', 'PacksKnown', '/t', 'REG_SZ', '/d', known.join(','), '/f']);
}
