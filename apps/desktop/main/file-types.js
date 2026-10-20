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
