// How the window checks are cut into chunks for the lanes side by side: by
// the seconds each part took in earlier runs on this computer, never by a
// guess (PM, 11 Nov 2026). An app longer than a lane's fair share is split
// into pieces of its parts, each piece a check process of its own, so no one
// app holds the gate up at its end.
//
// A piece names its parts (RUTBA_VERIFY_PART_IDS), except one, which runs
// every part of the app the others do not (RUTBA_VERIFY_SKIP_IDS): a part
// added since the last measure runs there, so nothing is left out. An app is
// split only when every check it ran last time was inside a part or a section;
// a check outside them would run in every piece and be counted twice.

import fs from 'node:fs';

/**
 * The part and section ids verify-apps.js names now, read from its text:
 * `part('<app>', () => verifySomething(` and `section('<app>', '<name>'`. A
 * part named another way is not among them, so it is never named in a piece
 * and runs in the piece that runs the rest (review of 11 Nov: a part measured
 * and since renamed had a piece of its own that ran nothing, every run).
 */
export function sourceIds(text) {
  const ids = new Set();
  for (const m of String(text).matchAll(/\bpart\(\s*'([\w-]+)'\s*,\s*(?:async\s*)?\(\s*\)\s*=>\s*(?:await\s+)?([A-Za-z_$][\w$]*)\s*\(/g)) {
    // A machine part is the machine lane's: no piece of the parallel lane is cut to run it.
    const line = text.slice(m.index, text.indexOf('\n', m.index) === -1 ? undefined : text.indexOf('\n', m.index));
    if (/\bmachine:\s*true\b/.test(line)) continue;
    ids.add(`${m[1]}#${m[2]}`);
  }
  for (const m of String(text).matchAll(/\bsection\(\s*'([\w-]+)'\s*,\s*'([A-Za-z][\w-]*)'/g)) ids.add(`${m[1]}#${m[2]}`);
  return ids;
}

/** The measured seconds: { parts: { id: seconds }, apps: { app: { of, inParts, seconds, ids } } }. */
export function readSeconds(file) {
  try {
    const j = JSON.parse(fs.readFileSync(file, 'utf8'));
    return { parts: j.parts || {}, apps: j.apps || {} };
  } catch {
    return { parts: {}, apps: {} };
  }
}

export function writeSeconds(file, seconds) {
  const sorted = (o) => Object.fromEntries(Object.entries(o).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)));
  fs.writeFileSync(file, `${JSON.stringify({ about: "Seconds each window check part took when it last ran clean in the parallel lane, and whether each app's checks were all inside its parts (tools/gate-plan.js).", parts: sorted(seconds.parts), apps: sorted(seconds.apps) }, null, 1)}\n`);
}

/**
 * Note what a parallel-lane run measured. `apps` is a parts file's apps
 * (verify-apps.js), `whole` the apps that ran whole (no parts asked for by
 * id), whose count says whether all their checks were in parts.
 */
export function noteSeconds(seconds, apps, { whole = [] } = {}) {
  for (const [app, v] of Object.entries(apps || {})) {
    let inParts = 0;
    const ran = [];
    for (const [id, p] of Object.entries(v?.parts || {})) {
      // '*' is the checks outside any part: they are not in one (review of 11 Nov: counted in, the guard never held).
      if (p.machine || id === '*' || id.endsWith('#*')) continue;
      inParts += p.of || 0;
      if (p.of > 0) ran.push(id);
      // Clean only: a part that failed or threw may have waited out a timeout, or stopped short.
      if (p.of > 0 && !p.threw && !(p.failures?.length || p.failed?.length) && Number.isFinite(p.seconds)) seconds.parts[id] = Math.round(p.seconds * 10) / 10;
    }
    if (whole.includes(app) && v?.of > 0) {
      // The parts this whole run ran are the app's parts now: one measured before and gone since is forgotten.
      for (const id of Object.keys(seconds.parts)) if (id.startsWith(`${app}#`) && !ran.includes(id)) delete seconds.parts[id];
      seconds.apps[app] = { of: v.of, inParts, seconds: Math.round(v.seconds || 0), ids: ran.sort() };
    }
  }
  return seconds;
}

/** At most this many pieces of one app, and none shorter than this many seconds of parts. */
export const MAX_PIECES = 4;
export const MIN_PIECE = 120;

/**
 * The chunks for the parallel lanes: [{ label, app, ids, skip, seconds }].
 * `timing` is each app's measured seconds as a whole chunk (timing.json),
 * `seconds` the parts' (readSeconds), `jobs` the lanes. A chunk never
 * measured runs first: not knowing how long it takes, it must not be the one
 * left running at the end. The rest run longest first.
 */
export function planChunks(apps, { timing = {}, seconds = { parts: {}, apps: {} }, jobs = 3, current = null } = {}) {
  const known = apps.filter((a) => Number.isFinite(timing[a]));
  const share = Math.max(MIN_PIECE * 2, Math.ceil(known.reduce((t, a) => t + timing[a], 0) / Math.max(1, jobs)));
  const out = [];
  for (const app of apps) {
    if (!Number.isFinite(timing[app])) { out.push({ label: app, app, ids: null, skip: null, seconds: null }); continue; }
    const whole = seconds.apps[app];
    // Only parts the last whole run of the app ran, and that the checks still name (`current`, from sourceIds).
    const parts = Object.entries(seconds.parts)
      .filter(([id]) => id.startsWith(`${app}#`) && (whole?.ids || []).includes(id) && (!current || current.has(id)))
      .sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : 1));
    const partsSum = parts.reduce((t, [, s]) => t + s, 0);
    // Every check inside a part, each part measured, and long enough to be worth more than one process.
    const splittable = whole && whole.of > 0 && whole.of === whole.inParts && parts.length >= 2 && timing[app] > share * 1.25;
    const k = splittable ? Math.min(MAX_PIECES, parts.length, Math.ceil(timing[app] / share), Math.floor(partsSum / MIN_PIECE)) : 1;
    if (k < 2) { out.push({ label: app, app, ids: null, skip: null, seconds: timing[app] }); continue; }
    // Longest part first into the lightest piece; of two as light, the one with fewer parts, so a part of no seconds never leaves a piece empty.
    let bins = Array.from({ length: k }, () => ({ ids: [], load: 0 }));
    const lightest = () => bins.reduce((m, x) => (x.load < m.load || (x.load === m.load && x.ids.length < m.ids.length) ? x : m), bins[0]);
    for (const [id, s] of parts) {
      const b = lightest();
      b.ids.push(id);
      b.load += s;
    }
    // A piece with no part would run the app whole beside the others: none is made.
    bins = bins.filter((b) => b.ids.length);
    if (bins.length < 2) { out.push({ label: app, app, ids: null, skip: null, seconds: timing[app] }); continue; }
    // The lightest piece runs the rest: what the others do not name, a part new since the last measure among it.
    const rest = lightest();
    const named = bins.filter((b) => b !== rest).flatMap((b) => b.ids);
    const overhead = Math.max(0, timing[app] - partsSum);
    bins.forEach((b, i) => out.push({ label: `${app}.${i + 1}`, app, ids: b === rest ? null : b.ids, skip: b === rest ? named : null, seconds: Math.round(b.load + overhead) }));
  }
  return out.sort((a, b) => (b.seconds == null) - (a.seconds == null) || (a.seconds == null ? 0 : b.seconds - a.seconds));
}

/**
 * The pieces of one app as the app's own record, as one chunk would have
 * written it. A piece that crashed (no parts file) makes the app's parts
 * unknown, so its run again is the app whole.
 */
export function mergePieces(rs, app, leftToMachine) {
  const crashed = rs.some((r) => !r.parts?.[app]);
  const one = (r) => r.parts?.[app] || { passed: r.sum.passed, of: r.sum.of, failed: r.sum.failed };
  const ps = rs.map(one);
  const cat = (key) => ps.flatMap((p) => p[key] || []);
  return {
    ok: rs.every((r) => r.ok),
    passed: ps.reduce((t, p) => t + (p.passed || 0), 0),
    of: ps.reduce((t, p) => t + (p.of || 0), 0),
    failed: cat('failed'),
    realFailed: cat('realFailed'),
    passedNames: cat('passedNames'),
    // A piece that crashed leaves its parts unknown: the app whole is run again (review of 11 Nov: [] here let a crashed app clear on another's run again).
    failedParts: crashed ? ['*'] : cat('failedParts'),
    parts: crashed ? {} : Object.assign({}, ...ps.map((p) => p.parts || {})),
    toMachine: [...new Set(rs.flatMap((r) => leftToMachine(r.seen)))],
    seconds: Math.round(rs.reduce((t, r) => t + r.seconds, 0)),
    pieces: rs.length,
    // A piece ended as hung: where it stopped, for the failure list.
    hung: rs.find((r) => r.hung)?.hung || null,
  };
}
