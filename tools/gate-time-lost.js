// What the waits that gave up cost, by the check file that asked for them.
//
// A wait that gives up says so in its lane's log, with how long it took and
// where it was asked for (verify-waits.js). On 12 Nov 233 of them cost 3,511
// s of lane time, half of it, and nothing in the gate's own output said so:
// they were found by reading the logs by hand. The gate now prints this table
// after the window checks, so a new slow wait shows the day it comes in.

const LINE = /\[wait\] gave up on (.+?) after (\d+) ms(?: \(([^)]*)\))?/;

/**
 * The check file a wait belongs to, from where it was asked for: the last
 * file named in "inner.js:50 < check.js:128" (a bare ":128" is the file
 * before it), so a shared helper's wait is put on the check that used it.
 */
export function checkFileOf(at) {
  // verify-apps.js runs every part, so it is named only when no other file is.
  let file = null;
  let runner = null;
  for (const step of String(at || '').split('<').map((s) => s.trim()).filter(Boolean)) {
    const m = /^([^:\s]+\.[cm]?js)?:\d+/.exec(step);
    if (!m?.[1]) continue;
    if (m[1] === 'verify-apps.js') runner = m[1];
    else file = m[1];
  }
  return file || runner || 'unknown';
}

/** A wait's name with its numbers taken out, so "the window at 1680 px" and "at 900 px" count as one. */
const kind = (what) => what.replace(/\d+(?:x\d+)?/g, 'N');

/** The waits that gave up in these logs' text, totalled by check file, the costliest first. */
export function timeLost(texts) {
  const by = new Map();
  let ms = 0;
  let waits = 0;
  for (const text of texts) {
    for (const line of String(text).split(/\r?\n/)) {
      const m = LINE.exec(line);
      if (!m) continue;
      const took = Number(m[2]);
      const file = checkFileOf(m[3]);
      const row = by.get(file) || { file, ms: 0, waits: 0, kinds: new Map() };
      row.ms += took;
      row.waits += 1;
      const k = kind(m[1]);
      row.kinds.set(k, (row.kinds.get(k) || 0) + 1);
      by.set(file, row);
      ms += took;
      waits += 1;
    }
  }
  const rows = [...by.values()].sort((a, b) => b.ms - a.ms).map((r) => ({
    file: r.file,
    ms: r.ms,
    waits: r.waits,
    most: [...r.kinds.entries()].sort((a, b) => b[1] - a[1])[0],
  }));
  return { ms, waits, rows };
}

/** The table as the gate prints it: the `top` costliest files and a line for the rest. */
export function timeLostLines({ ms, waits, rows }, top = 10) {
  if (!waits) return ['  time lost to waits that gave up: none'];
  const s = (n) => `${Math.round(n / 1000).toLocaleString('en-GB')} s`;
  const out = [`  time lost to waits that gave up: ${waits} wait${waits === 1 ? '' : 's'}, ${s(ms)} of lane time`];
  for (const r of rows.slice(0, top)) out.push(`    ${s(r.ms).padStart(8)} ${String(r.waits).padStart(4)}  ${r.file}  (most: ${r.most[0]}, ${r.most[1]})`);
  const rest = rows.slice(top);
  if (rest.length) out.push(`    ${s(rest.reduce((t, r) => t + r.ms, 0)).padStart(8)} ${String(rest.reduce((t, r) => t + r.waits, 0)).padStart(4)}  ${rest.length} other files`);
  return out;
}
