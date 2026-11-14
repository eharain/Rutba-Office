// The gate's failure list: every window check that failed, kept until it is
// closed, and a full run on the record refused while any is open.
//
// Owner, 11 Nov 2026: "it should find all the failures and then fix the
// failure before re running". The gate used to be run, closed by one failure,
// fixed there and run again for an hour, and closed by the next. Now each run
// adds what failed to the list (.gate/ledger.json) with its number, the line
// that judged it, its part, its lane and the wait that gave up before it, and
// marks one that passed only when run again alone as flaky.
//
// An entry is one check in one part of one app: the same words failing in two
// apps (a console catch-all) are two entries, each run again in its own app.
// It is closed by a commit made after it failed, once a run on a clean tree
// at or after that commit passed it (twice, for a flaky one); or by a reason
// written down, which the next record's Gate section prints (PM, 11 Nov).
// A check the run makes up itself (a part that threw, a console error, a
// window process that ended early) passes when its part runs clean.

import fs from 'node:fs';
import { saidLoad } from '../apps/desktop/main/verify-load.js';

/** An empty list. */
export const emptyLedger = () => ({ entries: {}, runs: [], next: 1 });

/**
 * The list on disk, or an empty one when there is none. A list that is there
 * and does not read throws: read as empty, it would let --record past every
 * failure it held.
 */
export function readLedger(file) {
  let text;
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch (err) {
    if (err.code === 'ENOENT') return emptyLedger();
    throw new Error(`the failure list ${file} could not be read: ${err.message}`);
  }
  let l;
  try {
    l = JSON.parse(text);
  } catch (err) {
    throw new Error(`the failure list ${file} does not read as JSON (${err.message}): mend it or move it aside by hand`);
  }
  return { entries: l.entries || {}, runs: l.runs || [], next: l.next || 1 };
}

/** Written whole or not at all: to a file beside it, then renamed over it. */
export function writeLedger(ledger, file) {
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, `${JSON.stringify(ledger, null, 1)}\n`);
  fs.renameSync(tmp, file);
}

/** What a window process that crashed leaves in the place of its counts (tools/gate.js windowRun). */
export const CRASH_PLACEHOLDER = 'the run ended before it counted';

/** What a hung lane is said to have done: no check result for so many minutes, after the last one it gave. */
export const hangDetail = (h) => `hung: no check result for ${h.minutes} minutes${h.after ? ` after "${h.after}"` : ', none at all'}, and ended`;

const keyOf = (app, part, name) => `${app}|${part}|${name}`;
/**
 * The name a check gives itself when its part threw part way (verify-*.js: "the check ran to the end", "the ...
 * checks ran to the end", and most often "the ... checks ran", about nine in ten of them). Each is named only when it
 * fails, so its part running clean is its pass; "ran" alone was missed, and those entries could never be closed
 * by a commit (09.0381, 21.0083, 07.0520 and 18.0248 on 14 Nov 2026).
 */
export const RAN_TO_END = /\bran( to the end)?$/;

const partOf = (app, pid) => (!pid || pid === '*' ? `${app}#*` : pid);

/**
 * What a run's parts files say: `apps` as the gate keeps them (each with
 * `parts` from verify-apps.js), `machine` the machine lane's own (null when
 * it did not run), `machineCrashed` when it ran and wrote no parts file,
 * `again` the run again's parts files and `cleared` the apps whose run again
 * cleared them. Answers { failures, passed, clean }:
 *   failures  each failed check { key, name, id, app, part, lane, at, cause, detail, synthetic, flaky, again }
 *   passed    `<app>|<name>` of each check that passed the first time
 *   clean     `<app>|<part>` of each part that ran checks, none failing, without throwing
 * A pass in the run again is no proof: it is what makes a failure flaky.
 */
export function runFindings({ apps = {}, machine = null, machineCrashed = false, machineHung = null, machineScope = [], again = [], cleared = new Set(), narrowed = new Set() } = {}) {
  const failures = [];
  const passed = new Set();
  const clean = new Set();
  const againPassed = new Set();
  const againClean = new Set();
  const againFailures = [];
  const read = (lane, app, v, into, passInto, cleanInto) => {
    const parts = Object.entries(v?.parts || {});
    let allClean = parts.length > 0;
    for (const [pid, p] of parts) {
      const part = partOf(app, pid);
      for (const n of p.passedNames || []) passInto?.add(`${app}|${n}`);
      const fails = p.failures || (p.failed || []).map((name) => ({ name, id: 'new' }));
      for (const f of fails) into.push({ ...f, key: keyOf(app, part, f.name), app, part, lane, synthetic: Boolean(f.synthetic) });
      if (p.of > 0 && !fails.length && !p.threw) cleanInto?.add(`${app}|${part}`);
      else allClean = false;
    }
    if (parts.length) {
      // The app whole is clean only when the app ran whole: some of its parts, asked for by id, prove nothing of the rest (review of 11 Nov).
      if (allClean && !narrowed.has(app)) cleanInto?.add(`${app}|${app}#*`);
      return;
    }
    // No parts: a process that ended early, or a parts file from before parts were kept.
    for (const n of v?.passedNames || []) passInto?.add(`${app}|${n}`);
    for (const name of v?.failed || []) {
      const crash = name === CRASH_PLACEHOLDER;
      const nm = crash ? `${app}: the window process ran to its end and wrote its counts` : name;
      // Ended as hung (gate.js CHECK_MINUTES): said with the last check it gave, so the entry says where it stopped.
      const hang = crash && v?.hung ? { hang: true, detail: hangDetail(v.hung) } : {};
      into.push({ name: nm, id: 'new', key: keyOf(app, `${app}#*`, nm), app, part: `${app}#*`, lane, synthetic: crash, crash, ...hang });
    }
  };
  for (const [app, v] of Object.entries(apps || {})) read('parallel', app, v, failures, passed, clean);
  for (const [app, v] of Object.entries(machine || {})) read('machine', app, v, failures, passed, clean);
  if (machineCrashed) {
    const name = 'machine lane: the window process ran to its end and wrote its counts';
    failures.push({ name, id: 'new', key: keyOf('machine', 'machine#lane', name), app: 'machine', part: 'machine#lane', lane: 'machine', synthetic: true, crash: true, rerun: machineScope.map((a) => `${a}#*`), ...(machineHung ? { hang: true, detail: hangDetail(machineHung) } : {}) });
    // What the machine lane said failed before it hung, each an entry of its own: under its own app when its number says which
    // (gate.js appOfCheck), run again and proved as that app's; else in the lane, proved by the lane running clean.
    for (const f of machineHung?.failed || []) {
      const name = typeof f === 'string' ? f : f.name;
      const app = typeof f === 'string' ? null : f.app;
      if (app) failures.push({ name, id: 'new', key: keyOf(app, `${app}#*`, name), app, part: `${app}#*`, lane: 'machine', synthetic: false });
      else failures.push({ name, id: 'new', key: keyOf('machine', 'machine#lane', name), app: 'machine', part: 'machine#lane', lane: 'machine', synthetic: true, rerun: machineScope.map((a) => `${a}#*`) });
    }
  }
  // The machine lane ran to its end with checks in it, not an empty parts file from a narrowed run.
  const ranIn = (v) => v?.of ?? Object.values(v?.parts || {}).reduce((n, p) => n + (p.of || 0), 0);
  if (machine && Object.values(machine).reduce((n, v) => n + ranIn(v), 0) > 0) clean.add('machine|machine#lane');
  for (const file of again) for (const [app, v] of Object.entries(file || {})) read('again', app, v, againFailures, againPassed, againClean);
  const firstKeys = new Set(failures.map((f) => f.key));
  // Flaky: failed beside the others, passed alone. A part that threw is flaky when its part ran clean alone (a wait that gave up, most often).
  // So is a check that only says its part threw ("... ran", "... ran to the end"): it is named only when it fails, so a run
  // that passes never names it, and its part running clean alone is the pass (14.0043, 12 Nov 2026).
  // Its own part, by its id: the app whole counted clean says nothing of a part the run again did not hold (review of 13 Nov).
  for (const f of failures) f.flaky = f.crash ? cleared.has(f.app) : f.synthetic || RAN_TO_END.test(f.name) ? againClean.has(`${f.app}|${f.part}`) : againPassed.has(`${f.app}|${f.name}`);
  // What failed only in the run again (a part that got further there): an entry of its own.
  for (const f of againFailures) if (!firstKeys.has(f.key)) { failures.push({ ...f, again: true, flaky: false }); firstKeys.add(f.key); }
  return { failures, passed, clean };
}

const pad = (n) => String(n).padStart(4, '0');

/**
 * Add a run to the list. Each failure opens an entry or adds to its own (a
 * closed one that fails again is opened again); each open entry that passed
 * in this run gains a pass, marked clean when the run's tree held nothing
 * beyond its commit. `run` is { head, time, label, scope, clean }.
 */
export function recordRun(ledger, run, { failures = [], passed = new Set(), clean = new Set() } = {}) {
  ledger.next ||= 1;
  const time = run.time || new Date().toISOString();
  const cleanTree = run.clean !== false;
  ledger.runs.push({ head: run.head, time, label: run.label || null, scope: run.scope || null, clean: cleanTree, failed: failures.length });
  if (ledger.runs.length > 200) ledger.runs.splice(0, ledger.runs.length - 200);
  const idTaken = (id, key) => Object.values(ledger.entries).some((e) => e.id === id && e.key !== key);
  const failedNow = new Set();
  for (const f of failures) {
    failedNow.add(f.key);
    const e = ledger.entries[f.key];
    const seen = { head: run.head, time, lane: f.lane, part: f.part, at: f.at || null, cause: f.cause || null, detail: f.detail || null, clearedAlone: Boolean(f.flaky), onlyAgain: Boolean(f.again) };
    const number = f.id && f.id !== 'new' && !f.synthetic && !idTaken(f.id, f.key) ? f.id : null;
    if (!e) {
      ledger.entries[f.key] = { key: f.key, id: number || `L.${pad(ledger.next++)}`, name: f.name, app: f.app, part: f.part, at: f.at || null, synthetic: Boolean(f.synthetic), ...(f.rerun ? { rerun: f.rerun } : {}), status: 'open', failures: 1, flaky: Boolean(f.flaky), first: seen, last: seen, passedAt: [], closed: null, reopened: 0 };
      continue;
    }
    e.failures += 1;
    e.last = seen;
    e.flaky = e.flaky || Boolean(f.flaky);
    if (number && e.id.startsWith('L.')) e.id = number;
    if (f.at) e.at = f.at;
    // A failure after a pass makes that pass no proof.
    e.passedAt = [];
    if (e.status === 'closed') { e.status = 'open'; e.reopened += 1; e.closed = null; }
  }
  for (const e of Object.values(ledger.entries)) {
    if (e.status !== 'open' || failedNow.has(e.key)) continue;
    // An engine entry is proved only on the host it last failed on: a Linux failure is not closed by a pass on Windows (review of 11 Nov).
    // Or on the host its file belongs to now, for an entry whose file was moved to fix it (gate-hosts.js engineFindings).
    const apps = e.app === 'engine' && e.last?.lane ? [`engine@${e.last.lane}`, 'engine@home'] : [e.app];
    // A check that only says its part threw ("... ran", "... ran to the end") is never named by a run that passes: its part running clean is its pass.
    const ok = apps.some((app) => (e.synthetic || RAN_TO_END.test(e.name) ? clean.has(`${app}|${e.part}`) : passed.has(`${app}|${e.name}`)));
    if (ok) e.passedAt.push({ head: run.head, time, label: run.label || null, clean: cleanTree });
  }
  return ledger;
}

/** The entries still open, the most failed first. */
export function openEntries(ledger) {
  return Object.values(ledger.entries).filter((e) => e.status === 'open').sort((a, b) => b.failures - a.failures || String(a.id).localeCompare(String(b.id)));
}

/**
 * An entry by its number (`18.0142`, `L.0003`), its key, or its name or the
 * start of its name when only one entry has it. Answers { entry }, or
 * { many } when the key names more than one, or {} for none.
 */
export function findEntry(ledger, key) {
  const all = Object.values(ledger.entries);
  const one = (list) => (list.length === 1 ? { entry: list[0] } : list.length ? { many: list } : null);
  return one(all.filter((e) => e.id === key)) || (ledger.entries[key] ? { entry: ledger.entries[key] } : null) || one(all.filter((e) => e.name === key)) || one(all.filter((e) => e.name.startsWith(key))) || {};
}

/** The owner's word quoted in a reason: `owner: "..."`, or the owner named before a quotation of ten characters or more. */
export const OWNER_QUOTED = /\bowner\b[^"“]{0,40}["“][^"”]{10,}["”]/i;

/**
 * An engine test's entry moved with its test to another file, the fix for it
 * (12 Nov 2026: the Locker's large-file tests went to a file the lab machine
 * runs). Its number, failures and flakiness stay; its name and part follow the
 * test; the passes before the move are no proof, so it is closed only by passes
 * in the new file, as before. `has(file, title)` says whether a file holds a
 * test of that title now. Answers { ok, said, from, name }.
 */
export function moveEntry(ledger, key, to, { has = () => false, time = new Date().toISOString() } = {}) {
  const found = findEntry(ledger, key);
  if (found.many) return { ok: false, said: `${key} names ${found.many.length} entries: move each by its number` };
  const e = found.entry;
  if (!e) return { ok: false, said: `no entry ${key} in the failure list` };
  if (e.status === 'closed') return { ok: false, said: `${e.id} is closed already` };
  if (e.app !== 'engine' || !String(e.part).startsWith('test#') || e.synthetic) return { ok: false, said: `${e.id} is not an engine test's: only those move with their file` };
  const from = e.part.slice('test#'.length);
  const title = e.name.startsWith(`${from} > `) ? e.name.slice(from.length + 3) : null;
  if (!title) return { ok: false, said: `${e.id}'s name does not start with its file, ${from}` };
  if (!to || to === from) return { ok: false, said: `name the file the test is in now (--to tests/<file>.test.js), not ${from}` };
  if (has(from, title)) return { ok: false, said: `${from} still has "${title}": move the test first` };
  if (!has(to, title)) return { ok: false, said: `${to} has no test "${title}"` };
  const name = `${to} > ${title}`;
  const part = `test#${to}`;
  const nk = keyOf('engine', part, name);
  if (ledger.entries[nk]) return { ok: false, said: `${to} has an entry for that test already (${ledger.entries[nk].id})` };
  delete ledger.entries[e.key];
  ledger.entries[nk] = { ...e, key: nk, name, part, at: null, passedAt: [], moved: [...(e.moved || []), { from: e.part, to: part, time }] };
  return { ok: true, said: `${e.id} moved from ${from} to ${to}: passes there prove it from now`, from: e.name, name };
}

/**
 * Close an entry: `by` a commit made after it last failed, proved by a run on
 * a clean tree at or after that commit that passed it (two such runs for a
 * flaky entry); or with a `reason` written down, which for a flaky entry must
 * quote the owner. `descends(head, commit)` says whether `commit` is in the
 * history of `head`. Answers { ok, said }.
 */
export function closeEntry(ledger, key, { by = null, reason = null, descends = () => false, time = new Date().toISOString() } = {}) {
  const found = findEntry(ledger, key);
  if (found.many) return { ok: false, said: `${key} names ${found.many.length} entries (${found.many.slice(0, 4).map((e) => e.id).join(', ')}): close each by its number` };
  const e = found.entry;
  if (!e) return { ok: false, said: `no entry ${key} in the failure list` };
  if (e.status === 'closed') return { ok: false, said: `${e.id} is closed already` };
  if (by) {
    if (e.last?.head && (e.last.head === by || descends(e.last.head, by))) return { ok: false, said: `${e.id} last failed at ${e.last.head.slice(0, 8)}, which already holds ${by.slice(0, 8)}: the fix is a commit made after the failure` };
    const proofs = e.passedAt.filter((p) => p.clean !== false && (p.head === by || descends(p.head, by)));
    const need = e.flaky ? 2 : 1;
    if (proofs.length < need) return { ok: false, said: `${e.id} has passed ${proofs.length} of the ${need} run${need === 1 ? '' : 's'} it needs on a clean tree at or after ${by.slice(0, 8)}${e.flaky ? ' (it is flaky: two)' : ''}: commit the fix, run its part (npm run gate -- verify:apps --failed), and close it after` };
    e.status = 'closed';
    e.closed = { by, proof: proofs.at(-1), time };
    return { ok: true, said: `${e.id} closed by ${by.slice(0, 8)}, passed at ${proofs.at(-1).head.slice(0, 8)}` };
  }
  if (reason && reason.trim().length >= 10) {
    // A flaky entry has a cause still in the code: a reason alone closes it
    // only with the owner's word quoted (PM, 11 Nov 2026).
    if (e.flaky && !OWNER_QUOTED.test(reason)) return { ok: false, said: `${e.id} failed and then passed alone: close it with --by <the commit that fixes its cause>, or quote the owner's word in the reason (owner: "...")` };
    e.status = 'closed';
    e.closed = { reason: reason.trim(), time, recorded: null };
    return { ok: true, said: `${e.id} closed: ${reason.trim()} (said in the next record's Gate section)` };
  }
  return { ok: false, said: 'close with --by <commit> (a run after it passed the check) or --reason "<why, in a sentence>"' };
}

/**
 * The entries closed by a reason, not a commit, that no record has said yet:
 * a way round the list, so each is printed in the record's Gate section.
 */
export function reasonClosures(ledger) {
  return Object.values(ledger.entries).filter((e) => e.status === 'closed' && e.closed?.reason && !e.closed.recorded).map((e) => ({ id: e.id, name: e.name, reason: e.closed.reason }));
}

/** Mark those closures said, by the version whose record printed them. */
export function markRecorded(ledger, version) {
  for (const e of Object.values(ledger.entries)) if (e.status === 'closed' && e.closed?.reason && !e.closed.recorded) e.closed.recorded = version;
  return ledger;
}

/** The parts to run again for the open entries: each entry's own part, or its app whole. */
export function openParts(ledger) {
  const ids = new Set();
  for (const e of openEntries(ledger)) {
    if (e.rerun?.length) for (const id of e.rerun) ids.add(id);
    else if (e.part && e.part.includes('#')) ids.add(e.part);
    else if (e.app) ids.add(`${e.app}#*`);
  }
  return [...ids];
}

/** The list as lines for the log: number, where, part and name, then why. */
export function ledgerLines(entries, { max = 40 } = {}) {
  const out = [];
  for (const e of entries.slice(0, max)) {
    // Passes on a clean tree only: those are what closing it by a commit counts.
    const proofs = e.passedAt.filter((p) => p.clean !== false);
    const tags = [e.flaky ? 'flaky: passed alone' : null, e.last?.onlyAgain ? 'failed in the run again' : null, e.synthetic ? 'said by the run itself' : null, e.reopened ? `failed again after a close ×${e.reopened}` : null, proofs.length ? `passed since ×${proofs.length} on a clean tree, last at ${proofs.at(-1).head.slice(0, 8)}` : null].filter(Boolean);
    out.push(`  [${e.id}] ${e.name.slice(0, 140)}`);
    out.push(`       ${[e.at && `at ${e.at}`, e.part && `part ${e.part}`, e.last?.lane && `${e.last.lane} lane`, `failed ${e.failures}×`, ...tags].filter(Boolean).join(' · ')}`);
    if (e.last?.cause?.what) out.push(`       after the wait for ${e.last.cause.what} gave up at ${e.last.cause.ms} ms${e.last.cause.at ? ` (${e.last.cause.at})` : ''}${e.last.cause.load ? `, ${saidLoad(e.last.cause.load)}` : ''}`);
    if (e.last?.detail) out.push(`       said: ${String(e.last.detail).slice(0, 200)}`);
  }
  if (entries.length > max) out.push(`  and ${entries.length - max} more: npm run gate -- ledger`);
  return out;
}

/** The list as a Markdown page, .gate/failures.md, for reading and handing on. */
export function ledgerMarkdown(ledger, { head = '' } = {}) {
  const open = openEntries(ledger);
  const rows = open.map((e) => `| ${e.id} | ${e.app || ''} | ${e.part || ''} | ${e.at || ''} | ${e.failures} | ${e.flaky ? 'yes' : ''} | ${e.name.replace(/\|/g, '/').slice(0, 160)} |`);
  return [
    `# Gate failure list${head ? ` at ${head.slice(0, 8)}` : ''}`,
    '',
    `${open.length} open. A full run on the record (--record) is refused until each is closed: \`npm run gate -- ledger close <number> --by <commit>\` once a run on a clean tree after the fix passed it (twice for a flaky one), or \`--reason "..."\`, which the record's Gate section prints.`,
    '',
    '| Number | App | Part | Judged at | Failed | Flaky | Check |',
    '| --- | --- | --- | --- | --- | --- | --- |',
    ...rows,
    '',
  ].join('\n');
}
