// What a speed check measures (D-2026-11-13-20, the owner: "Yes, own
// processor time (Recommended)"; D-2026-11-13-25: "Either one under the
// limit (Recommended)"). A check that times work Office does passes when the
// clock, or the processor time Office's own processes spent on it, is under
// its limit: on a computer doing other work as well, the clock measures that
// work too, and on work spread over several cores the processor time adds
// the cores up. Both figures are in the detail. The limits' numbers are as
// before. Typing lag (pwfast) and a frame's paint
// (Sheets) are what a person waits for, so they stay on the clock.
//
// Beside every speed check a fixed reference task is timed by the clock and
// written into the check's detail with how busy the processors were and how
// much memory was free. It says how the computer ran at that moment; it is
// for information only and never counts toward a pass.

import crypto from 'node:crypto';
import os from 'node:os';
// The default import, read when called: a unit test gives its own metrics and has no Electron app.
import electron from 'electron';

/** Each of Office's processes now (main, windows, GPU, helpers): pid to processor time in ms since it started. */
function processTimes(metrics = () => electron.app.getAppMetrics()) {
  const out = new Map();
  for (const m of metrics()) {
    const s = m.cpu?.cumulativeCPUUsage;
    if (Number.isFinite(s)) out.set(m.pid, s * 1000);
  }
  return out;
}

/**
 * Start counting Office's own processor time. The answer's `end()` says how
 * much its processes spent since: each one there at the start by what it
 * added, each one started since by all it used (a window opened for the
 * check), and how many of those there at the start ended in between (what
 * those spent after the start is not counted, so the count is never more than
 * was spent). A process that started and ended inside the interval is not
 * seen at all. Programs Office starts as child processes are not Office's
 * processes here, so a check whose work runs in one stays on the clock.
 *
 * `measured` is false when the count cannot be trusted: no process reported
 * its time, or one named in `need` (the main process always, and the window
 * doing the work) did not. The processor time is not judged then, and only
 * the clock can pass the check (review of 13 Nov: 0 ms would pass).
 */
export function cpuSince({ metrics, mainPid = process.pid } = {}) {
  const before = processTimes(metrics);
  const t0 = performance.now();
  return {
    end({ need = [] } = {}) {
      const after = processTimes(metrics);
      let cpu = 0;
      for (const [pid, ms] of after) cpu += Math.max(0, ms - (before.get(pid) ?? 0));
      const ended = [...before.keys()].filter((pid) => !after.has(pid)).length;
      const missing = [mainPid, ...need].filter((pid) => pid != null && !after.has(pid));
      const measured = after.size > 0 && before.has(mainPid) && missing.length === 0;
      return { cpuMs: Math.round(cpu), ms: Math.round(performance.now() - t0), ended, measured };
    },
  };
}

/** The fixed reference task: the same hashing every time, timed by the clock, in ms. Its bytes are made at the first call. */
let referenceBytes = null;
export function referenceMs() {
  referenceBytes ||= Buffer.alloc(32 * 1024 * 1024, 7);
  const t0 = performance.now();
  for (let i = 0; i < 2; i++) crypto.createHash('sha256').update(referenceBytes).digest();
  return Math.round(performance.now() - t0);
}

/** The share of the processors busy over `ms`, from the system's own counters, as a whole percent. */
export async function busyNow(ms = 500) {
  const sample = () => os.cpus().reduce((t, c) => ({ idle: t.idle + c.times.idle, all: t.all + Object.values(c.times).reduce((a, b) => a + b, 0) }), { idle: 0, all: 0 });
  const a = sample();
  await new Promise((r) => setTimeout(r, ms));
  const b = sample();
  const all = b.all - a.all;
  return all > 0 ? Math.round((1 - (b.idle - a.idle) / all) * 100) : null;
}

/**
 * How the computer ran beside a speed check, for its detail: the reference
 * task's time, how busy the processors were, and the memory free. Never part
 * of a pass.
 */
export async function machineNow() {
  const busy = await busyNow();
  return { referenceMs: referenceMs(), busy, freeMb: Math.round(os.freemem() / 1024 / 1024) };
}

/** The detail's words for how the computer ran (machineNow). */
export function machineWords(machine) {
  return `the reference task took ${machine.referenceMs} ms, the processors ${machine.busy ?? '?'}% busy, ${machine.freeMb} MB free (information only)`;
}

/**
 * Whether a measurement is under its limit (D-2026-11-13-25): by the clock, or
 * by Office's own processor time when every process it needed reported it.
 * A measurement that does not say it was measured is not: the processor time
 * counts only when `measured` is true, so a 0 never passes on its own.
 */
export function within({ cpuMs, ms, measured = false }, limitMs) {
  return ms < limitMs || (measured === true && cpuMs < limitMs);
}

/** The detail's words for a measurement: both figures that are judged, and the machine. Not measured unless it says so, as within(). */
export function costWords({ cpuMs, ms, ended = 0, measured = false }, machine) {
  const cpu = measured === true
    ? `${cpuMs} ms of Office's own processor time${ended ? ` (${ended} of its processes ended in between)` : ''}`
    : `Office's own processor time could not be measured (a process did not report it), so the clock alone was judged`;
  return `${ms} ms by the clock, ${cpu}${machine ? `; ${machineWords(machine)}` : ''}`;
}
