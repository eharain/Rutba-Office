// The gate's reform after gates 6 to 9 (owner, 11 Nov 2026: "do not run this
// once again without proper reforms"), as the Fable 5.2 review set it out:
// timing budgets only in the machine lane, one floor for every wait on a
// window, a shared-lane export waited on while it moves and cancelled when
// it does not, and key lag judged on the keys as a whole.
import test from 'node:test';
import assert from 'node:assert/strict';
import { sectionsIn } from '../tools/verify-sections.js';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { keyLagVerdict } from '../apps/desktop/main/verify-key-lag.js';
import { waitForExport } from '../apps/desktop/main/verify-export-wait.js';
import { until, waitFor, UI_WAIT } from '../apps/desktop/main/verify-waits.js';

const MAIN = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'apps', 'desktop', 'main');
const read = (f) => fs.readFileSync(path.join(MAIN, f), 'utf8');
const VERIFY = fs.readdirSync(MAIN).filter((f) => /^verify-.*\.js$/.test(f));

/* ── key lag ─────────────────────────────────────────────────────────── */

test('key lag is judged on the keys as a whole: gate 9\'s single late keys pass, a held window does not', () => {
  // machine.log 11 Nov: one key at 255 ms of eighteen, one at 312 of twenty-three, medians near 70 and 28.
  const g9a = [70, 62, 75, 255, 66, 71, 80, 68, 74, 69, 72, 77, 65, 70, 73, 71, 68, 69];
  const g9b = [28, 26, 30, 312, 27, 29, 31, 25, 28, 30, 27, 26, 29, 28, 31, 27, 30, 26, 28, 29, 27, 28, 30];
  assert.equal(keyLagVerdict(g9a).ok, true, keyLagVerdict(g9a).said);
  assert.equal(keyLagVerdict(g9b).ok, true, keyLagVerdict(g9b).said);
  assert.equal(keyLagVerdict([10, 11, 8, 7, 207, 141, 72, 18, 8, 7, 6, 7]).ok, true, 'one in twelve over 200 ms');
  assert.equal(keyLagVerdict([10, 11, 8, 7, 207, 241, 72, 18, 8, 7, 6, 7]).ok, false, 'two in twelve over 200 ms is the work holding the windows');
  assert.equal(keyLagVerdict([20, 22, 640, 21]).ok, false, 'no key may take half a second');
  assert.equal(keyLagVerdict([150, 160, 140, 155, 145]).ok, false, 'a median of 150 ms is held, whatever the slowest');
  assert.equal(keyLagVerdict([20, null, 22]).ok, false, 'a key that never came');
  assert.match(keyLagVerdict(g9b).said, /median 28 ms .* 1 of 23 at 200 ms or more \(2 allowed\), slowest 312 ms/);
});

/* ── waiting for an export ───────────────────────────────────────────── */

/** A window whose export box shows `lines` one after another, as `script()` says. */
function fakeWin(script) {
  const state = { cancelled: 0, shown: null };
  const win = {
    webContents: {
      executeJavaScript: async (code) => {
        if (code.includes('find((x) => /^Cancel$/')) { state.cancelled++; state.shown = null; return true; }
        if (code.includes("Boolean(document.querySelector('.vd-export-busy'))")) return state.shown !== null;
        state.shown = script(state);
        return state.shown;
      },
    },
  };
  return { win, state };
}

test('an export that moves, however slowly, is waited for to its end', async () => {
  let n = 0;
  const { win, state } = fakeWin(() => `Encoding 0:0${Math.min(9, Math.floor(n++ / 3))} of 0:09 | `);
  const r = await waitForExport(win, async () => n > 20, { quietMs: 3000, capMs: 20000, tickMs: 10 });
  assert.equal(r.done, true);
  assert.equal(state.cancelled, 0);
});

test('an export that stops moving is cancelled, its box seen to close, and the check told why (gate 9 left one running into the next)', async () => {
  const { win, state } = fakeWin((s) => (s.cancelled ? null : 'Encoding 0:00.6 of 0:09.6 | '));
  const r = await waitForExport(win, async () => false, { quietMs: 150, capMs: 10000, tickMs: 10 });
  assert.equal(r.done, false);
  assert.equal(state.cancelled, 1);
  assert.match(r.stopped, /did not move for 0 seconds, and was cancelled|did not move for \d+ seconds, and was cancelled/);
  assert.equal(r.shown, 'Encoding 0:00.6 of 0:09.6 | ');
});

test('an export that runs past its cap is cancelled too, and one whose box closed with no file ends the wait', async () => {
  let n = 0;
  const moving = fakeWin(() => `Encoding ${n++} | `);
  const long = await waitForExport(moving.win, async () => false, { quietMs: 5000, capMs: 150, tickMs: 10 });
  assert.match(long.stopped, /ran past 0 seconds, and was cancelled/);
  assert.equal(moving.state.cancelled, 1);
  let k = 0;
  const closed = fakeWin(() => (k++ < 3 ? 'Encoding 0:01 | ' : null));
  const t0 = Date.now();
  const gone = await waitForExport(closed.win, async () => false, { quietMs: 5000, capMs: 10000, tickMs: 50 });
  assert.equal(gone.done, false);
  assert.match(gone.stopped, /box closed and no file came/);
  assert.ok(Date.now() - t0 < 5000);
});

test('a box not drawn yet is waited for, not taken for one that closed (review of 11 Nov: beside other lanes it can take seconds)', async () => {
  let k = 0;
  // Nothing for the first 4 s of ticks, then the export runs and its file comes.
  const late = fakeWin(() => (k++ < 80 ? null : `Encoding ${k} | `));
  const r = await waitForExport(late.win, async () => k > 100, { quietMs: 5000, capMs: 20000, tickMs: 50, openMs: 15000 });
  assert.equal(r.done, true, r.stopped);
  const never = fakeWin(() => null);
  const n = await waitForExport(never.win, async () => false, { quietMs: 5000, capMs: 20000, tickMs: 20, openMs: 300 });
  assert.match(n.stopped, /did not open in 0 seconds/);
});

test('the video checks wait for their exports this way, not by the project\'s length plus a few seconds', () => {
  for (const f of ['verify-intro-video.js', 'verify-video-editor.js', 'verify-recorder.js']) {
    const src = read(f);
    assert.match(src, /waitForExport\(/, f);
    assert.doesNotMatch(src, /(length|projectLength) \* 1000\) \+ \d+/, `${f} still gives an export its length and a few seconds`);
  }
});

/* ── one floor for the waits on a window ─────────────────────────────── */

test('a wait asked for less than UI_WAIT is given UI_WAIT, and ends as soon as what it waits for comes', async () => {
  assert.equal(UI_WAIT, 15000);
  assert.equal(waitFor(4000), UI_WAIT);
  assert.equal(waitFor(60000), 60000);
  assert.equal(waitFor(4000, { exact: true }), 4000);
  // Asked for 100 ms, it holds on past that (gate 8 closed on a 5,000 ms wait beaten by a loaded machine).
  const t0 = Date.now();
  assert.equal(await until(() => Date.now() - t0 > 300, 'a menu', 100, { tick: 10 }), true);
  const lines = [];
  await assert.rejects(until(() => false, 'a measured wait', 120, { exact: true, tick: 10, log: (l) => lines.push(l) }), /waited 120 ms for a measured wait/);
  assert.match(lines[0], /\[wait\] gave up on a measured wait after \d+ ms/);
});

test('every check waits through the one until: verify-apps.js takes it from verify-waits.js and no check file has a loop of its own', () => {
  assert.match(read('verify-apps.js'), /import \{ until \} from '\.\/verify-waits\.js';/);
  assert.doesNotMatch(read('verify-apps.js'), /async function until\(/);
  // verify-edit.js is the editing pass's own harness, outside verify:apps.
  const own = VERIFY.filter((f) => !['verify-waits.js', 'verify-edit.js'].includes(f)).filter((f) => /function until\(/.test(read(f)) || /const until = async[^\n]*\n?[^\n]*(?<!h\.)\bfor \(;;\)/.test(read(f)));
  assert.deepEqual(own, []);
});

test('a wait kept to its own figure ({ exact: true }) is one whose length is what is checked, and there are none in the shared lane', () => {
  const exact = VERIFY.filter((f) => f !== 'verify-waits.js').flatMap((f) => (read(f).match(/exact: true/g) || []).map(() => f));
  assert.deepEqual(exact, []);
});

/* ── budgets only in the machine lane ────────────────────────────────── */

/** The `part(...)` lines of verify-apps.js that run `name`, directly or through a closure of verify-apps.js that calls it. */
function partsRunning(name) {
  const src = read('verify-apps.js');
  const lines = src.split('\n').filter((l) => /^\s*await part\(/.test(l));
  const direct = lines.filter((l) => new RegExp(`=> ${name}\\(`).test(l));
  if (direct.length) return direct;
  // A closure: `const sheetViews = async () => { ... verifyViews(...) ... };`
  const closures = [...src.matchAll(/const (\w+) = async \(\) => \{([\s\S]*?)\n {2}\};/g)].filter((m) => new RegExp(`\\b${name}\\(`).test(m[2])).map((m) => m[1]);
  return lines.filter((l) => closures.some((c) => new RegExp(`=> ${c}\\(`).test(l)));
}

/** The functions a check file exports. */
const exportsOf = (f) => [...read(f).matchAll(/export (?:async )?function (\w+)/g)].map((m) => m[1]);

// The check files whose checks assert a time (review of 11 Nov and the audit
// of every check file): each runs in the machine lane, alone. A new file
// with a `_MS` budget constant joins this list by itself.
const TIMED = ['verify-mail-index-worker.js', 'verify-pwfast.js', 'verify-doc-thread.js', 'verify-perf-windows.js', 'verify-video-mp4.js', 'verify-views.js', 'verify-layoutviews.js', 'verify-ocr-pdf.js', 'verify-windows-codec.js#verifyWindowsCodec', 'verify-broadcast-live.js'];
// Deadlines that are not budgets: a browser list fetched once a run, sixty seconds; how often the
// processors are read for the load beside a wait that gave up (verify-load.js), which asserts no time.
const NOT_BUDGETS = new Set(['verify-mail-signin.js', 'verify-load.js']);

test('every check that asserts a time runs in the machine lane, and only there', () => {
  const withMs = VERIFY.filter((f) => /^const [A-Z_]+_MS = \d/m.test(read(f)) && !NOT_BUDGETS.has(f));
  for (const f of withMs) assert.ok(TIMED.some((t) => t.split('#')[0] === f), `${f} has a _MS budget: run it in the machine lane and list it here`);
  // `file#fn` where only that function of the file is timed.
  for (const entry of TIMED) {
    const [f, only] = entry.split('#');
    const fns = (only ? [only] : exportsOf(f)).filter((fn) => partsRunning(fn).length);
    assert.ok(fns.length, `${f}: no part runs it`);
    for (const fn of fns) {
      for (const line of partsRunning(fn)) assert.match(line, /\{ machine: true \}\);\s*$/, `${f} (${fn}) runs in the shared lane: ${line.trim()}`);
    }
  }
});

test('no part marked for the machine lane sits inside a section, where the machine lane never reaches it', () => {
  // The sections as the scanner finds them (tools/verify-sections.js): no part at all sits in one now (tests/gate-sections.test.js).
  const sections = sectionsIn(read('verify-apps.js').replace(/\r\n/g, '\n'));
  assert.ok(sections.length >= 30);
  const nested = sections.filter((s) => /\{ machine: true \}/.test(s.body)).map((s) => `${s.app}#${s.name}`);
  assert.deepEqual(nested, []);
});

test('a run of some apps, some parts, all but some parts or the machine lane keeps itself alive with a hidden window', () => {
  // Without it the last part's windows closing ends the application, and the parts after never run
  // (a run of RUTBA_VERIFY_PART_IDS alone ended after its first part, 14 Nov 2026).
  assert.match(read('verify-apps.js'), /if \(appsOnly\.length \|\| partIds\.length \|\| skipIds\.length \|\| lane === 'machine'\) \{\s*const \{ BrowserWindow \} = await import\('electron'\);\s*new BrowserWindow\(\{ show: false/);
});
