// After a seek, a check waits for the Video preview to say it painted the new
// instant with each clip's decoder on it, not a fixed time (D-2026-11-13-20,
// the owner: "Wait on signals (Recommended)").
import test from 'node:test';
import assert from 'node:assert/strict';
import { seekVideoPainted } from '../apps/desktop/main/verify-waits.js';

/** A Video window as the check sees it: the canvas's data-time and data-settled, and its scrub bar. */
function windowWith(canvas, { onSeek = () => {} } = {}) {
  const seeks = [];
  const js = async (code) => {
    if (code.includes('dataset.time ??')) return canvas.time ?? '';
    if (code.includes('.vd-scrub')) { seeks.push(code); onSeek(canvas); return 1; }
    // The painted test, run against this canvas.
    const was = /at !== ("[^"]*")/.exec(code)[1];
    const t = Number(/Math\.abs\(Number\(at\) - ([\d.]+)\)/.exec(code)[1]);
    const at = canvas.time;
    return (at !== JSON.parse(was) || Math.abs(Number(at) - t) < 1e-6) && canvas.settled === '1';
  };
  return { js, seeks };
}

/** A wait that asks a few times, as until does, without the minute's outer limit. */
const quickWait = async (cond, what, _ms, { tries = 5 } = {}) => {
  for (let i = 0; i < tries; i++) if (await cond()) return true;
  throw new Error(`never: ${what}`);
};

test('a seek answers once the preview has painted a new instant and settled', async () => {
  const canvas = { time: '1', settled: '1' };
  const w = windowWith(canvas, { onSeek: (c) => { c.time = '2.5'; c.settled = '1'; } });
  assert.equal(await seekVideoPainted(w.js, 2.5, { waitFor: quickWait }), true);
  assert.equal(w.seeks.length, 1);
});

test('a new instant painted but not settled is waited for: the decoder has not landed yet', async () => {
  const canvas = { time: '1', settled: '1' };
  let asked = 0;
  const w = windowWith(canvas, { onSeek: (c) => { c.time = '2.5'; c.settled = '0'; } });
  const wait = async (cond, what, ms) => quickWait(async () => { asked += 1; if (asked === 3) canvas.settled = '1'; return cond(); }, what, ms);
  assert.equal(await seekVideoPainted(w.js, 2.5, { waitFor: wait }), true);
  assert.ok(asked >= 3, 'it went on asking until the frame settled');
});

test('a seek to the instant already shown is enough once it is settled: nothing new is painted for it', async () => {
  const canvas = { time: '3', settled: '1' };
  const w = windowWith(canvas);
  assert.equal(await seekVideoPainted(w.js, 3, { waitFor: quickWait }), true);
});

test('a signal that never comes answers false, and does not throw: the check reads the frame and says what it found', async () => {
  const canvas = { time: '1', settled: '1' };
  const w = windowWith(canvas, { onSeek: (c) => { c.settled = '0'; } });
  assert.equal(await seekVideoPainted(w.js, 4, { waitFor: quickWait }), false);
});
