// A camera the studio still wants is asked for again
// (renderer/recording/camera-retry.js): every few seconds after it fails,
// for as long as it is wanted, never once it was refused, and a wait for the
// next try ends at once when the camera is let go. A capture that opened
// but gives no frame is opened again, a camera each time after a longer
// wait, a screen or window once.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  openWhileWanted, cancellableWait, retryCamera, framelessWaitMs, reopenLimit, CAMERA_RETRY_MS,
} from '../apps/desktop/renderer/recording/camera-retry.js';

const failure = (name) => Object.assign(new Error(name), { name });

/** A wait that is over at once, noting how long it was asked to be. */
const instant = (log) => (ms) => { log.push(ms); return { done: Promise.resolve(true), cancel() {} }; };

test('a camera that will not start is asked for again until it does, its trouble shown each time', async () => {
  const errors = ['NotReadableError', 'NotFoundError', 'AbortError'];
  let tries = 0;
  const shown = [];
  const waits = [];
  const stream = await openWhileWanted(async () => {
    tries += 1;
    if (errors.length) throw failure(errors.shift());
    return 'stream';
  }, { onFail: (err) => shown.push(err.name), pause: instant(waits) });
  assert.equal(stream, 'stream');
  assert.equal(tries, 4);
  assert.deepEqual(shown, ['NotReadableError', 'NotFoundError', 'AbortError']);
  assert.deepEqual(waits, [CAMERA_RETRY_MS, CAMERA_RETRY_MS, CAMERA_RETRY_MS], 'a few seconds between tries');
  assert.ok(CAMERA_RETRY_MS >= 2000 && CAMERA_RETRY_MS <= 6000);
});

test('a camera that was refused is not asked for again', async () => {
  let tries = 0;
  const shown = [];
  const stream = await openWhileWanted(async () => { tries += 1; throw failure('NotAllowedError'); }, { onFail: (err) => shown.push(err.name), pause: instant([]) });
  assert.equal(stream, null);
  assert.equal(tries, 1);
  assert.deepEqual(shown, ['NotAllowedError']);
  assert.equal(retryCamera(failure('SecurityError')), false);
  assert.equal(retryCamera(failure('NotReadableError')), true);
  assert.equal(retryCamera(failure('OverconstrainedError')), true, 'a camera unplugged may come back');
  assert.equal(retryCamera(undefined), true);
});

test('a camera no longer wanted is not asked for again, and its wait ends at once', async () => {
  let wanted = true;
  let tries = 0;
  let pending = null;
  const opening = openWhileWanted(async () => { tries += 1; throw failure('NotReadableError'); }, {
    wanted: () => wanted,
    onWait: (w) => { pending = w; },
    delayMs: 60000,
  });
  await new Promise((r) => setImmediate(r));
  assert.ok(pending, 'waiting for the next try');
  wanted = false;
  pending.cancel();
  assert.equal(await pending.done, false, 'the wait ended early');
  assert.equal(await opening, null);
  assert.equal(tries, 1);
});

test('a failure after the camera was let go is not shown', async () => {
  let wanted = true;
  const shown = [];
  const r = await openWhileWanted(async () => { wanted = false; throw failure('NotReadableError'); }, { wanted: () => wanted, onFail: (e) => shown.push(e), pause: instant([]) });
  assert.equal(r, null);
  assert.deepEqual(shown, []);
});

test('a cancellable wait clears its timer either way', async () => {
  const timers = new Map();
  let next = 1;
  const set = (fn, ms) => { const id = next++; timers.set(id, { fn, ms }); return id; };
  const clear = (id) => { timers.delete(id); };
  const a = cancellableWait(500, { set, clear });
  assert.equal(timers.size, 1);
  a.cancel();
  assert.equal(timers.size, 0, 'cancelled: no timer left');
  assert.equal(await a.done, false);
  const b = cancellableWait(500, { set, clear });
  const [[id, timer]] = [...timers];
  assert.equal(timer.ms, 500);
  timer.fn();
  timers.delete(id);
  assert.equal(await b.done, true, 'ran out');
  b.cancel();
  assert.equal(timers.size, 0, 'a cancel after it ran out clears nothing more');
});

test('a capture with no frame is waited on longer each time it is opened again, a camera without end, a window once', () => {
  assert.equal(framelessWaitMs(0), 3000, 'three seconds the first time, as before');
  assert.equal(framelessWaitMs(1), 6000);
  assert.equal(framelessWaitMs(2), 12000);
  assert.equal(framelessWaitMs(3), 20000);
  assert.equal(framelessWaitMs(12), 20000, 'never longer than twenty seconds');
  assert.equal(framelessWaitMs(undefined), 3000);
  assert.equal(reopenLimit(true), Infinity);
  assert.equal(reopenLimit(false), 1);
});
