// The throwaway recording that wakes a fresh window's video encoder before
// Export to Video's real one (renderer/recording/warm-encoder.js): it runs
// until the encoder gives its first piece and no longer, keeps drawing the
// whole while, gives up at its cap or on Cancel, is always stopped, records
// a small canvas of its own and never the export's tracks, and an empty
// result is never taken for a recording.
import test from 'node:test';
import assert from 'node:assert/strict';
import { warmEncoder, warmOnOwnCanvas, recordedSomething, WARM_CAP_MS } from '../apps/desktop/renderer/recording/warm-encoder.js';

/** A recorder whose encoder gives its first data `readyMs` after it starts, on a clock the test moves. */
function fakeRecorder(clock, readyMs, { failStart = false } = {}) {
  const rec = {
    state: 'inactive', slice: null, stops: 0, ondataavailable: null, onstop: null, onerror: null,
    start(slice) {
      if (failStart) throw new Error('not supported');
      rec.state = 'recording';
      rec.slice = slice;
      rec.startedAt = clock.now;
    },
    stop() {
      rec.stops += 1;
      rec.state = 'inactive';
      queueMicrotask(() => rec.onstop?.());
    },
    tick() {
      if (rec.state !== 'recording') return;
      const ready = clock.now - rec.startedAt >= readyMs;
      rec.ondataavailable?.({ data: { size: ready ? 2048 : 0 } });
    },
  };
  return rec;
}

function clockFor(getRec) {
  const clock = { now: 0 };
  clock.wait = async (ms) => { clock.now += ms; getRec()?.tick(); };
  return clock;
}

test('the encoder is waited on until its first data, drawn on all the while, then stopped', async () => {
  let rec = null;
  const clock = clockFor(() => rec);
  let draws = 0;
  const answered = await warmEncoder(() => (rec = fakeRecorder(clock, 3000)), { draw: () => { draws += 1; }, now: () => clock.now, wait: clock.wait });
  assert.equal(answered, true);
  assert.equal(rec.slice, 100, 'asked for its data a tenth of a second at a time');
  assert.equal(rec.state, 'inactive');
  assert.equal(rec.stops, 1);
  assert.ok(clock.now >= 3000 && clock.now < 3200, `stopped as soon as it answered, at ${clock.now} ms`);
  assert.ok(draws >= 85, `a frame drawn every step (${draws})`);
  assert.equal(rec.ondataavailable, null, 'its data is let go');
});

test('an encoder that never answers is given up on at the cap, and still stopped', async () => {
  let rec = null;
  const clock = clockFor(() => rec);
  const answered = await warmEncoder(() => (rec = fakeRecorder(clock, Infinity)), { now: () => clock.now, wait: clock.wait });
  assert.equal(answered, false);
  assert.ok(clock.now >= WARM_CAP_MS && clock.now < WARM_CAP_MS + 100, `given up at ${clock.now} ms`);
  assert.equal(rec.state, 'inactive');
  assert.equal(rec.stops, 1);
});

test('Cancel ends the wait at once', async () => {
  let rec = null;
  const clock = clockFor(() => rec);
  const answered = await warmEncoder(() => (rec = fakeRecorder(clock, 5000)), { cancelled: () => clock.now >= 500, now: () => clock.now, wait: clock.wait });
  assert.equal(answered, false);
  assert.ok(clock.now < 600, `stopped at ${clock.now} ms`);
  assert.equal(rec.state, 'inactive');
});

test('a recorder that cannot be made or started does not stop the real recording', async () => {
  assert.equal(await warmEncoder(() => { throw new Error('no such type'); }), false);
  let rec = null;
  const clock = clockFor(() => rec);
  assert.equal(await warmEncoder(() => (rec = fakeRecorder(clock, 0, { failStart: true })), { now: () => clock.now, wait: clock.wait }), false);
  assert.equal(rec.stops, 0, 'one that never started is not stopped');
});

test('the warm-up records a canvas of its own, drawn and asked for a frame each step, and lets its track go', async () => {
  const clock = { now: 0 };
  let rec = null;
  let frames = 0;
  const track = { stopped: false, requestFrame() { frames += 1; }, stop() { track.stopped = true; } };
  const stream = { getVideoTracks: () => [track], getTracks: () => [track] };
  const fills = [];
  const made = {
    width: 0, height: 0, rate: null,
    getContext: () => ({ set fillStyle(v) { fills.push(v); }, fillRect() {} }),
    captureStream(rate) { made.rate = rate; return stream; },
  };
  class Recorder {
    constructor(s, o) { rec = fakeRecorder(clock, 1000); rec.stream = s; rec.options = o; return rec; }
  }
  const answered = await warmOnOwnCanvas('video/mp4;codecs=avc1', {
    canvas: () => made, Recorder, now: () => clock.now, wait: async (ms) => { clock.now += ms; rec?.tick(); },
  });
  assert.equal(answered, true);
  assert.equal(rec.stream, stream, 'the recorder is on the canvas made for it');
  assert.equal(rec.options.mimeType, 'video/mp4;codecs=avc1', 'in the type the real recording is made in');
  assert.deepEqual([made.width, made.height, made.rate], [480, 270, 0]);
  assert.ok(frames >= 30, `a frame asked for each step (${frames})`);
  assert.ok(fills.includes('#222') && fills.includes('#444'), 'each frame differs from the last');
  assert.equal(rec.state, 'inactive');
  assert.equal(track.stopped, true, 'its canvas track is let go');
});

test('a warm-up with no recorder, or whose canvas fails, leaves the real recording to go ahead', async () => {
  assert.equal(await warmOnOwnCanvas('video/webm', { Recorder: null }), false);
  assert.equal(await warmOnOwnCanvas('video/webm', { Recorder: class {}, canvas: () => { throw new Error('no canvas'); } }), false);
});

test('an empty recording is not a recording', () => {
  assert.equal(recordedSomething(new Uint8Array(0)), false);
  assert.equal(recordedSomething(null), false);
  assert.equal(recordedSomething(new Uint8Array(3)), true);
  assert.equal(recordedSomething(new ArrayBuffer(0)), false);
  assert.equal(recordedSomething({ size: 10 }), true, 'a Blob by its size');
});
