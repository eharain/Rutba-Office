// The presenter view's clock (A4.1): Pause pauses the clock only, and the
// show goes on where it is; Resume takes up the time from where it was
// paused; Reset puts the clock back to nought and leaves the slide alone.
import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentService } from '../apps/desktop/main/present.js';

function service() {
  let at = 1_000_000;
  const sent = [];
  const present = createPresentService({ broadcast: (name, state) => sent.push({ name, state }), now: () => at });
  return { present, sent, tick: (ms) => { at += ms; }, at: () => at };
}

test('Pause holds the clock and the show goes on, on the same slide', () => {
  const { present, tick, at } = service();
  present.set({ id: 'd', index: 0, running: true });
  present.set({ index: 3 });
  tick(65_000);
  const paused = present.set({ pause: true });
  assert.equal(paused.running, true, 'the show is not ended');
  assert.equal(paused.index, 3, 'the slide stays');
  assert.equal(paused.pausedAt - paused.startedAt, 65_000, 'held at 1:05');
  tick(30_000);
  assert.equal(present.state().pausedAt - present.state().startedAt, 65_000, 'still 1:05 half a minute later');
  // Moving on while paused moves the show, not the clock.
  const moved = present.set({ index: 4 });
  assert.equal(moved.pausedAt - moved.startedAt, 65_000);
  const going = present.set({ pause: false });
  assert.equal(going.pausedAt, null);
  assert.equal(at() - going.startedAt, 65_000, 'the time is taken up from where it was paused');
});

test('Reset puts the clock back to nought, held or going as it was, and leaves the slide alone', () => {
  const { present, tick, at } = service();
  present.set({ id: 'd', index: 0, running: true });
  present.set({ index: 5, step: 2 });
  tick(90_000);
  const reset = present.set({ resetClock: true });
  assert.equal(reset.index, 5);
  assert.equal(reset.step, 2);
  assert.equal(reset.running, true);
  assert.equal(at() - reset.startedAt, 0);
  tick(10_000);
  present.set({ pause: true });
  tick(10_000);
  const held = present.set({ resetClock: true });
  assert.equal(held.pausedAt - held.startedAt, 0, 'paused at 0:00');
  assert.ok(held.pausedAt != null, 'still paused');
  assert.equal(held.index, 5);
});

test('ending the show still ends it, and a new show starts a new clock', () => {
  const { present, tick } = service();
  present.set({ id: 'd', index: 0, running: true });
  tick(5_000);
  present.set({ pause: true });
  const ended = present.set({ running: false });
  assert.equal(ended.startedAt, null);
  assert.equal(ended.pausedAt, null);
  tick(5_000);
  const again = present.set({ running: true });
  assert.equal(again.pausedAt, null);
  assert.ok(again.startedAt > 0);
  // Pause and Reset with no show leave everything as it was.
  present.set({ running: false });
  assert.deepEqual(present.set({ pause: true }).pausedAt, null);
  assert.equal(present.set({ resetClock: true }).startedAt, null);
});

test('every change is broadcast whole, so both windows agree', () => {
  const { present, sent } = service();
  present.set({ id: 'd', running: true });
  present.set({ pause: true });
  assert.equal(sent.length, 2);
  assert.equal(sent[1].name, 'present:state');
  assert.ok(!('pause' in sent[1].state), 'the request is not kept as state');
});
