// The slideshow's arithmetic: stepping through the folder, reading a
// length, and how long each item stays on screen — the part that does not
// need a window to be right.
import test from 'node:test';
import assert from 'node:assert/strict';
import { nextIndex, formatLength, advanceAfter } from '../apps/desktop/renderer/apps/pictures/show.js';

test('nextIndex steps through the folder and loops either way, always', () => {
  assert.equal(nextIndex(5, 2, 1), 3);
  assert.equal(nextIndex(5, 4, 1), 0, 'loops forward past the end');
  assert.equal(nextIndex(5, 0, -1), 4, 'loops backward past the start');
  assert.equal(nextIndex(1, 0, 1), 0, 'a single picture stays put');
  assert.equal(nextIndex(0, 0, 1), -1, 'nothing to show');
});

test('formatLength reads a duration the way a clip\'s tile badge does', () => {
  assert.equal(formatLength(7), '0:07');
  assert.equal(formatLength(185), '3:05');
  assert.equal(formatLength(3723), '1:02:03');
  assert.equal(formatLength(0), '0:00');
  assert.equal(formatLength(59), '0:59');
  assert.equal(formatLength(3600), '1:00:00');
  assert.equal(formatLength(NaN), '--:--');
  assert.equal(formatLength(Infinity), '--:--');
  assert.equal(formatLength(-1), '--:--');
});

test('advanceAfter gives a picture its chosen interval and leaves a clip up to a minute', () => {
  assert.equal(advanceAfter({ kind: 'still' }, 4), 4);
  assert.equal(advanceAfter({ kind: 'maybe-animated' }, 8), 8);
  assert.equal(advanceAfter({ kind: 'pdf' }, 15), 15, 'anything that is not a clip gets the chosen interval');
  assert.equal(advanceAfter({ kind: 'video', duration: 12 }, 4), 12, 'a short clip gets its own length');
  assert.equal(advanceAfter({ kind: 'audio', duration: 500 }, 4), 60, 'a long clip is capped at a minute');
  assert.equal(advanceAfter({ kind: 'video', duration: NaN }, 4), 60, 'an unknown duration is capped too');
  assert.equal(advanceAfter({ kind: 'video' }, 2), 60, 'no duration at all is the same as unknown');
});
