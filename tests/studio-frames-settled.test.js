// The preview says when the frame it painted is the one asked for: each clip
// showing at that instant has its decoder ready, not seeking, and on the
// instant. A check waits on it after a seek, rather than a fixed time
// (D-2026-11-13-20, the owner: "Wait on signals (Recommended)").
import test from 'node:test';
import assert from 'node:assert/strict';
import { framesSettledAt } from '../packages/studio/src/index.js';

const el = (o) => ({ paused: true, seeking: false, readyState: 4, currentTime: 0, ...o });
const planOf = (layers, videos) => ({ duration: 10, layers, videos });

test('a paused clip on the instant, ready and not seeking, is settled; seeking, not ready or elsewhere is not', () => {
  const layer = { type: 'video', videoKey: 'a', timing: { start: 2, end: 8 }, offset: 1 };
  // At 5 s on the timeline the clip shows 1 + (5 - 2) = 4 s of its source.
  const at = (o) => framesSettledAt(planOf([layer], { a: { el: el(o), duration: 20 } }), 5);
  assert.equal(at({ currentTime: 4 }), true);
  assert.equal(at({ currentTime: 4.05 }), true, 'within the painter\'s own 0.08 s');
  assert.equal(at({ currentTime: 3 }), false, 'still on the frame before the seek');
  assert.equal(at({ currentTime: 4, seeking: true }), false);
  assert.equal(at({ currentTime: 4, readyState: 1 }), false);
});

test('only what shows at the instant is waited for: a clip before or after it, a hidden one, or a driven export frame', () => {
  const far = el({ currentTime: 0, seeking: true });
  const layers = [
    { type: 'video', videoKey: 'early', timing: { start: 0, end: 1 } },
    { type: 'video', videoKey: 'hidden', visible: false },
    { type: 'video', videoKey: 'driven' },
    { type: 'text', text: 'a title' },
  ];
  const videos = { early: { el: far }, hidden: { el: far }, driven: { frame: null } };
  assert.equal(framesSettledAt(planOf(layers, videos), 5), true);
  assert.equal(framesSettledAt(null, 5), true);
});

test('a clip showing at the instant whose decoder has not loaded is not settled: nothing of it is painted yet', () => {
  const layer = { type: 'video', videoKey: 'none' };
  assert.equal(framesSettledAt(planOf([layer], {}), 5), false);
  assert.equal(framesSettledAt(planOf([layer], { none: { duration: 10 } }), 5), false, 'an entry with no element yet');
});

test('a clip near the end of its source is judged at the point the painter clamps it to', () => {
  const layer = { type: 'video', videoKey: 'a' };
  // 9.99 s asked of a 5 s source: the painter seeks to 5 - 0.05.
  assert.equal(framesSettledAt(planOf([layer], { a: { el: el({ currentTime: 4.95 }), duration: 5 } }), 9.99), true);
  assert.equal(framesSettledAt(planOf([layer], { a: { el: el({ currentTime: 2 }), duration: 5 } }), 9.99), false);
});

test('a playing clip is settled once ready and not seeking: it moves on by itself', () => {
  const layer = { type: 'video', videoKey: 'a' };
  assert.equal(framesSettledAt(planOf([layer], { a: { el: el({ paused: false, currentTime: 1 }), duration: 20 } }), 5), true);
});
