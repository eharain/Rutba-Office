// Go Live's bitrate follows the upload: down a quarter when a destination
// drops frames or its queue fills, not below the floor; back up a tenth
// after a calm spell, not above the target.
import test from 'node:test';
import assert from 'node:assert/strict';
import { BitrateGovernor } from '../packages/stream/src/adapt.js';

test('live bitrate: steady when every destination keeps up', () => {
  const g = new BitrateGovernor({ target: 4500 });
  for (let t = 0; t < 60000; t += 2000) assert.equal(g.tick([{ dropped: 0, queued: 0.1 }, { dropped: 0, queued: 0 }], t), 4500);
  assert.equal(g.lowered, false);
});

test('live bitrate: down a quarter for a destination dropping frames or filling up, to the floor and no lower', () => {
  const g = new BitrateGovernor({ target: 4000 });
  assert.equal(g.tick([{ dropped: 3, queued: 0 }, { dropped: 0, queued: 0 }], 0), 3000);
  assert.equal(g.tick([{ dropped: 0, queued: 0.8 }], 2000), 2250);
  for (let t = 4000; t < 40000; t += 2000) g.tick([{ dropped: 5, queued: 1 }], t);
  assert.equal(g.kbps, 1000);
  assert.equal(g.lowered, true);
});

test('live bitrate: back up a tenth after ten calm seconds, step by step, to the target', () => {
  const g = new BitrateGovernor({ target: 3000, recoverMs: 10000 });
  g.tick([{ dropped: 1 }], 0);
  assert.equal(g.kbps, 2250);
  assert.equal(g.tick([{ dropped: 0 }], 4000), 2250);
  assert.equal(g.tick([{ dropped: 0 }], 10000), 2475);
  assert.equal(g.tick([{ dropped: 0 }], 12000), 2475);
  let t = 20000;
  while (g.kbps < 3000 && t < 200000) { g.tick([{ dropped: 0 }], t); t += 10000; }
  assert.equal(g.kbps, 3000);
  assert.equal(g.lowered, false);
});
