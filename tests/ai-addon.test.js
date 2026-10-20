// AI Tools is an add-on (D-2026-10-20-19): with a choice made that does not
// add it, the tools answer nothing, whatever their own switch says; added,
// their switch decides as before; a copy with no choice made yet goes by
// the switch alone.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openNow, MODE_KEY } from '../apps/desktop/main/mcp/access.js';

test('AI Tools as an add-on: not added, refused with how to add it; added, the switch decides', () => {
  const off = openNow({ [MODE_KEY]: 'always', 'apps.choice': { off: [], on: [], addons: [] } });
  assert.equal(off.open, false);
  assert.match(off.reason, /Choose apps/);
  assert.equal(openNow({ [MODE_KEY]: 'always', 'apps.choice': { off: [], on: [], addons: ['ai'] } }).open, true);
  assert.equal(openNow({ [MODE_KEY]: 'off', 'apps.choice': { off: [], on: [], addons: ['ai'] } }).open, false);
  assert.equal(openNow({ [MODE_KEY]: 'always' }).open, true);
});
