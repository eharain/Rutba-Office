// The table a window registers for a pack's language (renderer/pack-table.js,
// L3): over Office's own catalogue for a language it carries built in.

import test from 'node:test';
import assert from 'node:assert/strict';
import { packTable } from '../apps/desktop/renderer/pack-table.js';

const builtIn = { ur: { Close: 'بند کریں', 'New in this Office': 'اس آفس میں نیا' } };

test('a pack for a built-in language is laid over Office\'s own words; any other pack is the whole table', () => {
  // After Update anyway, or a pack rolled back: the label new in this Office keeps Office's Urdu.
  assert.deepEqual(packTable('ur', { Close: 'بند' }, builtIn), { Close: 'بند', 'New in this Office': 'اس آفس میں نیا' });
  // Every entry stale: the window keeps Office's Urdu, not empty.
  assert.deepEqual(packTable('UR', {}, builtIn), builtIn.ur);
  assert.deepEqual(packTable('fr', { Close: 'Fermer' }, builtIn), { Close: 'Fermer' });
  assert.equal(packTable('ur', null, builtIn), null);
  assert.deepEqual(packTable('ur', { Close: 'x' }), { Close: 'x' });
});
