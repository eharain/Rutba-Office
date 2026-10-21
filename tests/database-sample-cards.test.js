// Data Files' start screen offers every sample business the templates hold
// (packages/templates/data-files), in their order and under the business's
// own name: a card for a sample that is not there, or a sample with no card,
// or a name that differs between the card and the database, fails here.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { SAMPLE_IDS, sampleAbout } from '../packages/database/src/index.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SOURCE = fs.readFileSync(path.join(ROOT, 'apps/desktop/renderer/apps/base.js'), 'utf8');
const CARDS = [...SOURCE.matchAll(/\{ id: '([a-z]+)', group: '([a-z]+)', icon: '[A-Za-z]+', business: '((?:[^'\\]|\\.)*)'/g)].map((m) => ({ id: m[1], group: m[2], business: m[3].replace(/\\'/g, "'") }));
const GROUPS = [...SOURCE.matchAll(/\{ id: '([a-z]+)', name: \(\) => t\(/g)].map((m) => m[1]);

test('a card for every sample, in the samples\' order, and none for a sample that is not there', () => {
  assert.deepEqual(CARDS.map((c) => c.id), SAMPLE_IDS);
});

test('each card names the business as its sample does, and sits in a group the screen shows', () => {
  for (const c of CARDS) {
    assert.equal(c.business, sampleAbout(c.id).name, c.id);
    assert.ok(GROUPS.includes(c.group), `${c.id}: ${c.group}`);
  }
});
