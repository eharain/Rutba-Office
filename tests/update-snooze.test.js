// "Not now" on the update prompt: one version, one day, never a newer one.
import test from 'node:test';
import assert from 'node:assert/strict';
import { snoozeFor, isSnoozed, SNOOZE_MS } from '../apps/desktop/main/update-snooze.js';

test('a snooze covers the version it was given for a day and no other', () => {
  const now = 1_000_000;
  const entry = snoozeFor('1.13.0', now);
  assert.deepEqual(entry, { version: '1.13.0', until: now + SNOOZE_MS });
  assert.equal(isSnoozed(entry, '1.13.0', now + 1), true);
  assert.equal(isSnoozed(entry, '1.13.0', now + SNOOZE_MS - 1), true);
  assert.equal(isSnoozed(entry, '1.13.0', now + SNOOZE_MS), false, 'the day is up');
  assert.equal(isSnoozed(entry, '1.14.0', now + 1), false, 'a newer release is a new question');
  assert.equal(isSnoozed(null, '1.13.0', now), false);
  assert.equal(isSnoozed(entry, null, now), false);
  assert.equal(snoozeFor(null), null);
});
