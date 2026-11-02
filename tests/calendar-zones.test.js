// Zones by today's names: the runtime lists some by names the zone database
// has replaced, and Calendar showed "Asia/Calcutta" to someone in India.
import test from 'node:test';
import assert from 'node:assert/strict';
import { currentZone, ianaZone, offsetMinutes, CURRENT_ZONE } from '../packages/calendar/src/tz.js';

test('an old zone name reads as today\'s, and keeps its clock', () => {
  assert.equal(currentZone('Asia/Calcutta'), 'Asia/Kolkata');
  assert.equal(currentZone('Europe/Kiev'), 'Europe/Kyiv');
  assert.equal(currentZone('Asia/Saigon'), 'Asia/Ho_Chi_Minh');
  assert.equal(currentZone('Europe/London'), 'Europe/London', 'a current name is left as it is');
  const at = Date.UTC(2026, 10, 2, 12);
  for (const [old, now] of Object.entries(CURRENT_ZONE)) {
    assert.equal(offsetMinutes(now, at), offsetMinutes(old, at), `${now} keeps ${old}'s clock`);
  }
});

test('a file\'s zone, old name or new, comes back by today\'s name', () => {
  assert.equal(ianaZone('Asia/Calcutta'), 'Asia/Kolkata');
  assert.equal(ianaZone('Asia/Kolkata'), 'Asia/Kolkata');
  assert.equal(ianaZone('India Standard Time'), 'Asia/Kolkata');
});

test('every old name the runtime lists has today\'s name beside it, or is not old', () => {
  const listed = new Set(Intl.supportedValuesOf('timeZone'));
  for (const old of ['Asia/Calcutta', 'Asia/Saigon', 'Asia/Katmandu', 'Asia/Rangoon', 'Europe/Kiev']) {
    if (listed.has(old)) assert.notEqual(currentZone(old), old, `${old} is listed and must be renamed`);
  }
});
