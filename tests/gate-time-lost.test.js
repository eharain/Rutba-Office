// The gate's table of what the waits that gave up cost (12 Nov: 3,511 s of
// lane time, found only by reading the logs by hand).
import test from 'node:test';
import assert from 'node:assert/strict';
import { checkFileOf, timeLost, timeLostLines } from '../tools/gate-time-lost.js';

const LOG = [
  '     [wait] gave up on the window at 1680 px after 15039 ms (verify-ribbon.js:149)',
  '     [wait] gave up on the window at 900 px after 15012 ms (verify-ribbon.js:149)',
  'ok   [02.0001] ribbon: something passed',
  '     [wait] gave up on the greyed picture after 15060 ms (verify-image-layers.js:50 < verify-image-tools.js:138)',
  '     [wait] gave up on the main panel after 15072 ms (verify-recorder-app.js:76 < :195)',
  '     [wait] gave up on the tab after 4000 ms',
].join('\r\n');

test('a wait is put on the check file that asked for it, not on the shared helper it went through', () => {
  assert.equal(checkFileOf('verify-image-layers.js:50 < verify-image-tools.js:138'), 'verify-image-tools.js');
  assert.equal(checkFileOf('verify-recorder-app.js:76 < :195'), 'verify-recorder-app.js');
  assert.equal(checkFileOf('verify-ribbon.js:149'), 'verify-ribbon.js');
  assert.equal(checkFileOf('verify-word-tab-keys.js:51 < verify-apps.js:924'), 'verify-word-tab-keys.js');
  assert.equal(checkFileOf('verify-apps.js:884 < :7846'), 'verify-apps.js');
  assert.equal(checkFileOf(''), 'unknown');
});

test('the waits that gave up are totalled by file, the costliest first, with the wait most often to blame', () => {
  const lost = timeLost([LOG, '     [wait] gave up on the window at 1280 px after 15000 ms (verify-ribbon.js:149)']);
  assert.equal(lost.waits, 6);
  assert.equal(lost.ms, 15039 + 15012 + 15060 + 15072 + 4000 + 15000);
  assert.equal(lost.rows[0].file, 'verify-ribbon.js');
  assert.equal(lost.rows[0].waits, 3);
  assert.deepEqual(lost.rows[0].most, ['the window at N px', 3]);
  assert.deepEqual(lost.rows.map((r) => r.file), ['verify-ribbon.js', 'verify-recorder-app.js', 'verify-image-tools.js', 'unknown']);
});

test('the table says the total, a line a file, and the rest in one line', () => {
  const lines = timeLostLines(timeLost([LOG]), 2);
  assert.match(lines[0], /5 waits, 64 s of lane time/);
  assert.match(lines[1], /30 s +2 +verify-ribbon\.js +\(most: the window at N px, 2\)/);
  assert.match(lines[3], /2 other files/);
  assert.deepEqual(timeLostLines(timeLost(['nothing here'])), ['  time lost to waits that gave up: none']);
});
