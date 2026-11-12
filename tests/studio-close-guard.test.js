// When the studio's window holds a close to ask first
// (renderer/studio/close-guard.js): while it records, stops or is live, but
// never a close it asked for itself, as the Recorder does the moment the
// window it recorded for says it took the takes, while its page still reads
// as stopping.
import test from 'node:test';
import assert from 'node:assert/strict';
import { closeHeld, lateCloseAllowed } from '../apps/desktop/renderer/studio/close-guard.js';

test('a close is held to ask while a recording runs or stops, or the window is live', () => {
  assert.equal(closeHeld({ phase: 'recording' }), true);
  assert.equal(closeHeld({ phase: 'stopping' }), true);
  assert.equal(closeHeld({ phase: 'setup', live: true }), true);
});

test('nothing running: the window closes without a question', () => {
  assert.equal(closeHeld({ phase: 'setup' }), false);
  assert.equal(closeHeld({ phase: 'countdown' }), false);
  assert.equal(closeHeld(), false);
});

test('a close the window asked for itself is never held, even while its page still reads as stopping', () => {
  assert.equal(closeHeld({ phase: 'stopping', closing: true }), false, 'the takes were handed over: the Recorder closes');
  assert.equal(closeHeld({ phase: 'recording', live: true, closing: true }), false, 'Stop and close, answered');
});

test('a took that comes after the wait gave up closes the Recorder only while nothing runs', () => {
  assert.equal(lateCloseAllowed('setup'), true, 'the takes are shown here: the window that asked has them after all');
  for (const phase of ['countdown', 'starting', 'recording', 'stopping']) assert.equal(lateCloseAllowed(phase), false, phase);
});
