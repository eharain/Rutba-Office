// When the studio's window holds a close to ask first
// (renderer/studio/close-guard.js): while it records, stops or is live, but
// never a close it asked for itself, as the Recorder does the moment the
// window it recorded for says it took the takes, while its page still reads
// as stopping.
import test from 'node:test';
import assert from 'node:assert/strict';
import { closeHeld, lateCloseAllowed, sessionUnkept, arranged, offersSaved, setupChoices } from '../apps/desktop/renderer/studio/close-guard.js';

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

// A new session is written nowhere (D-2026-11-22-11): one the person arranged
// is asked about before it is left, and one nobody touched is not.
const NEW = { clean: true, id: '', name: 'New session' };
const SAVED = { id: 's1', name: 'Tuesday show' };
const TOOLS = { clean: true, id: '', tools: true };

test('leaving a new session the person arranged is asked about; a saved setup and an untouched session are not', () => {
  assert.equal(sessionUnkept({ setup: NEW, changed: true }), true);
  assert.equal(sessionUnkept({ setup: NEW, changed: false }), false, 'nothing was arranged: there is nothing to keep');
  assert.equal(sessionUnkept({ setup: SAVED, changed: true }), false, 'a saved setup is written as it changes');
  assert.equal(sessionUnkept({ setup: TOOLS, changed: true }), false, 'the tools\' own session is theirs');
  assert.equal(sessionUnkept({ setup: null, changed: true }), false, 'before the window has loaded');
  assert.equal(sessionUnkept(), false);
});

test('only what the person arranges counts as a change: a scene picked or a sound turned down does not', () => {
  const before = { scenes: [{ id: 'a', items: [] }], active: 'a' };
  const added = { scenes: [{ id: 'a', items: [{ id: 'cam' }] }], active: 'a' };
  assert.equal(arranged({ record: true, before, next: added }), true, 'a source added');
  assert.equal(arranged({ before, next: added }), true, 'a change is the person\'s unless it says otherwise');
  assert.equal(arranged({ record: false, before, next: added }), false, 'a change made while running: a scene picked, a source found again');
  assert.equal(arranged({ record: true, before, next: { scenes: [{ id: 'a', items: [] }], active: 'a' } }), false, 'a change that left everything as it was');
});

test('a start with setups saved says where they are, until the session is changed, something runs, or the line is closed', () => {
  const start = { shown: true, setup: NEW, changed: false, count: 3, busy: false };
  assert.equal(offersSaved(start), true);
  assert.equal(offersSaved({ ...start, count: 0 }), false, 'nothing is saved: there is nothing to point at');
  assert.equal(offersSaved({ ...start, changed: true }), false, 'the person is arranging this one');
  assert.equal(offersSaved({ ...start, busy: true }), false, 'a recording or a broadcast is not talked over');
  assert.equal(offersSaved({ ...start, shown: false }), false, 'closed, or a setup was opened, or the session was reset');
  assert.equal(offersSaved({ ...start, setup: SAVED }), false, 'a saved setup is open');
  assert.equal(offersSaved({ ...start, setup: TOOLS }), false);
  assert.equal(offersSaved(), false);
});

test('the Setups menu offers "Save this session as a setup" in a new session and "New setup" with a saved one open, never both', () => {
  assert.deepEqual(setupChoices(false), ['keep']);
  assert.deepEqual(setupChoices(true), ['new']);
});
