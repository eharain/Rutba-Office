// A call that comes in while Live Broadcast is on air is turned away as busy
// (2.28.0, the screen review): no ring reaches the air through the computer's
// sound, the caller is kept as a missed call, and Talk says who called. Asked
// of the broadcast itself; with no Live Broadcast to ask, calls ring as before.
import test from 'node:test';
import assert from 'node:assert/strict';
import { incomingCall, onAirNow } from '../apps/desktop/renderer/apps/talk/ring.js';

test('on air, a call is turned away as busy; off air it rings', () => {
  assert.equal(incomingCall({ callId: 'c1', onAir: true }), 'busy');
  assert.equal(incomingCall({ callId: 'c1', onAir: false }), 'ring');
  assert.equal(incomingCall({ callId: 'c1' }), 'ring', 'as before when nothing is said of the air');
  // The call already ringing when the broadcast began is told again: it stays as it is.
  assert.equal(incomingCall({ callId: 'c1', ringing: 'c1', onAir: true }), 'same');
  assert.equal(incomingCall({ callId: 'c2', ringing: 'c1', onAir: true }), 'busy');
});

test('whether it is on air is the broadcast\'s own word, and false when it cannot be asked', async () => {
  assert.equal(await onAirNow({ live: { status: async () => ({ live: true }) } }), true);
  assert.equal(await onAirNow({ live: { status: async () => ({ live: false }) } }), false);
  assert.equal(await onAirNow({ live: { status: async () => { throw new Error('Live Broadcast is not installed.'); } } }), false);
  assert.equal(await onAirNow({}), false);
  assert.equal(await onAirNow(null), false);
});
