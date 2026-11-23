// The microphone and the camera of a meeting (renderer/meeting/devices.js,
// and what mesh.js and meet-room.js do with them): what a person chooses
// before coming in, with nothing listening until they have joined; why a
// microphone or a camera could not be used, known so a screen can say it; the
// microphone tried again; and a meeting that cannot be joined refused before
// any microphone opens. Run on a stand-in for WebRTC: nothing here asks the
// computer for a microphone or a camera.

import test from 'node:test';
import assert from 'node:assert/strict';
import { deviceFault, Preview } from '../apps/desktop/renderer/meeting/devices.js';
import { Mesh } from '../apps/desktop/renderer/meeting/mesh.js';
import { MeetingRoom } from '../apps/desktop/renderer/meeting/meet-room.js';
import { fakeRtc } from './talk-fake-rtc.js';

const named = (name) => Object.assign(new Error(name), { name });

/** The stand-in for WebRTC with a microphone or a camera that fails as the browser says it does; `faults` can be changed as a check goes. */
function rtcWith(faults = {}) {
  const rtc = fakeRtc();
  const real = rtc.getUserMedia;
  rtc.faults = { microphone: null, camera: null, ...faults };
  rtc.getUserMedia = async (c) => {
    const fault = c.audio ? rtc.faults.microphone : rtc.faults.camera;
    if (fault) { rtc.asked.push(c); throw named(fault); }
    return real(c);
  };
  return rtc;
}
const meshOf = (rtc) => new Mesh({ me: 'a', rtc, setup: async () => ({ iceServers: [] }), signal: () => {}, onChange: () => {} });

test('why a microphone or a camera could not be used: the computer has none, or it could not be used', () => {
  for (const name of ['NotFoundError', 'DevicesNotFoundError', 'OverconstrainedError']) assert.equal(deviceFault(named(name)), 'none', name);
  for (const name of ['NotAllowedError', 'NotReadableError', 'AbortError', 'SecurityError', 'Error']) assert.equal(deviceFault(named(name)), 'blocked', name);
  assert.equal(deviceFault(null), 'blocked');
});

/* ── before coming in ── */

test('before coming in nothing listens: the microphone is only chosen, and the camera opens when it is turned on', async () => {
  const rtc = fakeRtc();
  let changes = 0;
  const pick = new Preview({ rtc, onChange: () => { changes += 1; } });
  assert.deepEqual([pick.micOn, pick.camera, pick.opening], [true, null, false]);
  pick.setMic(false);
  assert.equal(pick.micOn, false);
  pick.setMic(true);
  assert.equal(rtc.asked.length, 0, 'the microphone is not opened to choose it');

  const opening = pick.setCamera(true);
  assert.equal(pick.opening, true, 'the screen can show it is on its way');
  await opening;
  assert.equal(pick.opening, false);
  assert.equal(pick.camera.kind, 'video');
  assert.deepEqual(rtc.asked.map((c) => [Boolean(c.audio), Boolean(c.video)]), [[false, true]], 'only the camera was asked for');
  assert.ok(changes >= 4, 'the screen is told at each step');

  // Off again: the camera is let go.
  const camera = pick.camera;
  await pick.setCamera(false);
  assert.equal(camera.stopped, true);
  assert.equal(pick.camera, null);
});

test('what was chosen goes into the meeting: the same camera, and the microphone off if it was turned off', async () => {
  const rtc = fakeRtc();
  const pick = new Preview({ rtc });
  await pick.setCamera(true);
  pick.setMic(false);
  const camera = pick.camera;
  assert.deepEqual(pick.take(), { video: true, muted: true, camera });
  // The meeting has it now: closing the dialog does not stop it.
  pick.close();
  assert.equal(camera.stopped, false);
  assert.deepEqual(new Preview({ rtc }).take(), { video: false, muted: false, camera: null });
});

test('a dialog closed without joining lets the camera go, also one still opening', async () => {
  const rtc = fakeRtc();
  const pick = new Preview({ rtc });
  await pick.setCamera(true);
  const camera = pick.camera;
  pick.close();
  assert.equal(camera.stopped, true);
  assert.equal(pick.camera, null);

  // Closed while the camera was still opening: it is stopped as soon as it arrives.
  let got = null;
  const slow = { getUserMedia: async (c) => { const s = await rtc.getUserMedia(c); got = s.getVideoTracks()[0]; return s; } };
  const late = new Preview({ rtc: slow });
  const opening = late.setCamera(true);
  late.close();
  await opening;
  assert.equal(got.stopped, true);
  assert.equal(late.camera, null);

  // Joined while it was still opening: the meeting comes in without it, and it is not left on.
  got = null;
  const hasty = new Preview({ rtc: slow });
  const coming = hasty.setCamera(true);
  assert.deepEqual(hasty.take(), { video: false, muted: false, camera: null });
  await coming;
  assert.equal(got.stopped, true);
});

test('a camera that cannot be used before coming in is known as which, and the person still comes in', async () => {
  const none = new Preview({ rtc: rtcWith({ camera: 'NotFoundError' }) });
  await none.setCamera(true);
  assert.deepEqual([none.camera, none.cameraFault, none.opening], [null, 'none', false]);
  const blocked = new Preview({ rtc: rtcWith({ camera: 'NotAllowedError' }) });
  await blocked.setCamera(true);
  assert.equal(blocked.cameraFault, 'blocked');
  await blocked.setCamera(false);
  assert.equal(blocked.cameraFault, null, 'turned off again, nothing is wrong');
  assert.deepEqual(blocked.take(), { video: false, muted: false, camera: null });
});

/* ── in the meeting ── */

test('in the meeting: a microphone that is missing or could not be used is known as which, and the person listens', async () => {
  const none = meshOf(rtcWith({ microphone: 'NotFoundError' }));
  await none.start();
  assert.deepEqual([none.mic, none.micFault, none.muted], [null, 'none', true]);
  const blocked = meshOf(rtcWith({ microphone: 'NotAllowedError' }));
  await blocked.start();
  assert.deepEqual([blocked.mic, blocked.micFault, blocked.muted], [null, 'blocked', true]);
  const fine = meshOf(fakeRtc());
  await fine.start();
  assert.deepEqual([fine.micFault, fine.cameraFault, fine.muted], [null, null, false]);
  for (const m of [none, blocked, fine]) m.close();
});

test('the microphone tried again once it can be used: on, and heard by those already connected', async () => {
  const rtc = rtcWith({ microphone: 'NotAllowedError' });
  const meshes = {};
  for (const id of ['a', 'b']) meshes[id] = new Mesh({ me: id, rtc, setup: async () => ({ iceServers: [] }), signal: (to, data) => meshes[to]?.signalled(id, data), onChange: () => {} });
  const settle = async () => { for (let i = 0; i < 6; i++) for (const m of Object.values(meshes)) await m.line; };
  await meshes.a.start();
  assert.equal(meshes.a.micFault, 'blocked');
  // The person allows the microphone in the computer's settings.
  rtc.faults.microphone = null;
  await meshes.b.start();
  meshes.a.expect({ id: 'b', name: 'Ben' });
  meshes.b.call({ id: 'a', name: 'Amira' });
  await settle();
  const sound = [...meshes.a.peers.values()][0].pc.getTransceivers()[0];
  assert.equal(sound.sender.track, null, 'nothing to send yet');
  assert.equal(meshes.b.people()[0].muted, true, 'the others see them muted');

  assert.equal(await meshes.a.retryMic(), true);
  await settle();
  assert.deepEqual([meshes.a.micFault, meshes.a.muted, meshes.a.mic.kind], [null, false, 'audio']);
  assert.equal(sound.sender.track, meshes.a.mic, 'the microphone goes to those already connected');
  assert.equal(meshes.b.people()[0].muted, false, 'and they see it on');
  // Asked again with a microphone that works: nothing is opened twice.
  const asked = rtc.asked.length;
  assert.equal(await meshes.a.retryMic(), true);
  assert.equal(rtc.asked.length, asked);
  for (const m of Object.values(meshes)) m.close();

  // Still not usable: known again, and nothing changes.
  const still = meshOf(rtcWith({ microphone: 'NotAllowedError' }));
  await still.start();
  assert.equal(await still.retryMic(), false);
  assert.deepEqual([still.micFault, still.muted], ['blocked', true]);
  still.close();
  assert.equal(await still.retryMic(), false, 'a meeting that is over opens nothing');
});

test('coming in with what was chosen: muted from the first moment, and the camera already open used as it is', async () => {
  const rtc = fakeRtc();
  const camera = new rtc.Track('video');
  const mesh = meshOf(rtc);
  await mesh.start({ video: true, muted: true, camera });
  assert.equal(mesh.camera, camera);
  assert.deepEqual(rtc.asked.map((c) => Boolean(c.video)), [false], 'no second camera opened');
  assert.deepEqual([mesh.muted, mesh.mic.enabled], [true, false]);
  mesh.close();

  // A camera that ended meanwhile (unplugged) is opened afresh.
  const gone = Object.assign(new rtc.Track('video'), { readyState: 'ended' });
  const again = meshOf(rtc);
  await again.start({ video: true, camera: gone });
  assert.notEqual(again.camera, gone);
  assert.equal(again.camera.kind, 'video');
  assert.equal(again.muted, false);
  again.close();
});

test('a camera that could not be turned on in the meeting is known as which, until it is turned off again', async () => {
  const mesh = meshOf(rtcWith({ camera: 'NotReadableError' }));
  await mesh.start();
  await mesh.setCamera(true);
  assert.deepEqual([mesh.camera, mesh.cameraFault], [null, 'blocked']);
  await mesh.setCamera(false);
  assert.equal(mesh.cameraFault, null);
  mesh.close();
});

/* ── at the door ── */

const ID = 'meet-AAAAAAAAAAAAAAAAAAAA';
const AMIRA = { id: 'u-a', name: 'Amira Khan' };
const meetingOf = (more = {}) => ({ id: ID, name: 'Monday plan', starter: { id: 'u-ben', name: 'Ben Ortiz' }, live: false, locked: false, most: 4, people: [], ...more });

/** A room on Talk's service as a window sees it, the meeting looked at as `peek` says. */
function roomOf({ peek, rtc = fakeRtc() } = {}) {
  const calls = [];
  const talk = {
    meetPeek: async (a) => { calls.push('meetPeek'); return peek(a); },
    meetJoin: async () => { calls.push('meetJoin'); return { meeting: meetingOf() }; },
    meetLeave: async () => ({}),
    callSetup: async () => ({ iceServers: [] }),
    send: async () => ({}),
  };
  return { room: new MeetingRoom({ shell: { talk }, me: AMIRA, rtc, onChange: () => {} }), calls, rtc };
}

test('a full meeting is refused before any microphone opens, and a camera brought along is let go', async () => {
  const people = ['u-ben', 'u-c', 'u-d', 'u-e'].map((id) => ({ id, name: id }));
  const { room, calls, rtc } = roomOf({ peek: () => ({ meeting: meetingOf({ people }) }) });
  const camera = new rtc.Track('video');
  assert.deepEqual(await room.join(ID, { video: true, camera }), { error: 'full', most: 4 });
  assert.deepEqual(calls, ['meetPeek'], 'never asked to be let in');
  assert.equal(rtc.asked.length, 0, 'no microphone and no camera was asked for');
  assert.equal(camera.stopped, true);
  assert.deepEqual([room.state, room.why, room.mesh], ['gone', 'refused', null]);
});

test('a locked meeting, one that is over, one this person may not join and a Talk that cannot be asked are refused the same way', async () => {
  const cases = [
    [() => ({ meeting: meetingOf({ locked: true }) }), { error: 'locked', most: 4 }],
    [() => ({ error: 'no-meeting', id: ID }), { error: 'no-meeting', id: ID }],
    [() => ({ error: 'not-allowed', starter: { name: 'Ben Ortiz' }, id: ID }), { error: 'not-allowed', starter: { name: 'Ben Ortiz' }, id: ID }],
    [() => { throw new Error('gone'); }, { error: 'offline' }],
  ];
  for (const [peek, expected] of cases) {
    const { room, calls, rtc } = roomOf({ peek });
    assert.deepEqual(await room.join(ID), expected);
    assert.deepEqual(calls, ['meetPeek']);
    assert.equal(rtc.asked.length, 0);
    assert.deepEqual([room.state, room.why], ['gone', 'refused']);
  }
});

test('a meeting that can be joined is joined with what was chosen', async () => {
  const { room, calls, rtc } = roomOf({ peek: () => ({ meeting: meetingOf({ people: [{ id: 'u-ben', name: 'Ben Ortiz' }] }) }) });
  const camera = new rtc.Track('video');
  const r = await room.join(ID, { video: true, muted: true, camera });
  assert.equal(r.meeting.id, ID);
  assert.deepEqual(calls, ['meetPeek', 'meetJoin'], 'looked at, then asked to be let in');
  assert.equal(room.mesh.camera, camera);
  assert.equal(room.mesh.muted, true);
  assert.deepEqual(rtc.asked.map((c) => Boolean(c.audio)), [true], 'the microphone only: the camera came with them');
  room.leave();
  assert.equal(camera.stopped, true, 'and leaving lets it go');
});
