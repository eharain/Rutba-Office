// A group call that is on air (Office 2.27.0, the plan's phase 3): a seat
// whose sound and picture are given to it (a broadcast's programme, or a
// recording's), and the word that the call is on air or being recorded,
// heeded only from the person who set the meeting up, with each person told
// whether they are in the picture. Run on a stand-in for WebRTC and for
// Talk's service.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Mesh, readWire, OFF_AIR } from '../apps/desktop/renderer/meeting/mesh.js';
import { MeetingRoom } from '../apps/desktop/renderer/meeting/meet-room.js';
import { fakeRtc } from './talk-fake-rtc.js';

/** People of one call, their set-ups passed straight to each other; only `host` may say the call is on air. */
function callOf(names, host, rtc = fakeRtc()) {
  const meshes = {};
  for (const id of names) {
    meshes[id] = new Mesh({
      me: id,
      rtc,
      setup: async () => ({ iceServers: [] }),
      signal: (to, data) => { meshes[to]?.signalled(id, data); },
      mayAir: (from) => from === host,
    });
  }
  const settle = async () => { for (let i = 0; i < 6; i++) for (const m of Object.values(meshes)) await m.line; };
  return { meshes, rtc, settle };
}

test('the word that a call is on air is checked as it arrives', () => {
  assert.deepEqual(readWire('{"t":"air","on":true,"you":true,"how":"live"}'), { t: 'air', on: true, you: true, how: 'live' });
  assert.deepEqual(readWire('{"t":"air","on":true,"how":"record"}'), { t: 'air', on: true, you: false, how: 'record' });
  // Anything but true is not on air, nobody is in the picture of a call that is not, and an unknown way is the plain one.
  assert.deepEqual(readWire('{"t":"air","on":"yes","you":true,"how":"film"}'), { t: 'air', on: false, you: false, how: 'live' });
});

test('a seat given its sound and picture sends them, shows as a screen, and asks for no microphone or camera', async () => {
  const { meshes, rtc, settle } = callOf(['host', 'b'], 'host');
  const audio = new rtc.Track('audio');
  const video = new rtc.Track('video');
  meshes.host.provide({ audio, video });
  await meshes.b.start();
  meshes.b.call({ id: 'host' });
  await settle();
  assert.equal(rtc.asked.length, 1, 'only b asked for a microphone');
  const lines = [...meshes.host.peers.values()][0].pc.getTransceivers();
  assert.equal(lines.find((l) => l.receiver.track.kind === 'audio').sender.track, audio);
  assert.equal(lines.find((l) => l.receiver.track.kind === 'video').sender.track, video);
  assert.deepEqual(meshes.host.mine(), { muted: false, camera: false, sharing: true, hand: false });
  assert.deepEqual(meshes.b.people().map((p) => [p.id, p.sharing, p.muted]), [['host', true, false]]);
  // Given no sound, the seat says it is muted; closing stops what it was given.
  const quiet = new Mesh({ me: 'q', rtc, setup: async () => ({}), signal: () => {} });
  quiet.provide({ video: new rtc.Track('video') });
  assert.equal(quiet.mine().muted, true);
  meshes.host.close();
  assert.ok(audio.stopped && video.stopped);
});

test('on air: everyone is told, each whether they are in the picture, and only the meeting\'s own person is believed', async () => {
  const { meshes, settle } = callOf(['host', 'b', 'c'], 'host');
  meshes.host.provide({ audio: null, video: null });
  await meshes.b.start();
  await meshes.c.start();
  meshes.b.call({ id: 'host' });
  meshes.c.call({ id: 'host' });
  meshes.c.call({ id: 'b' });
  await settle();
  assert.deepEqual(meshes.b.air, OFF_AIR);

  meshes.host.setAir({ on: true, how: 'live', shown: ['b'] });
  assert.deepEqual(meshes.b.air, { on: true, you: true, how: 'live' });
  assert.deepEqual(meshes.c.air, { on: true, you: false, how: 'live' });

  // Someone else in the call says it is off, or on: nobody takes their word.
  meshes.c.setAir({ on: false });
  assert.deepEqual(meshes.b.air, { on: true, you: true, how: 'live' });
  meshes.b.setAir({ on: true, how: 'record', shown: ['host', 'c'] });
  assert.deepEqual(meshes.c.air, { on: true, you: false, how: 'live' });
  assert.deepEqual(meshes.host.air, OFF_AIR);

  meshes.host.setAir({ on: true, how: 'live', shown: ['c'] });
  assert.deepEqual([meshes.b.air.you, meshes.c.air.you], [false, true]);
  meshes.host.setAir({ on: false, how: 'live', shown: ['c'] });
  assert.deepEqual(meshes.c.air, { on: false, you: false, how: null });
});

test('someone who comes in while it is on air is told as they connect, and when the one who said so is gone it is off', async () => {
  const { meshes, settle } = callOf(['host', 'b', 'c'], 'host');
  meshes.host.provide({});
  await meshes.b.start();
  meshes.b.call({ id: 'host' });
  await settle();
  meshes.host.setAir({ on: true, how: 'record', shown: ['b', 'c'] });
  await meshes.c.start();
  meshes.c.call({ id: 'host' });
  meshes.c.call({ id: 'b' });
  await settle();
  assert.deepEqual(meshes.c.air, { on: true, you: true, how: 'record' });
  // The host's computer drops out of the call: nothing is on air until it says so again.
  await meshes.b.drop('host');
  assert.deepEqual(meshes.b.air, OFF_AIR);
  await meshes.c.drop('b');
  assert.deepEqual(meshes.c.air, { on: true, you: true, how: 'record' }, 'someone else leaving changes nothing');
  meshes.c.close();
  assert.deepEqual(meshes.c.air, OFF_AIR);
});

/* ── the meeting that holds such a call ── */

const LIVE = 'live-BBBBBBBBBBBBBBBBBBBB';
const PLAIN = 'meet-AAAAAAAAAAAAAAAAAAAA';
const AMIRA = { id: 'u-a', name: 'Amira Khan' };
const BEN = { id: 'u-ben', name: 'Ben Ortiz', contact: true };
const meetingOf = (id, more = {}) => ({ id, name: 'On air', starter: { id: 'u-a', name: 'Amira Khan' }, at: null, live: id.startsWith('live-'), recorded: false, locked: false, groupId: null, link: `https://introducer.example/m/${id}`, most: 4, people: [], ...more });

function roomOf({ meeting, me = AMIRA, feed = null, rtc = fakeRtc(), answers = {} } = {}) {
  const calls = [];
  const talk = {};
  for (const name of ['meetJoin', 'meetLeave', 'meetInvite', 'send', 'callSetup']) {
    talk[name] = async (args) => {
      calls.push({ name, args });
      if (answers[name]) return answers[name](args);
      if (name === 'callSetup') return { iceServers: [] };
      if (name === 'meetJoin') return { meeting };
      if (name === 'meetInvite') return { told: args.invite.length };
      return {};
    };
  }
  const room = new MeetingRoom({ shell: { talk }, me, rtc, feed });
  const settle = async () => { for (let i = 0; i < 6; i++) { await room.mesh?.line; await new Promise((r) => setImmediate(r)); } };
  return { room, calls, rtc, settle };
}

test('a meeting held from a broadcast: the programme is the seat, no microphone is asked for, and it comes back the same way', async () => {
  const rtc = fakeRtc();
  let fed = 0;
  const feed = async () => { fed += 1; return { audio: new rtc.Track('audio'), video: new rtc.Track('video') }; };
  const { room, settle } = roomOf({ meeting: meetingOf(LIVE, { people: [BEN] }), feed, rtc });
  await room.join(LIVE);
  await settle();
  assert.equal(rtc.asked.length, 0);
  assert.equal(fed, 1);
  assert.equal(room.mesh.given, true);
  assert.deepEqual(room.people()[0].sharing, true);
  // The connection to the introducer was made again: the seat is fed anew.
  room.handle({ type: 'welcome' });
  await settle();
  await settle();
  assert.equal(fed, 2);
  assert.equal(rtc.asked.length, 0);
  // A feed that fails still gets the person in: they send nothing.
  const broken = roomOf({ meeting: meetingOf(LIVE), feed: async () => { throw new Error('no programme'); } });
  assert.equal((await broken.room.join(LIVE)).meeting.id, LIVE);
  assert.equal(broken.room.mesh.mine().muted, true);
});

test('only the person who set the meeting up says it is on air, and only in a meeting set up for it', async () => {
  const mine = roomOf({ meeting: meetingOf(LIVE) });
  await mine.room.join(LIVE);
  assert.equal(mine.room.setAir({ on: true, how: 'live', shown: ['u-ben'] }), true);
  assert.deepEqual([mine.room.mesh.airing.on, mine.room.mesh.airing.how, [...mine.room.mesh.airing.shown]], [true, 'live', ['u-ben']]);
  // Set up to be streamed, not to be recorded: it cannot be said to be recorded.
  assert.equal(mine.room.setAir({ on: true, how: 'record' }), false);

  const plain = roomOf({ meeting: meetingOf(PLAIN) });
  await plain.room.join(PLAIN);
  assert.equal(plain.room.setAir({ on: true, how: 'live' }), false, 'a meeting set up without the switch has no going live');
  assert.equal(plain.room.mesh.airing, null);

  const recorded = roomOf({ meeting: meetingOf(PLAIN, { recorded: true }) });
  await recorded.room.join(PLAIN);
  assert.equal(recorded.room.setAir({ on: true, how: 'record', shown: [] }), true);
  assert.equal(recorded.room.setAir({ on: true, how: 'live' }), false);

  const guest = roomOf({ meeting: meetingOf(LIVE), me: { id: 'u-ben', name: 'Ben Ortiz' } });
  await guest.room.join(LIVE);
  assert.equal(guest.room.setAir({ on: true, how: 'live' }), false);
  assert.deepEqual(guest.room.air, { on: false, you: false, how: null });
});

test('what was said of being on air is said again after the connection was made anew, and more contacts can be asked in', async () => {
  const { room, calls, settle } = roomOf({ meeting: meetingOf(LIVE) });
  await room.join(LIVE);
  room.setAir({ on: true, how: 'live', shown: ['u-ben'] });
  const before = room.mesh;
  room.handle({ type: 'welcome' });
  await settle();
  await settle();
  assert.notEqual(room.mesh, before);
  assert.deepEqual([room.mesh.airing.on, [...room.mesh.airing.shown]], [true, ['u-ben']]);

  assert.deepEqual(await room.invite(['u-ben', 'u-cara']), { told: 2 });
  assert.deepEqual(calls.find((c) => c.name === 'meetInvite').args, { id: LIVE, invite: ['u-ben', 'u-cara'] });
  const off = roomOf({ meeting: meetingOf(LIVE), answers: { meetInvite: async () => { throw new Error('gone'); } } });
  await off.room.join(LIVE);
  assert.deepEqual(await off.room.invite(['u-ben']), { error: 'offline' });
});
