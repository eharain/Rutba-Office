// A meeting held on a Talk Server, as a window holds it
// (apps/desktop/renderer/meeting/server-call.js): the same seat as in a
// small meeting, with the server passing everyone's sound and picture on,
// all of it encrypted with the meeting's key, and the line beside them
// sealed with it. Run on a stand-in for the
// engine and for Talk's service (talk-fake-room.js) and for WebRTC.

import test from 'node:test';
import assert from 'node:assert/strict';
import { ServerCall, WIRE_TOPIC, engineKey, whyGone, whyNotConnected } from '../apps/desktop/renderer/meeting/server-call.js';
import { fakeRtc } from './talk-fake-rtc.js';
import { fakeWorld, ConnectionErrorReason, DisconnectReason, RoomEvent } from './talk-fake-room.js';

const KEY = Uint8Array.from({ length: 32 }, (_, i) => i + 1);
const MID = 'meet-AAAAAAAAAAAAAAAAAAAA';
const NAMES = { 'u-amira': 'Amira Khan', 'u-ben': 'Ben Ortiz', 'u-cara': 'Cara Lund' };
const nativeSocket = globalThis.WebSocket;

/** People of one meeting on one server; `u-amira` set it up. */
function meetingOf(ids, { world = fakeWorld({ names: NAMES }), rtc = fakeRtc(), starter = 'u-amira' } = {}) {
  const calls = {};
  const log = { says: [], muted: [], gone: [], changes: 0 };
  for (const id of ids) {
    calls[id] = new ServerCall({
      me: id,
      rtc,
      shell: world.shell(id, { role: id === starter ? 'host' : 'member' }),
      load: world.load,
      worker: world.worker,
      onChange: () => { log.changes += 1; },
      onSay: (from, text, sid) => log.says.push({ at: id, from, text, id: sid }),
      mayMute: (from) => from === starter,
      mayAir: (from) => from === starter,
      onMuted: () => log.muted.push(id),
      onGone: (why) => log.gone.push({ id, why }),
    });
  }
  // Until nothing is on its way: every set-up taken its turn, every word sealed, passed on and opened.
  const settle = async () => { for (let i = 0; i < 8; i += 1) { for (const c of Object.values(calls)) { await c.line; await c.outbox; await c.inbox; } await new Promise((r) => setTimeout(r, 0)); } };
  const enter = async (id, opts = {}) => { await calls[id].start(opts); const r = await calls[id].enter({ id: MID, key: KEY }); await settle(); return r; };
  return { calls, log, world, rtc, settle, enter };
}
const closeAll = async (m) => { for (const c of Object.values(m.calls)) c.close(); await m.settle(); };
const person = (call, id) => call.people().find((p) => p.id === id);

test('three people come in, and each sees and hears the other two through the server', async () => {
  const m = meetingOf(['u-amira', 'u-ben', 'u-cara']);
  assert.deepEqual(await m.enter('u-amira', { video: true }), { entered: true });
  assert.deepEqual(await m.enter('u-ben'), { entered: true });
  assert.deepEqual(await m.enter('u-cara', { video: true }), { entered: true });
  for (const [id, call] of Object.entries(m.calls)) {
    const others = call.people();
    assert.deepEqual(others.map((p) => p.id).sort(), Object.keys(m.calls).filter((x) => x !== id).sort());
    for (const p of others) {
      assert.equal(p.connected, true);
      assert.equal(p.name, NAMES[p.id]);
      assert.equal(p.stream.getAudioTracks().length, 1, `${id} hears ${p.id}`);
      assert.equal(p.stream.getVideoTracks().length, p.id === 'u-ben' ? 0 : 1, `${id} sees ${p.id}`);
      assert.equal(p.camera, p.id !== 'u-ben');
    }
  }
  // Nobody made a connection to anybody: the only way out of each window is its way in.
  assert.equal(m.rtc.pcs.length, 0);
  assert.equal(m.world.asked.filter((a) => a.do === 'serverEnter').length, 3);
  assert.deepEqual(m.world.asked.filter((a) => a.do === 'doorOpen').map((a) => a.target), ['/rtc/v1?access_token=u-amira', '/rtc/v1?access_token=u-ben', '/rtc/v1?access_token=u-cara']);
  await closeAll(m);
});

test('the room is made encrypted, with the meeting\'s key and the worker, before it connects', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira');
  await m.enter('u-ben');
  const room = m.calls['u-amira'].room;
  assert.equal(room.e2ee, true);
  assert.equal(room.options.encryption.worker, m.world.workers[0]);
  assert.equal(room.options.adaptiveStream, false);
  assert.equal(room.options.stopLocalTrackOnUnpublish, false);
  assert.equal(m.world.keys[0], Buffer.from(KEY).toString('base64url'));
  assert.equal(engineKey(KEY), Buffer.from(KEY).toString('base64url'));
  assert.equal(engineKey(Uint8Array.from({ length: 32 }, () => 255)).length, 43);
  assert.equal(m.calls['u-amira'].encrypted, true);
  assert.equal(m.calls['u-ben'].encrypted, true);
  await closeAll(m);
});

test('it is said to be encrypted only while every sound and picture that arrives is', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira');
  await m.enter('u-ben');
  assert.equal(m.calls['u-amira'].encrypted, true);
  m.world.plainTrack('u-amira', 'u-ben', 'video');
  assert.equal(m.calls['u-amira'].encrypted, false);
  assert.equal(m.calls['u-ben'].encrypted, true);
  await closeAll(m);
});

test('without a key of 32 bytes nothing is asked of anyone', async () => {
  const m = meetingOf(['u-amira']);
  await m.calls['u-amira'].start();
  for (const key of [null, 'a string', new Uint8Array(16)]) assert.deepEqual(await m.calls['u-amira'].enter({ id: MID, key }), { error: 'failed' });
  assert.equal(m.world.asked.length, 0);
  await closeAll(m);
});

test('muted, camera, a shared screen and a raised hand reach the others, and only as words sealed with the meeting\'s key', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira');
  await m.enter('u-ben');
  const { 'u-amira': amira, 'u-ben': ben } = m.calls;
  ben.setMuted(true);
  ben.setHand(true);
  await ben.setCamera(true);
  await m.settle();
  assert.deepEqual({ muted: person(amira, 'u-ben').muted, camera: person(amira, 'u-ben').camera, hand: person(amira, 'u-ben').hand }, { muted: true, camera: true, hand: true });
  assert.equal(person(amira, 'u-ben').stream.getVideoTracks().length, 1);
  assert.equal(ben.sent.audio.pub.isMuted, true);

  // A screen shown takes the camera's place, and the camera comes back when it stops.
  const screen = new m.rtc.Track('video');
  ben.share(new m.rtc.MediaStream([screen]));
  await m.settle();
  assert.equal(person(amira, 'u-ben').sharing, true);
  assert.deepEqual(person(amira, 'u-ben').stream.getVideoTracks(), [screen]);
  ben.share(null);
  await m.settle();
  assert.equal(person(amira, 'u-ben').sharing, false);
  assert.deepEqual(person(amira, 'u-ben').stream.getVideoTracks(), [ben.camera]);

  await ben.setCamera(false);
  ben.setMuted(false);
  await m.settle();
  assert.equal(person(amira, 'u-ben').stream.getVideoTracks().length, 0);
  assert.equal(person(amira, 'u-ben').muted, false);
  assert.equal(ben.sent.audio.pub.isMuted, false);
  assert.equal(ben.mic.enabled, true);

  ben.setHand(false);
  await m.settle();
  assert.equal(person(amira, 'u-ben').hand, false);

  // Nothing on the line can be read by whoever passes it on.
  assert.ok(m.world.carried.length > 4);
  for (const c of m.world.carried) assert.ok(!/muted|hand|state/.test(Buffer.from(c.bytes).toString('latin1')));

  // The same words written by the server itself are not heeded, whatever the engine says of them.
  m.world.inject('u-amira', 'u-ben', { t: 'state', muted: true, camera: true, sharing: true, hand: true });
  m.world.inject('u-amira', 'u-ben', { t: 'state', muted: true, camera: true, sharing: true, hand: true }, { encryption: 1, topic: WIRE_TOPIC });
  await m.settle();
  assert.equal(person(amira, 'u-ben').hand, false);
  // Nor what a person did say, passed on a second time: Ben's raised hand of a moment ago.
  const raised = m.world.carried.filter((c) => c.from === 'u-ben').at(-2);
  m.world.pass('u-amira', 'u-ben', raised.bytes);
  await m.settle();
  assert.equal(person(amira, 'u-ben').hand, false);
  // Nor what somebody else said, passed on as Ben's.
  amira.setHand(true);
  await m.settle();
  const hers = m.world.carried.filter((c) => c.from === 'u-amira').at(-1);
  m.world.pass('u-amira', 'u-ben', hers.bytes);
  await m.settle();
  assert.equal(person(amira, 'u-ben').hand, false);
  // Nor anything on another line than ours.
  ben.setHand(true);
  await m.settle();
  assert.equal(person(amira, 'u-ben').hand, true);
  await closeAll(m);
});

test('what was said to one person alone does not open as said to everyone, nor for anybody else', async () => {
  const m = meetingOf(['u-amira', 'u-ben', 'u-cara']);
  for (const id of Object.keys(m.calls)) await m.enter(id);
  // Amira tells Ben alone that he is in the picture.
  m.calls['u-amira'].setAir({ on: true, how: 'record', shown: ['u-ben'] });
  await m.settle();
  assert.deepEqual([m.calls['u-ben'].air.you, m.calls['u-cara'].air.you], [true, false]);
  const forBen = m.world.carried.filter((c) => c.from === 'u-amira' && c.to === 'u-ben').at(-1);
  m.calls['u-amira'].setAir({ on: false });
  await m.settle();
  assert.equal(m.calls['u-cara'].air.on, false);
  // The server puts what was Ben's before Cara: it does not open for her.
  m.world.pass('u-cara', 'u-amira', forBen.bytes);
  await m.settle();
  assert.deepEqual([m.calls['u-cara'].air.on, m.calls['u-cara'].air.you], [false, false]);
  await closeAll(m);
});

test('with another key than the others, nothing they say is read, and the window knows it cannot read the meeting', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira');
  await m.calls['u-ben'].start();
  assert.deepEqual(await m.calls['u-ben'].enter({ id: MID, key: Uint8Array.from({ length: 32 }, () => 9) }), { entered: true });
  await m.settle();
  m.calls['u-amira'].setHand(true);
  m.calls['u-amira'].say('Can you read this?');
  m.calls['u-ben'].say('Or this?');
  await m.settle();
  assert.equal(person(m.calls['u-ben'], 'u-amira').hand, false);
  assert.deepEqual(m.log.says, []);
  assert.equal(m.calls['u-ben'].unreadable, false);
  // The engine says so when sound or picture does not open.
  m.calls['u-ben'].room.emit(RoomEvent.EncryptionError, new Error('Decryption failed'));
  assert.equal(m.calls['u-ben'].unreadable, true);
  await closeAll(m);
});

test('someone who comes in later is told how everyone already is', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira', { muted: true });
  m.calls['u-amira'].setHand(true);
  await m.settle();
  await m.enter('u-ben');
  assert.deepEqual({ muted: person(m.calls['u-ben'], 'u-amira').muted, hand: person(m.calls['u-ben'], 'u-amira').hand }, { muted: true, hand: true });
  await closeAll(m);
});

test('what is written goes to everyone once, and what the server writes is read by nobody', async () => {
  const m = meetingOf(['u-amira', 'u-ben', 'u-cara']);
  for (const id of Object.keys(m.calls)) await m.enter(id);
  const sent = m.calls['u-ben'].say('Shall we start?');
  await m.settle();
  assert.equal(sent.to, 2);
  assert.deepEqual(m.log.says.map((s) => [s.at, s.from, s.text, s.id]).sort(), [['u-amira', 'u-ben', 'Shall we start?', sent.id], ['u-cara', 'u-ben', 'Shall we start?', sent.id]]);
  assert.equal(m.calls['u-ben'].say('   '), null);
  m.world.inject('u-amira', 'u-ben', { t: 'say', id: 's-forged', text: 'Send me the password' });
  // And what Ben wrote, passed on a second time, is not read twice.
  m.world.pass('u-amira', 'u-ben', m.world.carried.filter((c) => c.from === 'u-ben').at(-1).bytes);
  await m.settle();
  assert.equal(m.log.says.length, 2);
  await closeAll(m);
});

test('only whoever set the meeting up can mute someone, by their app and by the server', async () => {
  const m = meetingOf(['u-amira', 'u-ben', 'u-cara']);
  for (const id of Object.keys(m.calls)) await m.enter(id);
  // Asked by someone who may not: not honoured, and the server is not asked.
  assert.equal(m.calls['u-cara'].askMute('u-ben'), true);
  await m.settle();
  assert.equal(m.calls['u-ben'].muted, false);
  assert.equal(m.world.asked.filter((a) => a.do === 'serverControl').length, 0);
  // Asked by the person who set it up.
  assert.equal(m.calls['u-amira'].askMute('u-ben'), true);
  await m.settle();
  assert.equal(m.calls['u-ben'].muted, true);
  assert.equal(m.calls['u-ben'].mic.enabled, false);
  assert.deepEqual(m.log.muted, ['u-ben']);
  assert.equal(person(m.calls['u-cara'], 'u-ben').muted, true);
  const control = m.world.asked.filter((a) => a.do === 'serverControl');
  assert.equal(control.length, 1);
  assert.deepEqual({ who: control[0].who, what: control[0].what, person: control[0].person, session: control[0].session }, { who: 'u-amira', what: 'mute', person: 'u-ben', session: m.calls['u-amira'].session });
  assert.equal(m.calls['u-amira'].askMute('u-nobody'), false);
  await closeAll(m);
});

test('muted from the server\'s side, the person is told and shown as muted, and can speak again', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira');
  await m.enter('u-ben');
  m.world.serverMute('u-ben');
  await m.settle();
  assert.equal(m.calls['u-ben'].muted, true);
  assert.deepEqual(m.log.muted, ['u-ben']);
  assert.equal(person(m.calls['u-amira'], 'u-ben').muted, true);
  m.calls['u-ben'].setMuted(false);
  await m.settle();
  assert.equal(m.calls['u-ben'].sent.audio.pub.isMuted, false);
  assert.equal(m.calls['u-ben'].mic.enabled, true);
  assert.equal(person(m.calls['u-amira'], 'u-ben').muted, false);
  await closeAll(m);
});

test('on air or recorded is heeded only from whoever set the meeting up, and each is told whether they are in the picture', async () => {
  const m = meetingOf(['u-amira', 'u-ben', 'u-cara']);
  for (const id of Object.keys(m.calls)) await m.enter(id);
  m.calls['u-cara'].setAir({ on: true, how: 'live', shown: ['u-ben'] });
  await m.settle();
  assert.deepEqual(m.calls['u-ben'].air, { on: false, you: false, how: null });
  m.calls['u-amira'].setAir({ on: true, how: 'record', shown: ['u-ben'] });
  await m.settle();
  assert.deepEqual(m.calls['u-ben'].air, { on: true, you: true, how: 'record' });
  assert.deepEqual(m.calls['u-cara'].air, { on: true, you: false, how: 'record' });
  // Whoever said so leaves: it is not on air until someone who may says so again.
  m.calls['u-amira'].close();
  await m.settle();
  assert.deepEqual(m.calls['u-ben'].air, { on: false, you: false, how: null });
  await closeAll(m);
});

test('someone who leaves is gone from everyone\'s list; leaving gives the way in back and stops the worker', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira');
  await m.enter('u-ben');
  const ben = m.calls['u-ben'];
  const { mic } = ben;
  const session = ben.session;
  ben.close();
  await m.settle();
  assert.deepEqual(m.calls['u-amira'].people(), []);
  assert.equal(mic.stopped, true);
  assert.equal(m.world.workers[1].stopped, true);
  assert.deepEqual(m.world.asked.filter((a) => a.do === 'serverLeave' && a.who === 'u-ben').map((a) => a.session), [session]);
  assert.deepEqual(m.log.gone, [], 'leaving is not being put out');
  assert.notEqual(globalThis.WebSocket, nativeSocket, 'the other meeting is still open');
  m.calls['u-amira'].close();
  await m.settle();
  assert.equal(globalThis.WebSocket, nativeSocket);
});

test('put out, ended, or opened on another computer: the window is told which', async () => {
  for (const [reason, why] of [[DisconnectReason.PARTICIPANT_REMOVED, 'removed'], [DisconnectReason.ROOM_DELETED, 'ended'], [DisconnectReason.DUPLICATE_IDENTITY, 'elsewhere'], [DisconnectReason.SERVER_SHUTDOWN, 'lost'], [undefined, 'lost']]) {
    const m = meetingOf(['u-amira', 'u-ben']);
    await m.enter('u-amira');
    await m.enter('u-ben');
    m.world.putOut('u-ben', reason);
    assert.deepEqual(m.log.gone, [{ id: 'u-ben', why }]);
    await closeAll(m);
  }
  assert.equal(whyGone({ DisconnectReason }, DisconnectReason.CLIENT_INITIATED), 'left');
  assert.equal(globalThis.WebSocket, nativeSocket);
});

test('while the connection to the server is being made again nobody is shown as reached, and nothing is said into the gap', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira');
  await m.enter('u-ben');
  const ben = m.calls['u-ben'];
  ben.room.emit('reconnecting');
  assert.equal(person(ben, 'u-amira').connected, false);
  assert.equal(ben.say('Can you hear me?').to, 0);
  ben.setHand(true);
  ben.room.emit('reconnected');
  await m.settle();
  assert.equal(person(ben, 'u-amira').connected, true);
  assert.equal(person(m.calls['u-amira'], 'u-ben').hand, true, 'said again once the connection is back');
  await closeAll(m);
});

test('refused before the door: said as the service said it, with nothing left open', async () => {
  for (const refusal of [{ error: 'no-pass' }, { error: 'server-gone', state: 'updating' }, { error: 'pin' }, { error: 'unreachable' }, { error: 'offline' }]) {
    const m = meetingOf(['u-amira']);
    m.world.enterRefusal = refusal;
    assert.deepEqual(await m.enter('u-amira'), refusal);
    assert.equal(m.world.workers.length, 0);
    assert.equal(globalThis.WebSocket, nativeSocket);
    // The seat is still theirs: they can be let in by another way without the microphone being asked for again.
    assert.ok(m.calls['u-amira'].mic);
    assert.deepEqual(await m.enter('u-amira'), { entered: true });
    await closeAll(m);
  }
});

test('refused at the room: a full meeting, a pass no longer taken, a door that cannot be reached', async () => {
  const cases = [
    [{ status: 429, why: 'refused', reason: 'room has exceeded its max participants' }, 'full'],
    [{ status: 403, why: 'refused', reason: 'not-admitted' }, 'not-admitted'],
    [{ status: 401, why: 'refused', reason: 'invalid token' }, 'not-admitted'],
    [{ status: 503, why: 'refused', reason: 'not-serving' }, 'not-serving'],
    // What LiveKit 1.13 answers through the door when the meeting already holds as many as it may.
    [{ status: 500, why: 'refused', reason: 'connection closed by media' }, 'turned-away'],
    [{ why: 'unreachable' }, 'unreachable'],
    [{ why: 'pin' }, 'pin'],
    [{ why: 'certificate' }, 'certificate'],
  ];
  for (const [doorRefusal, why] of cases) {
    const m = meetingOf(['u-amira']);
    m.world.doorRefusal = doorRefusal;
    assert.deepEqual(await m.enter('u-amira'), { error: why });
    await m.settle();
    assert.equal(m.world.workers[0].stopped, true);
    assert.equal(m.world.asked.filter((a) => a.do === 'serverLeave').length, 1, 'the way in is given back');
    assert.equal(globalThis.WebSocket, nativeSocket);
    assert.equal(m.calls['u-amira'].room, null);
    assert.deepEqual(m.log.gone, []);
    await closeAll(m);
  }
});

test('what the engine says of a refusal is read when the door said nothing', () => {
  const lib = { ConnectionErrorReason };
  assert.equal(whyNotConnected(lib, { reason: ConnectionErrorReason.NotAllowed, message: 'no' }, null), 'not-admitted');
  assert.equal(whyNotConnected(lib, { reason: ConnectionErrorReason.ServerUnreachable, message: '' }, null), 'unreachable');
  assert.equal(whyNotConnected(lib, { reason: ConnectionErrorReason.Timeout, message: '' }, null), 'unreachable');
  assert.equal(whyNotConnected(lib, { reason: ConnectionErrorReason.Cancelled, message: '' }, null), 'left');
  assert.equal(whyNotConnected(lib, { reason: ConnectionErrorReason.LeaveRequest, message: 'room is full' }, null), 'full');
  assert.equal(whyNotConnected(lib, new Error('something else'), null), 'failed');
  assert.equal(whyNotConnected(null, null, null), 'failed');
});

test('someone who gives up while still coming in is never left in the room', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira');
  m.world.holdDoor = true;
  await m.calls['u-ben'].start();
  const coming = m.calls['u-ben'].enter({ id: MID, key: KEY });
  await new Promise((r) => setTimeout(r, 5));
  m.calls['u-ben'].close();
  m.world.holdDoor = false;
  m.world.letIn();
  assert.deepEqual(await coming, { error: 'left' });
  await m.settle();
  assert.deepEqual(m.calls['u-amira'].people(), []);
  assert.equal(m.world.inside.has('u-ben'), false);
  assert.equal(m.world.asked.filter((a) => a.do === 'serverLeave' && a.who === 'u-ben').length, 1);
  await closeAll(m);
  assert.equal(globalThis.WebSocket, nativeSocket);
});

test('a seat given its sound and picture sends them in place of a microphone and a camera', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  const audio = new m.rtc.Track('audio');
  const video = new m.rtc.Track('video');
  m.calls['u-amira'].provide({ audio, video });
  assert.deepEqual(await m.calls['u-amira'].enter({ id: MID, key: KEY }), { entered: true });
  await m.enter('u-ben');
  const amira = person(m.calls['u-ben'], 'u-amira');
  assert.deepEqual(amira.stream.getAudioTracks(), [audio]);
  assert.deepEqual(amira.stream.getVideoTracks(), [video]);
  assert.equal(amira.sharing, true);
  assert.equal(m.rtc.asked.filter((c) => c.audio).length, 1, 'only the other person was asked for a microphone');
  await closeAll(m);
});

test('the room left with the seat kept: the microphone is not stopped, and the same seat comes in again', async () => {
  const m = meetingOf(['u-amira', 'u-ben']);
  await m.enter('u-amira');
  await m.enter('u-ben');
  const ben = m.calls['u-ben'];
  const { mic } = ben;
  ben.release();
  for (const id of [...ben.peers.keys()]) ben.unlink(id);
  await m.settle();
  assert.deepEqual(m.calls['u-amira'].people(), []);
  assert.notEqual(mic.stopped, true, 'the microphone is the seat\'s, not the room\'s');
  assert.deepEqual(await ben.enter({ id: MID, key: KEY }), { entered: true });
  await m.settle();
  assert.deepEqual(person(m.calls['u-amira'], 'u-ben').stream.getAudioTracks(), [mic]);
  // Counting starts again with the new stay, and is taken: what Ben says now is read.
  ben.setHand(true);
  await m.settle();
  assert.equal(person(m.calls['u-amira'], 'u-ben').hand, true);
  await closeAll(m);
});
