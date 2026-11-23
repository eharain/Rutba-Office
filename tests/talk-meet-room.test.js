// One small meeting as the window in it holds it (Office 2.27,
// renderer/meeting/meet-room.js): the introducer lets this person in and
// passes each pair's set-up; who comes and goes, the lock, being taken out
// and the end come from it; what is written stays between the computers.
// Run on a stand-in for Talk's service and for WebRTC.

import test from 'node:test';
import assert from 'node:assert/strict';
import { MeetingRoom } from '../apps/desktop/renderer/meeting/meet-room.js';
import { fakeRtc } from './talk-fake-rtc.js';

const ID = 'meet-AAAAAAAAAAAAAAAAAAAA';
const AMIRA = { id: 'u-a', name: 'Amira Khan' };
const BEN = { id: 'u-ben', name: 'Ben Ortiz', contact: true };
const CARA = { id: 'u-cara', name: 'Cara Lund', contact: true };
const meetingOf = (more = {}) => ({ id: ID, name: 'Monday plan', starter: { id: 'u-ben', name: 'Ben Ortiz' }, at: null, live: false, locked: false, groupId: null, link: `https://introducer.example/m/${ID}`, most: 4, people: [], ...more });

/** Talk's service as a window sees it, answering as `answers` says; every call is noted. */
function fakeShell(answers = {}) {
  const calls = [];
  const talk = {};
  for (const name of ['meetJoin', 'meetLeave', 'meetLock', 'meetRemove', 'meetEnd', 'send', 'callSetup']) {
    talk[name] = async (args) => {
      calls.push({ name, args });
      if (answers[name]) return answers[name](args);
      if (name === 'callSetup') return { iceServers: [{ urls: ['stun:stun.example'] }] };
      if (name === 'meetJoin') return { meeting: meetingOf() };
      return {};
    };
  }
  return { shell: { talk }, calls, of: (name) => calls.filter((c) => c.name === name).map((c) => c.args) };
}
function roomOf({ answers = {}, me = AMIRA, rtc = fakeRtc() } = {}) {
  const { shell, calls, of } = fakeShell(answers);
  const seen = { changes: 0, muted: 0 };
  const room = new MeetingRoom({ shell, me, rtc, onChange: () => { seen.changes += 1; }, onMuted: () => { seen.muted += 1; } });
  const settle = async () => { for (let i = 0; i < 6; i++) { await room.mesh?.line; await new Promise((r) => setImmediate(r)); } };
  return { room, shell, calls, of, seen, rtc, settle };
}
/** The other end of a connection this room made, answered by hand. */
function answerFor(rtc, room, from) {
  const other = new rtc.PeerConnection({});
  other.createDataChannel('x');
  room.handle({ type: 'meet.signal', id: ID, from, data: { kind: 'answer', sdp: `answer:${other.n}` } });
  return other;
}

test('coming in: the microphone first, then the introducer, then an offer to each person already there', async () => {
  const { room, calls, of, rtc, settle } = roomOf({ answers: { meetJoin: () => ({ meeting: meetingOf({ people: [BEN, CARA] }) }) } });
  const r = await room.join(ID);
  await settle();
  assert.equal(r.meeting.id, ID);
  assert.equal(room.state, 'in');
  assert.deepEqual(rtc.asked[0], { audio: { echoCancellation: true, noiseSuppression: true }, video: false });
  assert.ok(calls.findIndex((c) => c.name === 'meetJoin') >= 0);
  assert.deepEqual(of('meetJoin'), [{ id: ID }]);
  const offers = of('send');
  assert.deepEqual(offers.map((m) => [m.type, m.id, m.to, m.data.kind]), [['meet.signal', ID, 'u-ben', 'offer'], ['meet.signal', ID, 'u-cara', 'offer']]);

  const people = room.people();
  assert.deepEqual(people.map((p) => [p.id, p.name, p.me, p.starter]), [['u-a', 'Amira Khan', true, false], ['u-ben', 'Ben Ortiz', false, true], ['u-cara', 'Cara Lund', false, false]]);
  assert.equal(people[0].connected, true);
  assert.equal(people[1].connected, false, 'not until their answer');
  answerFor(rtc, room, 'u-ben');
  await settle();
  assert.equal(room.people()[1].connected, true);
  assert.equal(room.mine, false);
});

test('a refusal leaves nothing open: no connection, and the microphone let go', async () => {
  const { room, rtc } = roomOf({ answers: { meetJoin: () => ({ error: 'full', most: 4 }) } });
  const r = await room.join(ID);
  assert.deepEqual(r, { error: 'full', most: 4 });
  assert.equal(room.state, 'gone');
  assert.equal(room.why, 'refused');
  assert.equal(room.mesh, null);
  assert.equal(rtc.pcs.length, 0);
  assert.deepEqual(room.people(), []);
});

test('a service that cannot be reached is a refusal, not a crash', async () => {
  const { room } = roomOf({ answers: { meetJoin: () => { throw new Error('gone'); } } });
  assert.deepEqual(await room.join(ID), { error: 'offline' });
  assert.equal(room.state, 'gone');
});

test('someone who comes in after is waited for, answered when they offer, and let go when they leave', async () => {
  const { room, of, rtc, settle } = roomOf();
  await room.join(ID);
  room.handle({ type: 'meet.person', id: ID, event: 'joined', person: CARA });
  assert.deepEqual(room.people().map((p) => [p.name, p.connected]), [['Amira Khan', true], ['Cara Lund', false]]);

  const theirs = new rtc.PeerConnection({});
  theirs.createDataChannel('call');
  room.handle({ type: 'meet.signal', id: ID, from: 'u-cara', data: { kind: 'offer', sdp: `offer:${theirs.n}` } });
  await settle();
  assert.deepEqual(of('send').map((m) => [m.to, m.data.kind]), [['u-cara', 'answer']]);
  assert.equal(room.people()[1].name, 'Cara Lund');

  room.handle({ type: 'meet.person', id: ID, event: 'left', person: CARA });
  await settle();
  assert.deepEqual(room.people().map((p) => p.id), ['u-a']);
  // About this person themself, or about nobody: nothing.
  room.handle({ type: 'meet.person', id: ID, event: 'joined', person: AMIRA });
  room.handle({ type: 'meet.person', id: ID, event: 'joined' });
  assert.equal(room.people().length, 1);
});

test('what the introducer says while this person is still coming in is kept for the moment they are in', async () => {
  let let_in;
  const { room, settle } = roomOf({ answers: { meetJoin: () => new Promise((resolve) => { let_in = resolve; }) } });
  const joining = room.join(ID);
  for (let i = 0; i < 20 && !let_in; i++) await new Promise((r) => setImmediate(r));
  room.handle({ type: 'meet.person', id: ID, event: 'joined', person: CARA });
  room.handle({ type: 'meet.person', id: 'meet-ZZZZZZZZZZZZZZZZZZZZ', event: 'joined', person: BEN });
  assert.deepEqual(room.people(), [], 'not in yet');
  let_in({ meeting: meetingOf() });
  await joining;
  await settle();
  assert.deepEqual(room.people().map((p) => p.id), ['u-a', 'u-cara'], 'and nothing of another meeting');
});

test('another meeting is not this one: its people, set-ups and end change nothing here', async () => {
  const { room, of } = roomOf();
  await room.join(ID);
  const other = 'meet-ZZZZZZZZZZZZZZZZZZZZ';
  room.handle({ type: 'meet.person', id: other, event: 'joined', person: CARA });
  room.handle({ type: 'meet.signal', id: other, from: 'u-cara', data: { kind: 'offer', sdp: 'offer:1' } });
  room.handle({ type: 'meet.changed', meeting: { id: other, locked: true } });
  room.handle({ type: 'meet.ended', id: other });
  room.handle({ type: 'meet.removed', id: other, why: 'removed' });
  room.handle(null);
  assert.equal(room.state, 'in');
  assert.equal(room.meeting.locked, false);
  assert.equal(room.people().length, 1);
  assert.equal(of('send').length, 0);
});

test('locked, taken out, joined on another computer, ended: each says what happened', async () => {
  const start = async () => { const x = roomOf(); await x.room.join(ID); return x; };

  let x = await start();
  x.room.handle({ type: 'meet.changed', meeting: { id: ID, locked: true } });
  assert.equal(x.room.meeting.locked, true);
  assert.equal(x.room.meeting.name, 'Monday plan', 'the rest of the meeting stays');

  x.room.handle({ type: 'meet.removed', id: ID, why: 'removed', by: { id: 'u-ben', name: 'Ben Ortiz' } });
  assert.deepEqual([x.room.state, x.room.why, x.room.by.name], ['gone', 'removed', 'Ben Ortiz']);
  assert.equal(x.room.mesh.closed, true);
  assert.equal(x.of('meetLeave').length, 0, 'the introducer already has them out');
  x.room.handle({ type: 'meet.ended', id: ID });
  assert.equal(x.room.why, 'removed', 'gone once');

  x = await start();
  x.room.handle({ type: 'meet.removed', id: ID, why: 'elsewhere' });
  assert.deepEqual([x.room.state, x.room.why], ['gone', 'elsewhere']);

  x = await start();
  x.room.handle({ type: 'meet.ended', id: ID, by: { id: 'u-ben', name: 'Ben Ortiz' } });
  assert.deepEqual([x.room.state, x.room.why, x.room.by.name], ['gone', 'ended', 'Ben Ortiz']);
});

test('leaving tells the introducer once and lets go of the microphone', async () => {
  const { room, of } = roomOf();
  await room.join(ID);
  const mic = room.mesh.mic;
  room.leave();
  room.leave();
  assert.deepEqual(of('meetLeave'), [{ id: ID }]);
  assert.deepEqual([room.state, room.why], ['gone', 'left']);
  assert.equal(mic.stopped, true);
});

test('what is written is kept only here: mine, what is heard once, and how much is not yet seen', async () => {
  const { room, rtc, settle } = roomOf({ answers: { meetJoin: () => ({ meeting: meetingOf({ people: [BEN] }) }) } });
  await room.join(ID);
  await settle();
  answerFor(rtc, room, 'u-ben');
  await settle();

  const sent = room.say('Shall we start?');
  assert.equal(sent.to, 1);
  assert.deepEqual(JSON.parse(room.mesh.peers.get('u-ben').channel.sent.at(-1)), { t: 'say', id: sent.id, text: 'Shall we start?' });
  assert.equal(room.say('  '), null);

  const line = room.mesh.peers.get('u-ben').channel;
  line.onmessage({ data: JSON.stringify({ t: 'say', id: 's-1', text: 'Yes, go on.' }) });
  line.onmessage({ data: JSON.stringify({ t: 'say', id: 's-1', text: 'Yes, go on.' }) });
  assert.deepEqual(room.says.map((s) => [s.name, s.text, s.mine]), [['Amira Khan', 'Shall we start?', true], ['Ben Ortiz', 'Yes, go on.', false]]);
  assert.equal(room.unseen, 1);
  room.seen();
  assert.equal(room.unseen, 0);
});

test('only the person who set the meeting up is obeyed when they mute someone', async () => {
  // Ben set it up; Cara is in it too.
  const { room, rtc, settle, seen } = roomOf({ answers: { meetJoin: () => ({ meeting: meetingOf({ people: [BEN, CARA] }) }) } });
  await room.join(ID);
  await settle();
  answerFor(rtc, room, 'u-ben');
  answerFor(rtc, room, 'u-cara');
  await settle();
  room.mesh.peers.get('u-cara').channel.onmessage({ data: '{"t":"mute"}' });
  assert.equal(room.mesh.muted, false, 'not from Cara');
  assert.equal(seen.muted, 0);
  room.mesh.peers.get('u-ben').channel.onmessage({ data: '{"t":"mute"}' });
  assert.equal(room.mesh.muted, true);
  assert.equal(seen.muted, 1);
  assert.equal(room.mute('u-cara'), false, 'and this person cannot mute anyone: it is not their meeting');
});

test('the person who set the meeting up can mute, take out, lock and end', async () => {
  const mine = meetingOf({ starter: { id: 'u-a', name: 'Amira Khan' }, people: [BEN] });
  const { room, of, rtc, settle } = roomOf({ answers: {
    meetJoin: () => ({ meeting: mine }),
    meetLock: ({ locked }) => ({ meeting: { ...mine, locked } }),
    meetRemove: ({ who }) => (who === 'u-ben' ? { meeting: mine } : { error: 'nobody' }),
    meetEnd: () => ({ ended: true }),
  } });
  await room.join(ID);
  await settle();
  answerFor(rtc, room, 'u-ben');
  await settle();
  assert.equal(room.mine, true);
  assert.equal(room.people()[0].starter, true);

  assert.equal(room.mute('u-ben'), true);
  assert.equal(room.mesh.peers.get('u-ben').channel.sent.at(-1), '{"t":"mute"}');

  assert.equal((await room.lock(true)).meeting.locked, true);
  assert.equal(room.meeting.locked, true);
  assert.deepEqual(of('meetLock'), [{ id: ID, locked: true }]);

  assert.ok((await room.remove('u-ben')).meeting);
  assert.deepEqual(await room.remove('u-zed'), { error: 'nobody' });
  assert.deepEqual(of('meetRemove'), [{ id: ID, who: 'u-ben' }, { id: ID, who: 'u-zed' }]);

  assert.deepEqual(await room.end(), { ended: true });
  assert.deepEqual([room.state, room.why], ['gone', 'ended']);
  assert.deepEqual(of('meetEnd'), [{ id: ID }]);
});

test('a refusal to lock or end changes nothing here', async () => {
  const { room } = roomOf({ answers: { meetLock: () => ({ error: 'not-yours' }), meetEnd: () => ({ error: 'not-yours' }) } });
  await room.join(ID);
  assert.deepEqual(await room.lock(true), { error: 'not-yours' });
  assert.equal(room.meeting.locked, false);
  assert.deepEqual(await room.end(), { error: 'not-yours' });
  assert.equal(room.state, 'in');
});

test('a connection to the introducer made again brings this person back into the meeting, as they were', async () => {
  const { room, of, settle } = roomOf();
  await room.join(ID, { video: true });
  room.mesh.setMuted(true);
  room.heard('u-ben', 'before the break', 's-1');
  const first = room.mesh;

  room.handle({ type: 'welcome' });
  for (let i = 0; i < 20 && of('meetJoin').length < 2; i++) await new Promise((r) => setImmediate(r));
  await settle();
  assert.equal(of('meetJoin').length, 2);
  assert.equal(first.closed, true);
  assert.notEqual(room.mesh, first);
  assert.equal(room.state, 'in');
  assert.equal(room.mesh.muted, true, 'still muted');
  assert.ok(room.mesh.camera, 'still with the camera');
  assert.equal(room.says.length, 1, 'what was written is still there');
});

test('a meeting that is gone when the connection comes back is left, and says it was lost', async () => {
  let n = 0;
  const { room, of } = roomOf({ answers: { meetJoin: () => (++n === 1 ? { meeting: meetingOf() } : { error: 'no-meeting' }) } });
  await room.join(ID);
  // What the window is told the moment the room is gone is already the reason.
  const whys = [];
  room.onChange = (x) => { if (x.state === 'gone') whys.push(x.why); };
  room.handle({ type: 'welcome' });
  for (let i = 0; i < 50 && room.state !== 'gone'; i++) await new Promise((r) => setImmediate(r));
  assert.equal(of('meetJoin').length, 2);
  assert.deepEqual([room.state, room.why], ['gone', 'lost']);
  assert.deepEqual([...new Set(whys)], ['lost']);
});

/* ── giving up while still coming in, and a failed connection tried again (2.27.0) ── */

const until = async (ready) => { for (let i = 0; i < 50 && !ready(); i++) await new Promise((r) => setImmediate(r)); assert.ok(ready(), 'reached'); };

test('giving up while the introducer is asked: let in all the same, so taken out again, and nothing left open', async () => {
  let let_in = null;
  const { room, of, rtc, settle } = roomOf({ answers: { meetJoin: () => new Promise((resolve) => { let_in = resolve; }) } });
  const joining = room.join(ID);
  await until(() => let_in);
  const mic = room.mesh.mic;
  room.leave();
  assert.deepEqual(of('meetLeave'), [], 'not in yet, so nothing to leave yet');
  let_in({ meeting: meetingOf({ people: [BEN] }) });
  assert.deepEqual(await joining, { error: 'left' });
  await settle();
  assert.deepEqual([room.state, room.why], ['gone', 'left']);
  assert.equal(room.mesh, null);
  assert.equal(mic.stopped, true);
  assert.deepEqual(of('meetLeave'), [{ id: ID }], 'never left in the meeting unseen and unheard');
  assert.deepEqual(of('send'), [], 'and nobody is offered a connection');
  assert.equal(rtc.pcs.length, 0);
  assert.deepEqual(room.people(), []);
});

test('giving up, then refused anyway: nothing to leave', async () => {
  let let_in = null;
  const { room, of } = roomOf({ answers: { meetJoin: () => new Promise((resolve) => { let_in = resolve; }) } });
  const joining = room.join(ID);
  await until(() => let_in);
  room.leave();
  let_in({ error: 'full', most: 4 });
  assert.deepEqual(await joining, { error: 'left' });
  assert.deepEqual(of('meetLeave'), []);
  assert.equal(room.why, 'left', 'what the person did, not a refusal they never saw');
});

test('giving up while the microphone is asked for: the introducer is never asked, and the microphone is let go', async () => {
  const rtc = fakeRtc();
  const real = rtc.getUserMedia;
  let open = null;
  let mic = null;
  rtc.getUserMedia = (c) => new Promise((resolve) => { open = async () => { const s = await real(c); mic = s.getAudioTracks()[0]; resolve(s); }; });
  const { room, of } = roomOf({ rtc });
  const joining = room.join(ID);
  await until(() => open);
  room.leave();
  await open();
  assert.deepEqual(await joining, { error: 'left' });
  assert.deepEqual(of('meetJoin'), []);
  assert.deepEqual(of('meetLeave'), []);
  assert.equal(mic.stopped, true);
  assert.equal(room.mesh, null);
});

test('the meeting is named while this person comes in, and giving up while it is looked at opens no microphone', async () => {
  let let_in = null;
  const { room, shell } = roomOf({ answers: { meetJoin: () => new Promise((resolve) => { let_in = resolve; }) } });
  shell.talk.meetPeek = async () => ({ meeting: meetingOf({ people: [BEN] }) });
  assert.equal(room.wanted, '');
  const joining = room.join(ID);
  await until(() => let_in);
  assert.equal(room.wanted, 'Monday plan');
  let_in({ meeting: meetingOf() });
  assert.equal((await joining).meeting.id, ID);

  let seen = null;
  const other = roomOf();
  other.shell.talk.meetPeek = () => new Promise((resolve) => { seen = resolve; });
  const coming = other.room.join(ID);
  await until(() => seen);
  other.room.leave();
  seen({ meeting: meetingOf({ people: [BEN] }) });
  assert.deepEqual(await coming, { error: 'left' });
  assert.equal(other.rtc.asked.length, 0, 'no microphone asked for');
  assert.deepEqual(other.of('meetJoin'), []);
  assert.equal(other.room.mesh, null);
});

test('a connection to one person that failed is tried again from the room', async () => {
  const { room, of, settle } = roomOf({ answers: { meetJoin: () => ({ meeting: meetingOf({ people: [BEN] }) }) } });
  assert.equal(room.redial('u-ben'), false, 'not in a meeting yet');
  await room.join(ID);
  await settle();
  assert.equal(of('send').length, 1);
  room.mesh.peers.get('u-ben').pc.fail();
  assert.deepEqual(room.people().map((p) => [p.id, Boolean(p.failed)]), [['u-a', false], ['u-ben', true]]);
  assert.equal(room.redial('u-ben'), true);
  await settle();
  // "u-a" sorts before "u-ben": this computer offers anew.
  assert.deepEqual(of('send').map((m) => [m.to, m.data.kind]), [['u-ben', 'offer'], ['u-ben', 'offer']]);
  assert.equal(room.people()[1].failed, false);
  assert.equal(room.redial('u-nobody'), false);
});
