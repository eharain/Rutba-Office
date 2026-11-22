// The guests of a Live Broadcast or Recorder window
// (apps/desktop/renderer/studio/guest-session.js, Office 2.27.0): the small
// Talk meeting the window holds. What it holds to: the meeting is set up to
// be streamed live or to be recorded, and no other is ever held; this
// computer's seat is the programme, with no microphone or camera asked for;
// someone who comes in is never put on the frame by it; and everyone is told
// when it is on air, each whether they are in the picture. Run on stand-ins
// for Talk's service, for WebRTC and for the window.

import test from 'node:test';
import assert from 'node:assert/strict';
import { GuestSession } from '../apps/desktop/renderer/studio/guest-session.js';
import { emptyStack, normaliseStack, addItem, applyScene } from '../packages/studio/src/stack.js';
import { GUEST, SEATS, GUEST_SCENES, withGuests, seatsShown } from '../packages/studio/src/guests.js';
import { fakeRtc } from './talk-fake-rtc.js';

const ID = 'live-AAAAAAAAAAAAAAAAAAAA';
const AMIRA = { id: 'u-a', name: 'Amira Khan' };
const BEN = { id: 'u-ben', name: 'Ben Ortiz' };
const CARA = { id: 'u-cara', name: 'Cara Lund' };
const DAN = { id: 'u-dan', name: 'Dan Moyo' };
const EVE = { id: 'u-eve', name: 'Eve Park' };
const ONLINE = { phase: 'online', me: AMIRA, contacts: [BEN, CARA, DAN, EVE] };
const meetingOf = (more = {}) => ({ id: ID, name: 'Live broadcast with Amira', starter: AMIRA, at: null, live: true, recorded: false, locked: false, link: `https://introducer.example/m/${ID}`, most: 4, people: [], ...more });

function base() {
  let s = emptyStack();
  ({ stack: s } = addItem(s, { kind: 'screen', ref: 'screen:1', name: 'Screen 1' }));
  ({ stack: s } = addItem(s, { kind: 'camera', ref: 'cam-a', name: 'USB camera' }));
  return s;
}
const scene = (s, n) => applyScene(s, s.scenes.find((sc) => sc.key === GUEST_SCENES[n - 1]));

/** Talk's service as a window sees it, answering as `answers` says; every call is noted. */
function fakeShell(answers = {}) {
  const calls = [];
  const talk = {};
  for (const name of ['meetCreate', 'meetJoin', 'meetLeave', 'meetInvite', 'meetRemove', 'meetEnd', 'send', 'callSetup']) {
    talk[name] = async (args) => {
      calls.push({ name, args });
      if (answers[name]) return answers[name](args);
      if (name === 'callSetup') return { iceServers: [] };
      if (name === 'meetCreate') return { meeting: meetingOf({ live: Boolean(args.live), recorded: Boolean(args.recorded) }) };
      if (name === 'meetJoin') return { meeting: meetingOf() };
      if (name === 'meetInvite') return { told: args.invite.length };
      return {};
    };
  }
  return { shell: { talk }, calls, of: (name) => calls.filter((c) => c.name === name).map((c) => c.args) };
}

/**
 * A studio window standing in, with its guest session. A change the session
 * asks for is taken at once; the frames come back to the session at once
 * too, or, with `win.late`, only when the test says `push()`, as a window
 * that has not drawn again yet.
 */
function studioOf({ how = 'live', answers = {}, stack = base(), talk = ONLINE } = {}) {
  const rtc = fakeRtc();
  const { shell, calls, of } = fakeShell(answers);
  const win = { stack, air: null, onAir: false, late: false, changes: 0 };
  const said = [];
  const guests = { seats: {}, people: [], synced: 0, set(seats, people) { this.seats = seats; this.people = people; }, sync() { this.synced += 1; } };
  const fed = [];
  let session = null;
  const push = () => session.sync({ stack: win.stack, airStack: win.air || win.stack, onAir: win.onAir });
  session = new GuestSession({
    shell,
    how,
    guests,
    rtc,
    feed: async () => { const f = { audio: new rtc.Track('audio'), video: new rtc.Track('video') }; fed.push(f); return f; },
    change: (fn) => {
      win.stack = normaliseStack(fn(win.stack));
      if (win.air) win.air = normaliseStack(fn(win.air));
      win.changes += 1;
      if (!win.late) push();
    },
    say: (what, d) => said.push([what, what === 'came' ? d.name : d.why]),
  });
  session.setTalk(talk);
  push();
  const settle = async () => { for (let i = 0; i < 6; i++) { await session.room?.mesh?.line; await new Promise((r) => setImmediate(r)); } };
  /** Someone comes into the meeting and connects; answers what their computer hears from this one. */
  const comes = async (person) => {
    session.handle({ type: 'meet.person', id: ID, event: 'joined', person });
    const theirs = new rtc.PeerConnection({});
    const line = theirs.createDataChannel('call');
    const heard = [];
    line.onmessage = (e) => heard.push(JSON.parse(e.data));
    session.handle({ type: 'meet.signal', id: ID, from: person.id, data: { kind: 'offer', sdp: `offer:${theirs.n}` } });
    await settle();
    const answer = of('send').filter((m) => m.to === person.id && m.data.kind === 'answer').pop();
    await theirs.setRemoteDescription({ type: 'answer', sdp: answer.data.sdp });
    await settle();
    return { heard, air: () => heard.filter((m) => m.t === 'air').pop() || null, airs: () => heard.filter((m) => m.t === 'air').length };
  };
  const goes = async (person) => { session.handle({ type: 'meet.person', id: ID, event: 'left', person }); await settle(); };
  return { session, shell, calls, of, win, said, guests, fed, rtc, push, settle, comes, goes };
}

test('asking contacts in sets a meeting up for what the window does and holds it, the programme as this seat', async () => {
  const live = studioOf();
  assert.deepEqual(await live.session.invite(['u-ben', 'u-ben', 'u-cara'], { name: 'Live broadcast with Amira' }), { ok: true, meeting: meetingOf() });
  assert.deepEqual(live.of('meetCreate'), [{ name: 'Live broadcast with Amira', live: true, recorded: false, invite: ['u-ben', 'u-cara'], now: true }]);
  assert.deepEqual(live.of('meetJoin'), [{ id: ID }]);
  assert.equal(live.rtc.asked.length, 0, 'no microphone and no camera is asked for');
  assert.equal(live.fed.length, 1, 'the programme is this seat');
  assert.equal(live.session.room.state, 'in');
  // The three seats and their scenes are there, and nobody is on the frame.
  assert.deepEqual(live.win.stack.items.filter((it) => it.kind === GUEST).map((it) => it.ref).sort(), [...SEATS]);
  assert.deepEqual(live.win.stack.scenes.map((sc) => sc.key).filter((k) => GUEST_SCENES.includes(k)), [...GUEST_SCENES]);
  assert.deepEqual(seatsShown(live.win.stack), []);
  // More are asked into the meeting there is.
  assert.deepEqual(await live.session.invite(['u-dan']), { ok: true, told: 1 });
  assert.deepEqual(live.of('meetInvite'), [{ id: ID, invite: ['u-dan'] }]);
  assert.equal(live.of('meetCreate').length, 1);

  const rec = studioOf({ how: 'record', answers: { meetJoin: () => ({ meeting: meetingOf({ live: false, recorded: true }) }) } });
  assert.equal((await rec.session.invite(['u-ben'], { name: 'Recording with Amira' })).ok, true);
  assert.deepEqual(rec.of('meetCreate'), [{ name: 'Recording with Amira', live: false, recorded: true, invite: ['u-ben'], now: true }]);

  assert.deepEqual(await live.session.invite([]), { error: 'nobody' });
});

test('a meeting not set up for it, or someone else\'s, is never held, and nothing is asked of Talk while it is not connected', async () => {
  const plain = studioOf({ answers: { meetJoin: () => ({ meeting: meetingOf({ live: false }) }) } });
  assert.deepEqual(await plain.session.hold(ID), { error: 'not-live' });
  assert.deepEqual(plain.of('meetLeave'), [{ id: ID }]);
  assert.equal(plain.session.room, null);
  assert.deepEqual(plain.win.stack, base(), 'the setup is left as it was');

  const rec = studioOf({ how: 'record' });
  assert.deepEqual(await rec.session.hold(ID), { error: 'not-recorded' }, 'a meeting to be streamed is not one to be recorded');
  assert.equal(rec.session.room, null);

  const theirs = studioOf({ answers: { meetJoin: () => ({ meeting: meetingOf({ starter: BEN }) }) } });
  assert.deepEqual(await theirs.session.hold(ID), { error: 'not-yours' });
  assert.deepEqual(theirs.of('meetLeave'), [{ id: ID }]);

  const refused = studioOf({ answers: { meetJoin: () => ({ error: 'no-meeting' }) } });
  assert.deepEqual(await refused.session.hold(ID), { error: 'no-meeting' });
  assert.equal(refused.session.room, null);
  assert.deepEqual(refused.said, [], 'the refusal is what is said, once, by whoever asked');

  const off = studioOf({ talk: { phase: 'connecting', me: AMIRA, contacts: [] } });
  assert.deepEqual(await off.session.invite(['u-ben']), { error: 'offline' });
  assert.deepEqual(await off.session.hold(ID), { error: 'offline' });
  assert.equal(off.calls.length, 0);
});

test('someone who comes in takes a seat off air, and everyone is told when it is on air and who is in the picture', async () => {
  const { session, win, said, guests, push, comes } = studioOf();
  await session.hold(ID);
  const ben = await comes(BEN);
  assert.deepEqual(session.seats, { 'seat-1': BEN });
  assert.deepEqual(guests.seats, { 'seat-1': BEN });
  assert.deepEqual(said, [['came', 'Ben Ortiz']]);
  assert.equal(session.where('u-ben'), 'off-air');
  assert.deepEqual(ben.air(), { t: 'air', on: false, you: false, how: 'live' });

  // On air with Ben not in the scene: he is told it is on air, and that he is not in it.
  win.onAir = true;
  push();
  assert.deepEqual(ben.air(), { t: 'air', on: true, you: false, how: 'live' });
  assert.equal(session.where('u-ben'), 'off-air');
  const cara = await comes(CARA);
  assert.deepEqual(session.seats, { 'seat-1': BEN, 'seat-2': CARA });
  assert.deepEqual(cara.air(), { t: 'air', on: true, you: false, how: 'live' }, 'told as she connects');

  // The scene for one guest shows the first to come.
  win.stack = scene(win.stack, 1);
  push();
  assert.deepEqual(ben.air(), { t: 'air', on: true, you: true, how: 'live' });
  assert.deepEqual(cara.air(), { t: 'air', on: true, you: false, how: 'live' });
  assert.deepEqual([session.where('u-ben'), session.where('u-cara')], ['on-air', 'off-air']);
  assert.deepEqual(session.shown(), ['u-ben']);

  // Said once per change, not at every look.
  const before = ben.airs();
  push();
  push();
  assert.equal(ben.airs(), before);

  // What is held on air is what counts, not what is being arranged.
  win.air = win.stack;
  win.stack = scene(win.stack, 2);
  push();
  assert.deepEqual(cara.air().you, false, 'Cara is in the arrangement, not yet on air');
  win.air = null;
  push();
  assert.deepEqual(cara.air().you, true);

  win.onAir = false;
  push();
  assert.deepEqual(ben.air(), { t: 'air', on: false, you: false, how: 'live' });
  assert.equal(session.where('u-ben'), 'off-air');
});

test('a recording says it is a recording', async () => {
  const { session, win, push, comes } = studioOf({ how: 'record', answers: { meetJoin: () => ({ meeting: meetingOf({ live: false, recorded: true }) }) } });
  await session.hold(ID);
  const ben = await comes(BEN);
  win.stack = scene(win.stack, 1);
  win.onAir = true;
  push();
  assert.deepEqual(ben.air(), { t: 'air', on: true, you: true, how: 'record' });
});

test('a seat that is on the frame is taken off it before anyone sits in it, however often the meeting changes meanwhile', async () => {
  // The scene for three guests is up and on air, with nobody in it yet.
  const ready = scene(withGuests(base()), 3);
  const { session, win, guests, push, comes } = studioOf({ stack: ready });
  win.air = ready;
  win.onAir = true;
  push();
  await session.hold(ID);
  win.late = true;
  win.changes = 0;
  const ben = await comes(BEN);
  // Ben is in the meeting and connected, but the window has not drawn again: he sits nowhere yet.
  assert.deepEqual(session.seats, {});
  assert.deepEqual(guests.seats, {});
  assert.equal(win.changes, 1, 'the seat was taken off the frame once');
  assert.deepEqual(seatsShown(win.stack), ['seat-2', 'seat-3']);
  assert.deepEqual(seatsShown(win.air), ['seat-2', 'seat-3'], 'and off what is held on air');
  session.step();
  session.step();
  assert.deepEqual(guests.seats, {}, 'still nowhere');
  assert.equal(win.changes, 1);

  push();
  assert.deepEqual(session.seats, { 'seat-1': BEN });
  assert.deepEqual(session.shown(), [], 'seated, and not in the picture');
  assert.equal(session.where('u-ben'), 'off-air');
  assert.deepEqual(ben.air(), { t: 'air', on: true, you: false, how: 'live' });
});

test('three seats: a fourth person waits unseated, and takes the seat someone leaves', async () => {
  const { session, said, comes, goes } = studioOf();
  await session.hold(ID);
  await comes(BEN);
  await comes(CARA);
  await comes(DAN);
  await comes(EVE);
  assert.deepEqual(session.seats, { 'seat-1': BEN, 'seat-2': CARA, 'seat-3': DAN });
  assert.equal(session.where('u-eve'), 'unseated');
  assert.equal(session.people().length, 4);
  await goes(CARA);
  assert.deepEqual(session.seats, { 'seat-1': BEN, 'seat-2': EVE, 'seat-3': DAN }, 'the others keep their seats');
  assert.deepEqual(said.map((s) => s[1]), ['Ben Ortiz', 'Cara Lund', 'Dan Moyo', 'Eve Park']);
});

test('the person who holds it mutes, takes out and ends; leaving lets the meeting go on, and its code goes with it to Talk', async () => {
  const a = studioOf();
  await a.session.hold(ID);
  const ben = await a.comes(BEN);
  assert.equal(a.session.mute('u-ben'), true);
  assert.ok(ben.heard.some((m) => m.t === 'mute'));
  await a.session.remove('u-ben');
  assert.deepEqual(a.of('meetRemove'), [{ id: ID, who: 'u-ben' }]);
  assert.deepEqual(await a.session.end(), {});
  assert.deepEqual(a.of('meetEnd'), [{ id: ID }]);
  assert.equal(a.session.room, null);
  assert.deepEqual(a.session.seats, {});
  assert.deepEqual(a.guests.seats, {});
  assert.deepEqual(a.said, [['came', 'Ben Ortiz']], 'ending it oneself needs no telling');
  assert.deepEqual(await a.session.end(), { error: 'no-meeting' });

  const b = studioOf();
  await b.session.hold(ID);
  await b.comes(BEN);
  assert.equal(b.session.leave(), ID);
  assert.deepEqual(b.of('meetLeave'), [{ id: ID }]);
  assert.equal(b.of('meetEnd').length, 0, 'the meeting goes on for the others');
  assert.equal(b.session.room, null);
  assert.deepEqual(b.guests.seats, {});
  assert.equal(b.session.leave(), null);

  // Ended from elsewhere: the person is told.
  const c = studioOf();
  await c.session.hold(ID);
  await c.comes(BEN);
  c.session.handle({ type: 'meet.ended', id: ID, by: AMIRA });
  assert.equal(c.session.room, null);
  assert.deepEqual(c.said, [['came', 'Ben Ortiz'], ['gone', 'ended']]);
  assert.deepEqual(c.guests.seats, {});
  // What is said of some other meeting is nothing to this window.
  c.session.handle({ type: 'meet.ended', id: 'meet-ZZZZZZZZZZZZZZZZZZZZ' });
  c.session.handle({ type: 'conv.message', id: ID });
});

test('the connection made again keeps everyone in their seat, and they are told again what is on air', async () => {
  let people = [];
  const { session, win, said, push, comes, settle, rtc, of } = studioOf({ answers: { meetJoin: () => ({ meeting: meetingOf({ people }) }) } });
  await session.hold(ID);
  await comes(BEN);
  win.stack = scene(win.stack, 1);
  win.onAir = true;
  push();
  assert.deepEqual(session.shown(), ['u-ben']);

  people = [BEN];
  session.handle({ type: 'welcome' });
  await settle();
  assert.equal(session.room.state, 'in');
  assert.equal(of('meetJoin').length, 2);
  assert.deepEqual(session.seats, { 'seat-1': BEN });
  assert.deepEqual(seatsShown(win.stack), ['seat-1'], 'the scene is as it was');
  assert.deepEqual(said, [['came', 'Ben Ortiz']], 'nobody is announced twice');
  assert.equal(session.where('u-ben'), 'connecting');

  // Ben's computer answers the new offer: he hears again that he is on air.
  const theirs = new rtc.PeerConnection({});
  theirs.createDataChannel('x');
  session.handle({ type: 'meet.signal', id: ID, from: 'u-ben', data: { kind: 'answer', sdp: `answer:${theirs.n}` } });
  await settle();
  const mine = [...session.room.mesh.peers.values()][0];
  assert.ok(mine.channel.sent.map((d) => JSON.parse(d)).some((m) => m.t === 'air' && m.on && m.you), JSON.stringify(mine.channel.sent));
  assert.equal(session.where('u-ben'), 'on-air');
});

test('another setup opened while guests are here is given the seats, and a window with no meeting is left alone', async () => {
  const { session, win, push, comes } = studioOf();
  push();
  assert.deepEqual(win.stack, base(), 'no seats until there is a meeting');
  await session.hold(ID);
  await comes(BEN);
  win.stack = base();
  push();
  assert.deepEqual(win.stack.items.filter((it) => it.kind === GUEST).length, 3);
  assert.deepEqual(session.seats, { 'seat-1': BEN });
  assert.equal(session.where('u-ben'), 'off-air');
});
