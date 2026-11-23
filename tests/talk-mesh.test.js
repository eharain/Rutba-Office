// The small group call (Office 2.27, renderer/apps/talk/mesh.js): up to four,
// each computer connected to each other. Whoever comes in offers to everyone
// already there. How each person is (muted, camera, showing a screen, hand
// up) and what is written travel on a line between each pair, never through
// a server. Run on a stand-in for WebRTC.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Mesh, readWire, MESH_MOST, SAY_MOST } from '../apps/desktop/renderer/apps/talk/mesh.js';
import { fakeRtc } from './talk-fake-rtc.js';

/** People of one call, their set-ups passed straight to each other, as the introducer would. */
function callOf(names, { rtc = fakeRtc(), mayMute = () => false, setup = null } = {}) {
  const meshes = {};
  const log = { signals: [], says: [], muted: [], changes: 0 };
  for (const id of names) {
    meshes[id] = new Mesh({
      me: id,
      rtc,
      setup: async () => (setup ? setup(id) : { iceServers: [{ urls: ['stun:stun.example'] }] }),
      signal: (to, data) => { log.signals.push({ from: id, to, data }); meshes[to]?.signalled(id, data); },
      onChange: () => { log.changes += 1; },
      onSay: (from, text, sid) => log.says.push({ at: id, from, text, id: sid }),
      mayMute: (from) => mayMute(id, from),
      onMuted: () => log.muted.push(id),
    });
  }
  const settle = async () => { for (let i = 0; i < 6; i++) for (const m of Object.values(meshes)) await m.line; };
  return { meshes, log, rtc, settle };
}

test('what travels between two computers is checked as it arrives', () => {
  assert.deepEqual(readWire('{"t":"state","muted":true,"camera":1,"hand":true}'), { t: 'state', muted: true, camera: false, sharing: false, hand: true });
  assert.deepEqual(readWire('{"t":"say","id":"s1","text":"hello"}'), { t: 'say', id: 's1', text: 'hello' });
  assert.deepEqual(readWire('{"t":"mute"}'), { t: 'mute' });
  assert.equal(readWire(JSON.stringify({ t: 'say', id: 's1', text: 'x'.repeat(SAY_MOST + 50) })).text.length, SAY_MOST);
  for (const bad of ['not json', 'null', '7', '{"t":"say","id":"s1","text":"   "}', '{"t":"say","text":"no id"}', JSON.stringify({ t: 'say', id: 'x'.repeat(65), text: 'hi' }), '{"t":"say","id":7,"text":"hi"}', '{"t":"remove"}', '{"t":"lock"}', '{}']) {
    assert.equal(readWire(bad), null, bad);
  }
  assert.equal(MESH_MOST, 4);
});

test('whoever comes in offers to those already there, and the pair is connected', async () => {
  const { meshes, log, settle } = callOf(['a', 'b']);
  await meshes.a.start();
  await meshes.b.start();
  meshes.a.expect({ id: 'b', name: 'Ben' });
  meshes.b.call({ id: 'a', name: 'Amira' });
  await settle();

  assert.deepEqual(log.signals.map((s) => `${s.from}>${s.to}:${s.data.kind}`), ['b>a:offer', 'a>b:answer']);
  const [ben] = meshes.a.people();
  const [amira] = meshes.b.people();
  assert.deepEqual({ id: ben.id, name: ben.name, connected: ben.connected }, { id: 'b', name: 'Ben', connected: true });
  assert.deepEqual({ id: amira.id, name: amira.name, connected: amira.connected }, { id: 'a', name: 'Amira', connected: true });
  assert.equal(meshes.a.names.size, 0, 'no longer only waited for');
  // Both ends carry a sound and a picture line, the microphone on the first.
  for (const m of [meshes.a, meshes.b]) {
    const trs = [...m.peers.values()][0].pc.getTransceivers();
    assert.deepEqual(trs.map((t) => [t.receiver.track.kind, t.direction, t.sender.track?.kind ?? null]), [['audio', 'sendrecv', 'audio'], ['video', 'sendrecv', null]]);
  }
});

test('nobody offers to themselves, and a set-up from nobody is dropped', async () => {
  const { meshes, log, settle, rtc } = callOf(['a', 'b']);
  await meshes.a.start();
  meshes.a.call({ id: 'a' });
  meshes.a.expect({ id: 'a' });
  meshes.a.signalled('a', { kind: 'offer', sdp: 'offer:1' });
  meshes.a.signalled('b', null);
  meshes.a.signalled('b', { kind: 'answer', sdp: 'answer:9' });
  meshes.a.signalled('b', { kind: 'candidate', candidate: { c: 1 } });
  meshes.a.signalled('b', { kind: 'offer' });
  await settle();
  assert.equal(rtc.pcs.length, 0);
  assert.equal(log.signals.length, 0);
  assert.equal(meshes.a.names.size, 0);
});

test('three people: the third offers to both, and each has two connections', async () => {
  const { meshes, settle } = callOf(['a', 'b', 'c']);
  for (const m of Object.values(meshes)) await m.start();
  meshes.b.call({ id: 'a' });
  await settle();
  meshes.c.call({ id: 'a' });
  meshes.c.call({ id: 'b' });
  await settle();
  for (const m of Object.values(meshes)) {
    assert.equal(m.people().length, 2);
    assert.ok(m.people().every((p) => p.connected));
  }
});

test('how a person is right now reaches the others as soon as the pair is connected, and when it changes', async () => {
  const { meshes, settle } = callOf(['a', 'b']);
  await meshes.a.start();
  await meshes.b.start();
  meshes.b.setMuted(true);
  meshes.b.call({ id: 'a' });
  await settle();
  const ben = () => meshes.a.people()[0];
  assert.equal(ben().muted, true, 'told on connecting');
  assert.equal(ben().hand, false);

  meshes.b.setHand(true);
  meshes.b.setMuted(false);
  assert.deepEqual({ muted: ben().muted, hand: ben().hand, camera: ben().camera, sharing: ben().sharing }, { muted: false, hand: true, camera: false, sharing: false });
  assert.equal(meshes.b.mic.enabled, true);
  meshes.b.setMuted(true);
  assert.equal(meshes.b.mic.enabled, false, 'muted is the microphone off, not only a mark');
});

test('without a microphone a person listens, and says so', async () => {
  const { meshes, settle } = callOf(['a', 'b'], { rtc: fakeRtc({ microphone: false }) });
  await meshes.a.start();
  await meshes.b.start();
  assert.equal(meshes.a.mine().muted, true);
  meshes.a.setMuted(false);
  assert.equal(meshes.a.mine().muted, true, 'there is nothing to unmute');
  meshes.b.call({ id: 'a' });
  await settle();
  assert.equal(meshes.b.people()[0].muted, true);
});

test('the camera and a shown screen come and go without the set-up being made again', async () => {
  const { meshes, settle, rtc, log } = callOf(['a', 'b']);
  await meshes.a.start();
  await meshes.b.start({ video: true });
  meshes.b.call({ id: 'a' });
  await settle();
  const made = rtc.pcs.length;
  const said = log.signals.length;
  const video = () => [...meshes.b.peers.values()][0].pc.getTransceivers()[1].sender.track;
  const ben = () => meshes.a.people()[0];
  const cam = meshes.b.camera;
  assert.equal(video(), cam);
  assert.equal(ben().camera, true);

  const screen = new rtc.MediaStream([new rtc.Track('video')]);
  meshes.b.share(screen);
  await settle();
  assert.equal(video(), screen.getVideoTracks()[0], 'the screen in place of the camera');
  assert.equal(ben().sharing, true);

  screen.getVideoTracks()[0].onended();
  await settle();
  assert.equal(video(), cam, 'the camera back when the screen stops');
  assert.equal(ben().sharing, false);

  await meshes.b.setCamera(false);
  await settle();
  assert.equal(video(), null);
  assert.equal(cam.stopped, true);
  assert.equal(ben().camera, false);
  assert.equal(rtc.pcs.length, made, 'no new connection');
  assert.equal(log.signals.length, said, 'no new set-up');
});

test('what is written goes to everyone connected, once, and is cut to what one message holds', async () => {
  const { meshes, settle, log } = callOf(['a', 'b', 'c']);
  for (const m of Object.values(meshes)) await m.start();
  meshes.b.call({ id: 'a' });
  meshes.c.call({ id: 'a' });
  meshes.c.call({ id: 'b' });
  await settle();
  const sent = meshes.a.say('Shall we start?');
  assert.equal(sent.to, 2);
  assert.deepEqual(log.says.map((s) => [s.at, s.from, s.text, s.id]), [['b', 'a', 'Shall we start?', sent.id], ['c', 'a', 'Shall we start?', sent.id]]);
  assert.equal(meshes.a.say('   '), null);
  assert.equal(meshes.a.say('x'.repeat(SAY_MOST + 10)).text.length, SAY_MOST);
  assert.notEqual(meshes.a.say('one').id, meshes.a.say('two').id);
});

test('a request to mute is honoured only from whoever may ask', async () => {
  // Only "a" may mute people (in a meeting: the person who set it up).
  const { meshes, settle, log } = callOf(['a', 'b', 'c'], { mayMute: (me, from) => from === 'a' });
  for (const m of Object.values(meshes)) await m.start();
  meshes.b.call({ id: 'a' });
  meshes.c.call({ id: 'a' });
  meshes.c.call({ id: 'b' });
  await settle();

  assert.equal(meshes.c.askMute('b'), true, 'sent, but it is for b to honour');
  assert.equal(meshes.b.muted, false);
  assert.deepEqual(log.muted, []);

  assert.equal(meshes.a.askMute('b'), true);
  assert.equal(meshes.b.muted, true);
  assert.equal(meshes.b.mic.enabled, false);
  assert.deepEqual(log.muted, ['b']);
  assert.equal(meshes.a.people().find((p) => p.id === 'b').muted, true, 'and everyone sees it');
  assert.equal(meshes.c.people().find((p) => p.id === 'b').muted, true);

  meshes.a.askMute('b');
  assert.deepEqual(log.muted, ['b'], 'already muted: nothing more');
  assert.equal(meshes.a.askMute('nobody'), false);
  // Nothing on the line can unmute someone: only they can.
  [...meshes.a.peers.values()].find((p) => p.id === 'b').channel.send('{"t":"unmute"}');
  assert.equal(meshes.b.muted, true);
});

test('a way to reach someone that arrives before their answer is kept until it can be used', async () => {
  const rtc = fakeRtc();
  const held = [];
  const b = new Mesh({ me: 'b', rtc, setup: async () => ({ iceServers: [] }), signal: (to, data) => held.push(data) });
  await b.start();
  b.call({ id: 'a' });
  await b.line;
  const pc = rtc.pcs[0];
  b.signalled('a', { kind: 'candidate', candidate: { c: 1 } });
  b.signalled('a', { kind: 'candidate', candidate: null });
  await b.line;
  assert.deepEqual(pc.candidates, [], 'held: there is no answer yet');
  // The other side, made by hand.
  const other = new rtc.PeerConnection({});
  other.createDataChannel('x');
  b.signalled('a', { kind: 'answer', sdp: `answer:${other.n}` });
  b.signalled('a', { kind: 'candidate', candidate: { c: 2 } });
  await b.line;
  assert.deepEqual(pc.candidates, [{ c: 1 }, null, { c: 2 }]);
  // This side's own ways are handed on as they are found.
  pc.onicecandidate({ candidate: { toJSON: () => ({ c: 9 }) } });
  pc.onicecandidate({ candidate: null });
  assert.deepEqual(held.filter((d) => d.kind === 'candidate').map((d) => d.candidate), [{ c: 9 }, null]);
});

test("each connection uses the relays the set-up names, and the other side's own when they offer one", async () => {
  const relay = { urls: ['turn:relay.example'], username: 'u', credential: 'c' };
  const { meshes, settle, log, rtc } = callOf(['a', 'b'], { setup: (id) => (id === 'b' ? { iceServers: [{ urls: ['stun:one.example'] }], offerRelay: relay } : { iceServers: [{ urls: ['stun:two.example'] }] }) });
  await meshes.a.start();
  await meshes.b.start();
  meshes.b.call({ id: 'a' });
  await settle();
  assert.deepEqual(log.signals[0].data.relay, relay, "the one who offers hands over their own relay");
  assert.deepEqual(rtc.pcs[0].config.iceServers, [{ urls: ['stun:one.example'] }]);
  assert.deepEqual(rtc.pcs[1].config.iceServers, [{ urls: ['stun:two.example'] }, relay]);
});

test('someone gone is let go, and a connection that fails says so', async () => {
  const { meshes, settle } = callOf(['a', 'b', 'c']);
  for (const m of Object.values(meshes)) await m.start();
  meshes.b.call({ id: 'a' });
  meshes.c.call({ id: 'a' });
  await settle();
  const pcOfB = [...meshes.a.peers.values()].find((p) => p.id === 'b').pc;
  const pcOfC = [...meshes.a.peers.values()].find((p) => p.id === 'c').pc;
  pcOfC.fail();
  assert.deepEqual(meshes.a.people().map((p) => [p.id, p.connected, p.failed]), [['b', true, false], ['c', false, true]]);

  meshes.a.drop('b');
  await settle();
  assert.equal(pcOfB.closed, true);
  assert.deepEqual(meshes.a.people().map((p) => p.id), ['c']);
  // Someone only waited for, who left before offering.
  meshes.a.expect({ id: 'd', name: 'Dana' });
  meshes.a.drop('d');
  await settle();
  assert.equal(meshes.a.names.size, 0);
});

test('someone who comes in again is connected anew, not twice', async () => {
  const { meshes, settle, rtc } = callOf(['a', 'b']);
  await meshes.a.start();
  await meshes.b.start();
  meshes.b.call({ id: 'a' });
  await settle();
  const first = [...meshes.a.peers.values()][0].pc;
  meshes.b.call({ id: 'a' });
  await settle();
  assert.equal(first.closed, true);
  assert.equal(meshes.a.people().length, 1);
  assert.equal(meshes.a.people()[0].connected, true);
  assert.equal(rtc.pcs.filter((p) => !p.closed).length, 2);
});

test('leaving closes every connection and lets go of the microphone, camera and screen', async () => {
  const { meshes, settle, rtc, log } = callOf(['a', 'b']);
  await meshes.a.start({ video: true });
  await meshes.b.start();
  meshes.b.call({ id: 'a' });
  await settle();
  const screen = new rtc.MediaStream([new rtc.Track('video')]);
  meshes.a.share(screen);
  const { mic, camera } = meshes.a;
  const pc = [...meshes.a.peers.values()][0].pc;
  meshes.a.close();
  assert.equal(pc.closed, true);
  assert.ok(mic.stopped && camera.stopped && screen.getVideoTracks()[0].stopped);
  assert.equal(meshes.a.people().length, 0);
  const before = log.changes;
  meshes.a.setHand(true);
  meshes.a.signalled('b', { kind: 'offer', sdp: 'offer:1' });
  await settle();
  assert.equal(log.changes, before, 'a call that is over tells its screen nothing more');
  assert.equal(meshes.a.people().length, 0, 'and takes no new set-up');
  meshes.a.close();
});

/* ── a connection that failed, tried again (2.27.0) ── */

/** Two people connected, then the connection between them failed at both ends. */
async function failedPair() {
  const it = callOf(['a', 'b']);
  await it.meshes.a.start();
  await it.meshes.b.start();
  it.meshes.b.call({ id: 'a', name: 'Amira' });
  await it.settle();
  it.meshes.a.peers.get('b').pc.fail();
  it.meshes.b.peers.get('a').pc.fail();
  assert.deepEqual([it.meshes.a.people()[0].failed, it.meshes.b.people()[0].failed], [true, true]);
  it.log.signals.length = 0;
  return it;
}

test('trying again: the one whose id sorts first offers anew, and the pair is connected', async () => {
  const { meshes, log, settle, rtc } = await failedPair();
  assert.equal(meshes.a.redial('b'), true);
  assert.equal(meshes.a.people()[0].failed, false, 'no longer said as failed while it is tried');
  await settle();
  assert.deepEqual(log.signals.map((s) => `${s.from}>${s.to}:${s.data.kind}`), ['a>b:offer', 'b>a:answer']);
  assert.deepEqual(meshes.a.people().map((p) => [p.id, p.connected, p.failed]), [['b', true, false]]);
  assert.deepEqual(meshes.b.people().map((p) => [p.id, p.connected, p.failed]), [['a', true, false]]);
  assert.equal(rtc.pcs.filter((p) => !p.closed).length, 2, 'the failed connections are closed, not left beside the new ones');
});

test('trying again from the other side asks the first to offer, so both pressing at once never meet half way', async () => {
  const { meshes, log, settle } = await failedPair();
  assert.equal(meshes.b.redial('a'), true);
  await settle();
  assert.deepEqual(log.signals.map((s) => `${s.from}>${s.to}:${s.data.kind}`), ['b>a:again', 'a>b:offer', 'b>a:answer']);
  assert.ok(meshes.a.people()[0].connected && meshes.b.people()[0].connected);

  // Both at once: one offer still, from the same side.
  const both = await failedPair();
  both.meshes.b.redial('a');
  both.meshes.a.redial('b');
  await both.settle();
  assert.ok(both.log.signals.every((s) => s.data.kind !== 'offer' || s.from === 'a'), 'only the first ever offers');
  assert.ok(both.meshes.a.people()[0].connected && both.meshes.b.people()[0].connected);
});

test('nothing is tried for someone connected or not there, and asking again of a connected pair changes nothing', async () => {
  const { meshes, log, settle, rtc } = callOf(['a', 'b']);
  await meshes.a.start();
  await meshes.b.start();
  meshes.b.call({ id: 'a' });
  await settle();
  const before = log.signals.length;
  const made = rtc.pcs.length;
  assert.equal(meshes.a.redial('b'), false);
  assert.equal(meshes.b.redial('a'), false);
  assert.equal(meshes.a.redial('nobody'), false);
  meshes.a.signalled('b', { kind: 'again' });
  meshes.a.signalled('nobody', { kind: 'again' });
  await settle();
  assert.equal(log.signals.length, before);
  assert.equal(rtc.pcs.length, made, 'no connection made anew');
});

test('asking again with no answer says it failed once more, after a wait', async () => {
  const { meshes, settle } = callOf(['a', 'b']);
  await meshes.a.start();
  await meshes.b.start();
  meshes.b.call({ id: 'a' });
  await settle();
  // Failed at one end only: the other still holds it connected, and so does not offer.
  meshes.b.peers.get('a').pc.fail();
  assert.equal(meshes.b.redial('a', { wait: 15 }), true);
  assert.equal(meshes.b.people()[0].failed, false);
  await new Promise((r) => setTimeout(r, 60));
  assert.deepEqual(meshes.b.people().map((p) => [p.connected, p.failed]), [[false, true]]);
});

test('a call closed while its microphone was still being opened keeps no microphone and no camera', async () => {
  const rtc = fakeRtc();
  const real = rtc.getUserMedia;
  let open = null;
  let mic = null;
  rtc.getUserMedia = (c) => new Promise((resolve) => { open = async () => { const s = await real(c); mic = s.getAudioTracks()[0]; resolve(s); }; });
  const { meshes } = callOf(['a'], { rtc });
  const camera = new rtc.Track('video');
  const starting = meshes.a.start({ camera });
  meshes.a.close();
  await open();
  await starting;
  assert.equal(mic.stopped, true);
  assert.equal(camera.stopped, true);
  assert.deepEqual([meshes.a.mic, meshes.a.camera], [null, null]);
});
