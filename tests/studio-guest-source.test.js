// Guests from Talk as the studio draws and mixes them
// (apps/desktop/renderer/studio/guest-source.js, Office 2.27.0): a seat's
// picture or name card, and where each guest's sound goes. What it holds to:
// the person at the computer always hears a guest, a guest's sound reaches
// what is recorded and streamed only while their seat is shown on air with
// Sound on, and what is sent back to the guests never has their own sound
// in it. Run on stand-ins for the page and for Web Audio.

import test from 'node:test';
import assert from 'node:assert/strict';
import { GuestSources } from '../apps/desktop/renderer/studio/guest-source.js';

class FakeStream {
  constructor(tracks = []) { this.tracks = tracks; }
  getAudioTracks() { return this.tracks.filter((t) => t.kind === 'audio'); }
  getVideoTracks() { return this.tracks.filter((t) => t.kind === 'video'); }
}
globalThis.MediaStream ||= FakeStream;

const track = (kind) => ({ kind, clone() { return { kind, copy: true }; } });

/** Web Audio standing in: every node remembers what it is joined to. */
function sound() {
  const node = (kind, more = {}) => {
    const n = { kind, to: new Set(), gain: { value: 1 }, connect(x) { n.to.add(x); return x; }, disconnect(x) { if (x) n.to.delete(x); else n.to.clear(); }, ...more };
    return n;
  };
  const speakers = node('speakers');
  const ctx = {
    destination: speakers,
    createGain: () => node('gain'),
    createMediaStreamSource: (stream) => node('source', { stream }),
    createMediaStreamDestination: () => node('dest', { stream: new FakeStream([track('audio')]) }),
  };
  const local = node('gain');
  const air = node('dest');
  return { ctx, speakers, local, air };
}

/** The page standing in: elements that remember where they were put. */
function page() {
  const made = [];
  const el = (tag) => {
    const e = { tag, style: {}, children: [], videoWidth: 0, listeners: {}, appendChild(c) { e.children.push(c); c.parent = e; }, remove() { if (e.parent) e.parent.children = e.parent.children.filter((c) => c !== e); e.removed = true; }, setAttribute() {}, addEventListener(name, fn) { e.listeners[name] = fn; }, play: async () => {}, getContext: () => null };
    made.push(e);
    return e;
  };
  const body = el('body');
  const doc = { body, createElement: el, getElementById: (id) => made.find((e) => e.id === id) || null };
  return { doc, made };
}

function studio() {
  const s = sound();
  const p = page();
  let changes = 0;
  const guests = new GuestSources({ audio: () => s.ctx, local: () => s.local, air: () => s.air, onChange: () => { changes += 1; }, doc: p.doc });
  return { guests, ...s, ...p, changes: () => changes };
}

const seat = (ref, on = false) => ({ id: ref.replace('seat', 'guest'), kind: 'guest', ref, on });
const person = (id, name, more = {}) => ({ id, name, stream: new FakeStream([track('audio'), track('video')]), connected: true, camera: true, sharing: false, muted: false, ...more });
const frame = (shown, soundMode = 'on') => ({ sound: soundMode, items: ['seat-1', 'seat-2', 'seat-3'].map((ref) => seat(ref, shown.includes(ref))) });

test('a seat nobody sits in draws nothing; a guest with no picture is a card with their name; with a camera, their picture', () => {
  const { guests } = studio();
  assert.equal(guests.drawable(seat('seat-1')), null);
  assert.equal(guests.stateOf(seat('seat-1')), 'off');
  assert.equal(guests.streamOf(seat('seat-1')), null);

  const ben = person('u-ben', 'Ben Ortiz', { camera: false });
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [ben]);
  const card = guests.drawable(seat('seat-1'));
  assert.equal(card.tag, 'canvas');
  assert.equal(card.width, 640);
  assert.equal(guests.drawable(seat('seat-1')), card, 'the same card, not made again every frame');
  assert.equal(guests.streamOf(seat('seat-1')), null);
  assert.equal(guests.stateOf(seat('seat-1')), 'live');
  assert.equal(guests.nameOf(seat('seat-1')), 'Ben Ortiz');

  // The camera comes on, and its first frame arrives.
  ben.camera = true;
  const video = guests.held.get('u-ben').video;
  assert.equal(guests.drawable(seat('seat-1')), card, 'no frame yet: the card still');
  video.videoWidth = 1280;
  assert.equal(guests.drawable(seat('seat-1')), video);
  assert.equal(guests.streamOf(seat('seat-1')), ben.stream);
  assert.equal(video.muted, true, 'the video is silent here: the sound goes through the graph');
  assert.equal(video.srcObject, ben.stream);
  // Still connecting says so.
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [{ ...ben, connected: false }]);
  assert.equal(guests.stateOf(seat('seat-1')), 'opening');
});

test('a guest is always heard at this computer, and reaches what is recorded and streamed only while shown on air with Sound on', () => {
  const { guests, speakers, air } = studio();
  const ben = person('u-ben', 'Ben Ortiz');
  const cara = person('u-cara', 'Cara Lund');
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' }, 'seat-2': { id: 'u-cara', name: 'Cara Lund' } }, [ben, cara]);
  const b = guests.held.get('u-ben');
  const c = guests.held.get('u-cara');
  for (const h of [b, c]) {
    assert.ok(h.src.to.has(h.hear) && h.hear.to.has(speakers), 'heard here');
    assert.equal(h.hear.gain.value, 1);
    assert.ok(h.src.to.has(h.air) && h.air.to.has(air), 'joined to the air, closed');
    assert.equal(h.air.gain.value, 0, 'off air until shown');
  }
  assert.equal(guests.heard(), false);

  guests.sync(frame(['seat-1']));
  assert.deepEqual([b.air.gain.value, c.air.gain.value], [1, 0]);
  assert.deepEqual([guests.onAir('u-ben'), guests.onAir('u-cara')], [true, false]);
  assert.equal(guests.heard(), true);

  // Sound off, or Back soon's music: no guest reaches the air, shown or not.
  guests.sync(frame(['seat-1', 'seat-2'], 'off'));
  assert.deepEqual([b.air.gain.value, c.air.gain.value], [0, 0]);
  guests.sync(frame(['seat-1', 'seat-2'], 'music'));
  assert.deepEqual([b.air.gain.value, c.air.gain.value], [0, 0]);
  assert.equal(guests.heard(), false);
  guests.sync(frame(['seat-2']));
  assert.deepEqual([b.air.gain.value, c.air.gain.value], [0, 1]);
  assert.equal(b.hear.gain.value, 1, 'still heard here');
});

test('someone who takes a seat that is on air stays off it until the frame says so again', () => {
  const { guests } = studio();
  guests.sync(frame(['seat-1']));
  // The seat was taken off the frame as the person sat down (guests.js): the frame comes first here too.
  guests.sync(frame([]));
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [person('u-ben', 'Ben Ortiz')]);
  assert.equal(guests.held.get('u-ben').air.gain.value, 0);
  guests.sync(frame(['seat-1']));
  assert.equal(guests.held.get('u-ben').air.gain.value, 1);
});

test('someone who leaves is let go: their video out of the page, their sound out of the graph', () => {
  const { guests, air, speakers } = studio();
  const ben = person('u-ben', 'Ben Ortiz');
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [ben]);
  guests.sync(frame(['seat-1']));
  const h = guests.held.get('u-ben');
  const { video, src, hear, air: gate } = h;
  guests.set({}, []);
  assert.equal(guests.held.size, 0);
  assert.equal(video.removed, true);
  assert.equal(video.srcObject, null);
  assert.equal(src.to.size + hear.to.size + gate.to.size, 0);
  assert.ok(!gate.to.has(air) && !hear.to.has(speakers));
  assert.equal(guests.drawable(seat('seat-1')), null);
  // The connection made again gives a new stream: the old one is let go and the new one held.
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [ben]);
  const first = guests.held.get('u-ben');
  const again = person('u-ben', 'Ben Ortiz');
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [again]);
  assert.notEqual(guests.held.get('u-ben'), first);
  assert.equal(guests.held.get('u-ben').stream, again.stream);
  assert.equal(guests.held.get('u-ben').air.gain.value, 1, 'and it is on air as the seat is');
});

test('a guest whose sound has not arrived yet is joined when it does', () => {
  const { guests } = studio();
  const ben = person('u-ben', 'Ben Ortiz');
  ben.stream.tracks = [];
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [ben]);
  assert.equal(guests.held.get('u-ben').src, null);
  ben.stream.tracks = [track('audio')];
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [ben]);
  assert.ok(guests.held.get('u-ben').src);
  const src = guests.held.get('u-ben').src;
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [ben]);
  assert.equal(guests.held.get('u-ben').src, src, 'not joined twice');
});

test('what is sent back to the guests is the programme and this computer\'s own sound, never the guests\'', () => {
  const { guests, local, air } = studio();
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [person('u-ben', 'Ben Ortiz')]);
  guests.sync(frame(['seat-1']));
  const picture = track('video');
  const asked = [];
  const canvas = { captureStream: (fps) => { asked.push(fps); return new FakeStream([picture]); } };
  const sent = guests.feed(canvas);
  assert.equal(sent.video, picture);
  assert.deepEqual(asked, [15]);
  assert.equal(sent.audio.copy, true, 'a copy: the call stops what it is given');
  assert.ok(local.to.has(guests.back), 'this computer\'s own sound');
  const h = guests.held.get('u-ben');
  assert.ok(!h.src.to.has(guests.back) && !h.air.to.has(guests.back) && !h.hear.to.has(guests.back), 'no guest in it');
  assert.ok(!air.to.has(guests.back), 'and not what goes on air, which has the guests');
  // Asked for again (the connection made anew): the same way back, fresh tracks.
  const back = guests.back;
  guests.feed(canvas);
  assert.equal(guests.back, back);
  // No canvas, or no sound to be had: what there is still goes.
  assert.deepEqual(guests.feed(null).video, null);
});

test('closed, it holds nobody and draws nothing', () => {
  const { guests } = studio();
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [person('u-ben', 'Ben Ortiz')]);
  guests.dispose();
  assert.equal(guests.held.size, 0);
  assert.equal(guests.drawable(seat('seat-1')), null);
  guests.set({ 'seat-1': { id: 'u-ben', name: 'Ben Ortiz' } }, [person('u-ben', 'Ben Ortiz')]);
  assert.equal(guests.held.size, 0, 'nothing is opened after it closed');
});
