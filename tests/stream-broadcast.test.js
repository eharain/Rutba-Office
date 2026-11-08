// A broadcast with streams joining and leaving (packages/stream/src/broadcast.js),
// the bar's rules (bar.js) and the platform list. The publishers are stand-ins
// that record what they were sent: no network.

import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { Broadcast } from '../packages/stream/src/broadcast.js';
import { PLATFORMS, parseRtmpUrl } from '../packages/stream/src/endpoints.js';
import { barModel, toggleStream, toggleGroup, reconcileOn, shapesOn, uploadMbps, onChanges } from '../packages/stream/src/bar.js';
import { isKeyBody } from '../packages/stream/src/flv.js';

class FakePublisher extends EventEmitter {
  constructor(o, log) {
    super();
    this.o = o;
    this.log = log;
    this.state = 'idle';
    this.sent = 0;
    this.dropped = 0;
    this.socket = { writableLength: 0 };
    this.maxQueueBytes = 100;
    this.items = [];
    log.push(this);
  }
  async start() {
    if (this.o.key === 'refuse') throw new Error('The server refused the stream: bad key');
    if (this.o.key === 'hang') await new Promise(() => {});
    this.state = 'live';
  }
  sendMetadata(b) { this.items.push(['meta', b]); }
  sendVideo(b, ts, { config = false } = {}) { this.items.push([config ? 'vconfig' : isKeyBody(b) ? 'key' : 'inter', ts]); return true; }
  sendAudio(b, ts) { this.items.push(['audio', ts]); return true; }
  close() { if (this.state === 'closed') return; this.state = 'closed'; this.emit('close'); }
  drop() { this.state = 'closed'; this.emit('error', new Error('line gone')); this.emit('close'); }
}

const tick = () => new Promise((r) => setTimeout(r, 5));
const data = (n = 4) => new Uint8Array(n).fill(1);

function open(extra = {}) {
  const pubs = [];
  const seen = [];
  const b = new Broadcast({
    video: { width: 1920, height: 1080, fps: 30, kbps: 6000 },
    tall: { width: 1080, height: 1920, fps: 30, kbps: 6000 },
    audio: { kbps: 128 },
    makePublisher: (o) => new FakePublisher(o, pubs),
    onChange: (s) => seen.push(s),
    backoffMs: [5, 5, 5],
    tickMs: 100000,
    ...extra,
  });
  return { b, pubs, seen, stream: (id, shape = 'wide', key = `k-${id}`, more = {}) => b.add({ id, name: id, url: 'rtmp://x/a', key, shape, ...more }) };
}

test('a broadcast opens with nothing going, and stays open', () => {
  const { b } = open();
  assert.equal(b.status().live, true);
  assert.deepEqual(b.status().destinations, []);
  b.close();
  assert.equal(b.status().live, false);
});

test('a stream joining late is given the metadata and the configurations first, and the encoder is asked for a key frame', async () => {
  const { b, pubs, seen, stream } = open();
  stream('a');
  await tick();
  b.send({ kind: 'video', timestamp: 100, data: data(), config: true });
  b.send({ kind: 'audio', timestamp: 100, data: data(), config: true });
  b.send({ kind: 'video', timestamp: 133, data: data(), key: true });
  seen.length = 0;
  stream('b');
  await tick();
  const late = pubs[1];
  assert.deepEqual(late.items.map((x) => x[0]), ['meta', 'vconfig', 'audio']);
  assert.equal(late.items[1][1], 133, 'the configurations at the broadcast\'s current time');
  assert.ok(seen.some((s) => s.keyFor?.includes('wide')), 'the window is told to make a key frame');
  assert.equal(seen.filter((s) => s.keyFor).length, 1, 'once');
  b.close();
});

test('stopping one stream leaves the others going and is not retried', async () => {
  const { b, pubs, stream } = open();
  stream('a');
  stream('b');
  await tick();
  assert.equal(b.remove('a'), true);
  assert.equal(pubs[0].state, 'closed');
  await tick();
  assert.equal(pubs.length, 2, 'nothing tried again');
  b.send({ kind: 'video', timestamp: 200, data: data(), key: true });
  assert.deepEqual(pubs[1].items.at(-1), ['key', 200], 'the other carries on');
  assert.equal(pubs[0].items.filter((x) => x[0] === 'key').length, 0);
  assert.equal(b.running, true);
  b.close();
});

test('every stream off is still a broadcast, and one turned back on goes out again', async () => {
  const { b, pubs, stream } = open();
  stream('a');
  await tick();
  b.send({ kind: 'video', timestamp: 10, data: data(), config: true });
  b.remove('a');
  assert.deepEqual(b.status().destinations, []);
  assert.equal(b.status().live, true);
  stream('a');
  await tick();
  assert.equal(pubs.length, 2, 'a new connection');
  assert.deepEqual(pubs[1].items.map((x) => x[0]), ['meta', 'vconfig']);
  assert.equal(b.status().destinations[0].state, 'live');
  b.close();
});

test('wide and tall go out apart: each stream is sent only its own shape, and the sound to both', async () => {
  const { b, pubs, stream } = open();
  stream('w', 'wide');
  stream('t', 'tall');
  await tick();
  b.send({ kind: 'video', shape: 'wide', timestamp: 1, data: data(), key: true });
  b.send({ kind: 'video', shape: 'tall', timestamp: 2, data: data(), key: true });
  b.send({ kind: 'audio', timestamp: 3, data: data() });
  assert.deepEqual(pubs[0].items.filter((x) => x[0] !== 'meta'), [['key', 1], ['audio', 3]]);
  assert.deepEqual(pubs[1].items.filter((x) => x[0] !== 'meta'), [['key', 2], ['audio', 3]]);
  b.close();
});

test('each shape has its own metadata and its own bitrate: a slow tall stream leaves wide alone', async () => {
  const { b, pubs, stream } = open();
  stream('w', 'wide');
  stream('t', 'tall');
  await tick();
  const meta = (p) => JSON.stringify([...p.items.find((x) => x[0] === 'meta')[1]]);
  assert.notEqual(meta(pubs[0]), meta(pubs[1]));
  pubs[1].dropped = 5;
  b.tick();
  assert.deepEqual(b.rates(), { wide: { kbps: 6000, target: 6000 }, tall: { kbps: 4500, target: 6000 } });
  b.close();
});

test('a platform\'s ceiling holds its shape down while it is on, and lifts when it leaves', async () => {
  const { b, stream } = open({ video: { width: 1920, height: 1080, fps: 60, kbps: 9000 } });
  stream('y', 'wide');
  stream('k', 'wide', 'k-k', { maxKbps: 8000 });
  await tick();
  assert.deepEqual(b.rates().wide, { kbps: 8000, target: 8000 });
  b.remove('k');
  assert.equal(b.rates().wide.target, 9000);
  b.close();
});

test('a refused key is said once and not tried again; a dropped line is tried again', async () => {
  const { b, pubs, stream } = open();
  stream('r', 'wide', 'refuse');
  await tick();
  await tick();
  const d = b.status().destinations[0];
  assert.equal(d.state, 'error');
  assert.equal(d.refused, true);
  assert.equal(pubs.length, 1);
  stream('l');
  await tick();
  pubs[1].drop();
  assert.equal(b.status().destinations.find((x) => x.id === 'l').state, 'reconnecting');
  await new Promise((r) => setTimeout(r, 30));
  assert.equal(pubs.length, 3, 'a new connection');
  assert.equal(b.status().destinations.find((x) => x.id === 'l').state, 'live');
  b.close();
});

test('a refused stream can be started afresh, as after a new key', async () => {
  const { b, pubs, stream } = open();
  stream('r', 'wide', 'refuse');
  await tick();
  stream('r', 'wide', 'better');
  await tick();
  assert.equal(pubs.at(-1).o.key, 'better');
  assert.equal(b.status().destinations[0].state, 'live');
  b.close();
});

test('adding a stream already going does not open a second connection', async () => {
  const { b, pubs, stream } = open();
  stream('a');
  stream('a');
  await tick();
  assert.equal(pubs.length, 1);
  b.close();
});

test('a tall stream on a broadcast with no tall picture is sent the wide one', async () => {
  const { b, pubs, stream } = open({ tall: null });
  stream('t', 'tall');
  await tick();
  b.send({ kind: 'video', timestamp: 1, data: data(), key: true });
  assert.deepEqual(pubs[0].items.at(-1), ['key', 1]);
  b.close();
});

test('the end closes everything and refuses more', async () => {
  const { b, pubs, stream } = open();
  stream('a');
  await tick();
  b.close();
  assert.equal(pubs[0].state, 'closed');
  assert.throws(() => stream('b'));
  assert.deepEqual(b.send({ kind: 'audio', timestamp: 1, data: data() }), { live: 0 });
});

/* ── the bar ─────────────────────────────────────────────────────────── */

const streams = [
  { id: 'yt', shape: 'wide', hasKey: true },
  { id: 'fb', shape: 'wide', hasKey: true },
  { id: 'tw', shape: 'wide', hasKey: false },
  { id: 'tk', shape: 'tall', hasKey: true },
];

test('an icon turns its stream on and off; one with no key cannot be turned on', () => {
  assert.deepEqual(toggleStream([], streams, 'yt'), ['yt']);
  assert.deepEqual(toggleStream(['yt'], streams, 'yt'), []);
  assert.deepEqual(toggleStream([], streams, 'tw'), []);
});

test('a removed stream drops out of the choices', () => {
  assert.deepEqual(reconcileOn(['yt', 'gone', 'tw'], streams), ['yt']);
});

test('Wide or Tall turns its whole group off, and again brings back the ones that were on', () => {
  const first = toggleGroup(['yt', 'tk'], streams, 'wide');
  assert.deepEqual(first.on, ['tk']);
  assert.deepEqual(first.memory.wide, ['yt']);
  const back = toggleGroup(first.on, streams, 'wide', first.memory);
  assert.deepEqual(back.on, ['yt', 'tk']);
  assert.deepEqual(back.memory.wide, []);
  assert.deepEqual(toggleGroup([], streams, 'tall').on, [], 'nothing to bring back');
  assert.deepEqual(toggleGroup(first.on, streams, 'tall', first.memory).on, [], 'the other group has its own');
});

test('the bar says what each stream is doing, and only for streams on the broadcast', () => {
  const m = barModel({
    streams, on: ['yt', 'fb', 'tk'], open: true,
    status: [{ id: 'yt', state: 'live' }, { id: 'fb', state: 'reconnecting' }, { id: 'tk', state: 'error', detail: 'no' }],
  });
  assert.deepEqual(m.tiles.map((x) => x.state), ['live', 'retrying', 'nokey', 'refused']);
  assert.equal(m.air, 'live');
  assert.equal(m.liveCount, 1);
  assert.equal(m.wide.on, true);
  assert.equal(m.tall.tiles.length, 1);
});

test('Go live needs something on; all off during a broadcast is offline, not ended', () => {
  assert.equal(barModel({ streams, on: [] }).canGo, false);
  assert.equal(barModel({ streams, on: [] }).air, 'ready');
  assert.equal(barModel({ streams, on: ['tw'] }).canGo, false, 'a stream with no key is not on');
  assert.equal(barModel({ streams, on: ['yt'] }).canGo, true);
  assert.equal(barModel({ streams, on: [], open: true }).air, 'offline');
  assert.equal(barModel({ streams, on: ['yt'], open: true, status: [{ id: 'yt', state: 'connecting' }] }).air, 'connecting');
  assert.equal(barModel({ streams, on: ['yt'], open: true, status: [{ id: 'yt', state: 'error' }] }).air, 'offline', 'refused and nothing else trying');
});

test('only the shapes with a stream on run an encoder; the upload counts both', () => {
  assert.deepEqual(shapesOn(streams, ['yt']), ['wide']);
  assert.deepEqual(shapesOn(streams, ['yt', 'tk']), ['wide', 'tall']);
  assert.deepEqual(shapesOn(streams, ['tk']), ['tall']);
  assert.deepEqual(shapesOn(streams, ['tw']), []);
  const m = barModel({ streams, on: ['yt', 'tk'] });
  assert.equal(uploadMbps(m.tiles, { wide: 6000, tall: 4000 }).toFixed(1), '11.8');
});

test('what changes when the choices change', () => {
  assert.deepEqual(onChanges(['a', 'b'], ['b', 'c']), { start: ['c'], stop: ['a'] });
});

/* ── the platforms ───────────────────────────────────────────────────── */

test('the platform list: named, marked, and only the three known addresses filled in', () => {
  assert.deepEqual(Object.keys(PLATFORMS), ['youtube', 'facebook', 'twitch', 'tiktok', 'instagram', 'linkedin', 'x', 'kick', 'rumble', 'custom']);
  for (const [id, p] of Object.entries(PLATFORMS)) {
    assert.ok(p.name && p.keyHelp && p.mark && /^#[0-9a-f]{6}$/.test(p.colour), id);
    if (p.url) assert.ok(parseRtmpUrl(p.url)?.secure, `${id} is encrypted`);
  }
  assert.deepEqual(Object.entries(PLATFORMS).filter(([, p]) => p.url).map(([id]) => id), ['youtube', 'facebook', 'twitch']);
  assert.deepEqual(Object.entries(PLATFORMS).filter(([, p]) => p.tall).map(([id]) => id), ['tiktok', 'instagram']);
  assert.equal(PLATFORMS.kick.maxKbps, 8000);
});
