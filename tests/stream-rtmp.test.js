// Live streaming (packages/stream): AMF0, the FLV packets, RTMP addresses, and
// a whole publish to a server on this computer that answers as YouTube's
// does: the handshake, connect, createStream, publish, then the metadata,
// the configurations and the frames, checked as the server read them. No
// network beyond this computer's loopback.

import test from 'node:test';
import assert from 'node:assert/strict';
import net from 'node:net';
import crypto from 'node:crypto';
import { encodeAmf0, decodeAmf0, EcmaArray, videoBody, audioBody, metadataBody, isKeyBody, RtmpPublisher, ChunkReader, chunkMessage, parseRtmpUrl, maskKey, liveBitrate, PLATFORMS } from '../packages/stream/src/index.js';

test('AMF0: commands and metadata are written and read back as they were', () => {
  const buf = encodeAmf0('connect', 1, { app: 'live2', tcUrl: 'rtmp://x/live2', n: 3.5, ok: true, none: null }, null, new EcmaArray({ width: 1920, stereo: true }));
  const [name, txn, obj, nul, meta] = decodeAmf0(buf);
  assert.equal(name, 'connect');
  assert.equal(txn, 1);
  assert.deepEqual(obj, { app: 'live2', tcUrl: 'rtmp://x/live2', n: 3.5, ok: true, none: null });
  assert.equal(nul, null);
  assert.deepEqual(meta, { width: 1920, stereo: true });
  const long = 'x'.repeat(70000);
  assert.equal(decodeAmf0(encodeAmf0(long))[0], long, 'a long string');
});

test('FLV bodies: AVC configuration and frames, AAC configuration and frames', () => {
  const cfg = videoBody({ data: new Uint8Array([1, 0x4d, 0, 0x28]), config: true });
  assert.deepEqual([...cfg.subarray(0, 5)], [0x17, 0, 0, 0, 0], 'a key frame, AVC, the sequence header');
  assert.ok(isKeyBody(cfg));
  const inter = videoBody({ data: new Uint8Array([0, 0, 0, 1, 0x41]) });
  assert.deepEqual([...inter.subarray(0, 2)], [0x27, 1], 'an inter frame, AVC, a NALU');
  assert.ok(!isKeyBody(inter));
  const key = videoBody({ data: new Uint8Array([0, 0, 0, 1, 0x65]), key: true });
  assert.equal(key[0], 0x17);
  assert.deepEqual([...audioBody({ data: new Uint8Array([0x11, 0x90]), config: true })], [0xaf, 0, 0x11, 0x90]);
  assert.deepEqual([...audioBody({ data: new Uint8Array([9]) })], [0xaf, 1, 9]);
  const meta = decodeAmf0(metadataBody({ width: 1280, height: 720, fps: 30, videoKbps: 4000, audioKbps: 128 }));
  assert.equal(meta[0], '@setDataFrame');
  assert.equal(meta[1], 'onMetaData');
  assert.equal(meta[2].videocodecid, 7);
  assert.equal(meta[2].audiocodecid, 10);
  assert.equal(meta[2].width, 1280);
});

test('RTMP addresses are taken apart, keys are masked, and the platforms are known', () => {
  assert.deepEqual(parseRtmpUrl('rtmps://a.rtmps.youtube.com:443/live2'), { secure: true, host: 'a.rtmps.youtube.com', port: 443, app: 'live2', tcUrl: 'rtmps://a.rtmps.youtube.com:443/live2' });
  assert.deepEqual(parseRtmpUrl('rtmp://example.org/live/'), { secure: false, host: 'example.org', port: 1935, app: 'live', tcUrl: 'rtmp://example.org/live' });
  assert.equal(parseRtmpUrl('rtmps://live-api-s.facebook.com:443/rtmp/').app, 'rtmp');
  assert.equal(parseRtmpUrl('https://example.org/live'), null);
  assert.equal(parseRtmpUrl('rtmp://example.org'), null, 'no app');
  assert.equal(maskKey('abcd-efgh-1234'), '••••1234');
  assert.equal(liveBitrate(1080, 60), 9000);
  assert.equal(liveBitrate(720, 30), 4000);
  for (const p of ['youtube', 'facebook', 'twitch']) assert.ok(parseRtmpUrl(PLATFORMS[p].url)?.secure, `${p} is reached over RTMPS`);
});

test('chunks: a long message split at the chunk size and read back whole, through a peer chunk size change', () => {
  const body = crypto.randomBytes(10000);
  const wire = chunkMessage({ csid: 6, type: 9, streamId: 1, timestamp: 1234, body, chunkSize: 4096 });
  const r = new ChunkReader();
  r.chunkSize = 4096;
  // Fed a few bytes at a time, as a socket hands them over.
  const got = [];
  for (let i = 0; i < wire.length; i += 777) got.push(...r.push(wire.subarray(i, i + 777)));
  assert.equal(got.length, 1);
  assert.equal(got[0].type, 9);
  assert.equal(got[0].streamId, 1);
  assert.equal(got[0].timestamp, 1234);
  assert.ok(got[0].body.equals(body));
  const late = chunkMessage({ csid: 6, type: 9, streamId: 1, timestamp: 0x1000000 + 5, body: Buffer.from('later') });
  assert.equal(r.push(late)[0].timestamp, 0x1000005, 'an extended timestamp');
});

/** A server on this computer that answers a publisher as YouTube's ingest does, and keeps what it was sent. */
function fakeIngest({ refuse = false } = {}) {
  const seen = { commands: [], data: [], video: [], audio: [], clientChunkSize: 128, pong: false, acks: 0 };
  const server = net.createServer((sock) => {
    let stage = 'c0c1';
    let buf = Buffer.alloc(0);
    const reader = new ChunkReader();
    const send = (csid, type, body, streamId = 0) => sock.write(chunkMessage({ csid, type, body, streamId, chunkSize: 128 }));
    const reply = (name, txn, ...args) => send(3, 20, encodeAmf0(name, txn, null, ...args));
    sock.on('data', (d) => {
      buf = Buffer.concat([buf, d]);
      if (stage === 'c0c1') {
        if (buf.length < 1537) return;
        const c1 = buf.subarray(1, 1537);
        const s1 = crypto.randomBytes(1536);
        sock.write(Buffer.concat([Buffer.from([3]), s1, c1]));
        buf = buf.subarray(1537);
        stage = 'c2';
      }
      if (stage === 'c2') {
        if (buf.length < 1536) return;
        buf = buf.subarray(1536);
        stage = 'chunks';
      }
      if (stage !== 'chunks' || !buf.length) return;
      const messages = reader.push(buf);
      buf = Buffer.alloc(0);
      for (const m of messages) {
        if (m.type === 1) { reader.chunkSize = m.body.readUInt32BE(0); seen.clientChunkSize = reader.chunkSize; }
        else if (m.type === 3) seen.acks += 1;
        else if (m.type === 4 && m.body.readUInt16BE(0) === 7) seen.pong = true;
        else if (m.type === 20) {
          const [name, txn, , ...rest] = decodeAmf0(m.body);
          seen.commands.push({ name, args: rest, streamId: m.streamId });
          if (name === 'connect') {
            seen.connect = decodeAmf0(m.body)[2];
            const w = Buffer.alloc(4); w.writeUInt32BE(5000); send(2, 5, w);
            const pb = Buffer.alloc(5); pb.writeUInt32BE(5000); pb[4] = 2; send(2, 6, pb);
            const ping = Buffer.alloc(6); ping.writeUInt16BE(6, 0); ping.writeUInt32BE(42, 2); send(2, 4, ping);
            reply('_result', txn, { code: 'NetConnection.Connect.Success' });
          } else if (name === 'createStream') reply('_result', txn, 1);
          else if (name === 'publish') {
            seen.key = rest[0];
            send(5, 20, encodeAmf0('onStatus', 0, null, refuse ? { level: 'error', code: 'NetStream.Publish.BadName', description: 'Wrong key' } : { level: 'status', code: 'NetStream.Publish.Start' }), 1);
          }
        } else if (m.type === 18) seen.data.push(decodeAmf0(m.body));
        else if (m.type === 9) seen.video.push({ ts: m.timestamp, body: m.body, streamId: m.streamId });
        else if (m.type === 8) seen.audio.push({ ts: m.timestamp, body: m.body });
      }
    });
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve({ server, port: server.address().port, seen })));
}

test('a whole publish to a local ingest: handshake, connect, createStream, publish with the key, then metadata, configurations and frames in order', async () => {
  const { server, port, seen } = await fakeIngest();
  const p = new RtmpPublisher({ url: `rtmp://127.0.0.1:${port}/live2`, key: 'test-key-1234', timeoutMs: 5000 });
  const states = [];
  p.on('state', (s) => states.push(s.state));
  await p.start();
  assert.equal(p.state, 'live');
  p.sendMetadata(metadataBody({ width: 1280, height: 720, fps: 30, videoKbps: 4000, audioKbps: 128 }));
  p.sendVideo(videoBody({ data: Buffer.from([1, 0x4d, 0, 0x28, 0xff]), config: true }), 0, { config: true });
  p.sendAudio(audioBody({ data: Buffer.from([0x11, 0x90]), config: true }), 0);
  // An inter frame before the first key frame is not sent: a viewer could not decode it.
  assert.equal(p.sendVideo(videoBody({ data: Buffer.alloc(300, 2) }), 10), false);
  for (let i = 0; i < 6; i++) {
    p.sendVideo(videoBody({ data: Buffer.alloc(i === 0 ? 9000 : 1200, i), key: i === 0 }), 33 * i);
    p.sendAudio(audioBody({ data: Buffer.alloc(200, i) }), 21 * i);
  }
  await new Promise((r) => setTimeout(r, 300));
  p.close();
  await new Promise((r) => setTimeout(r, 100));
  server.close();

  assert.deepEqual(states.slice(0, 1), ['connecting']);
  assert.ok(states.includes('live'));
  assert.equal(seen.clientChunkSize, 4096, 'the chunk size set first');
  assert.equal(seen.connect.app, 'live2');
  assert.equal(seen.connect.tcUrl, `rtmp://127.0.0.1:${port}/live2`);
  assert.deepEqual(seen.commands.map((c) => c.name).slice(0, 5), ['connect', 'releaseStream', 'FCPublish', 'createStream', 'publish']);
  assert.equal(seen.key, 'test-key-1234');
  assert.equal(seen.commands.find((c) => c.name === 'publish').args[1], 'live');
  assert.ok(seen.pong, 'the ping was answered');
  assert.equal(seen.data.length, 1);
  assert.equal(seen.data[0][2].width, 1280);
  assert.equal(seen.video.length, 7, 'the configuration and six frames, the early inter frame dropped');
  assert.deepEqual([...seen.video[0].body.subarray(0, 2)], [0x17, 0], 'the AVC configuration first');
  assert.equal(seen.video[1].body[0], 0x17, 'then a key frame');
  assert.equal(seen.video[1].body.length, 9005, 'whole, across chunks');
  assert.ok(seen.video.every((v) => v.streamId === 1), 'on the stream created');
  assert.deepEqual(seen.video.slice(1).map((v) => v.ts), [0, 33, 66, 99, 132, 165]);
  assert.equal(seen.audio.length, 7);
  assert.deepEqual([...seen.audio[0].body.subarray(0, 2)], [0xaf, 0], 'the AAC configuration first');
  assert.ok(seen.commands.some((c) => c.name === 'FCUnpublish'), 'unpublished when closed');
});

test('a refused key: the publisher says so and closes', async () => {
  const { server, port } = await fakeIngest({ refuse: true });
  const p = new RtmpPublisher({ url: `rtmp://127.0.0.1:${port}/live2`, key: 'wrong', timeoutMs: 5000 });
  p.on('error', () => {});
  await assert.rejects(p.start(), /refused the stream: Wrong key/);
  assert.equal(p.state, 'closed');
  server.close();
});

test('no server there: the publisher says it could not connect', async () => {
  const probe = net.createServer();
  const port = await new Promise((r) => probe.listen(0, '127.0.0.1', () => r(probe.address().port)));
  probe.close();
  const p = new RtmpPublisher({ url: `rtmp://127.0.0.1:${port}/live2`, key: 'k', timeoutMs: 3000 });
  p.on('error', () => {});
  await assert.rejects(p.start());
});
