// Network cameras (packages/netcam), the parts: a typed address taken apart
// with its sign-in lifted out, the sign-ins cameras ask for (checked against
// the examples in the standards), what a camera says it sends (SDP), one RTP
// packet, and H.264 put back together from packets into frames a decoder
// takes. No network.

import test from 'node:test';
import assert from 'node:assert/strict';
import { parseCameraAddress, cameraNameFrom, sameCamera, parseChallenges, chooseChallenge, authorization, parseSdp, controlUrl, spropSets, parseRtp, H264Assembler, spsInfo, annexB, splitAnnexB } from '../packages/netcam/src/index.js';
import { packetise, makeSps, makePps, makeSlice } from './netcam-server.js';

test('a camera address: the sign-in typed inside it is lifted out and never kept in the address', () => {
  const a = parseCameraAddress('rtsp://admin:s3cr%40t@192.168.1.64:554/Streaming/Channels/101?transportmode=unicast');
  assert.equal(a.url, 'rtsp://192.168.1.64/Streaming/Channels/101?transportmode=unicast');
  assert.equal(a.user, 'admin');
  assert.equal(a.password, 's3cr@t');
  assert.equal(a.host, '192.168.1.64');
  assert.equal(a.port, 554);
  assert.equal(a.secure, false);
  assert.ok(!a.url.includes('s3cr'), 'no password in the kept address');

  assert.equal(parseCameraAddress('  10.0.0.9/stream1 ').url, 'rtsp://10.0.0.9/stream1', 'rtsp:// is assumed');
  assert.equal(parseCameraAddress('rtsp://cam.example:8554/live').port, 8554);
  assert.equal(parseCameraAddress('rtsp://cam.example:8554/live').url, 'rtsp://cam.example:8554/live');
  const s = parseCameraAddress('rtsps://cam.example/live');
  assert.deepEqual([s.secure, s.port, s.url], [true, 322, 'rtsps://cam.example/live']);
  const six = parseCameraAddress('rtsp://[fe80::1]:554/a');
  assert.equal(six.host, 'fe80::1');
  assert.equal(six.url, 'rtsp://[fe80::1]/a');
  assert.equal(parseCameraAddress('rtsp://cam.example').url, 'rtsp://cam.example/', 'no path is the root');

  for (const bad of ['', '   ', 'http://cam.example/live', 'rtmp://cam.example/live', 'rtsp://', 'rtsp://cam example/live', 'rtsp://cam.example:99999/x', `rtsp://cam.example/${'x'.repeat(2100)}`, null, undefined]) {
    assert.equal(parseCameraAddress(bad), null, `not an address: ${String(bad).slice(0, 30)}`);
  }
  assert.equal(cameraNameFrom('rtsp://u:p@porch.example/1'), 'porch.example');
  assert.ok(sameCamera('rtsp://a:b@10.0.0.9:554/s', '10.0.0.9/s'));
  assert.ok(!sameCamera('rtsp://10.0.0.9/s', 'rtsp://10.0.0.9/t'));
});

test('sign-in: Digest answers match the examples in RFC 2617, RFC 7616 and RFC 2069', () => {
  // RFC 2617 section 3.5.
  const [old] = parseChallenges('Digest realm="testrealm@host.com", qop="auth,auth-int", nonce="dcd98b7102dd2f0e8b11d0f600bfb0c093", opaque="5ccc069c403ebaf9f0171e9517f40e41"');
  assert.deepEqual([old.scheme, old.qop, old.algorithm, old.stale], ['digest', 'auth', 'MD5', false]);
  const line = authorization({ challenge: old, user: 'Mufasa', password: 'Circle Of Life', method: 'GET', uri: '/dir/index.html', count: 1, cnonce: '0a4f113b' });
  assert.match(line, /^Digest username="Mufasa", realm="testrealm@host\.com", nonce="dcd98b7102dd2f0e8b11d0f600bfb0c093", uri="\/dir\/index\.html", response="6629fae49393a05397450978507c4ef1", qop=auth, nc=00000001, cnonce="0a4f113b", opaque="5ccc069c403ebaf9f0171e9517f40e41"$/);

  // RFC 7616 section 3.9.1, both algorithms.
  const both = parseChallenges([
    'Digest realm="http-auth@example.org", qop="auth, auth-int", algorithm=SHA-256, nonce="7ypf/xlj9XXwfDPEoM4URrv/xwf94BcCAzFZH4GiTo0v", opaque="FQhe/qaU925kfnzjCev0ciny7QMkPqMAFRtzCUYo5tdS"',
    'Digest realm="http-auth@example.org", qop="auth, auth-int", algorithm=MD5, nonce="7ypf/xlj9XXwfDPEoM4URrv/xwf94BcCAzFZH4GiTo0v", opaque="FQhe/qaU925kfnzjCev0ciny7QMkPqMAFRtzCUYo5tdS"',
  ]);
  assert.equal(both.length, 2);
  const cnonce = 'f2/wE4q74E6zIJEtWaHKaf5wv/H5QzzpXusqGemxURZJ';
  const sha = authorization({ challenge: both[0], user: 'Mufasa', password: 'Circle of Life', method: 'GET', uri: '/dir/index.html', count: 1, cnonce });
  assert.ok(sha.includes('response="753927fa0e85d155564e2e272a28d1802ca10daf4496794697cf8db5856cb6c1"'), sha);
  assert.ok(sha.includes('algorithm=SHA-256'));
  const md = authorization({ challenge: both[1], user: 'Mufasa', password: 'Circle of Life', method: 'GET', uri: '/dir/index.html', count: 1, cnonce });
  assert.ok(md.includes('response="8ca523f5e9506fed4657c9700eebdbec"'), md);
  assert.equal(chooseChallenge([both[1], both[0]]).algorithm, 'SHA-256', 'the stronger one is answered');

  // RFC 2069 (no qop), as corrected by its erratum.
  const [plain] = parseChallenges('Digest realm="testrealm@host.com", nonce="dcd98b7102dd2f0e8b11d0f600bfb0c093", opaque="5ccc069c403ebaf9f0171e9517f40e41"');
  assert.equal(plain.qop, '');
  const oldest = authorization({ challenge: plain, user: 'Mufasa', password: 'CircleOfLife', method: 'GET', uri: '/dir/index.html' });
  assert.ok(oldest.includes('response="1949323746fe6a43ef61f9606e7febea"'), oldest);
  assert.ok(!oldest.includes('qop'), 'no qop is answered without one');
});

test('sign-in: Basic only when nothing stronger is offered, and what cannot be answered is not', () => {
  const offered = parseChallenges(['Basic realm="cam"', 'Digest realm="cam", nonce="abc"']);
  assert.equal(chooseChallenge(offered).scheme, 'digest');
  const basic = chooseChallenge(parseChallenges('Basic realm="cam"'));
  assert.equal(authorization({ challenge: basic, user: 'admin', password: 'pä:ss', method: 'DESCRIBE', uri: 'rtsp://x/' }), `Basic ${Buffer.from('admin:pä:ss', 'utf8').toString('base64')}`);
  assert.equal(chooseChallenge(parseChallenges('Negotiate')), null);
  assert.equal(chooseChallenge(parseChallenges('Digest realm="cam", nonce="n", algorithm=SHA-512-256')), null, 'an algorithm that is not known');
  assert.equal(chooseChallenge(parseChallenges('Digest realm="cam", nonce="n", qop="auth-int"')), null, 'only a kind of answer that is not given');
  assert.equal(chooseChallenge(parseChallenges('Digest realm="cam"')), null, 'no nonce');
  assert.equal(chooseChallenge([]), null);
  assert.equal(parseChallenges('Digest realm="cam", nonce="n", stale=TRUE')[0].stale, true);
  const quoted = authorization({ challenge: parseChallenges('Digest realm="a\\"b", nonce="n"')[0], user: 'u"x', password: 'p', method: 'PLAY', uri: 'rtsp://x/' });
  assert.ok(quoted.includes('username="u\\"x"') && quoted.includes('realm="a\\"b"'), 'quotes inside a value are escaped');
  const sess = authorization({ challenge: parseChallenges('Digest realm="cam", nonce="n", qop="auth", algorithm=MD5-sess')[0], user: 'u', password: 'p', method: 'PLAY', uri: 'rtsp://x/', count: 11, cnonce: 'c' });
  assert.ok(sess.includes('nc=0000000b') && sess.includes('algorithm=MD5-SESS'));
});

test('SDP: the streams a camera offers, their codecs, parameter sets and the addresses they are set up by', () => {
  const sps = makeSps({ width: 1280, height: 720 });
  const pps = makePps();
  const sdp = parseSdp(
    [
      'v=0',
      'o=- 1 1 IN IP4 10.0.0.9',
      's=Media Presentation',
      'a=control:rtsp://10.0.0.9/Streaming/Channels/101/',
      'm=video 0 RTP/AVP 96',
      'a=rtpmap:96 H264/90000',
      `a=fmtp:96 profile-level-id=420029; packetization-mode=1; sprop-parameter-sets=${sps.toString('base64')},${pps.toString('base64')}`,
      'a=control:trackID=1',
      'm=audio 0 RTP/AVP 8',
      'a=control:trackID=2',
      'm=audio 0 RTP/AVP 104',
      'a=rtpmap:104 mpeg4-generic/16000/2',
      'a=control:rtsp://10.0.0.9/Streaming/Channels/101/trackID=3',
      'm=application 0 RTP/AVP 107',
      'a=rtpmap:107 vnd.onvif.metadata/90000',
      'a=control:trackID=4',
    ].join('\r\n'),
  );
  assert.equal(sdp.control, 'rtsp://10.0.0.9/Streaming/Channels/101/');
  assert.equal(sdp.tracks.length, 3, 'the metadata stream is not a picture or a sound');
  const [video, alaw, aac] = sdp.tracks;
  assert.deepEqual([video.kind, video.codec, video.clock, video.payload, video.control, video.transport], ['video', 'H264', 90000, 96, 'trackID=1', 'RTP/AVP']);
  assert.equal(video.fmtp['packetization-mode'], '1');
  assert.deepEqual(spropSets(video), [sps, pps]);
  assert.deepEqual([alaw.codec, alaw.clock, alaw.channels], ['PCMA', 8000, 1], 'a fixed payload number needs no rtpmap');
  assert.deepEqual([aac.codec, aac.clock, aac.channels], ['MPEG4-GENERIC', 16000, 2]);
  assert.deepEqual(spropSets(alaw), []);
  assert.deepEqual(spropSets({ fmtp: { 'sprop-parameter-sets': 'not base64!,' } }), []);
  assert.deepEqual(parseSdp('').tracks, []);
  assert.deepEqual(parseSdp('m=video x\r\nm=video 0 RTP/AVP 300\r\ngarbage').tracks, []);

  const base = 'rtsp://10.0.0.9/live';
  assert.equal(controlUrl(base, 'trackID=1'), 'rtsp://10.0.0.9/live/trackID=1');
  assert.equal(controlUrl(`${base}/`, 'trackID=1'), 'rtsp://10.0.0.9/live/trackID=1');
  assert.equal(controlUrl(base, '*'), base);
  assert.equal(controlUrl(base, ''), base);
  assert.equal(controlUrl(base, 'rtsp://10.0.0.9/other/track1'), 'rtsp://10.0.0.9/other/track1');
  assert.equal(controlUrl(base, '/abs/track1'), 'rtsp://10.0.0.9/abs/track1');
});

test('RTP: a packet is its fields and its payload, with extensions, sources and padding stepped over', () => {
  const [plain] = packetise([Buffer.from([0x41, 1, 2, 3])], { timestamp: 0xfffffff0, sequence: 65535 });
  const p = parseRtp(plain);
  assert.deepEqual([p.payloadType, p.marker, p.sequence, p.timestamp, p.ssrc], [96, true, 65535, 0xfffffff0, 0x1234abcd]);
  assert.deepEqual([...p.payload], [0x41, 1, 2, 3]);

  const full = Buffer.concat([
    Buffer.from([0x80 | 0x20 | 0x10 | 1, 96, 0, 7, 0, 0, 0, 9, 0, 0, 0, 1]),
    Buffer.from([0, 0, 0, 5]), // one contributing source
    Buffer.from([0xbe, 0xde, 0, 1, 9, 9, 9, 9]), // an extension of one word
    Buffer.from([0x65, 0xaa]),
    Buffer.from([0, 0, 3]), // three bytes of padding
  ]);
  const q = parseRtp(full);
  assert.deepEqual([q.marker, q.sequence, q.timestamp], [false, 7, 9]);
  assert.deepEqual([...q.payload], [0x65, 0xaa]);

  assert.equal(parseRtp(Buffer.alloc(11)), null, 'too short');
  assert.equal(parseRtp(Buffer.from([0x40, 96, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1])), null, 'not version 2');
  assert.equal(parseRtp(Buffer.from([0x90, 96, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1])), null, 'an extension that is not there');
  assert.equal(parseRtp(Buffer.from([0xa0, 96, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 200])), null, 'more padding than packet');
  assert.equal(parseRtp(null), null);
});

test('H.264 parameter sets: the codec name and the picture size, from sets written here and from one a real encoder wrote', () => {
  for (const [width, height] of [[1280, 720], [1920, 1080], [640, 360], [704, 576], [2560, 1440], [3840, 2160]]) {
    const info = spsInfo(makeSps({ width, height }));
    assert.deepEqual([info.width, info.height], [width, height], `${width} by ${height}`);
    assert.equal(info.codec, 'avc1.42c01f');
    const high = spsInfo(makeSps({ width, height, profile: 100, level: 40, constraints: 0 }));
    assert.deepEqual([high.codec, high.width, high.height], ['avc1.640028', width, height], `high profile, ${width} by ${height}`);
  }
  // An x264 High profile set for 1920 by 1080, with the escape bytes its timing information needs.
  const real = Buffer.from('6764002aacd9407802 27e5c04400000300040000030 0f03c60c658'.replace(/ /g, ''), 'hex');
  assert.deepEqual(spsInfo(real), { codec: 'avc1.64002a', width: 1920, height: 1080 });
  assert.equal(spsInfo(Buffer.from([0x68, 1, 2, 3])), null, 'a picture parameter set is not one');
  assert.equal(spsInfo(Buffer.from([0x67, 1])), null);
  assert.deepEqual(spsInfo(Buffer.from([0x67, 0x42, 0xc0, 0x1f])), { codec: 'avc1.42c01f', width: 0, height: 0 }, 'cut short: the name alone');
  assert.deepEqual(spsInfo(Buffer.from([0x67, 0x42, 0xc0, 0x1f, 0, 0, 0, 0, 0, 0, 0, 0])), { codec: 'avc1.42c01f', width: 0, height: 0 }, 'nonsense: the name alone');
});

test('Annex B: units are joined after start codes and taken apart again', () => {
  const nals = [Buffer.from([0x67, 1, 2]), Buffer.from([0x68, 3]), Buffer.from([0x65, 4, 5, 6])];
  const joined = annexB(nals);
  assert.deepEqual([...joined.subarray(0, 5)], [0, 0, 0, 1, 0x67]);
  assert.deepEqual(splitAnnexB(joined), nals);
  assert.deepEqual(splitAnnexB(Buffer.from([0, 0, 1, 0x67, 9, 0, 0, 0, 1, 0x68, 8])), [Buffer.from([0x67, 9]), Buffer.from([0x68, 8])], 'short and long start codes');
  assert.deepEqual(splitAnnexB(Buffer.from([1, 2, 3])), []);
});

const feed = (asm, packets) => packets.flatMap((p) => asm.push(parseRtp(p)));

test('H.264 from packets: single units, a unit cut across packets and units sharing a packet come back as whole frames', () => {
  const sps = makeSps({ width: 1280, height: 720 });
  const pps = makePps();
  const idr = makeSlice({ key: true, size: 5000, fill: 0x11 });
  const inter = makeSlice({ size: 3000, fill: 0x22 });
  const small = makeSlice({ size: 60, fill: 0x33 });
  const asm = new H264Assembler();
  let sequence = 10;
  const send = (nals, opts) => {
    const packets = packetise(nals, { sequence, ...opts });
    sequence += packets.length;
    return feed(asm, packets);
  };

  const keyPackets = packetise([sps, pps, idr], { timestamp: 9000, sequence, stap: true });
  sequence += keyPackets.length;
  assert.ok(keyPackets.length > 4, 'the key frame is cut across packets');
  assert.equal(feed(asm, keyPackets.slice(0, -1)).length, 0, 'nothing until the frame is whole');
  const [first] = feed(asm, keyPackets.slice(-1));
  assert.equal(first.key, true);
  assert.equal(first.timestamp, 0, 'time starts at the first frame');
  assert.deepEqual(first.config, { codec: 'avc1.42c01f', width: 1280, height: 720 });
  assert.deepEqual(splitAnnexB(first.data), [sps, pps, idr], 'the units, whole and in order');

  const [second] = send([inter], { timestamp: 9000 + 3000 });
  assert.deepEqual([second.key, second.timestamp, second.config], [false, 33333, undefined]);
  assert.deepEqual(splitAnnexB(second.data), [inter]);

  const [third] = send([small], { timestamp: 9000 + 6000 });
  assert.deepEqual(splitAnnexB(third.data), [small], 'a small unit travels in one packet');
  assert.equal(third.timestamp, 66667);
});

test('H.264 from packets: nothing before the first key frame, and parameter sets from the SDP or an earlier packet lead every key frame', () => {
  const sps = makeSps({ width: 640, height: 360 });
  const pps = makePps();
  const idr = makeSlice({ key: true });
  const inter = makeSlice();

  // Sets named in the SDP only.
  const fromSdp = new H264Assembler({ sets: [sps, pps] });
  assert.equal(feed(fromSdp, packetise([inter], { timestamp: 0, sequence: 1 })).length, 0, 'a frame that leans on one never seen is not handed on');
  const [key] = feed(fromSdp, packetise([idr], { timestamp: 3000, sequence: 2 }));
  assert.deepEqual(splitAnnexB(key.data), [sps, pps, idr]);
  assert.equal(key.config.width, 640);

  // Sets sent in the stream by themselves, with their own time and marker, before the key frame.
  const inBand = new H264Assembler();
  assert.equal(feed(inBand, packetise([sps], { timestamp: 0, sequence: 1 })).length, 0);
  assert.equal(feed(inBand, packetise([pps], { timestamp: 0, sequence: 2 })).length, 0);
  const [later] = feed(inBand, packetise([idr], { timestamp: 10, sequence: 3 }));
  assert.deepEqual(splitAnnexB(later.data), [sps, pps, idr]);

  // No sets anywhere: the key frame cannot be decoded, so it is not handed on.
  const none = new H264Assembler();
  assert.equal(feed(none, packetise([idr], { timestamp: 0, sequence: 1 })).length, 0);
  assert.equal(feed(none, packetise([inter], { timestamp: 3000, sequence: 2 })).length, 0);

  // A change of picture size is told once, on the key frame that brings it.
  const bigger = makeSps({ width: 1920, height: 1080 });
  const [again] = feed(fromSdp, packetise([idr], { timestamp: 6000, sequence: 3 }));
  assert.equal(again.config, undefined, 'the same parameters are not told twice');
  const [changed] = feed(fromSdp, packetise([bigger, pps, idr], { timestamp: 9000, sequence: 4 }));
  assert.deepEqual([changed.config.width, changed.config.height], [1920, 1080]);
});

test('H.264 from packets: a frame ends at the marker or at the next time, delimiters are dropped, and frames keep their order in time across the clock wrapping', () => {
  const sets = [makeSps({ width: 640, height: 360 }), makePps()];
  const asm = new H264Assembler({ sets });
  const idr = makeSlice({ key: true });
  const a = makeSlice({ fill: 1 });
  const b = makeSlice({ fill: 2 });
  const delimiter = Buffer.from([0x09, 0xf0]);
  const sei = Buffer.from([0x06, 5, 1, 0x80]);
  const near = 0xffffffff - 1500;
  let sequence = 65530;
  const packets = (nals, timestamp, extra = {}) => {
    const made = packetise(nals, { timestamp, sequence, ...extra });
    sequence = (sequence + made.length) & 0xffff;
    return made;
  };

  const [k] = feed(asm, packets([delimiter, sei, idr], near));
  assert.deepEqual(splitAnnexB(k.data), [...sets, sei, idr], 'the delimiter is dropped and the sets lead');

  // No marker on this camera: the frame is known to be whole when the next one starts.
  const noMarker = packets([a], (near + 3000) >>> 0);
  noMarker[0][1] &= 0x7f;
  assert.equal(feed(asm, noMarker).length, 0);
  const out = feed(asm, packets([b], (near + 6000) >>> 0));
  assert.equal(out.length, 2, 'the frame before and this one');
  assert.deepEqual(out.map((u) => u.timestamp), [33333, 66667], 'time runs on through the 32-bit wrap and the sequence wrap');
  assert.deepEqual(splitAnnexB(out[0].data), [a]);
  assert.deepEqual(splitAnnexB(out[1].data), [b]);
});

test('H.264 from packets: after a lost packet nothing is handed on until the next key frame', () => {
  const sets = [makeSps({ width: 640, height: 360 }), makePps()];
  const asm = new H264Assembler({ sets });
  const idr = makeSlice({ key: true, size: 4000 });
  const inter = makeSlice({ size: 4000 });
  assert.equal(feed(asm, packetise([idr], { timestamp: 0, sequence: 1 })).length, 1);
  const cut = packetise([inter], { timestamp: 3000, sequence: 5 });
  assert.ok(cut.length >= 4);
  cut.splice(1, 1); // one piece of the frame never arrives
  assert.equal(feed(asm, cut).length, 0, 'the broken frame');
  assert.equal(feed(asm, packetise([inter], { timestamp: 6000, sequence: 5 + cut.length + 1 })).length, 0, 'a whole frame that leans on the broken one');
  const [key] = feed(asm, packetise([idr], { timestamp: 9000, sequence: 5 + cut.length + 1 + 4 }));
  assert.equal(key.key, true, 'the next key frame starts the picture again');
  assert.equal(key.timestamp, 100000);

  // The same packet twice is taken once.
  const twice = packetise([makeSlice({ size: 30 })], { timestamp: 12000, sequence: 5 + cut.length + 1 + 4 + 4 });
  assert.equal(feed(asm, [twice[0], twice[0]]).length, 1);
});

test('H.264 from packets: packet kinds that are not read and broken packets are passed over without harm', () => {
  const asm = new H264Assembler({ sets: [makeSps({ width: 640, height: 360 }), makePps()] });
  const rtp = (payload, sequence, timestamp = 0, marker = false) => {
    const [p] = packetise([Buffer.from(payload)], { timestamp, sequence });
    if (!marker) p[1] &= 0x7f;
    return p;
  };
  const out = feed(asm, [
    rtp([25, 0, 0, 0, 2, 0x65, 1], 1), // STAP-B
    rtp([29, 0x85, 0, 0, 1, 2], 2), // FU-B
    rtp([24, 0, 9, 0x65], 3), // a STAP-A that claims more than it holds
    rtp([24, 0, 0], 4), // a STAP-A with an empty unit
    rtp([28], 5), // an FU-A with no header
    rtp([28, 0x45, 1, 2], 6), // the end of a unit whose start was never seen
    rtp([0x65, 7, 7, 7], 7, 0, true),
  ]);
  assert.equal(out.length, 1);
  assert.deepEqual(splitAnnexB(out[0].data).at(-1), Buffer.from([0x65, 7, 7, 7]));
  assert.deepEqual(asm.push(null), []);
  assert.deepEqual(asm.push({ payload: Buffer.alloc(0), sequence: 8, timestamp: 0 }), []);
});
