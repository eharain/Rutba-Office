// Network cameras (packages/netcam), the reader whole: it reads a camera that
// the test itself serves on this computer's loopback (tests/netcam-server.js),
// through the questions a camera is asked, the sign-ins it may demand, the
// picture arriving down the same connection, and every way it can go wrong.
// There is no real camera here; what these prove is the reader against the
// standard as this camera speaks it.

import test from 'node:test';
import assert from 'node:assert/strict';
import net from 'node:net';
import { once } from 'node:events';
import { RtspReader, CameraError, splitAnnexB } from '../packages/netcam/src/index.js';
import { startCamera, interleave, makeSps, makePps, makeSlice } from './netcam-server.js';

const SPS = makeSps({ width: 1280, height: 720 });
const PPS = makePps();

async function until(check, ms = 3000) {
  const end = Date.now() + ms;
  while (!check()) {
    if (Date.now() > end) throw new Error('waited too long');
    await new Promise((r) => setTimeout(r, 10));
  }
}

/** A reader on a camera, with what it emitted. */
function reading(camera, options = {}) {
  const reader = new RtspReader({ url: camera.url, timeoutMs: 2000, ...options });
  const seen = { units: [], configs: [], ended: undefined };
  reader.on('unit', (u) => seen.units.push(u));
  reader.on('config', (c) => seen.configs.push(c));
  reader.on('end', (e) => {
    seen.ended = e;
  });
  return { reader, seen };
}

const fails = async (promise, code) => {
  await assert.rejects(promise, (err) => {
    assert.ok(err instanceof CameraError, `a CameraError, not ${err}`);
    assert.equal(err.code, code);
    return true;
  });
};

test('a camera is asked, set up over the one connection and read: frames arrive whole, and stopping tells the camera', async () => {
  const camera = await startCamera();
  try {
    const { reader, seen } = reading(camera);
    assert.deepEqual(await reader.start(), { codec: 'H264', sound: false });
    assert.deepEqual(camera.requests.map((r) => r.method), ['OPTIONS', 'DESCRIBE', 'SETUP', 'PLAY']);
    const [, describe, setup, play] = camera.requests;
    assert.equal(describe.headers.accept, 'application/sdp');
    assert.equal(describe.headers['user-agent'], 'Rutba Office');
    assert.equal(setup.uri, `${camera.url}/trackID=0`);
    assert.equal(setup.headers.transport, 'RTP/AVP/TCP;unicast;interleaved=0-1', 'the picture is asked for over this connection, never to a port of its own');
    assert.equal(play.uri, `${camera.url}/`);
    assert.ok(play.headers.session, 'PLAY carries the session');
    assert.ok(!describe.headers.authorization, 'no sign-in is sent to a camera that asks for none');

    const idr = makeSlice({ key: true, size: 6000, fill: 0x10 });
    const inter = makeSlice({ size: 2500, fill: 0x20 });
    camera.send([SPS, PPS, idr], { timestamp: 1000, stap: true });
    camera.send([inter], { timestamp: 4000 });
    await until(() => seen.units.length === 2);
    assert.deepEqual(seen.configs, [{ codec: 'avc1.42c01f', width: 1280, height: 720 }]);
    assert.deepEqual([seen.units[0].key, seen.units[0].timestamp, seen.units[1].key, seen.units[1].timestamp], [true, 0, false, 33333]);
    assert.deepEqual(splitAnnexB(seen.units[0].data), [SPS, PPS, idr]);
    assert.deepEqual(splitAnnexB(seen.units[1].data), [inter]);

    reader.stop();
    assert.equal(seen.ended, null, 'ended by being stopped, with no fault');
    await until(() => camera.requests.some((r) => r.method === 'TEARDOWN'));
    const bye = camera.requests.find((r) => r.method === 'TEARDOWN');
    assert.equal(bye.headers.session, play.headers.session);
    await until(() => camera.players.size === 0);
    reader.stop();
    await assert.rejects(reader.start(), /once/);
  } finally {
    await camera.close();
  }
});

test('answers and packets that arrive in small pieces, on other channels, without numbers and with whole addresses are read the same', async () => {
  const camera = await startCamera({ chop: 7, channels: 4, noCseq: true, absolute: true, base: 'none', sound: true, sprop: [SPS, PPS] });
  try {
    const { reader, seen } = reading(camera);
    assert.deepEqual(await reader.start(), { codec: 'H264', sound: true });
    assert.equal(camera.requests[2].uri, `${camera.url}/trackID=0`);
    assert.equal(camera.requests[3].uri, camera.url, 'with no base given, the address asked for is the one played');
    const idr = makeSlice({ key: true, size: 3000 });
    // A report on the channel beside the picture's, a packet on a channel never set up, then the picture.
    camera.raw(interleave(5, Buffer.from([0x80, 200, 0, 1, 0, 0, 0, 0])));
    camera.raw(interleave(9, Buffer.alloc(40, 1)));
    camera.send([idr], { timestamp: 500 });
    await until(() => seen.units.length === 1);
    assert.deepEqual(splitAnnexB(seen.units[0].data), [SPS, PPS, idr], 'the sets named in the SDP lead the key frame');
    reader.stop();
  } finally {
    await camera.close();
  }
});

test('sign-in: Digest (MD5, SHA-256, with and without qop) and Basic are answered, and Digest is chosen over Basic', async () => {
  const cases = [
    { scheme: 'digest', algorithm: 'MD5' },
    { scheme: 'digest', algorithm: 'SHA-256' },
    { scheme: 'digest', algorithm: 'MD5', qop: false },
    { scheme: 'digest', algorithm: 'MD5', also: 'basic' },
    { scheme: 'basic' },
    { scheme: 'digest', algorithm: 'MD5', options: true },
  ];
  for (const auth of cases) {
    const camera = await startCamera({ auth: { ...auth, user: 'admin', password: 'p@ss word' } });
    try {
      const { reader } = reading(camera, { user: 'admin', password: 'p@ss word' });
      await reader.start();
      const signed = camera.requests.filter((r) => r.headers.authorization);
      assert.ok(signed.length >= 3, 'DESCRIBE, SETUP and PLAY are signed');
      const want = auth.scheme === 'basic' ? /^Basic / : /^Digest /;
      for (const r of signed) assert.match(r.headers.authorization, want, JSON.stringify(auth));
      if (auth.scheme === 'digest' && auth.qop !== false) {
        assert.deepEqual(signed.slice(0, 3).map((r) => /nc=(\w+)/.exec(r.headers.authorization)[1]), ['00000001', '00000002', '00000003'], 'each answer to the same nonce is counted');
      }
      const unsigned = camera.requests.filter((r) => !r.headers.authorization);
      assert.ok(unsigned.length >= 1 && unsigned.length <= 2, 'the sign-in is sent only once the camera asks');
      reader.stop();
    } finally {
      await camera.close();
    }
  }
});

test('sign-in: none given, a wrong one, one typed inside the address, and a kind that is not known', async () => {
  const auth = { scheme: 'digest', algorithm: 'MD5', user: 'admin', password: 'right' };
  const camera = await startCamera({ auth });
  try {
    await fails(reading(camera).reader.start(), 'sign-in-needed');
    await fails(reading(camera, { user: 'admin', password: 'wrong' }).reader.start(), 'sign-in-refused');
    await fails(reading(camera, { user: 'other', password: 'right' }).reader.start(), 'sign-in-refused');
    const inside = reading(camera, { url: camera.url.replace('rtsp://', 'rtsp://admin:right@') });
    await inside.reader.start();
    assert.ok(camera.requests.every((r) => !r.uri.includes('right')), 'the password never travels in an address');
    inside.reader.stop();
  } finally {
    await camera.close();
  }
  const basic = await startCamera({ auth: { scheme: 'basic', user: 'admin', password: 'right' } });
  try {
    await fails(reading(basic, { user: 'admin', password: 'wrong' }).reader.start(), 'sign-in-refused');
    assert.equal(basic.requests.filter((r) => r.headers.authorization).length, 1, 'a refused sign-in is not sent again and again');
  } finally {
    await basic.close();
  }
  const odd = await startCamera({ auth: { scheme: 'Negotiate', user: 'a', password: 'b' } });
  try {
    await fails(reading(odd, { user: 'a', password: 'b' }).reader.start(), 'sign-in-unsupported');
  } finally {
    await odd.close();
  }
});

test('what a camera may answer instead of a picture: no such stream, H.265 only, another kind of video, no picture over this connection', async () => {
  const cases = [
    [{}, { path: '/nothing' }, 'not-found'],
    [{ codec: 'H265' }, {}, 'h265'],
    [{ codec: 'JPEG' }, {}, 'codec'],
    [{ mode: '2' }, {}, 'codec'],
    [{ setup: 461 }, {}, 'no-tcp'],
    [{ setup: 500 }, {}, 'protocol'],
  ];
  for (const [options, ask, code] of cases) {
    const camera = await startCamera(options);
    try {
      const url = ask.path ? `rtsp://127.0.0.1:${camera.port}${ask.path}` : camera.url;
      await fails(reading(camera, { url }).reader.start(), code);
    } finally {
      await camera.close();
    }
  }
});

test('an address that is not a camera: nothing there, something that never answers, something that is not RTSP, and no address at all', async () => {
  // A port nothing listens on: taken, then let go.
  const spare = net.createServer();
  await new Promise((r) => spare.listen(0, '127.0.0.1', r));
  const { port } = spare.address();
  await new Promise((r) => spare.close(r));
  await fails(new RtspReader({ url: `rtsp://127.0.0.1:${port}/x`, timeoutMs: 2000 }).start(), 'unreachable');

  const sockets = new Set();
  const silent = net.createServer((s) => {
    sockets.add(s);
    s.on('error', () => {});
  });
  await new Promise((r) => silent.listen(0, '127.0.0.1', r));
  await fails(new RtspReader({ url: `rtsp://127.0.0.1:${silent.address().port}/x`, timeoutMs: 150 }).start(), 'timeout');

  const web = net.createServer((s) => {
    sockets.add(s);
    s.on('error', () => {});
    s.once('data', () => s.write('HTTP/1.1 400 Bad Request\r\nContent-Length: 0\r\n\r\n'));
  });
  await new Promise((r) => web.listen(0, '127.0.0.1', r));
  await fails(new RtspReader({ url: `rtsp://127.0.0.1:${web.address().port}/x`, timeoutMs: 2000 }).start(), 'protocol');

  const noise = net.createServer((s) => {
    sockets.add(s);
    s.on('error', () => {});
    s.once('data', () => s.write(Buffer.from([1, 2, 3, 4, 5, 6, 7, 8])));
  });
  await new Promise((r) => noise.listen(0, '127.0.0.1', r));
  await fails(new RtspReader({ url: `rtsp://127.0.0.1:${noise.address().port}/x`, timeoutMs: 2000 }).start(), 'protocol');

  const endless = net.createServer((s) => {
    sockets.add(s);
    s.on('error', () => {});
    s.once('data', () => s.write(`RTSP/1.0 200 OK\r\nX: ${'y'.repeat(70000)}`));
  });
  await new Promise((r) => endless.listen(0, '127.0.0.1', r));
  await fails(new RtspReader({ url: `rtsp://127.0.0.1:${endless.address().port}/x`, timeoutMs: 2000 }).start(), 'protocol');

  const hangsUp = net.createServer((s) => {
    sockets.add(s);
    s.on('error', () => {});
    s.once('data', () => s.destroy());
  });
  await new Promise((r) => hangsUp.listen(0, '127.0.0.1', r));
  await fails(new RtspReader({ url: `rtsp://127.0.0.1:${hangsUp.address().port}/x`, timeoutMs: 2000 }).start(), 'closed');

  await fails(new RtspReader({ url: 'http://127.0.0.1/x' }).start(), 'address');
  await fails(new RtspReader({ url: '' }).start(), 'address');
  await fails(
    new RtspReader({
      url: 'rtsp://127.0.0.1/x',
      connect: () => {
        throw Object.assign(new Error('no route'), { code: 'ENETUNREACH' });
      },
    }).start(),
    'unreachable',
  );

  for (const s of sockets) s.destroy();
  for (const server of [silent, web, noise, endless, hangsUp]) await new Promise((r) => server.close(r));
});

test('a camera that moves its stream is followed on the same machine only', async () => {
  const camera = await startCamera({ moved: { from: '/old', to: '/stream1' } });
  try {
    const { reader } = reading(camera, { url: `rtsp://127.0.0.1:${camera.port}/old` });
    await reader.start();
    assert.deepEqual(camera.requests.filter((r) => r.method === 'DESCRIBE').map((r) => r.uri), [`rtsp://127.0.0.1:${camera.port}/old`, camera.url]);
    reader.stop();
  } finally {
    await camera.close();
  }
  const away = await startCamera({ moved: { from: '/stream1', to: 'rtsp://elsewhere.example/stream1' }, auth: { scheme: 'basic', user: 'a', password: 'b' } });
  try {
    await fails(reading(away, { user: 'a', password: 'b' }).reader.start(), 'redirect');
  } finally {
    await away.close();
  }
  const round = await startCamera({ moved: { from: '/stream1', to: '/stream1' } });
  try {
    await fails(reading(round).reader.start(), 'redirect');
    assert.equal(round.requests.filter((r) => r.method === 'DESCRIBE').length, 4, 'a camera that sends the reader round in circles is left after a few turns');
  } finally {
    await round.close();
  }
});

test('while reading: the session is kept awake, the camera’s own questions are answered, and a picture that stops or a camera that goes ends the reading', async () => {
  const camera = await startCamera({ timeout: 2, sprop: [SPS, PPS] });
  try {
    const { reader, seen } = reading(camera, { quietMs: 60000 });
    await reader.start();
    camera.raw('OPTIONS * RTSP/1.0\r\nCSeq: 7\r\n\r\n');
    camera.raw('ANNOUNCE * RTSP/1.0\r\nCSeq: 8\r\n\r\n');
    await until(() => camera.requests.filter((r) => r.method === '').length === 2);
    assert.deepEqual(camera.requests.filter((r) => r.method === '').map((r) => r.line), ['RTSP/1.0 200 OK', 'RTSP/1.0 405 Method Not Allowed']);
    assert.deepEqual(camera.requests.filter((r) => r.method === '').map((r) => r.headers.cseq), ['7', '8']);
    await until(() => camera.requests.some((r) => r.method === 'GET_PARAMETER'), 4000);
    const keep = camera.requests.find((r) => r.method === 'GET_PARAMETER');
    assert.ok(keep.headers.session, 'the question that keeps the session awake names it');
    assert.equal(seen.ended, undefined);
    camera.cut();
    await until(() => seen.ended !== undefined);
    assert.equal(seen.ended.code, 'closed');
  } finally {
    await camera.close();
  }

  const plain = await startCamera({ timeout: 2, publicLine: 'OPTIONS, DESCRIBE, SETUP, PLAY, TEARDOWN', sprop: [SPS, PPS] });
  try {
    const { reader, seen } = reading(plain, { quietMs: 60000 });
    await reader.start();
    await until(() => plain.requests.filter((r) => r.method === 'OPTIONS').length >= 2, 4000);
    assert.ok(plain.requests.filter((r) => r.method === 'OPTIONS').at(-1).headers.session, 'a camera with no GET_PARAMETER is kept awake by OPTIONS');
    reader.stop();
    assert.equal(seen.ended, null);
  } finally {
    await plain.close();
  }

  const quiet = await startCamera({ sprop: [SPS, PPS] });
  try {
    const { reader, seen } = reading(quiet, { quietMs: 250 });
    const [, ended] = await Promise.all([reader.start(), once(reader, 'end')]);
    assert.equal(ended[0].code, 'stalled');
    assert.equal(seen.units.length, 0);
    await until(() => quiet.players.size === 0);
  } finally {
    await quiet.close();
  }
});

test('stopping while the camera is still being asked ends the asking', async () => {
  const camera = await startCamera();
  try {
    const { reader, seen } = reading(camera);
    camera.onRequest = (req) => {
      if (req.method === 'DESCRIBE') reader.stop();
    };
    await fails(reader.start(), 'stopped');
    assert.equal(seen.ended, undefined, 'a reading that never began has no end to tell');
    assert.ok(!camera.requests.some((r) => r.method === 'SETUP'));
  } finally {
    await camera.close();
  }
});
