// The door of a Talk Server as Office reaches it (apps/desktop/main/talk-door.js
// and talk-frames.js): what is checked before a pass or a join token goes
// anywhere, and the signalling connection made by hand on the checked
// connection. Against a stand-in door on this computer (talk-door-stand-in.js).

import test from 'node:test';
import assert from 'node:assert/strict';
import tls from 'node:tls';
import { OP, acceptKey, closePayload, encodeFrame, frameReader } from '../apps/desktop/main/talk-frames.js';
import { DoorError, askDoor, connectDoor, exchange, isLoopback, openSocket, readDoor } from '../apps/desktop/main/talk-door.js';
import { selfSigned, standInDoor } from './talk-door-stand-in.js';

const doors = [];
test.after(async () => { for (const d of doors) await d.close(); });
const door = async (o) => { const d = await standInDoor(o); doors.push(d); return d; };
const until = async (fn, ms = 4000) => { const t = Date.now(); while (!fn()) { if (Date.now() - t > ms) throw new Error('waited too long'); await new Promise((r) => setTimeout(r, 5)); } };
const code = async (p) => { try { await p; } catch (err) { assert.ok(err instanceof DoorError, String(err)); return err.code; } return null; };

// ── the frames ───────────────────────────────────────────────────────────────

test('a frame written is the frame read, at every length that changes its head, masked and not, in pieces of any size', () => {
  for (const mask of [true, false]) {
    for (const n of [0, 1, 125, 126, 127, 65535, 65536, 70000]) {
      const sent = Buffer.alloc(n);
      for (let i = 0; i < n; i += 1) sent[i] = (i * 7 + n) & 255;
      const frame = encodeFrame(OP.binary, sent, { mask });
      for (const piece of [frame.length, 1, 3, 1000]) {
        if (piece === 1 && n > 1000) continue;
        const got = [];
        const r = frameReader({ onMessage: (data, binary) => got.push({ data, binary }), onError: (why) => assert.fail(why) });
        for (let at = 0; at < frame.length; at += piece) r.push(frame.subarray(at, at + piece));
        assert.equal(got.length, 1, `${n} bytes in pieces of ${piece}`);
        assert.equal(got[0].binary, true);
        assert.ok(got[0].data.equals(sent), `${n} bytes, mask ${mask}`);
      }
    }
  }
});

test('text is told as text, a message in parts is joined, and a ping between the parts is answered in its place', () => {
  const got = [];
  const pings = [];
  const r = frameReader({ onMessage: (data, binary) => got.push(binary ? data : data.toString('utf8')), onPing: (d) => pings.push(d.toString()), onError: (why) => assert.fail(why) });
  r.push(Buffer.concat([
    encodeFrame(OP.text, 'salām', { mask: false }),
    encodeFrame(OP.text, 'one ', { mask: false, fin: false }),
    encodeFrame(OP.ping, 'are you there', { mask: false }),
    encodeFrame(OP.more, 'two ', { mask: false, fin: false }),
    encodeFrame(OP.more, 'three', { mask: false }),
  ]));
  assert.deepEqual(got, ['salām', 'one two three']);
  assert.deepEqual(pings, ['are you there']);
});

test('a close is told with its code and reason, and nothing after it is read', () => {
  const seen = [];
  const r = frameReader({ onMessage: () => seen.push('message'), onClose: (c, reason) => seen.push([c, reason]) });
  r.push(Buffer.concat([encodeFrame(OP.close, closePayload(4001, 'over'), { mask: false }), encodeFrame(OP.text, 'late', { mask: false })]));
  r.push(encodeFrame(OP.text, 'later', { mask: false }));
  assert.deepEqual(seen, [[4001, 'over']]);
  const bare = [];
  frameReader({ onMessage: () => {}, onClose: (c, reason) => bare.push([c, reason]) }).push(encodeFrame(OP.close, Buffer.alloc(0), { mask: false }));
  assert.deepEqual(bare, [[1005, '']]);
});

test('what is not a WebSocket\'s is refused once, and nothing more is read', () => {
  const why = (bytes, o = {}) => { const errs = []; const r = frameReader({ onMessage: () => errs.push('message'), onError: (w) => errs.push(w), ...o }); r.push(bytes); r.push(encodeFrame(OP.text, 'more', { mask: false })); return errs; };
  const reserved = encodeFrame(OP.text, 'x', { mask: false });
  reserved[0] |= 0x40;
  assert.deepEqual(why(reserved), ['extension'], 'a compressed frame, which was never offered');
  assert.deepEqual(why(encodeFrame(OP.more, 'x', { mask: false })), ['continuation']);
  assert.deepEqual(why(Buffer.concat([encodeFrame(OP.text, 'a', { mask: false, fin: false }), encodeFrame(OP.text, 'b', { mask: false })])), ['interleaved']);
  assert.deepEqual(why(encodeFrame(OP.ping, Buffer.alloc(126), { mask: false })), ['control']);
  assert.deepEqual(why(encodeFrame(OP.ping, 'x', { mask: false, fin: false })), ['control']);
  assert.deepEqual(why(encodeFrame(3, 'x', { mask: false })), ['opcode']);
  assert.deepEqual(why(encodeFrame(OP.binary, Buffer.alloc(2000), { mask: false }), { most: 1000 }), ['too-large']);
  // Too large in its parts, each of them small enough.
  assert.deepEqual(why(Buffer.concat([encodeFrame(OP.binary, Buffer.alloc(600), { mask: false, fin: false }), encodeFrame(OP.more, Buffer.alloc(600), { mask: false })]), { most: 1000 }), ['too-large']);
});

test('the answer to a key is the one RFC 6455 gives', () => {
  assert.equal(acceptKey('dGhlIHNhbXBsZSBub25jZQ=='), 's3pPLMBiTxaQ9kYGzzhZRbK+xOo=');
});

test('the frames are the ones RFC 6455 writes out, so the reader and the writer do not only agree with each other', () => {
  // Section 5.7: "Hello" unmasked, masked, and in two parts; then the heads of 256 and 65536 bytes.
  assert.ok(encodeFrame(OP.text, 'Hello', { mask: false }).equals(Buffer.from([0x81, 0x05, 0x48, 0x65, 0x6c, 0x6c, 0x6f])));
  const got = [];
  const reader = () => frameReader({ onMessage: (data, binary) => got.push([data.toString('utf8'), binary]), onError: (why) => assert.fail(why) });
  reader().push(Buffer.from([0x81, 0x85, 0x37, 0xfa, 0x21, 0x3d, 0x7f, 0x9f, 0x4d, 0x51, 0x58]));
  const parts = reader();
  parts.push(Buffer.from([0x01, 0x03, 0x48, 0x65, 0x6c]));
  parts.push(Buffer.from([0x80, 0x02, 0x6c, 0x6f]));
  assert.deepEqual(got, [['Hello', false], ['Hello', false]]);
  assert.equal(encodeFrame(OP.binary, Buffer.alloc(256), { mask: false }).subarray(0, 4).toString('hex'), '827e0100');
  assert.equal(encodeFrame(OP.binary, Buffer.alloc(65536), { mask: false }).subarray(0, 10).toString('hex'), '827f0000000000010000');
});

// ── a door as the introducer hands it out ────────────────────────────────────

const PIN = 'ab'.repeat(32);

test('a door is read strictly: where to connect, and what must be seen there', () => {
  assert.deepEqual(readDoor({ family: 4, url: 'wss://203.0.113.7:7443', address: '203.0.113.7', media: 'udp' }, PIN), { secure: true, address: '203.0.113.7', port: 7443, family: 4, name: null, pin: PIN, media: 'udp', host: '203.0.113.7:7443' });
  const named = readDoor({ family: 6, url: 'wss://talk.example.org', address: '2001:db8::7', name: 'Talk.Example.org', media: 'tcp' });
  assert.deepEqual(named, { secure: true, address: '2001:db8::7', port: 443, family: 6, name: 'talk.example.org', pin: null, media: 'tcp', host: 'talk.example.org:443' });
  assert.equal(readDoor({ url: 'wss://[2001:db8::7]:7443', address: '2001:db8::7' }, PIN.toUpperCase()).host, '[2001:db8::7]:7443');
  // The address is where Office connects, whatever the url's own host says.
  assert.equal(readDoor({ url: 'wss://talk.example.org:7443', address: '203.0.113.9', name: 'talk.example.org' }).address, '203.0.113.9');
});

test('a door that is not one is refused, not guessed at', () => {
  const bad = (d, pin) => { try { readDoor(d, pin); } catch (err) { return err.code; } return null; };
  assert.equal(bad({ url: 'https://203.0.113.7:7443', address: '203.0.113.7' }, PIN), 'door');
  assert.equal(bad({ url: 'wss://203.0.113.7:7443/elsewhere', address: '203.0.113.7' }, PIN), 'door');
  assert.equal(bad({ url: 'wss://203.0.113.7:7443?x=1', address: '203.0.113.7' }, PIN), 'door');
  assert.equal(bad({ url: 'wss://user:word@203.0.113.7:7443', address: '203.0.113.7' }, PIN), 'door');
  assert.equal(bad({ url: 'wss://talk.example.org', address: 'talk.example.org' }), 'door', 'an address that is a name would be looked up afresh');
  assert.equal(bad({ url: 'wss://203.0.113.7:7443', address: '203.0.113.7', name: '203.0.113.7' }), 'door');
  assert.equal(bad({ url: 'wss://203.0.113.7:7443', address: '203.0.113.7' }, 'abc'), 'door', 'a pin that is not one');
  assert.equal(bad({ url: 'not an address', address: '203.0.113.7' }), 'door');
  assert.equal(bad(null), 'door');
});

test('a door that protects nothing is used on this computer only', () => {
  const bad = (d, pin) => { try { readDoor(d, pin); } catch (err) { return err.code; } return null; };
  assert.equal(bad({ url: 'ws://203.0.113.7:7880', address: '203.0.113.7' }), 'plain');
  assert.equal(bad({ url: 'ws://192.168.1.20:7880', address: '192.168.1.20' }), 'plain');
  assert.equal(bad({ url: 'ws://127.0.0.1:7880', address: '203.0.113.7' }), 'plain', 'it is the address that is connected to');
  assert.equal(readDoor({ url: 'ws://127.0.0.1:7880', address: '127.0.0.1' }).secure, false);
  assert.equal(readDoor({ url: 'ws://[::1]:7880', address: '::1' }).secure, false);
  // A pin or a name with no TLS to carry them proves nothing.
  assert.equal(bad({ url: 'ws://127.0.0.1:7880', address: '127.0.0.1' }, PIN), 'door');
  assert.equal(bad({ url: 'ws://127.0.0.1:7880', address: '127.0.0.1', name: 'talk.example.org' }), 'door');
  assert.equal(isLoopback('127.9.9.9'), true);
  assert.equal(isLoopback('::ffff:127.0.0.1'), true);
  assert.equal(isLoopback('128.0.0.1'), false);
  assert.equal(isLoopback('localhost'), false);
});

// ── the connection, checked ──────────────────────────────────────────────────

test('with a pin, Office goes on only when the certificate shown has exactly that pin', async () => {
  const cert = selfSigned();
  const d = await door({ tls: cert });
  const at = { url: d.url, address: '127.0.0.1' };
  const socket = await connectDoor(readDoor(at, cert.pin));
  assert.equal(socket.encrypted, true);
  socket.destroy();
  // One digit off, and another server's certificate.
  const off = cert.pin.slice(0, 63) + (cert.pin.endsWith('0') ? '1' : '0');
  assert.equal(await code(connectDoor(readDoor(at, off))), 'pin');
  assert.equal(await code(connectDoor(readDoor(at, selfSigned().pin))), 'pin');
  assert.equal(await code(askDoor(readDoor(at, off), { path: '/admit', body: { pass: 'p' } })), 'pin');
  assert.deepEqual(d.hits, [], 'nothing was sent to a server that did not show the pinned certificate');
});

test('with a pin nothing else about the certificate counts: not its names, not who signed it', async () => {
  // A certificate for another address altogether, shown at this one.
  const cert = selfSigned({ addresses: ['203.0.113.7'] });
  const d = await door({ tls: cert });
  const socket = await connectDoor(readDoor({ url: d.url, address: '127.0.0.1' }, cert.pin));
  socket.destroy();
});

test('with no pin, a certificate nobody vouches for is refused', async () => {
  const cert = selfSigned();
  const d = await door({ tls: cert });
  assert.equal(await code(connectDoor(readDoor({ url: d.url, address: '127.0.0.1' }))), 'certificate');
  assert.equal(await code(connectDoor(readDoor({ url: `wss://talk.example.org:${d.port}`, address: '127.0.0.1', name: 'talk.example.org' }))), 'certificate');
  assert.deepEqual(d.hits, []);
});

test('with a name, the certificate is checked for that name at the address the introducer gave', async () => {
  const cert = selfSigned({ names: ['talk.example.org'], addresses: [] });
  const d = await door({ tls: cert });
  // The stand-in's certificate vouched for, as a public one is by those Office already trusts.
  const trusting = { connectTls: (o) => tls.connect({ ...o, ca: [cert.cert] }) };
  const good = readDoor({ url: `wss://talk.example.org:${d.port}`, address: '127.0.0.1', name: 'talk.example.org' });
  const socket = await connectDoor(good, trusting);
  assert.equal(socket.authorized, true);
  assert.equal(socket.servername, 'talk.example.org');
  socket.destroy();
  // The same certificate under another name, and at its bare address.
  assert.equal(await code(connectDoor(readDoor({ url: `wss://other.example.org:${d.port}`, address: '127.0.0.1', name: 'other.example.org' }), trusting)), 'certificate');
  assert.equal(await code(connectDoor(readDoor({ url: d.url, address: '127.0.0.1' }), trusting)), 'certificate');
});

test('a door nothing answers at cannot be reached, and a wait has an end', async () => {
  const d = await standInDoor();
  await d.close();
  assert.equal(await code(connectDoor(readDoor({ url: d.url, address: '127.0.0.1' }))), 'unreachable');
  const never = { connectPlain: () => ({ once: () => {}, destroy: () => {} }) };
  assert.equal(await code(connectDoor(readDoor({ url: 'ws://127.0.0.1:9', address: '127.0.0.1' }), { ...never, timeoutMs: 30 })), 'unreachable');
});

// ── the pass exchanged ───────────────────────────────────────────────────────

test('a pass is exchanged at the door for the join token, on the checked connection', async () => {
  const cert = selfSigned();
  const d = await door({ tls: cert, admit: (b) => ({ status: 200, body: { token: `join-for-${b.pass}`, room: 'r1', identity: 'u-a', role: 'host', e2ee: true, live: false, control: 'ctl' } }) });
  const got = await exchange({ doors: [{ url: d.url, address: '127.0.0.1' }], pin: cert.pin, pass: 'pass-1' });
  assert.equal(got.admitted.token, 'join-for-pass-1');
  assert.equal(got.admitted.control, 'ctl');
  assert.equal(got.door.port, d.port);
  assert.deepEqual(d.hits, [{ method: 'POST', path: '/admit', body: { pass: 'pass-1' }, host: `127.0.0.1:${d.port}` }]);
});

test('the next door is tried when the first does not answer, and the first that answers is the one used', async () => {
  const cert = selfSigned({ addresses: ['127.0.0.1', '::1'] });
  const gone = await standInDoor({ tls: cert });
  await gone.close();
  const d = await door({ tls: cert });
  const got = await exchange({ doors: [{ url: gone.url, address: '127.0.0.1' }, { url: d.url, address: '127.0.0.1' }, { url: 'wss://203.0.113.1:1', address: '203.0.113.1' }], pin: cert.pin, pass: 'p' });
  assert.equal(got.door.port, d.port);
  assert.equal(d.hits.length, 1);
});

test('a door\'s refusal is told as the door said it, and its other doors are not asked the same', async () => {
  const cert = selfSigned();
  for (const [status, body, want, state, max] of [
    [409, { error: 'full', max: 25 }, 'full', undefined, 25],
    [409, { error: 'full', max: 'many' }, 'full'],
    [403, { error: 'expired' }, 'expired'],
    [403, { error: 'used' }, 'used'],
    [403, { error: 'stale' }, 'stale'],
    [403, { error: 'ended' }, 'ended'],
    [403, { error: 'removed' }, 'removed'],
    [403, { error: 'signature' }, 'signature'],
    [503, { error: 'not-serving', state: 'paused' }, 'not-serving', 'paused'],
    [500, { error: 'Not A Code!' }, 'refused'],
    [200, { room: 'r1' }, 'refused'],
  ]) {
    const a = await standInDoor({ tls: cert, admit: () => ({ status, body }) });
    const b = await standInDoor({ tls: cert });
    try {
      await assert.rejects(exchange({ doors: [{ url: a.url, address: '127.0.0.1' }, { url: b.url, address: '127.0.0.1' }], pin: cert.pin, pass: 'p' }), (err) => {
        assert.ok(err instanceof DoorError);
        assert.equal(err.code, want);
        assert.equal(err.extra.state, state);
        assert.equal(err.extra.max, max, 'how many a full room holds, when the door says it in a number');
        return true;
      });
      assert.equal(b.hits.length, 0, want);
    } finally { await a.close(); await b.close(); }
  }
});

test('when no door can be used, the reason that matters most is the one told', async () => {
  const cert = selfSigned();
  const d = await door({ tls: cert });
  const gone = await standInDoor({ tls: cert });
  await gone.close();
  const other = selfSigned().pin;
  // A wrong certificate at one door and silence at another: it is the certificate that is said.
  assert.equal(await code(exchange({ doors: [{ url: d.url, address: '127.0.0.1' }, { url: gone.url, address: '127.0.0.1' }], pin: other, pass: 'p' })), 'pin');
  assert.equal(await code(exchange({ doors: [{ url: gone.url, address: '127.0.0.1' }, { url: d.url, address: '127.0.0.1' }], pin: other, pass: 'p' })), 'pin');
  assert.equal(await code(exchange({ doors: [{ url: 'ws://203.0.113.7:7880', address: '203.0.113.7' }], pass: 'p' })), 'plain');
  assert.equal(await code(exchange({ doors: [{ url: gone.url, address: '127.0.0.1' }], pin: cert.pin, pass: 'p' })), 'unreachable');
  assert.equal(await code(exchange({ doors: [], pin: cert.pin, pass: 'p' })), 'unreachable');
  assert.deepEqual(d.hits, [], 'the pass went to no door that showed another certificate');
});

test('a door that sends Office elsewhere is not followed', async () => {
  const elsewhere = await door();
  const d = await door({ admit: () => ({ status: 307, headers: { location: `http://127.0.0.1:${elsewhere.port}/admit` }, body: {} }) });
  const at = readDoor({ url: d.url, address: '127.0.0.1' });
  const got = await askDoor(at, { path: '/admit', body: { pass: 'p' } });
  assert.equal(got.status, 307);
  assert.equal(await code(exchange({ doors: [{ url: d.url, address: '127.0.0.1' }], pass: 'p' })), 'refused');
  assert.deepEqual(elsewhere.hits, []);
});

// ── the signalling connection ────────────────────────────────────────────────

const opened = (d, target, o = {}) => {
  const seen = { open: null, messages: [], close: null };
  const ws = openSocket(d, target, {
    open: (e) => { seen.open = e; },
    message: (data, binary) => seen.messages.push(binary ? data : data.toString('utf8')),
    close: (e) => { assert.equal(seen.close, null, 'a close is told once'); seen.close = e; },
  }, o);
  return { ws, seen };
};

test('the signalling connection goes through the pinned door, both ways, text and bytes of any size', async () => {
  const cert = selfSigned();
  const d = await door({ tls: cert });
  const at = readDoor({ url: d.url, address: '127.0.0.1' }, cert.pin);
  const { ws, seen } = opened(at, '/rtc/v1?access_token=join-token&auto_subscribe=1');
  assert.equal(ws.send('too early', false), false, 'nothing is sent before it is open');
  await until(() => seen.open);
  assert.equal(d.sockets[0].target, '/rtc/v1?access_token=join-token&auto_subscribe=1');
  const big = Buffer.alloc(200000);
  for (let i = 0; i < big.length; i += 1) big[i] = i & 255;
  assert.equal(ws.send('hello', false), true);
  assert.equal(ws.send(Buffer.from([1, 2, 3])), true);
  assert.equal(ws.send(big), true);
  await until(() => seen.messages.length === 3);
  assert.equal(seen.messages[0], 'hello');
  assert.ok(Buffer.from([1, 2, 3]).equals(seen.messages[1]));
  assert.ok(big.equals(seen.messages[2]));
  assert.deepEqual(d.sockets[0].got.map((g) => g.binary), [false, true, true]);
  // The server asks whether Office is still there, and is answered.
  d.sockets[0].ping(Buffer.from('still there?'));
  await until(() => d.sockets[0].pongs.length === 1);
  assert.equal(d.sockets[0].pongs[0].toString(), 'still there?');
  // Office closes: the server is told why, and the window once.
  ws.close(1000, 'left');
  await until(() => d.sockets[0].closed);
  assert.deepEqual(d.sockets[0].closed, { code: 1000, reason: 'left' });
  assert.deepEqual(seen.close, { code: 1000, reason: 'left' });
  assert.equal(ws.send('after', false), false);
});

test('a close from the server is told with its code, and a protocol asked for is the one answered', async () => {
  const d = await door();
  const { seen } = opened(readDoor({ url: d.url, address: '127.0.0.1' }), '/rtc', { protocols: ['lk-v1', 'not a protocol'] });
  await until(() => seen.open);
  assert.equal(seen.open.protocol, 'lk-v1');
  assert.deepEqual(d.sockets[0].protocols, ['lk-v1']);
  d.sockets[0].close(4004, 'room closed');
  await until(() => seen.close);
  assert.deepEqual(seen.close, { code: 4004, reason: 'room closed' });
});

test('a connection the engine does not take up is told as closed, with what it answered', async () => {
  const d = await door({ rtc: 'refuse' });
  const { seen } = opened(readDoor({ url: d.url, address: '127.0.0.1' }), '/rtc?access_token=old');
  await until(() => seen.close);
  assert.equal(seen.open, null);
  assert.equal(seen.close.code, 1006);
  assert.equal(seen.close.status, 401);
  assert.equal(seen.close.why, 'refused');
  assert.equal(seen.close.reason, 'invalid token');
});

test('only the engine\'s signalling is asked for through a door, and a wrong certificate closes before anything is sent', async () => {
  const cert = selfSigned();
  const d = await door({ tls: cert });
  const at = readDoor({ url: d.url, address: '127.0.0.1' }, cert.pin);
  for (const target of ['/admit', '/control', '/rtcx', '/', '', '/rtc/v1?a=1\r\nx-injected: 1', 'http://elsewhere.example/rtc',
    // A step back up out of the engine's part, however it is written, and what a request line cannot carry.
    '/rtc/../admit', '/rtc/./v1', '/rtc/%2e%2e/admit', '/rtc//admit', '/rtc/v1/', '/rtc?name=علي', '/rtc?x=\u0001', '/rtc?a=1#b', '/rtc?a b', null, undefined]) {
    const { seen } = opened(at, target);
    await until(() => seen.close);
    assert.equal(seen.close.why, 'path', JSON.stringify(target));
  }
  const { seen } = opened(readDoor({ url: d.url, address: '127.0.0.1' }, selfSigned().pin), '/rtc');
  await until(() => seen.close);
  assert.equal(seen.close.why, 'pin');
  assert.deepEqual(d.hits, []);
});

test('closing before the door has answered connects to nothing more and tells the close once', async () => {
  const d = await door();
  const { ws, seen } = opened(readDoor({ url: d.url, address: '127.0.0.1' }), '/rtc');
  ws.close();
  await new Promise((r) => setTimeout(r, 60));
  assert.deepEqual(seen.close, { code: 1000, reason: '' });
  assert.equal(seen.open, null);
  assert.equal(d.sockets.length, 0);
});
