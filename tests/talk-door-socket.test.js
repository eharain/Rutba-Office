// A meeting on a Talk Server: the WebSocket the meeting's engine is handed
// in a window (apps/desktop/renderer/apps/talk/door-socket.js). The shell
// here is a stand-in for Talk's service: it notes what is asked and says
// what the test tells it to.

import test from 'node:test';
import assert from 'node:assert/strict';
import { DoorSocket, doorsOf, lastRefusal, sessionOf, useDoorSockets } from '../apps/desktop/renderer/apps/talk/door-socket.js';

const tick = () => new Promise((r) => setTimeout(r, 0));

function shellOf({ open = undefined } = {}) {
  const hears = [];
  const asked = [];
  let next = 1;
  const shell = {
    on: (event, fn) => { hears.push({ event, fn }); return () => {}; },
    talk: {
      doorOpen: async (p) => { asked.push({ do: 'open', ...p }); return open ? open(p) : { socket: next++ }; },
      doorSend: async (p) => { asked.push({ do: 'send', ...p }); return { sent: true }; },
      doorClose: async (p) => { asked.push({ do: 'close', ...p }); return { closed: true }; },
    },
  };
  const say = (msg) => { for (const h of hears) if (h.event === 'talk:door') h.fn(msg); };
  return { shell, asked, say, hears };
}

/** A socket, with everything it tells noted in order. */
function socketOf(shell, url = 'wss://s3.talk-door.invalid/rtc/v1?access_token=t', protocols = undefined) {
  const s = new DoorSocket(shell, sessionOf(url), protocols);
  const seen = [];
  s.onopen = () => seen.push('open');
  s.onmessage = (e) => seen.push({ message: e.data });
  s.onerror = () => seen.push('error');
  s.onclose = (e) => seen.push({ close: e.code, reason: e.reason, clean: e.wasClean });
  return { s, seen };
}

test('only the address of a way in is read as one', () => {
  assert.deepEqual(sessionOf('wss://s3.talk-door.invalid/rtc/v1?access_token=t&x=1'), { session: 3, target: '/rtc/v1?access_token=t&x=1' });
  assert.deepEqual(sessionOf('ws://s12.talk-door.invalid/rtc'), { session: 12, target: '/rtc' });
  for (const other of ['wss://talk-door.invalid/rtc', 'wss://s3.talk-door.invalid.example.com/rtc', 'https://s3.talk-door.invalid/rtc', 'wss://sx.talk-door.invalid/rtc', 'wss://introducer.rutba.io/ws', 'not an address', '']) {
    assert.equal(sessionOf(other), null, other);
  }
});

test('it opens through the service, with the way in, the path and the protocols asked for', async () => {
  const { shell, asked, say } = shellOf();
  const { s, seen } = socketOf(shell, 'wss://s3.talk-door.invalid/rtc/v1?access_token=t', ['a', 'b']);
  assert.equal(s.readyState, DoorSocket.CONNECTING);
  await tick();
  assert.deepEqual(asked, [{ do: 'open', session: 3, target: '/rtc/v1?access_token=t', protocols: ['a', 'b'] }]);
  say({ socket: 1, type: 'open', protocol: 'a' });
  assert.equal(s.readyState, DoorSocket.OPEN);
  assert.equal(s.protocol, 'a');
  assert.deepEqual(seen, ['open']);
});

test('what the server says arrives as text, as bytes, or as a blob, as the engine asked', async () => {
  const { shell, say } = shellOf();
  const { s, seen } = socketOf(shell);
  await tick();
  say({ socket: 1, type: 'open' });
  say({ socket: 1, type: 'message', data: 'hello', binary: false });
  s.binaryType = 'arraybuffer';
  const whole = new Uint8Array([9, 1, 2, 3, 9]);
  say({ socket: 1, type: 'message', data: whole.subarray(1, 4), binary: true });
  s.binaryType = 'blob';
  say({ socket: 1, type: 'message', data: new Uint8Array([7, 8]), binary: true });
  assert.equal(seen[1].message, 'hello');
  assert.ok(seen[2].message instanceof ArrayBuffer);
  assert.deepEqual([...new Uint8Array(seen[2].message)], [1, 2, 3]);
  assert.ok(seen[3].message instanceof Blob);
  assert.deepEqual([...new Uint8Array(await seen[3].message.arrayBuffer())], [7, 8]);
});

test('an event listener hears what a handler hears', async () => {
  const { shell, say } = shellOf();
  const s = new DoorSocket(shell, sessionOf('wss://s1.talk-door.invalid/rtc'));
  const heard = [];
  for (const type of ['open', 'message', 'close', 'error']) s.addEventListener(type, (e) => heard.push(e.type));
  await tick();
  say({ socket: 1, type: 'open' });
  say({ socket: 1, type: 'message', data: 'x', binary: false });
  say({ socket: 1, type: 'close', code: 1000, reason: '' });
  assert.deepEqual(heard, ['open', 'message', 'close']);
});

test('what is said before the socket knows its number is not lost', async () => {
  let answer;
  const { shell, say } = shellOf({ open: () => new Promise((r) => { answer = r; }) });
  const { s, seen } = socketOf(shell);
  await tick();
  say({ socket: 5, type: 'open', protocol: '' });
  say({ socket: 5, type: 'message', data: 'first', binary: false });
  assert.deepEqual(seen, []);
  answer({ socket: 5 });
  await tick();
  assert.deepEqual(seen, ['open', { message: 'first' }]);
  assert.equal(s.readyState, DoorSocket.OPEN);
});

test('what is said of a socket already closed is dropped, and never kept', async () => {
  const { shell, say } = shellOf();
  const { s } = socketOf(shell);
  await tick();
  say({ socket: 1, type: 'open' });
  s.close();
  for (let i = 0; i < 200; i += 1) say({ socket: 1, type: 'close', code: 1000, reason: '' });
  assert.equal(doorsOf(shell).early.size, 0);
  assert.equal(doorsOf(shell).open.size, 0);
});

test('it sends text as text and bytes as a copy of them, and nothing before it is open', async () => {
  const { shell, asked, say } = shellOf();
  const { s } = socketOf(shell);
  assert.throws(() => s.send('too soon'), { name: 'InvalidStateError' });
  await tick();
  say({ socket: 1, type: 'open' });
  s.send('hello');
  const bytes = new Uint8Array([1, 2, 3, 4]);
  s.send(bytes.subarray(1, 3));
  s.send(new Uint8Array([5, 6]).buffer);
  bytes[1] = 99;
  assert.throws(() => s.send({ not: 'bytes' }), TypeError);
  const sent = asked.filter((a) => a.do === 'send');
  assert.deepEqual(sent[0], { do: 'send', socket: 1, data: 'hello', binary: false });
  assert.deepEqual([...sent[1].data], [2, 3]);
  assert.equal(sent[1].binary, true);
  assert.deepEqual([...sent[2].data], [5, 6]);
});

test('closing asks the service once and is told once', async () => {
  const { shell, asked, say } = shellOf();
  const { s, seen } = socketOf(shell);
  await tick();
  say({ socket: 1, type: 'open' });
  s.close(1000, 'left');
  s.close();
  say({ socket: 1, type: 'close', code: 1000, reason: 'left' });
  assert.deepEqual(asked.filter((a) => a.do === 'close'), [{ do: 'close', socket: 1, code: 1000, reason: 'left' }]);
  assert.deepEqual(seen, ['open', { close: 1000, reason: 'left', clean: true }]);
  assert.equal(s.readyState, DoorSocket.CLOSED);
  assert.throws(() => s.send('after'), { name: 'InvalidStateError' });
});

test('closed before the service had answered, the connection is closed as soon as it has', async () => {
  let answer;
  const { shell, asked } = shellOf({ open: () => new Promise((r) => { answer = r; }) });
  const { s, seen } = socketOf(shell);
  await tick();
  s.close(1000, 'changed my mind');
  assert.equal(s.readyState, DoorSocket.CLOSING);
  answer({ socket: 8 });
  await tick();
  assert.deepEqual(asked.filter((a) => a.do === 'close'), [{ do: 'close', socket: 8, code: 1000, reason: 'changed my mind' }]);
  assert.deepEqual(seen, [{ close: 1000, reason: 'changed my mind', clean: true }]);
});

test('a connection the server ends is a clean close; one that breaks is an error first', async () => {
  const a = shellOf();
  const one = socketOf(a.shell);
  await tick();
  a.say({ socket: 1, type: 'open' });
  a.say({ socket: 1, type: 'close', code: 1000, reason: 'bye' });
  assert.deepEqual(one.seen, ['open', { close: 1000, reason: 'bye', clean: true }]);

  const b = shellOf();
  const two = socketOf(b.shell);
  await tick();
  b.say({ socket: 1, type: 'open' });
  b.say({ socket: 1, type: 'close', code: 1006, reason: '' });
  assert.deepEqual(two.seen, ['open', 'error', { close: 1006, reason: '', clean: false }]);
});

test('a connection the server does not take up is an error, and why it was refused is kept', async () => {
  const { shell, say } = shellOf();
  const { s, seen } = socketOf(shell);
  await tick();
  say({ socket: 1, type: 'close', code: 1006, reason: 'room is full', status: 403, why: 'refused' });
  assert.deepEqual(seen, ['error', { close: 1006, reason: 'room is full', clean: false }]);
  assert.deepEqual(s.refusal, { status: 403, why: 'refused', reason: 'room is full' });
});

test('with no way in left, or a service that fails, it ends as a connection that never opened', async () => {
  const a = shellOf({ open: () => ({ error: 'no-session' }) });
  const one = socketOf(a.shell);
  await tick();
  assert.deepEqual(one.seen, ['error', { close: 1006, reason: '', clean: false }]);
  assert.equal(one.s.refusal.why, 'no-session');

  const b = shellOf({ open: () => { throw new Error('gone'); } });
  const two = socketOf(b.shell);
  await tick();
  assert.deepEqual(two.seen, ['error', { close: 1006, reason: '', clean: false }]);
  assert.equal(doorsOf(b.shell).waiting, 0);
});

test('two sockets of one window each hear their own, through one listener', async () => {
  const { shell, say, hears } = shellOf();
  const one = socketOf(shell, 'wss://s1.talk-door.invalid/rtc');
  const two = socketOf(shell, 'wss://s2.talk-door.invalid/rtc');
  await tick();
  say({ socket: 2, type: 'open' });
  say({ socket: 2, type: 'message', data: 'for two', binary: false });
  say({ socket: 1, type: 'open' });
  assert.deepEqual(one.seen, ['open']);
  assert.deepEqual(two.seen, ['open', { message: 'for two' }]);
  assert.equal(hears.length, 1);
});

test('while a meeting is open, its way in gets a door socket and every other address the one of the browser', async () => {
  const { shell, asked } = shellOf();
  const made = [];
  class Native { constructor(url, protocols) { made.push({ url, protocols, args: arguments.length }); } }
  Object.assign(Native, { CONNECTING: 0, OPEN: 1, CLOSING: 2, CLOSED: 3 });
  const scope = { WebSocket: Native };
  const back = useDoorSockets(shell, 4, scope);
  const backOther = useDoorSockets(shell, 9, scope);
  assert.notEqual(scope.WebSocket, Native);
  assert.equal(scope.WebSocket.OPEN, 1);
  assert.equal(scope.WebSocket.CONNECTING, 0);

  const door = new scope.WebSocket('wss://s4.talk-door.invalid/rtc/v1?access_token=t');
  assert.ok(door instanceof DoorSocket);
  await tick();
  assert.deepEqual(asked, [{ do: 'open', session: 4, target: '/rtc/v1?access_token=t', protocols: [] }]);

  // A way in this window does not hold is nobody's: it is not carried.
  assert.ok(new scope.WebSocket('wss://s5.talk-door.invalid/rtc') instanceof Native);
  const other = new scope.WebSocket('wss://introducer.rutba.io/ws');
  assert.ok(other instanceof Native);
  new scope.WebSocket('wss://introducer.rutba.io/ws', ['p']);
  assert.deepEqual(made.slice(1), [{ url: 'wss://introducer.rutba.io/ws', protocols: undefined, args: 1 }, { url: 'wss://introducer.rutba.io/ws', protocols: ['p'], args: 2 }]);

  back();
  back();
  assert.notEqual(scope.WebSocket, Native, 'one meeting still open');
  assert.ok(new scope.WebSocket('wss://s4.talk-door.invalid/rtc') instanceof Native, 'the way in given back is carried no more');
  backOther();
  assert.equal(scope.WebSocket, Native);
});

test('why a connection was refused can be read once by the window, for the way in it was', async () => {
  const { shell, say } = shellOf();
  socketOf(shell);
  await tick();
  say({ socket: 1, type: 'close', code: 1006, reason: 'no room', status: 403, why: 'refused' });
  assert.equal(lastRefusal(shell, 7), null);
  assert.deepEqual(lastRefusal(shell, 3), { status: 403, why: 'refused', reason: 'no room' });
  assert.equal(lastRefusal(shell, 3), null);
});
