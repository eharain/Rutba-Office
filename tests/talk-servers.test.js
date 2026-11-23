// A meeting held on a Talk Server, as Talk's service asks for it
// (apps/desktop/main/talk-servers.js): what is asked of the introducer, the
// pass exchanged at the door, the way in a window is given, and the
// signalling carried through the main side. The introducer here is a
// stand-in that answers as the contract says; the door is the stand-in of
// talk-door-stand-in.js, reached by the real code of talk-door.js.

import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { createServers, restBase, DOOR_HOST } from '../apps/desktop/main/talk-servers.js';
import { DoorError } from '../apps/desktop/main/talk-door.js';
import { selfSigned, standInDoor } from './talk-door-stand-in.js';

const doors = [];
test.after(async () => { for (const d of doors) await d.close(); });
const until = async (fn, ms = 4000) => { const t = Date.now(); while (!fn()) { if (Date.now() - t > ms) throw new Error('waited too long'); await new Promise((r) => setTimeout(r, 5)); } };

/** A window as the service knows one: a number, and its closing. */
const windowOf = (id) => Object.assign(new EventEmitter(), { id });

/**
 * The introducer's `/servers/`: `routes` answers `"METHOD path"` with
 * [status, body] or a function of the request's body.
 */
function setUp({ routes = {}, tokens = ['token-1'], renew = async () => 'kept', door = undefined } = {}) {
  const asked = [];
  const told = [];
  let token = tokens[0];
  const fetchImpl = async (url, init) => {
    const u = new URL(url);
    const key = `${init.method} ${u.pathname}`;
    const body = init.body ? JSON.parse(init.body) : undefined;
    asked.push({ key, body, auth: init.headers.authorization, origin: u.origin, redirect: init.redirect });
    const r = routes[key];
    if (!r) return { status: 404, text: async () => JSON.stringify({ error: 'not-found' }) };
    const [status, answer] = typeof r === 'function' ? r(body, asked.filter((a) => a.key === key).length) : r;
    return { status, text: async () => (typeof answer === 'string' ? answer : JSON.stringify(answer)) };
  };
  const servers = createServers({
    server: () => 'wss://introducer.rutba.io/ws',
    token: () => token,
    renew: async (force) => { const r = await renew(force); if (r === 'fresh') token = tokens[1] || token; return r; },
    fetchImpl,
    send: (win, event, payload) => told.push({ win: win?.id, event, ...payload }),
    ...(door ? { door } : {}),
  });
  return { servers, api: servers.api, asked, told };
}

const MEETING = { id: 'meet-AAAAAAAAAAAAAAAAAAAA', live: false, e2ee: true, maxPeople: 25, state: 'open', created: 1 };

test('the introducer is asked by https at the address of its socket, and nowhere a link or a person names', () => {
  assert.equal(restBase('wss://introducer.rutba.io/ws'), 'https://introducer.rutba.io');
  assert.equal(restBase('ws://127.0.0.1:4700/ws'), 'http://127.0.0.1:4700');
  assert.equal(restBase('https://introducer.rutba.io/ws'), null);
  assert.equal(restBase('not an address'), null);
  assert.equal(restBase(''), null);
});

test('each question carries the person\'s sign-in, follows no redirect, and is asked under /servers/', async () => {
  const { api, asked } = setUp({ routes: { 'GET /servers/mine': [200, { plan: { id: 'free', maxPeople: 25, maxServers: 1, rutbaServers: false, recording: false }, servers: [{ id: 'srv_1', name: 'Home', state: 'active', answering: true, canCarry: true, why: null, pin: 'ab'.repeat(32), doors: [{ url: 'wss://203.0.113.7:7443', address: '203.0.113.7' }], settings: { who: 'named', named: ['u-b'], uses: {}, limits: {} } }] }] } });
  const mine = await api.serverMine();
  assert.deepEqual(asked, [{ key: 'GET /servers/mine', body: undefined, auth: 'Bearer token-1', origin: 'https://introducer.rutba.io', redirect: 'error' }]);
  assert.deepEqual(mine, { plan: { id: 'free', maxPeople: 25, maxServers: 1, rutbaServers: false, recording: false }, servers: [{ id: 'srv_1', name: 'Home', state: 'active', answering: true, canCarry: true, why: null, who: 'named', named: ['u-b'] }] });
  assert.equal(JSON.stringify(mine).includes('203.0.113.7'), false, 'a window is told no door');
  assert.equal(JSON.stringify(mine).includes('abab'), false, 'and no pin');
});

test('a token that ran out is renewed and the question asked once more; a sign-in that ended is said', async () => {
  const fresh = setUp({ tokens: ['old', 'new'], renew: async () => 'fresh', routes: { 'GET /servers/mine': (_b, n) => (n === 1 ? [401, { error: 'sign-in', reason: 'expired' }] : [200, { plan: {}, servers: [] }]) } });
  assert.deepEqual((await fresh.api.serverMine()).servers, []);
  assert.deepEqual(fresh.asked.map((a) => a.auth), ['Bearer old', 'Bearer new']);

  const still = setUp({ renew: async () => 'kept', routes: { 'GET /servers/mine': [401, { error: 'sign-in' }] } });
  assert.deepEqual(await still.api.serverMine(), { error: 'sign-in' });
  assert.equal(still.asked.length, 2, 'asked twice and no more');

  const ended = setUp({ renew: async () => 'ended', routes: { 'GET /servers/mine': [401, { error: 'sign-in' }] } });
  assert.deepEqual(await ended.api.serverMine(), { error: 'sign-in' });
  assert.equal(ended.asked.length, 1);

  const out = createServers({ server: () => 'wss://introducer.rutba.io/ws', token: () => null, renew: async () => 'kept', fetchImpl: async () => assert.fail('asked with no sign-in') });
  assert.deepEqual(await out.api.serverMine(), { error: 'sign-in' });
});

test('an introducer out of reach, one with no servers, and an answer that is no answer each have their own word', async () => {
  const offline = createServers({ server: () => 'wss://introducer.rutba.io/ws', token: () => 't', renew: async () => 'kept', fetchImpl: async () => { throw new TypeError('fetch failed'); } });
  assert.deepEqual(await offline.api.serverMine(), { error: 'offline' });
  assert.deepEqual(await setUp().api.serverMine(), { error: 'no-service' }, 'the whole of /servers/ is absent');
  assert.deepEqual(await setUp({ routes: { 'GET /servers/mine': [404, '<html>not here</html>'] } }).api.serverMine(), { error: 'no-service' });
  assert.deepEqual(await setUp({ routes: { 'GET /servers/mine': [500, { error: 'failed' }] } }).api.serverMine(), { error: 'failed' });
  assert.deepEqual(await setUp({ routes: { 'GET /servers/mine': [502, '<html>bad gateway</html>'] } }).api.serverMine(), { error: 'failed' });
  assert.deepEqual(await setUp({ routes: { 'GET /servers/mine': [429, { error: 'too-many' }] } }).api.serverMine(), { error: 'too-many' });
  assert.deepEqual(await setUp({ routes: { 'GET /servers/mine': [403, { error: 'banned' }] } }).api.serverMine(), { error: 'banned' });
  assert.deepEqual(await setUp({ routes: { 'GET /servers/mine': [200, 'not json'] } }).api.serverMine(), { error: 'refused' });
});

test('a server is asked for by the meeting\'s own code, and a window is told the meeting, never where it is', async () => {
  const { api, asked } = setUp({ routes: { 'POST /servers/meetings': [200, { meeting: MEETING, server: { id: 'srv_1', doors: [{ url: 'wss://203.0.113.7:7443', address: '203.0.113.7' }], pin: 'cd'.repeat(32) } }] } });
  const got = await api.serverCarry({ id: MEETING.id, live: false });
  assert.deepEqual(asked[0].body, { meetingId: MEETING.id, live: false });
  assert.deepEqual(got, { meeting: { id: MEETING.id, live: false, e2ee: true, maxPeople: 25, state: 'open' } });
  // What is not a meeting's code is not asked at all.
  for (const id of ['', 'short', 'has space in it', 'a/b/c/d/e/f/g', null, 'x'.repeat(65)]) assert.deepEqual(await api.serverCarry({ id }), { error: 'refused' });
  assert.equal(asked.length, 1);
});

test('with no server that can carry it, the person\'s own servers are named with the reason', async () => {
  const { api } = setUp({ routes: { 'POST /servers/meetings': [409, { error: 'no-server', own: [{ id: 'srv_1', name: 'Home', why: 'not-answering', doors: ['x'] }, { id: 'srv_2', name: 'Office', why: 'full' }] }] } });
  assert.deepEqual(await api.serverCarry({ id: MEETING.id }), { error: 'no-server', own: [{ id: 'srv_1', name: 'Home', why: 'not-answering' }, { id: 'srv_2', name: 'Office', why: 'full' }] });
  const taken = setUp({ routes: { 'POST /servers/meetings': [409, { error: 'meeting-taken' }] } });
  assert.deepEqual(await taken.api.serverCarry({ id: MEETING.id }), { error: 'meeting-taken' });
});

test('letting in, taking out, looking and ending are each one question about the meeting', async () => {
  const id = MEETING.id;
  const { api, asked } = setUp({ routes: {
    [`POST /servers/meetings/${id}/admit`]: [200, { meeting: MEETING, person: { id: 'u-b', kind: 'office' }, role: 'member' }],
    [`POST /servers/meetings/${id}/remove`]: [200, { meeting: MEETING }],
    [`GET /servers/meetings/${id}`]: [200, { meeting: MEETING }],
    [`POST /servers/meetings/${id}/end`]: [200, { meeting: MEETING.id, state: 'ended' }],
  } });
  assert.deepEqual(await api.serverAdmit({ id, person: 'u-b' }), { person: 'u-b', role: 'member' });
  assert.deepEqual(asked.at(-1).body, { person: { id: 'u-b', kind: 'office' }, role: 'member' }, 'only an Office account is let in from here');
  assert.deepEqual(await api.serverAdmit({ id, person: 'u-b', role: 'owner' }), { error: 'refused' });
  assert.deepEqual(await api.serverAdmit({ id, person: '' }), { error: 'refused' });
  assert.deepEqual(await api.serverRemove({ id, person: 'u-b' }), { removed: true });
  assert.deepEqual(asked.at(-1).body, { person: 'u-b' });
  assert.deepEqual((await api.serverInfo({ id })).meeting.state, 'open');
  assert.deepEqual(await api.serverEnd({ id }), { ended: true });
  assert.equal(asked.at(-1).body, undefined);
  assert.equal(asked.length, 4);
  // Somebody who is not its starter or a host is told the meeting is not there.
  const other = setUp({ routes: { [`POST /servers/meetings/${id}/admit`]: [404, { error: 'no-meeting' }], [`GET /servers/meetings/${id}`]: [404, { error: 'no-meeting' }] } });
  assert.deepEqual(await other.api.serverAdmit({ id, person: 'u-b' }), { error: 'no-meeting' });
  assert.deepEqual(await other.api.serverInfo({ id }), { error: 'no-meeting' });
});

// ── into the meeting ─────────────────────────────────────────────────────────

/** A door stand-in in place of talk-door.js, to see what it is given. */
function fakeDoor({ exchange } = {}) {
  const seen = { exchanged: [], asked: [], opened: [] };
  return {
    seen,
    door: {
      exchange: async (o) => { seen.exchanged.push(o); return exchange ? exchange(o, seen.exchanged.length) : { door: { tag: 'door-1' }, admitted: { token: 'join-1', room: 'r1', identity: 'u-a', role: 'member', e2ee: true, live: false } }; },
      askDoor: async (d, q) => { seen.asked.push({ d, ...q }); return { status: 200, body: { people: [{ id: 'u-a', name: 'Amira', muted: false, extra: 1 }] } }; },
      openSocket: (d, target, on, o) => {
        const s = { d, target, on, protocols: o?.protocols, sent: [], closed: null, send: (data, binary) => { s.sent.push({ data, binary }); return true; }, close: (code, reason) => { s.closed = { code, reason }; on.close({ code, reason }); } };
        seen.opened.push(s);
        return s;
      },
    },
  };
}
const passRoute = (id, pass = 'pass-1') => ({ [`POST /servers/meetings/${id}/pass`]: (_b, n) => [200, { pass: `${pass}-${n}`, expires: 2, server: { id: 'srv_1', doors: [{ url: 'wss://203.0.113.7:7443', address: '203.0.113.7' }], pin: 'cd'.repeat(32) }, meeting: MEETING, role: 'member' }] });

test('a person fetches their own pass, it is exchanged at the door, and the window is given a way in and no address', async () => {
  const f = fakeDoor();
  const { api, asked } = setUp({ routes: passRoute(MEETING.id), door: f.door });
  const win = windowOf(7);
  const got = await api.serverEnter({ id: MEETING.id }, win);
  assert.equal(asked[0].key, `POST /servers/meetings/${MEETING.id}/pass`);
  assert.equal(asked[0].body, undefined, 'the pass is asked for with no body: the person is who signed in');
  assert.deepEqual(f.seen.exchanged, [{ doors: [{ url: 'wss://203.0.113.7:7443', address: '203.0.113.7' }], pin: 'cd'.repeat(32), pass: 'pass-1-1' }]);
  assert.deepEqual(got, { session: 1, url: `wss://s1.${DOOR_HOST}`, token: 'join-1', room: 'r1', identity: 'u-a', role: 'member', host: false, e2ee: true, live: false, meeting: { id: MEETING.id, live: false, e2ee: true, maxPeople: 25, state: 'open' } });
  const text = JSON.stringify(got);
  for (const secret of ['203.0.113.7', 'cdcd', 'pass-1']) assert.equal(text.includes(secret), false, secret);
});

test('a host is told they are one, and their control token stays on the main side', async () => {
  const f = fakeDoor({ exchange: () => ({ door: { tag: 'door-1' }, admitted: { token: 'join-1', room: 'r1', identity: 'u-a', role: 'host', e2ee: true, live: true, control: 'control-secret' } }) });
  const { api } = setUp({ routes: passRoute(MEETING.id), door: f.door });
  const win = windowOf(7);
  const got = await api.serverEnter({ id: MEETING.id }, win);
  assert.equal(got.host, true);
  assert.equal(got.role, 'host');
  assert.equal(got.live, true);
  assert.equal(JSON.stringify(got).includes('control-secret'), false);
  assert.deepEqual(await api.serverControl({ session: got.session, do: 'roster' }, win), { people: [{ id: 'u-a', name: 'Amira', muted: false }] });
  assert.deepEqual(f.seen.asked[0], { d: { tag: 'door-1' }, method: 'POST', path: '/control', body: { control: 'control-secret', do: 'roster' } });
  assert.deepEqual(await api.serverControl({ session: got.session, do: 'mute', person: 'u-b' }, win), { done: true });
  assert.deepEqual(f.seen.asked[1].body, { control: 'control-secret', do: 'mute', person: 'u-b' });
  assert.deepEqual(await api.serverControl({ session: got.session, do: 'mute-all' }, win), { done: true });
  assert.deepEqual(await api.serverControl({ session: got.session, do: 'mute' }, win), { error: 'refused' });
  assert.deepEqual(await api.serverControl({ session: got.session, do: 'kick' }, win), { error: 'refused' });
  assert.equal(f.seen.asked.length, 3);
});

test('somebody who is not a host asks the door nothing', async () => {
  const f = fakeDoor();
  const { api } = setUp({ routes: passRoute(MEETING.id), door: f.door });
  const win = windowOf(7);
  const got = await api.serverEnter({ id: MEETING.id }, win);
  assert.deepEqual(await api.serverControl({ session: got.session, do: 'mute-all' }, win), { error: 'not-host' });
  assert.equal(f.seen.asked.length, 0);
});

test('a pass that ran out on the way is fetched once more, and no more than once', async () => {
  for (const why of ['expired', 'stale', 'used']) {
    const f = fakeDoor({ exchange: (_o, n) => { if (n === 1) throw new DoorError(why); return { door: {}, admitted: { token: 'join-2', room: 'r1', identity: 'u-a', role: 'member' } }; } });
    const { api, asked } = setUp({ routes: passRoute(MEETING.id), door: f.door });
    const got = await api.serverEnter({ id: MEETING.id }, windowOf(1));
    assert.equal(got.token, 'join-2', why);
    assert.deepEqual(f.seen.exchanged.map((e) => e.pass), ['pass-1-1', 'pass-1-2'], 'a new pass, not the old one again');
    assert.equal(asked.length, 2);

    const never = fakeDoor({ exchange: () => { throw new DoorError(why); } });
    const again = setUp({ routes: passRoute(MEETING.id), door: never.door });
    assert.deepEqual(await again.api.serverEnter({ id: MEETING.id }, windowOf(1)), { error: why });
    assert.equal(again.asked.length, 2);
    assert.deepEqual(again.servers.open(), { sessions: 0, sockets: 0 });
  }
});

test('every way of not getting in is told by its own word, and leaves nothing open', async () => {
  for (const [err, want] of [
    [new DoorError('pin'), { error: 'pin' }],
    [new DoorError('certificate'), { error: 'certificate' }],
    [new DoorError('plain'), { error: 'plain' }],
    [new DoorError('unreachable'), { error: 'unreachable' }],
    [new DoorError('ended'), { error: 'ended' }],
    [new DoorError('removed'), { error: 'removed' }],
    [new DoorError('not-serving', { state: 'paused' }), { error: 'not-serving', state: 'paused' }],
    [new Error('anything else'), { error: 'unreachable' }],
  ]) {
    const f = fakeDoor({ exchange: () => { throw err; } });
    const { api, servers, asked } = setUp({ routes: passRoute(MEETING.id), door: f.door });
    assert.deepEqual(await api.serverEnter({ id: MEETING.id }, windowOf(1)), want);
    assert.equal(asked.length, 1, 'no second pass for a refusal a new pass would not change');
    assert.deepEqual(servers.open(), { sessions: 0, sockets: 0 });
  }
  const id = MEETING.id;
  for (const [status, body, want] of [
    [403, { error: 'no-pass' }, { error: 'no-pass' }],
    [404, { error: 'no-meeting' }, { error: 'no-meeting' }],
    [409, { error: 'server-gone', state: 'not-answering' }, { error: 'server-gone', state: 'not-answering' }],
    [403, { error: 'banned' }, { error: 'banned' }],
    [200, { expires: 2 }, { error: 'refused' }],
  ]) {
    const f = fakeDoor();
    const { api } = setUp({ routes: { [`POST /servers/meetings/${id}/pass`]: [status, body] }, door: f.door });
    assert.deepEqual(await api.serverEnter({ id }, windowOf(1)), want);
    assert.equal(f.seen.exchanged.length, 0, 'no door is tried without a pass');
  }
});

// ── the signalling, carried for one window ───────────────────────────────────

test('the meeting\'s signalling is opened through the way in, and what happens on it is told to that window alone', async () => {
  const f = fakeDoor();
  const { api, told, servers } = setUp({ routes: passRoute(MEETING.id), door: f.door });
  const win = windowOf(7);
  const { session } = await api.serverEnter({ id: MEETING.id }, win);
  const { socket } = api.doorOpen({ session, target: '/rtc/v1?access_token=join-1', protocols: ['lk'] }, win);
  assert.equal(socket, 1);
  const s = f.seen.opened[0];
  assert.deepEqual([s.d, s.target, s.protocols], [{ tag: 'door-1' }, '/rtc/v1?access_token=join-1', ['lk']]);
  s.on.open({ protocol: 'lk' });
  s.on.message(Buffer.from('{"a":1}'), false);
  s.on.message(Buffer.from([9, 8, 7]), true);
  assert.deepEqual(told.slice(0, 2), [{ win: 7, event: 'talk:door', socket: 1, type: 'open', protocol: 'lk' }, { win: 7, event: 'talk:door', socket: 1, type: 'message', binary: false, data: '{"a":1}' }]);
  assert.deepEqual([told[2].type, told[2].binary, [...told[2].data]], ['message', true, [9, 8, 7]]);
  assert.ok(told[2].data instanceof Uint8Array && !Buffer.isBuffer(told[2].data), 'bytes go to the window as plain bytes');

  assert.deepEqual(api.doorSend({ socket, data: 'text', binary: false }, win), { sent: true });
  assert.deepEqual(api.doorSend({ socket, data: new Uint8Array([1, 2]) }, win), { sent: true });
  assert.deepEqual(api.doorSend({ socket, data: new Uint8Array([3, 4, 5]).buffer }, win), { sent: true });
  assert.deepEqual(api.doorSend({ socket, data: { not: 'bytes' } }, win), { sent: false });
  assert.deepEqual(s.sent.map((x) => [x.binary, Buffer.isBuffer(x.data) ? [...x.data] : x.data]), [[false, 'text'], [true, [1, 2]], [true, [3, 4, 5]]]);

  // The server closes it: told once, and the number means nothing after.
  s.on.close({ code: 4004, reason: 'room closed' });
  s.on.close({ code: 1006, reason: '' });
  assert.deepEqual(told.at(-1), { win: 7, event: 'talk:door', socket: 1, type: 'close', code: 4004, reason: 'room closed' });
  assert.equal(told.filter((t) => t.type === 'close').length, 1);
  assert.deepEqual(api.doorSend({ socket, data: 'late', binary: false }, win), { sent: false });
  assert.deepEqual(servers.open(), { sessions: 1, sockets: 0 });
  // The engine connects again on the same way in (it does when the network changes).
  assert.deepEqual(api.doorOpen({ session, target: '/rtc/v1?access_token=join-1&reconnect=1' }, win), { socket: 2 });
});

test('a way in belongs to the window that opened it: no other window can use it, send on it or close it', async () => {
  const f = fakeDoor({ exchange: () => ({ door: {}, admitted: { token: 'j', room: 'r', identity: 'u-a', role: 'host', control: 'ctl' } }) });
  const { api, servers } = setUp({ routes: passRoute(MEETING.id), door: f.door });
  const win = windowOf(7);
  const other = windowOf(8);
  const { session } = await api.serverEnter({ id: MEETING.id }, win);
  const { socket } = api.doorOpen({ session, target: '/rtc' }, win);
  assert.deepEqual(api.doorOpen({ session, target: '/rtc' }, other), { error: 'no-session' });
  assert.deepEqual(api.doorOpen({ session: 99, target: '/rtc' }, win), { error: 'no-session' });
  assert.deepEqual(api.doorOpen({ session, target: '/rtc' }, null), { error: 'no-session' });
  assert.deepEqual(await api.serverControl({ session, do: 'mute-all' }, other), { error: 'no-session' });
  assert.deepEqual(api.doorSend({ socket, data: 'x', binary: false }, other), { sent: false });
  assert.deepEqual(api.doorClose({ socket }, other), { closed: false });
  assert.deepEqual(api.serverLeave({ session }, other), { left: false });
  assert.equal(f.seen.opened.length, 1);
  assert.equal(f.seen.opened[0].closed, null);
  assert.deepEqual(servers.open(), { sessions: 1, sockets: 1 });
  // A window holds a few connections at once, never many.
  for (let i = 0; i < 3; i += 1) assert.ok(api.doorOpen({ session, target: '/rtc' }, win).socket);
  assert.deepEqual(api.doorOpen({ session, target: '/rtc' }, win), { error: 'too-many' });
});

test('leaving closes the way in and its connections; so does the window closing, and signing out', async () => {
  const f = fakeDoor();
  const { api, servers, told } = setUp({ routes: passRoute(MEETING.id), door: f.door });
  const win = windowOf(7);
  const a = await api.serverEnter({ id: MEETING.id }, win);
  api.doorOpen({ session: a.session, target: '/rtc' }, win);
  assert.deepEqual(api.serverLeave({ session: a.session }, win), { left: true });
  assert.deepEqual(f.seen.opened[0].closed, { code: 1001, reason: '' });
  assert.deepEqual(servers.open(), { sessions: 0, sockets: 0 });
  assert.equal(told.filter((t) => t.type === 'close').length, 0, 'the window that left is not told of its own leaving');
  assert.deepEqual(api.doorOpen({ session: a.session, target: '/rtc' }, win), { error: 'no-session' });

  const b = await api.serverEnter({ id: MEETING.id }, win);
  const c = await api.serverEnter({ id: MEETING.id }, windowOf(8));
  api.doorOpen({ session: b.session, target: '/rtc' }, win);
  assert.equal(win.listenerCount('closed'), 1, 'a window is watched once however often it comes in');
  win.emit('closed');
  assert.deepEqual(servers.open(), { sessions: 1, sockets: 0 });
  assert.ok(f.seen.opened[1].closed);
  assert.ok(c.session);
  servers.forget();
  assert.deepEqual(servers.open(), { sessions: 0, sockets: 0 });
});

test('a connection the window closes is told back to it as closed, as any close is', async () => {
  const f = fakeDoor();
  const { api, told } = setUp({ routes: passRoute(MEETING.id), door: f.door });
  const win = windowOf(7);
  const { session } = await api.serverEnter({ id: MEETING.id }, win);
  const { socket } = api.doorOpen({ session, target: '/rtc' }, win);
  assert.deepEqual(api.doorClose({ socket, code: 1000, reason: 'done' }, win), { closed: true });
  assert.deepEqual(told.at(-1), { win: 7, event: 'talk:door', socket, type: 'close', code: 1000, reason: 'done' });
  assert.deepEqual(api.doorClose({ socket }, win), { closed: false });
});

// ── the whole way, with the real door code ───────────────────────────────────

test('from the pass to the signalling, through a door that shows the pinned certificate', async () => {
  const cert = selfSigned();
  const d = await standInDoor({ tls: cert, admit: (b) => (b.pass === 'pass-1-1' ? { status: 200, body: { token: 'join-real', room: 'r1', identity: 'u-a', role: 'host', e2ee: true, live: false, control: 'ctl-real' } } : { status: 403, body: { error: 'signature' } }), control: (b) => (b.control === 'ctl-real' ? { status: 200, body: { people: [{ id: 'u-a', name: 'Amira', muted: true }] } } : { status: 403, body: { error: 'control' } }) });
  doors.push(d);
  const hand = (pin) => ({ [`POST /servers/meetings/${MEETING.id}/pass`]: (_b, n) => [200, { pass: `pass-1-${n}`, expires: 2, server: { id: 'srv_1', doors: [{ family: 4, url: d.url, address: '127.0.0.1', media: 'udp' }], pin }, meeting: MEETING, role: 'host' }] });
  const { api, told, servers } = setUp({ routes: hand(cert.pin) });
  const win = windowOf(7);
  const got = await api.serverEnter({ id: MEETING.id }, win);
  assert.equal(got.token, 'join-real');
  assert.equal(got.host, true);
  assert.deepEqual(await api.serverControl({ session: got.session, do: 'roster' }, win), { people: [{ id: 'u-a', name: 'Amira', muted: true }] });
  const { socket } = api.doorOpen({ session: got.session, target: `/rtc/v1?access_token=${got.token}` }, win);
  await until(() => told.some((t) => t.type === 'open'));
  assert.deepEqual(api.doorSend({ socket, data: new Uint8Array([5, 6, 7]) }, win), { sent: true });
  await until(() => told.some((t) => t.type === 'message'));
  assert.deepEqual([...told.find((t) => t.type === 'message').data], [5, 6, 7]);
  assert.equal(d.sockets[0].target, '/rtc/v1?access_token=join-real');
  api.serverLeave({ session: got.session }, win);
  await until(() => d.sockets[0].closed);
  assert.deepEqual(servers.open(), { sessions: 0, sockets: 0 });

  // The same server, and the introducer names another certificate: nothing is sent to it.
  const before = d.hits.length;
  const wrong = setUp({ routes: hand(selfSigned().pin) });
  assert.deepEqual(await wrong.api.serverEnter({ id: MEETING.id }, windowOf(9)), { error: 'pin' });
  assert.equal(d.hits.length, before);
});
