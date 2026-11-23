// A meeting held on a Talk Server from start to end, as the windows of the
// people in it hold it (apps/desktop/renderer/apps/talk/meet-room.js with
// meet-sitting.js): the person who set it up brings it to a server and lets
// the others in, each gets the meeting's key sealed for them alone, and
// whatever stops someone is kept as one word the screen has a sentence for.
// Run on stand-ins: the introducer's room for the meeting (below), the
// server and its engine (talk-fake-room.js) and WebRTC.

import test from 'node:test';
import assert from 'node:assert/strict';
import { MeetingRoom } from '../apps/desktop/renderer/apps/talk/meet-room.js';
import { sittingId } from '../apps/desktop/renderer/apps/talk/meet-sitting.js';
import { serverFaultText, serverStepText } from '../apps/desktop/renderer/apps/talk/meet-lines.js';
import { fakeRtc } from './talk-fake-rtc.js';
import { fakeWorld, DisconnectReason, RoomEvent } from './talk-fake-room.js';

const ID = 'meet-AAAAAAAAAAAAAAAAAAAA';
const NAMES = { 'u-amira': 'Amira Khan', 'u-ben': 'Ben Ortiz', 'u-cara': 'Cara Lund', 'u-dev': 'Dev Rao' };
const STARTER = 'u-amira';
const SITTING = new RegExp(`^${ID}-[0-9a-f]{12}$`);

/**
 * One meeting at the introducer and on a server, and the windows of the
 * people in it. `carried: false` makes it a small meeting, as on main.
 */
function meetingOf({ carried = true, most = 20 } = {}) {
  const world = fakeWorld({ names: NAMES });
  const rtc = fakeRtc();
  /** Who is in the introducer's room: id -> their MeetingRoom. */
  const inside = new Map();
  /** Everything said to Talk's service, in order: { who, do, ... }. */
  const asked = world.asked;
  /** Every set-up word passed through the introducer's room. */
  const passed = [];
  /** The server and what the introducer keeps of it. `carry` and `admit` answer once in place of the usual; `pass` answers every asking for a pass while it is set. */
  const server = { sitting: null, over: new Set(), admitted: new Set(), carry: null, admit: null, pass: null };
  const rooms = {};
  const seen = {};
  const person = (id) => ({ id, name: NAMES[id] });
  const view = (but = null) => ({ id: ID, name: 'All hands', starter: person(STARTER), at: null, live: false, locked: false, groupId: null, link: `https://introducer.example/m/${ID}`, most: carried ? most : 4, ...(carried ? { carried: 'server' } : {}), people: [...inside.keys()].filter((id) => id !== but).map(person) });
  const tellOthers = (but, msg) => { for (const [id, room] of inside) if (id !== but) room.handle(msg); };
  const leave = (id) => { if (inside.delete(id)) tellOthers(id, { type: 'meet.person', id: ID, event: 'left', person: person(id) }); };

  const shellFor = (id) => {
    const base = world.shell(id, { role: id === STARTER ? 'host' : 'member' });
    const note = (what, p) => asked.push({ who: id, do: what, ...p });
    const talk = {
      ...base.talk,
      callSetup: async () => ({ iceServers: [] }),
      meetPeek: async () => ({ meeting: view() }),
      meetJoin: async (p) => {
        note('meetJoin', p);
        const meeting = view(id);
        tellOthers(id, { type: 'meet.person', id: ID, event: 'joined', person: person(id) });
        inside.set(id, rooms[id]);
        return { meeting };
      },
      meetLeave: async (p) => { note('meetLeave', p); leave(id); return { left: true }; },
      meetLock: async () => ({}),
      meetRemove: async (p) => {
        note('meetRemove', p);
        const out = inside.get(p.who);
        leave(p.who);
        out?.handle({ type: 'meet.removed', id: ID, by: person(id) });
        return { removed: true };
      },
      meetEnd: async (p) => {
        note('meetEnd', p);
        const all = [...inside];
        inside.clear();
        for (const [other, room] of all) if (other !== id) room.handle({ type: 'meet.ended', id: ID, by: person(id) });
        return { ended: true };
      },
      // The introducer passes a set-up word on only between two people who are both in the meeting.
      send: async (msg) => {
        if (msg.type !== 'meet.signal') return {};
        passed.push({ from: id, to: msg.to, data: msg.data });
        if (inside.has(id) && inside.has(msg.to)) setImmediate(() => inside.get(msg.to)?.handle({ type: 'meet.signal', id: msg.id, from: id, data: msg.data }));
        return {};
      },
      serverCarry: async (p) => {
        note('serverCarry', p);
        if (server.carry) { const r = server.carry; server.carry = null; return r; }
        if (server.over.has(p.id)) return { error: 'meeting-taken' };
        server.sitting = p.id;
        return { meeting: { id: p.id, live: false, e2ee: true, maxPeople: most, state: 'on' } };
      },
      serverAdmit: async (p) => {
        note('serverAdmit', p);
        if (server.admit) { const r = server.admit; server.admit = null; return r; }
        if (p.id !== server.sitting) return { error: 'no-meeting' };
        server.admitted.add(`${p.id} ${p.person}`);
        return { person: p.person, role: p.role || 'member' };
      },
      serverRemove: async (p) => { note('serverRemove', p); server.admitted.delete(`${p.id} ${p.person}`); return { removed: true }; },
      serverEnd: async (p) => { note('serverEnd', p); server.over.add(p.id); if (server.sitting === p.id) server.sitting = null; return { ended: true }; },
      serverEnter: async (p) => {
        if (server.pass) { note('serverEnter', p); return server.pass; }
        if (!world.enterRefusal && (p.id !== server.sitting || (id !== STARTER && !server.admitted.has(`${p.id} ${id}`)))) { note('serverEnter', p); return { error: 'no-pass' }; }
        return base.talk.serverEnter(p);
      },
    };
    return { on: base.on, talk };
  };

  /** A window for this person, not yet in the meeting. */
  const window = (id, { feed = null } = {}) => {
    seen[id] = { changes: 0 };
    rooms[id] = new MeetingRoom({ shell: shellFor(id), me: person(id), rtc, feed, engine: { load: world.load, worker: world.worker }, onChange: () => { seen[id].changes += 1; } });
    return rooms[id];
  };
  // Until nothing is on its way: every set-up word passed, every key opened, every room come into.
  const settle = async () => {
    for (let i = 0; i < 14; i += 1) {
      // A key pair is made and a key sealed off this thread, and each word takes a turn to pass: a few turns each time.
      for (let k = 0; k < 4; k += 1) {
        for (const r of Object.values(rooms)) { await r.sitting?.turn; await r.mesh?.line; await r.mesh?.outbox; await r.mesh?.inbox; }
        await new Promise((r) => setImmediate(r));
      }
      await new Promise((r) => setTimeout(r, 0));
    }
  };
  const join = async (id, opts = {}) => { const r = await window(id, opts).join(ID); await settle(); return r; };
  const did = (what, who = null) => asked.filter((a) => a.do === what && (!who || a.who === who));
  const closeAll = async () => { for (const r of Object.values(rooms)) r.leave(); await settle(); };
  return { world, rtc, inside, asked, passed, server, rooms, seen, window, join, settle, did, closeAll };
}
const others = (room) => room.people().filter((p) => !p.me);

test('the person who set it up brings the meeting to a server, with an id for this sitting and a key of their own making', async () => {
  const m = meetingOf();
  const r = await m.join('u-amira');
  assert.equal(r.meeting.carried, 'server');
  const room = m.rooms['u-amira'];
  assert.equal(room.state, 'in');
  assert.equal(room.server, 'in');
  const carry = m.did('serverCarry');
  assert.equal(carry.length, 1);
  assert.match(carry[0].id, SITTING);
  assert.equal(room.sitting.mid, carry[0].id);
  assert.equal(m.did('serverEnter')[0].id, carry[0].id, 'into the room of that sitting');
  assert.ok(room.sitting.key instanceof Uint8Array && room.sitting.key.length === 32);
  assert.equal(m.world.keys[0], Buffer.from(room.sitting.key).toString('base64url'), 'the engine encrypts with that key');
  // Nobody made a connection to anybody.
  assert.equal(m.rtc.pcs.length, 0);
  assert.notEqual(sittingId(ID), sittingId(ID));
  await m.closeAll();
});

test('someone who joins asks the person who set it up, is let in at the server and handed the key sealed for them', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  const [amira, ben] = [m.rooms['u-amira'], m.rooms['u-ben']];
  assert.equal(ben.server, 'in');
  const mid = amira.sitting.mid;
  assert.deepEqual(m.did('serverAdmit').map((a) => [a.who, a.id, a.person]), [['u-amira', mid, 'u-ben']]);
  assert.equal(ben.sitting.mid, mid);
  assert.deepEqual([...ben.sitting.key], [...amira.sitting.key], 'the same key');
  assert.equal(m.world.keys[1], m.world.keys[0]);
  assert.deepEqual(m.passed.map((p) => [p.from, p.to, p.data.t]), [['u-amira', 'u-ben', 'srv-moved'], ['u-ben', 'u-amira', 'srv-ask'], ['u-amira', 'u-ben', 'srv-let']].slice(1), 'asked once, answered once');
  // The key itself is in nothing that went through the introducer.
  const through = JSON.stringify(m.passed);
  assert.equal(through.includes(m.world.keys[0]), false);
  assert.equal(through.includes(Buffer.from(amira.sitting.key).toString('base64')), false);
  // Each sees the other, connected through the server.
  assert.deepEqual(others(amira).map((p) => [p.id, p.name, p.connected, p.starter]), [['u-ben', 'Ben Ortiz', true, false]]);
  assert.deepEqual(others(ben).map((p) => [p.id, p.name, p.connected, p.starter]), [['u-amira', 'Amira Khan', true, true]]);
  assert.equal(m.rtc.pcs.length, 0);
  await m.closeAll();
});

test('what is written reaches everyone in the server\'s room, and a third person is let in like the second', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  await m.join('u-cara');
  assert.deepEqual(['u-amira', 'u-ben', 'u-cara'].map((id) => m.rooms[id].server), ['in', 'in', 'in']);
  m.rooms['u-cara'].say('Hello all');
  await m.settle();
  assert.deepEqual(m.rooms['u-amira'].says.map((s) => [s.from, s.text]), [['u-cara', 'Hello all']]);
  assert.deepEqual(m.rooms['u-ben'].says.map((s) => [s.name, s.text]), [['Cara Lund', 'Hello all']]);
  assert.equal(others(m.rooms['u-ben']).length, 2);
  await m.closeAll();
});

test('nobody can come in until the person who set it up is in the meeting: they wait, and are let in when that person arrives', async () => {
  const m = meetingOf();
  await m.join('u-ben');
  const ben = m.rooms['u-ben'];
  assert.equal(ben.state, 'in', 'in the meeting at the introducer');
  assert.equal(ben.server, 'waiting');
  assert.equal(m.passed.length, 0, 'nobody to ask');
  assert.equal(m.did('serverEnter').length, 0);
  assert.equal(serverStepText(ben.server, ben.meeting.starter), 'Waiting for Amira Khan, who lets people into this meeting.');

  await m.join('u-amira');
  assert.equal(m.rooms['u-amira'].server, 'in');
  assert.equal(ben.server, 'in');
  assert.ok(m.passed.some((p) => p.from === 'u-amira' && p.to === 'u-ben' && p.data.t === 'srv-moved'), 'told to ask now');
  assert.equal(ben.sitting.mid, m.rooms['u-amira'].sitting.mid);
  await m.closeAll();
});

test('only the person who set the meeting up lets people in: a key from anyone else is not taken', async () => {
  const m = meetingOf();
  await m.join('u-ben');
  await m.join('u-cara');
  const ben = m.rooms['u-ben'];
  // Cara's app says what only Amira's may.
  ben.handle({ type: 'meet.signal', id: ID, from: 'u-cara', data: { t: 'srv-let', mid: `${ID}-000000000000`, pub: 'AAAA', iv: 'AAAA', box: 'AAAA' } });
  ben.handle({ type: 'meet.signal', id: ID, from: 'u-cara', data: { t: 'srv-moved' } });
  ben.handle({ type: 'meet.signal', id: ID, from: 'u-cara', data: { t: 'srv-no', why: 'full' } });
  await m.settle();
  assert.equal(ben.server, 'waiting');
  assert.equal(ben.state, 'in');
  assert.equal(ben.sitting.key, null);
  assert.equal(m.did('serverEnter').length, 0);
  // And somebody who is not the starter lets nobody in.
  m.rooms['u-cara'].handle({ type: 'meet.signal', id: ID, from: 'u-ben', data: { t: 'srv-ask', pub: 'AAAA' } });
  await m.settle();
  assert.equal(m.did('serverAdmit').length, 0);
  await m.closeAll();
});

test('an answer sealed for another asking opens nothing', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  const let1 = m.passed.find((p) => p.data.t === 'srv-let');
  // Cara asks, and Amira's app is still at the server about her: she is handed what was sealed for Ben.
  m.server.admit = new Promise(() => {});
  await m.join('u-cara');
  const cara = m.rooms['u-cara'];
  assert.equal(cara.server, 'asking');
  cara.handle({ type: 'meet.signal', id: ID, from: 'u-amira', data: let1.data });
  await m.settle();
  assert.equal(cara.server, 'asking', 'still waiting for an answer of her own');
  assert.equal(cara.sitting.key, null);
  assert.equal(m.did('serverEnter', 'u-cara').length, 0);
  await m.closeAll();
});

test('taken out by the person who set it up: out at the introducer and out of the server\'s room, with no way back in', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  const [amira, ben] = [m.rooms['u-amira'], m.rooms['u-ben']];
  const mid = amira.sitting.mid;
  assert.deepEqual(await amira.remove('u-ben'), { removed: true });
  await m.settle();
  assert.deepEqual(m.did('serverRemove').map((a) => [a.id, a.person]), [[mid, 'u-ben']]);
  assert.equal(ben.state, 'gone');
  assert.equal(ben.why, 'removed');
  assert.equal(ben.by.name, 'Amira Khan');
  assert.equal(m.world.inside.has('u-ben'), false, 'out of the room on the server');
  assert.deepEqual(others(amira), []);
  await m.closeAll();
});

test('ended by the person who set it up: on the server first, then at the introducer, and everyone is out', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  const mid = m.rooms['u-amira'].sitting.mid;
  assert.deepEqual(await m.rooms['u-amira'].end(), { ended: true });
  await m.settle();
  const order = m.asked.filter((a) => a.do === 'serverEnd' || a.do === 'meetEnd').map((a) => a.do);
  assert.deepEqual(order, ['serverEnd', 'meetEnd']);
  assert.equal(m.did('serverEnd')[0].id, mid);
  assert.deepEqual([m.rooms['u-amira'].why, m.rooms['u-ben'].why], ['ended', 'ended']);
  assert.equal(m.world.inside.size, 0);
  assert.ok(m.world.workers.every((w) => w.stopped), 'every worker stopped');
});

test('no server can hold it: the person who set it up is told why, and a meeting nobody else is in is ended', async () => {
  const m = meetingOf();
  m.server.carry = { error: 'no-server', own: [{ id: 's-1', name: 'Office box', why: 'not-answering' }] };
  await m.join('u-amira');
  const room = m.rooms['u-amira'];
  assert.equal(room.state, 'gone');
  assert.equal(room.why, 'server');
  assert.deepEqual(room.fault, { why: 'no-server', mine: true, own: [{ id: 's-1', name: 'Office box', why: 'not-answering' }] });
  assert.equal(serverFaultText(room.fault, room.meeting.starter), 'No Talk Server can hold this meeting now. Office box is not answering. Check that it is running and connected.');
  assert.equal(m.did('meetEnd').length, 1, 'nobody who was asked waits in it');
  assert.equal(m.did('serverEnter').length, 0);
});

test('the service cannot be reached, has no servers, or wants a sign-in: each by its word', async () => {
  for (const [answer, why] of [[{ error: 'offline' }, 'offline'], [{ error: 'no-service' }, 'no-service'], [{ error: 'sign-in' }, 'sign-in'], [{ error: 'refused' }, 'refused'], [{}, 'failed']]) {
    const m = meetingOf();
    m.server.carry = answer;
    await m.join('u-amira');
    assert.equal(m.rooms['u-amira'].why, 'server');
    assert.equal(m.rooms['u-amira'].fault.why, why);
    assert.ok(serverFaultText(m.rooms['u-amira'].fault).length > 20);
  }
});

test('with others already in the meeting, a starter who cannot bring it to a server leaves it and does not end it', async () => {
  const m = meetingOf();
  await m.join('u-ben');
  m.server.carry = { error: 'no-server', own: [] };
  await m.join('u-amira');
  assert.equal(m.rooms['u-amira'].why, 'server');
  assert.equal(m.did('meetEnd').length, 0);
  assert.equal(m.did('meetLeave', 'u-amira').length, 1);
  assert.equal(m.rooms['u-ben'].state, 'in');
  assert.equal(m.rooms['u-ben'].server, 'waiting');
  await m.closeAll();
});

test('the door does not check out, or the server turns the connection away: out, with the word for it, and nothing left open', async () => {
  const cases = [
    [(w) => { w.enterRefusal = { error: 'pin' }; }, 'pin', 'Talk could not confirm that the Talk Server is the one this meeting is held on, so it did not connect.'],
    [(w) => { w.enterRefusal = { error: 'certificate' }; }, 'certificate', null],
    [(w) => { w.enterRefusal = { error: 'plain' }; }, 'plain', null],
    [(w) => { w.doorRefusal = { status: 500, why: 'refused', reason: 'connection closed by media' }; w.doorRefusalStays = true; }, 'turned-away', 'The Talk Server did not take the connection. The meeting may be full. Try again in a moment.'],
    [(w) => { w.doorRefusal = { status: 403, why: 'refused' }; }, 'not-admitted', 'The Talk Server did not let you in. Join again, so that Amira Khan can let you in.'],
    [(w) => { w.doorRefusal = { status: 429, why: 'refused' }; }, 'full', 'The meeting is full on its Talk Server.'],
  ];
  for (const [set, why, words] of cases) {
    const m = meetingOf();
    await m.join('u-amira');
    set(m.world);
    await m.join('u-ben');
    const ben = m.rooms['u-ben'];
    assert.equal(ben.state, 'gone', why);
    assert.equal(ben.why, 'server');
    assert.deepEqual(ben.fault, { why, mine: false });
    if (words) assert.equal(serverFaultText(ben.fault, ben.meeting.starter), words);
    assert.equal(m.did('meetLeave', 'u-ben').length, 1, 'out of the meeting at the introducer too');
    assert.equal(m.world.inside.has('u-ben'), false);
    assert.equal(m.rooms['u-amira'].server, 'in', 'the meeting goes on for the others');
    await m.closeAll();
  }
});

test('turned away by the engine after the door let them in: asked again by itself, where the door says if the meeting is full', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  m.world.doorRefusal = { status: 500, why: 'refused', reason: 'connection closed by media' };
  await m.join('u-ben');
  assert.equal(m.rooms['u-ben'].server, 'in', 'in at the second try');
  assert.equal(m.rooms['u-ben'].fault, null);
  assert.equal(m.passed.filter((p) => p.from === 'u-ben' && p.data.t === 'srv-ask').length, 2);
  await m.closeAll();
});

test('the meeting is full: the door says so with how many it holds, and the person is told once, with nothing tried again', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  m.server.pass = { error: 'full', max: 20 };
  await m.join('u-ben');
  m.server.pass = null;
  const ben = m.rooms['u-ben'];
  assert.equal(ben.state, 'gone');
  assert.equal(ben.why, 'server');
  assert.deepEqual(ben.fault, { why: 'full', mine: false, max: 20 });
  assert.equal(serverFaultText(ben.fault, ben.meeting.starter), 'This meeting is full: it holds 20 people.');
  assert.equal(m.passed.filter((p) => p.from === 'u-ben' && p.data.t === 'srv-ask').length, 1, 'asked once and no more');
  assert.equal(m.did('serverEnter', 'u-ben').length, 1);
  assert.equal(m.did('meetLeave', 'u-ben').length, 1, 'out of the meeting at the introducer too');
  assert.equal(m.rooms['u-amira'].server, 'in', 'the meeting goes on for the others');
  // A door that does not say how many still has a sentence.
  assert.equal(serverFaultText({ why: 'full', mine: false }), 'The meeting is full on its Talk Server.');
  await m.closeAll();
});

test('the meeting is over: said in the meeting\'s own words, and nothing is tried again', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  m.server.pass = { error: 'ended' };
  await m.join('u-ben');
  m.server.pass = null;
  const ben = m.rooms['u-ben'];
  assert.equal(ben.state, 'gone');
  assert.equal(ben.why, 'ended', 'as when the person who set it up ends it');
  assert.equal(ben.fault, null);
  assert.equal(m.passed.filter((p) => p.from === 'u-ben' && p.data.t === 'srv-ask').length, 1, 'asked once and no more');
  assert.equal(m.did('serverEnter', 'u-ben').length, 1);
  assert.equal(m.did('meetLeave', 'u-ben').length, 1);
  await m.closeAll();

  // Over under the person who set it up, as they come in: the meeting ends for everyone, so nobody waits in it.
  const n = meetingOf();
  n.server.pass = { error: 'ended' };
  await n.join('u-amira');
  n.server.pass = null;
  assert.equal(n.rooms['u-amira'].state, 'gone');
  assert.equal(n.rooms['u-amira'].why, 'ended');
  assert.equal(n.did('meetEnd', 'u-amira').length, 1);
  assert.equal(n.did('serverCarry', 'u-amira').length, 1, 'no server asked for a second time');
  await n.closeAll();
});

test('the meeting went on under a new id: the person who set it up is asked again, and tells it', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  m.world.enterRefusal = { error: 'ended', moved: true };
  await m.join('u-ben');
  assert.equal(m.rooms['u-ben'].server, 'in', 'in at the second asking');
  assert.equal(m.passed.filter((p) => p.from === 'u-ben' && p.data.t === 'srv-ask').length, 2);

  // Told so every time: asked twice more, then the meeting is over for this person.
  const n = meetingOf();
  await n.join('u-amira');
  n.server.pass = { error: 'ended', moved: true };
  await n.join('u-cara');
  // Three askings, each with a key made and sealed off this thread: waited for, however long they take.
  for (let i = 0; i < 40 && n.rooms['u-cara'].state !== 'gone'; i += 1) await n.settle();
  n.server.pass = null;
  assert.equal(n.rooms['u-cara'].state, 'gone');
  assert.equal(n.rooms['u-cara'].why, 'ended');
  assert.equal(n.passed.filter((p) => p.from === 'u-cara' && p.data.t === 'srv-ask').length, 3, 'once, and twice more');
  await m.closeAll();
  await n.closeAll();
});

test('a pass that ran out is asked for again by itself, and only then is the person told', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  m.world.enterRefusal = { error: 'expired' };
  await m.join('u-ben');
  assert.equal(m.rooms['u-ben'].server, 'in', 'in at the second try');
  assert.equal(m.passed.filter((p) => p.data.t === 'srv-ask').length, 2);

  // Every time: told after the tries are spent.
  const n = meetingOf();
  await n.join('u-amira');
  n.server.pass = { error: 'expired' };
  await n.join('u-cara');
  const cara = n.rooms['u-cara'];
  n.server.pass = null;
  assert.equal(cara.why, 'server');
  assert.equal(cara.fault.why, 'expired');
  assert.equal(serverFaultText(cara.fault), 'Your pass for this meeting ran out before it could be used. Try again.');
  assert.equal(n.passed.filter((p) => p.from === 'u-cara' && p.data.t === 'srv-ask').length, 3, 'once, and twice more');
  await m.closeAll();
  await n.closeAll();
});

test('the starter\'s app cannot let someone in: they are told so, by the word it sent', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  m.server.admit = { error: 'full' };
  await m.join('u-ben');
  const ben = m.rooms['u-ben'];
  assert.equal(ben.why, 'server');
  assert.deepEqual(ben.fault, { why: 'full', mine: false });
  assert.ok(m.passed.some((p) => p.data.t === 'srv-no' && p.data.why === 'full'));

  // A refusal with a word the apps do not share goes as 'refused'.
  const n = meetingOf();
  await n.join('u-amira');
  n.server.admit = { error: 'banned' };
  await n.join('u-ben');
  assert.equal(n.rooms['u-ben'].fault.why, 'refused');
  assert.equal(serverFaultText(n.rooms['u-ben'].fault, n.rooms['u-ben'].meeting.starter), 'Amira Khan\'s Talk could not let you into the meeting. Try joining again.');
  await m.closeAll();
  await n.closeAll();
});

test('the sitting is over on its server when someone asks: it is brought to one anew, and everybody is let into the new one', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  const amira = m.rooms['u-amira'];
  const first = amira.sitting.mid;
  const key = [...amira.sitting.key];
  // The server ended that sitting by itself (its own page, or everybody had been away).
  m.server.over.add(first);
  m.server.sitting = null;
  await m.join('u-cara');
  const second = amira.sitting.mid;
  assert.notEqual(second, first);
  assert.match(second, SITTING);
  assert.notDeepEqual([...amira.sitting.key], key, 'a new sitting has a new key');
  assert.deepEqual(['u-amira', 'u-ben', 'u-cara'].map((id) => [m.rooms[id].server, m.rooms[id].sitting.mid]), [['in', second], ['in', second], ['in', second]]);
  await m.closeAll();
});

test('the connection to the server is lost: in again with a new pass and the same key, with nobody asked', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  const ben = m.rooms['u-ben'];
  const asks = m.passed.length;
  m.world.putOut('u-ben', DisconnectReason.UNKNOWN_REASON);
  await m.settle();
  assert.equal(ben.state, 'in');
  assert.equal(ben.server, 'in');
  assert.equal(m.passed.length, asks, 'the key is the one they hold');
  assert.equal(m.did('serverEnter', 'u-ben').length, 2);
  assert.equal(others(m.rooms['u-amira']).length, 1);
  await m.closeAll();
});

test('the same person comes in on another computer: this one is out, as in a small meeting', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  m.world.putOut('u-ben', DisconnectReason.DUPLICATE_IDENTITY);
  await m.settle();
  assert.equal(m.rooms['u-ben'].state, 'gone');
  assert.equal(m.rooms['u-ben'].why, 'elsewhere');
  await m.closeAll();
});

test('the server ends the sitting under the person who set it up: the meeting is over for everyone', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  m.world.putOut('u-amira', DisconnectReason.ROOM_DELETED);
  await m.settle();
  assert.equal(m.rooms['u-amira'].why, 'ended');
  assert.equal(m.rooms['u-ben'].why, 'ended');
  assert.equal(m.did('meetEnd').length, 1);
});

test('the connection to the introducer is made again: back into its room, and the server\'s room is left as it is', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  const ben = m.rooms['u-ben'];
  const before = { enters: m.did('serverEnter').length, room: ben.mesh.room };
  m.inside.delete('u-ben');
  ben.handle({ type: 'welcome' });
  await m.settle();
  assert.equal(m.did('meetJoin', 'u-ben').length, 2);
  assert.equal(ben.state, 'in');
  assert.equal(ben.server, 'in');
  assert.equal(ben.mesh.room, before.room, 'the same room on the server');
  assert.equal(m.did('serverEnter').length, before.enters);
  assert.deepEqual(others(ben).map((p) => [p.id, p.connected]), [['u-amira', true]]);

  // The starter's connection made again: whoever asked meanwhile is told to ask again.
  const amira = m.rooms['u-amira'];
  const told = m.passed.filter((p) => p.data.t === 'srv-moved').length;
  m.inside.delete('u-amira');
  amira.handle({ type: 'welcome' });
  await m.settle();
  assert.equal(amira.server, 'in');
  assert.equal(m.did('serverCarry').length, 1, 'not brought to a server again');
  assert.equal(m.passed.filter((p) => p.data.t === 'srv-moved').length, told + 1);
  assert.equal(ben.server, 'in');
  await m.closeAll();
});

test('the person who set it up leaves and comes back: a new sitting, and those still in the meeting move to it', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  const first = m.rooms['u-amira'].sitting.mid;
  m.rooms['u-amira'].leave();
  await m.settle();
  const ben = m.rooms['u-ben'];
  assert.equal(ben.state, 'in', 'the meeting goes on');
  assert.equal(ben.server, 'in', 'and so does its room on the server');
  // Someone new meanwhile waits.
  await m.join('u-cara');
  assert.equal(m.rooms['u-cara'].server, 'waiting');

  await m.join('u-amira');
  const second = m.rooms['u-amira'].sitting.mid;
  assert.notEqual(second, first);
  assert.deepEqual(['u-ben', 'u-cara'].map((id) => [m.rooms[id].server, m.rooms[id].sitting.mid]), [['in', second], ['in', second]]);
  assert.equal(others(m.rooms['u-amira']).filter((p) => p.connected).length, 2);
  await m.closeAll();
});

test('nothing anybody sends opens here: the key is asked for once more, by itself', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  const ben = m.rooms['u-ben'];
  const asks = () => m.passed.filter((p) => p.from === 'u-ben' && p.data.t === 'srv-ask').length;
  assert.equal(asks(), 1);
  const err = new Error('decryption failed');
  ben.mesh.room.emit(RoomEvent.EncryptionError, err, { identity: 'u-amira' });
  await m.settle();
  assert.equal(asks(), 2);
  assert.equal(ben.server, 'in');
  // And only once: after that the screen says it.
  ben.mesh.room.emit(RoomEvent.EncryptionError, err, { identity: 'u-amira' });
  await m.settle();
  assert.equal(asks(), 2);
  assert.equal(ben.mesh.unreadable, true);
  assert.equal(ben.mesh.unreadableAll, true);
  assert.equal(ben.people().find((p) => p.id === 'u-amira').unreadable, true);
  await m.closeAll();
});

test('a meeting with no server is held exactly as before: a connection to each person, and nothing asked about servers', async () => {
  const m = meetingOf({ carried: false });
  await m.join('u-amira');
  await m.join('u-ben');
  const [amira, ben] = [m.rooms['u-amira'], m.rooms['u-ben']];
  assert.equal(amira.sitting, null);
  assert.equal(amira.server, null);
  assert.equal(ben.server, null);
  assert.equal(amira.mesh.constructor.name, 'Mesh');
  assert.equal(ben.meeting.most, 4);
  assert.deepEqual(m.asked.filter((a) => a.do.startsWith('server') || a.do.startsWith('door')), []);
  assert.ok(m.passed.some((p) => p.data.kind === 'offer'), 'set up between the two computers');
  assert.equal(m.rtc.pcs.length > 0, true);
  await m.closeAll();
});

test('a broadcast\'s or a recorder\'s seat does not hold a meeting on a server: refused with its own word, and out of it', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  const feed = async () => ({ audio: null, video: null });
  const r = await m.join('u-ben', { feed });
  assert.deepEqual(r, { error: 'carried' });
  assert.equal(m.rooms['u-ben'].state, 'gone');
  assert.equal(m.rooms['u-ben'].mesh, null);
  assert.equal(m.did('meetLeave', 'u-ben').length, 1);
  assert.equal(m.did('serverEnter', 'u-ben').length, 0);
  await m.closeAll();
});

test('leaving gives everything back: the room on the server, the way in, the worker', async () => {
  const m = meetingOf();
  await m.join('u-amira');
  await m.join('u-ben');
  m.rooms['u-ben'].leave();
  await m.settle();
  assert.equal(m.world.inside.has('u-ben'), false);
  assert.equal(m.did('serverLeave', 'u-ben').length, 1);
  assert.equal(m.did('meetLeave', 'u-ben').length, 1);
  assert.equal(m.rooms['u-ben'].sitting.stopped, true);
  assert.deepEqual(others(m.rooms['u-amira']), []);
  await m.closeAll();
  assert.ok(m.world.workers.every((w) => w.stopped));
});
