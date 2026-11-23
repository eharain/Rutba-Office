// Talk rings (Office 2.27.0, the owner's yes of 22 Nov 2026): Talk connects
// when Office starts unless the person turned that off, says what is worth a
// notification as 'talk:notify' (it shows none itself), keeps a call that is
// ringing for a window opened late, and shows the people blocked so one can
// be let back. On a stand-in introducer.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { standIn, BEN, CARA } from './talk-stand-in.js';
import { START_KEY, RING_KEY } from '../apps/desktop/main/talk.js';
import { NOTIFY_EVENT, NOTIFY_KINDS, NOTIFY_TEXT_MOST, createNotify, meetingInEvent } from '../apps/desktop/main/talk-notify.js';
import { EVENTS, EVENT_PACK, METHODS } from '../packages/office-shell/src/contract.js';

const made = [];
const talk = (o) => { const it = standIn(o); made.push(it); return it; };
test.after(() => { for (const it of made) it.clean(); });

const GROUP = { id: 'g_1', name: '', starter: 'u-a', created: 5, members: [{ id: 'u-a', name: 'Amira Khan' }, { id: 'u-ben', name: 'Ben Ortiz', contact: true }, { id: 'u-dan', name: 'Dan Abel' }] };
const CALL = { type: 'call.incoming', callId: 'call-0001', from: BEN, sdp: 'v=0', video: true };

test('Talk connects as Office starts, with no window asking, when someone is signed in', async () => {
  const it = talk();
  assert.equal(it.svc.state().startWithOffice, true);
  let started;
  await it.online(() => { started = it.svc.start(); });
  assert.deepEqual(started, { started: true });
  assert.equal(it.svc.state().phase, 'online');
  // Started twice, or a window then asking too, is still the one connection.
  it.svc.start();
  it.svc.resume();
  await it.tick();
  assert.equal(it.connections(), 1);
});

test('turned off, Talk does not connect at start, and connects when its window is opened', async () => {
  const it = talk({ settings: new Map([[START_KEY, false]]) });
  assert.equal(it.svc.state().startWithOffice, false);
  assert.deepEqual(it.svc.start(), { started: false });
  assert.equal(it.connections(), 0);
  assert.equal(it.svc.state().phase, 'signed-out');
  await it.online();
  assert.equal(it.connections(), 1);
});

test('with nobody signed in, starting does nothing', () => {
  const it = talk({ signedIn: false });
  assert.deepEqual(it.svc.start(), { started: false });
  assert.equal(it.connections(), 0);
});

test('the two choices are kept, told to every window, and nothing else is taken', async () => {
  const it = talk();
  assert.deepEqual([it.svc.state().startWithOffice, it.svc.state().ring], [true, true]);
  const state = it.svc.setOptions({ startWithOffice: false, ring: false, server: 'wss://elsewhere.example/ws' });
  assert.deepEqual([state.startWithOffice, state.ring], [false, false]);
  assert.deepEqual([it.settings.get(START_KEY), it.settings.get(RING_KEY)], [false, false]);
  assert.equal(it.settings.has('server'), false);
  assert.equal(it.told.at(-1).event, 'talk:state');
  assert.equal(it.told.at(-1).payload.ring, false);
  // Something that is not yes or no changes nothing.
  it.svc.setOptions({ ring: 'yes' });
  assert.equal(it.svc.state().ring, false);
  // A second service on the same settings starts as the person left it.
  const again = talk({ settings: it.settings });
  assert.deepEqual(again.svc.start(), { started: false });
});

test('a message is told once, with the name of who wrote it; one given again is not told again', async () => {
  const it = talk();
  await it.online();
  it.say({ type: 'messages', messages: [{ id: 'm-00000001', from: 'u-ben', text: 'Hello', at: 10 }] });
  it.say({ type: 'messages', messages: [{ id: 'm-00000001', from: 'u-ben', text: 'Hello', at: 10 }, { id: 'm-00000002', from: 'u-ben', text: 'x'.repeat(1000), at: 11 }] });
  await it.tick();
  const said = it.notified();
  assert.deepEqual(said.map((n) => [n.kind, n.id]), [['message', 'm-00000001'], ['message', 'm-00000002']]);
  assert.deepEqual(said[0], { kind: 'message', id: 'm-00000001', from: { id: 'u-ben', name: 'Ben Ortiz' }, group: null, text: 'Hello', at: 10 });
  assert.equal(said[1].text.length, NOTIFY_TEXT_MOST);
});

test('a message in a group names the group, by the others in it when it has no name, and someone who is not a contact by name', async () => {
  const it = talk({ after: [{ type: 'groups', groups: [GROUP] }] });
  await it.online();
  it.say({ type: 'group.messages', messages: [{ id: 'm-00000010', group: 'g_1', from: 'u-dan', text: 'Ready?', at: 20 }] });
  await it.tick();
  assert.deepEqual(it.notified(), [{ kind: 'message', id: 'm-00000010', from: { id: 'u-dan', name: 'Dan Abel' }, group: { id: 'g_1', name: 'Ben, Dan' }, text: 'Ready?', at: 20 }]);
});

test('a call that rings is told, kept for a window opened late, and told over when answered here', async () => {
  const it = talk();
  await it.online();
  it.say(CALL);
  // Said twice by a careless introducer: the one call.
  it.say(CALL);
  await it.tick();
  assert.deepEqual(it.notified(), [{ kind: 'call', id: 'call-0001', from: { id: 'u-ben', name: 'Ben Ortiz' }, video: true }]);
  // What a window opened now needs to show the call and answer it.
  assert.deepEqual(it.svc.ringing(), [CALL]);
  it.svc.send({ type: 'call.answer', callId: 'call-0001', sdp: 'v=0' });
  assert.deepEqual(it.notified().at(-1), { kind: 'call-over', id: 'call-0001', why: 'answered' });
  assert.deepEqual(it.svc.ringing(), []);
  assert.equal(it.svc.state().missed, 0);
});

test('a call turned away, answered on another computer, or missed says which; a missed one is told as missed, with a count', async () => {
  const it = talk();
  await it.online();
  it.say({ ...CALL, callId: 'call-0002' });
  await it.tick();
  it.svc.send({ type: 'call.decline', callId: 'call-0002' });
  assert.deepEqual(it.notified().at(-1), { kind: 'call-over', id: 'call-0002', why: 'declined' });

  it.say({ ...CALL, callId: 'call-0003' });
  it.say({ type: 'call.taken', callId: 'call-0003' });
  await it.tick();
  assert.deepEqual(it.notified().at(-1), { kind: 'call-over', id: 'call-0003', why: 'elsewhere' });
  assert.equal(it.svc.state().missed, 0);

  it.say({ ...CALL, callId: 'call-0004', video: false });
  it.say({ type: 'call.ended', callId: 'call-0004' });
  await it.tick();
  assert.deepEqual(it.notified().slice(-2).map((n) => [n.kind, n.id, n.why ?? null, n.count ?? null]), [['call-over', 'call-0004', 'missed', null], ['missed', 'u-ben', 'unanswered', 1]]);
  assert.deepEqual(it.notified().at(-1).from, { id: 'u-ben', name: 'Ben Ortiz' });

  // Calling again before the missed calls were looked at: the same entry, told again with what it now counts.
  it.say({ ...CALL, callId: 'call-0005' });
  await it.tick();
  it.svc.send({ type: 'call.decline', callId: 'call-0005', reason: 'busy' });
  const last = it.notified().slice(-2);
  assert.deepEqual([last[0].kind, last[0].why], ['call-over', 'missed']);
  assert.deepEqual([last[1].kind, last[1].id, last[1].callId, last[1].count, last[1].why], ['missed', 'u-ben', 'call-0005', 2, 'busy']);
  assert.equal(it.svc.state().missed, 1);
});

test('a call made while this computer was off is told as missed when the introducer says so', async () => {
  const it = talk({ after: [{ type: 'notices', notices: [{ id: 'n_1', kind: 'missed-call', callId: 'call-0100', from: { id: 'u-cara', name: 'Cara Lund' }, video: false, why: 'offline', at: 40, count: 3 }] }] });
  await it.online();
  await it.tick();
  assert.deepEqual(it.notified(), [{ kind: 'missed', id: 'u-cara', callId: 'call-0100', from: { id: 'u-cara', name: 'Cara Lund' }, video: false, why: 'offline', at: 40, count: 3 }]);
});

test('a call ringing when the connection is lost, or when the person signs out, no longer rings', async () => {
  const it = talk();
  await it.online();
  it.say({ ...CALL, callId: 'call-0006' });
  await it.tick();
  it.drop();
  await it.online();
  assert.deepEqual(it.notified().at(-1), { kind: 'call-over', id: 'call-0006', why: 'ended' });
  assert.deepEqual(it.svc.ringing(), []);
  it.say({ ...CALL, callId: 'call-0007' });
  await it.tick();
  await it.svc.signOut();
  assert.deepEqual(it.notified().at(-1), { kind: 'call-over', id: 'call-0007', why: 'ended' });
});

test('being asked into a meeting is told, with the live notice when it will be streamed, and told over when joined, put off or ended', async () => {
  const meeting = (id, live) => ({ id, name: 'Weekly', starter: { id: 'u-ben', name: 'Ben Ortiz' }, live, groupId: null, link: `https://introducer.example/m/${id}`, people: [] });
  const a = 'meet-aaaaaaaaaaaaaaaaaaaa';
  const b = 'live-bbbbbbbbbbbbbbbbbbbb';
  const c = 'meet-cccccccccccccccccccc';
  const it = talk({ answer: (msg) => (msg.type === 'meet.join' ? { type: 'meet.joined', ref: msg.ref, meeting: meeting(msg.id, false) } : null) });
  await it.online();
  it.say({ type: 'meet.invited', meeting: meeting(a, false), by: { id: 'u-ben', name: 'Ben Ortiz' } });
  it.say({ type: 'meet.invited', meeting: meeting(b, true), by: { id: 'u-ben', name: 'Ben Ortiz' } });
  it.say({ type: 'meet.invited', meeting: meeting(c, false), by: { id: 'u-ben', name: 'Ben Ortiz' } });
  await it.tick();
  assert.deepEqual(it.notified(), [a, b, c].map((id) => ({ kind: 'meeting', id, from: { id: 'u-ben', name: 'Ben Ortiz' }, name: 'Weekly', live: id === b, recorded: false, group: null })));
  await it.svc.meetJoin({ id: a });
  assert.deepEqual(it.notified().at(-1), { kind: 'meeting-over', id: a, why: 'joined' });
  it.svc.meetDismiss({ id: b });
  assert.deepEqual(it.notified().at(-1), { kind: 'meeting-over', id: b, why: 'dismissed' });
  // Put off twice says it once.
  it.svc.meetDismiss({ id: b });
  assert.equal(it.notified().filter((n) => n.kind === 'meeting-over').length, 2);
  it.say({ type: 'meet.ended', id: c });
  await it.tick();
  assert.deepEqual(it.notified().at(-1), { kind: 'meeting-over', id: c, why: 'ended' });
  assert.deepEqual(it.svc.state().meetings, []);
});

test('what is told is one of the kinds written down, and the windows may hear it', () => {
  assert.ok(EVENTS.includes(NOTIFY_EVENT));
  assert.equal(EVENT_PACK[NOTIFY_EVENT], 'talk');
  for (const m of ['setOptions', 'ringing', 'unblock']) assert.ok(METHODS.talk.includes(m), m);
  const heard = [];
  const notify = createNotify({ broadcast: (event, n) => heard.push([event, n]), person: (id) => ({ id, name: '' }), group: (id) => ({ id, name: '' }) });
  notify({ kind: 'something-else', id: 'x' });
  notify({ kind: 'message', from: 'u-ben', text: 'no id' });
  notify(null);
  assert.deepEqual(heard, []);
  for (const kind of NOTIFY_KINDS) notify({ kind, id: 'x', from: { id: 'u-ben', name: 'Ben' } });
  assert.deepEqual(heard.map(([event, n]) => [event, n.kind]), NOTIFY_KINDS.map((k) => [NOTIFY_EVENT, k]));
  // A window that fails to hear takes nothing down with it.
  createNotify({ broadcast: () => { throw new Error('a window on its way out'); }, person: (id) => ({ id, name: '' }), group: (id) => ({ id, name: '' }) })({ kind: 'call', id: 'x', from: 'u-ben' });
});

test('the people blocked are shown by name: the one they had here when blocked, or the one the introducer gives', async () => {
  const it = talk({
    welcome: { blocked: ['u-old', 'u-eve'], blockedPeople: [{ id: 'u-eve', name: 'Eve Marsh' }] },
    answer: (msg) => (msg.type === 'contact.block' ? { type: 'contact.blocked', id: msg.contactId } : msg.type === 'contact.unblock' ? { type: 'contact.unblocked', id: msg.contactId } : null),
  });
  await it.online();
  // Blocked from another computer before names were kept: listed, with no name.
  assert.deepEqual(it.svc.state().blockedPeople, [{ id: 'u-old', name: '', email: '' }, { id: 'u-eve', name: 'Eve Marsh', email: '' }]);
  assert.deepEqual(await it.svc.block({ contactId: 'u-ben' }), { blocked: true });
  assert.deepEqual(it.svc.state().blockedPeople.at(-1), { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com' });
  assert.equal(it.svc.state().contacts.some((c) => c.id === 'u-ben'), false);

  // Kept between two runs of Office, and an older introducer that names nobody takes no name away.
  const again = talk({ settings: it.settings, contacts: [CARA], welcome: { blocked: ['u-old', 'u-eve', 'u-ben'] } });
  await again.online();
  assert.deepEqual(again.svc.state().blockedPeople.map((p) => p.name), ['', 'Eve Marsh', 'Ben Ortiz']);
});

test('someone blocked is let back: the introducer is asked, the list loses them, and their name is no longer kept', async () => {
  const it = talk({
    welcome: { blocked: ['u-eve'], blockedPeople: [{ id: 'u-eve', name: 'Eve Marsh' }] },
    answer: (msg) => (msg.type === 'contact.unblock' ? (msg.contactId === 'u-eve' ? { type: 'contact.unblocked', id: msg.contactId } : { type: 'error', for: 'contact.unblock', ref: msg.ref, code: 'nobody' }) : null),
  });
  await it.online();
  assert.deepEqual(await it.svc.unblock({ contactId: 'u-eve' }), { unblocked: true });
  assert.ok(it.out.some((m) => m.type === 'contact.unblock' && m.contactId === 'u-eve'));
  assert.deepEqual(it.svc.state().blockedPeople, []);
  assert.deepEqual(it.settings.get('talk.blocked-people'), {});
  // Being let back does not make them a contact.
  assert.equal(it.svc.state().contacts.some((c) => c.id === 'u-eve'), false);
  assert.deepEqual(await it.svc.unblock({ contactId: 'u-nobody' }), { error: 'nobody' });
});

test('the Talk pack starts Talk as it is registered, and never when the person turned that off', async () => {
  const { register } = await import('../apps/desktop/main/talk-pack.js');
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'talk-pack-'));
  const told = [];
  try {
    // A sign-in is kept, and connecting at start is off: registered, started, and no connection is tried.
    const stores = { dir, settings: map(new Map([[START_KEY, false]])), secrets: map(new Map([['talk.token', 'amira']])) };
    const { talk: svc } = register({ stores, broadcast: (event, payload) => told.push({ event, payload }), packs: { service: () => null } });
    await new Promise((resolve) => setImmediate(resolve));
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(svc.state().phase, 'signed-out');
    assert.deepEqual(told.filter((t) => t.payload?.phase === 'connecting'), []);
    // Nobody signed in: the same, with connecting at start left on.
    const fresh = register({ stores: { dir, settings: map(new Map()), secrets: map(new Map()) }, broadcast: () => {}, packs: { service: () => null } }).talk;
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(fresh.state().phase, 'signed-out');
    assert.equal(fresh.state().startWithOffice, true);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a reminder of an event can offer its Talk meeting: the first link in the event\'s place and notes, and whether it is live', () => {
  const a = 'meet-aaaaaaaaaaaaaaaaaaaa';
  const b = 'live-bbbbbbbbbbbbbbbbbbbb';
  assert.deepEqual(meetingInEvent('Room 4', `Agenda attached.\nJoin: https://introducer.rutba.io/m/${a}`), { id: a, live: false });
  assert.deepEqual(meetingInEvent(`https://introducer.rutba.io/m/${b}`, `https://introducer.rutba.io/m/${a}`), { id: b, live: true });
  // Not a meeting's page, a code too long, or nothing at all.
  assert.equal(meetingInEvent('https://example.com/meet', `https://introducer.rutba.io/m/${a}x`, null, undefined), null);
  assert.equal(meetingInEvent(), null);
});

/* ── someone asking to be a contact, blocked from their request (2.27.0) ── */

test('blocking someone from their request names the request, so the introducer puts it away; a plain block names none', async () => {
  const it = talk({ answer: (msg) => (msg.type === 'contact.block' ? { type: 'contact.blocked', id: msg.contactId } : null) });
  await it.online();
  assert.deepEqual(await it.svc.block({ contactId: 'u-eve', requestId: 'r_7' }), { blocked: true });
  assert.deepEqual(await it.svc.block({ contactId: 'u-ben' }), { blocked: true });
  const sent = it.out.filter((m) => m.type === 'contact.block');
  assert.deepEqual(sent.map((m) => [m.contactId, m.requestId]), [['u-eve', 'r_7'], ['u-ben', undefined]]);
  assert.equal('requestId' in sent[1], false);
});
