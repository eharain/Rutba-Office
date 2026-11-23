// Small meetings (Office 2.27) as Talk's service keeps them, on a stand-in
// introducer: a meeting set up, looked at before joining, joined and left,
// the ones this person is asked into, the line a group's conversation keeps,
// and the link that opens Talk on a meeting.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTalkService } from '../apps/desktop/main/talk.js';
import { groupKey, CAN } from '../apps/desktop/main/talk-conversations.js';
import { meetingCode, MEETING_CODE } from '../apps/desktop/main/talk-meetings.js';
import { parseLink } from '@rutba/office-shell/links';
import { METHODS } from '@rutba/office-shell/contract';

const folders = [];
test.after(() => { for (const f of folders) fs.rmSync(f, { recursive: true, force: true }); });

const AMIRA = { id: 'u-a', name: 'Amira Khan', email: 'amira@example.com' };
const BEN = { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com', presence: 'online' };
const PLAIN = 'meet-AAAAAAAAAAAAAAAAAAAA';
const LIVE = 'live-BBBBBBBBBBBBBBBBBBBB';
const meetingOf = (id, more = {}) => ({ id, name: 'Monday plan', starter: { id: 'u-a', name: 'Amira Khan' }, at: null, created: 5, live: id.startsWith('live-'), locked: false, groupId: null, link: `https://introducer.example/m/${id}`, most: 4, people: [], ...more });

/** The Talk service on a fake introducer that answers as `answer(msg)` says. */
function service({ answer = () => null } = {}) {
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const out = [];
  const told = [];
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'talk-meetings-'));
  folders.push(dir);
  let ws = null;
  const say = (m) => setImmediate(() => ws.onmessage({ data: JSON.stringify(m) }));
  const svc = createTalkService({
    stores: { settings: map(new Map()), secrets: map(new Map([['talk.token', 'amira']])) },
    dataDir: dir,
    broadcast: (event, payload) => told.push({ event, payload }),
    connect: () => {
      ws = { readyState: 1, send: (text) => {
        const msg = JSON.parse(text);
        out.push(msg);
        if (msg.type === 'hello') { say({ type: 'welcome', me: AMIRA, contacts: [BEN], requests: { incoming: [], outgoing: [] } }); return; }
        for (const reply of [].concat(answer(msg) || [])) say(reply);
      }, close: () => {} };
      setImmediate(() => ws.onopen());
      return ws;
    },
  });
  /** Resolves once the windows have been told an introducer message of this type. */
  const heard = (type) => new Promise((resolve) => {
    const from = told.length;
    const look = () => (told.slice(from).some((x) => x.event === 'talk:event' && x.payload?.type === type) ? resolve() : setImmediate(look));
    look();
  });
  return { svc, out, say, told, heard };
}
const online = async (svc) => {
  svc.resume();
  for (let i = 0; i < 200 && svc.state().phase !== 'online'; i++) await new Promise((r) => setImmediate(r));
  assert.equal(svc.state().phase, 'online');
};

test("a meeting's code is read from its page's address, its link or the code, and from nothing else", () => {
  assert.equal(meetingCode(PLAIN), PLAIN);
  assert.equal(meetingCode(`https://introducer.rutba.io/m/${PLAIN}`), PLAIN);
  assert.equal(meetingCode(`  https://introducer.rutba.io/m/${LIVE}/?x=1 `), LIVE);
  assert.equal(meetingCode(`rutba-office://talk/meet/${LIVE}`), LIVE);
  assert.match(PLAIN, MEETING_CODE);
  for (const bad of ['', null, 'meet-short', `${PLAIN}A`, `room-${PLAIN.slice(5)}`, `https://introducer.rutba.io/i/${PLAIN}`, `https://introducer.rutba.io/m/x/${PLAIN}`, `javascript:${PLAIN}`, `${PLAIN} and more`]) {
    assert.equal(meetingCode(bad), null, String(bad));
  }
});

test('the link that opens Talk on a meeting, and near misses that open nothing', () => {
  assert.deepEqual(parseLink(`rutba-office://talk/meet/${PLAIN}`), { app: 'talk', query: { meet: PLAIN } });
  assert.deepEqual(parseLink(` rutba-office://talk/meet/${LIVE}/ `), { app: 'talk', query: { meet: LIVE } });
  for (const bad of [`rutba-office://talk/meet/${PLAIN}A`, 'rutba-office://talk/meet/meet-short', `rutba-office://talk/meet/${PLAIN}?server=evil.example`, `rutba-office://talk/meet/${PLAIN}/x`, `rutba-office://meet/${PLAIN}`, `rutba-office://talk/meet/room-${PLAIN.slice(5)}`]) {
    assert.equal(parseLink(bad), null, bad);
  }
});

test('the hello says this Office holds meetings, and the windows can ask for each thing a meeting needs', async () => {
  const { svc, out } = service();
  await online(svc);
  assert.ok(out[0].can.includes('meetings'));
  assert.ok(CAN.includes('meetings'));
  for (const m of ['meetCreate', 'meetInvite', 'meetPeek', 'meetJoin', 'meetLeave', 'meetLock', 'meetRemove', 'meetEnd', 'meetDismiss']) {
    assert.ok(METHODS.talk.includes(m), `${m} is in the contract`);
    assert.equal(typeof svc[m], 'function', `${m} is in the service`);
  }
});

test('a meeting set up: its name, time, the contacts asked, and whether it will be streamed live go as chosen', async () => {
  const { svc, out } = service({ answer: (msg) => (msg.type === 'meet.create' ? { type: 'meet.created', ref: msg.ref, meeting: meetingOf(msg.live ? LIVE : PLAIN, { name: msg.name, at: msg.at ?? null }) } : null) });
  await online(svc);
  const planned = await svc.meetCreate({ name: 'Monday plan', at: 1_800_000_000_000, live: true });
  assert.equal(planned.meeting.id, LIVE);
  assert.equal(planned.meeting.live, true);
  const sent = out.find((m) => m.type === 'meet.create');
  assert.deepEqual({ name: sent.name, at: sent.at, live: sent.live, invite: sent.invite, tell: sent.tell }, { name: 'Monday plan', at: 1_800_000_000_000, live: true, invite: [], tell: false });
  assert.equal('groupId' in sent, false);

  const now = await svc.meetCreate({ name: 'Quick word', invite: ['u-ben'], now: true });
  assert.equal(now.meeting.id, PLAIN);
  const second = out.filter((m) => m.type === 'meet.create')[1];
  assert.deepEqual({ invite: second.invite, tell: second.tell, live: second.live }, { invite: ['u-ben'], tell: true, live: false });
  assert.equal('at' in second, false, 'a meeting started now has no time');
  // A small meeting says nothing of servers to the introducer: what it is sent is what it always was.
  for (const sent of out.filter((m) => m.type === 'meet.create')) { assert.equal('carried' in sent, false); assert.equal('most' in sent, false); }
});

test('a meeting to be held on a Talk Server is asked for as one, with how many it is to hold', async () => {
  const { svc, out } = service({ answer: (msg) => (msg.type === 'meet.create' ? { type: 'meet.created', ref: msg.ref, meeting: meetingOf(PLAIN, { name: msg.name, ...(msg.carried ? { carried: msg.carried, most: msg.most ?? 500 } : {}) }) } : null) });
  await online(svc);
  const made = await svc.meetCreate({ name: 'All hands', invite: ['u-ben'], now: true, carried: 'server', most: 25 });
  assert.equal(made.meeting.carried, 'server');
  assert.equal(made.meeting.most, 25);
  const sent = out.filter((m) => m.type === 'meet.create');
  assert.deepEqual({ carried: sent[0].carried, most: sent[0].most, live: sent[0].live, tell: sent[0].tell }, { carried: 'server', most: 25, live: false, tell: true });

  // How many is a whole number or is left to the introducer; anything else about how it is held is not passed on.
  await svc.meetCreate({ name: 'All hands', carried: 'server', most: '25' });
  await svc.meetCreate({ name: 'All hands', carried: 'server', most: 1 });
  await svc.meetCreate({ name: 'All hands', carried: 'server' });
  await svc.meetCreate({ name: 'All hands', carried: 'mesh', most: 25 });
  await svc.meetCreate({ name: 'All hands', carried: true, most: 25 });
  await svc.meetCreate({ name: 'All hands', most: 25 });
  const rest = out.filter((m) => m.type === 'meet.create').slice(1).map((m) => [m.carried ?? null, m.most ?? null]);
  assert.deepEqual(rest, [['server', null], ['server', null], ['server', null], [null, null], [null, null], [null, null]]);
});

test('a refusal at set-up comes back as its reason', async () => {
  const { svc } = service({ answer: (msg) => (msg.type === 'meet.create' ? { type: 'error', for: 'meet.create', code: msg.name ? 'not-a-contact' : 'name', ref: msg.ref } : null) });
  await online(svc);
  assert.deepEqual(await svc.meetCreate({ name: '' }), { error: 'name' });
  assert.deepEqual(await svc.meetCreate({ name: 'x', invite: ['u-zed'] }), { error: 'not-a-contact' });
});

test('a meeting is looked at before joining, by the code alone, and a stranger is told whose it is', async () => {
  const { svc, out } = service({ answer: (msg) => {
    if (msg.type !== 'meet.peek') return null;
    return msg.id === PLAIN ? { type: 'meet.info', ref: msg.ref, meeting: meetingOf(PLAIN, { people: [{ id: 'u-ben', name: 'Ben Ortiz', contact: true }] }) } : { type: 'error', for: 'meet.peek', code: 'not-allowed', starter: { name: 'Dana Ruiz' }, ref: msg.ref };
  } });
  await online(svc);
  const seen = await svc.meetPeek({ link: `https://somewhere-else.example/m/${PLAIN}` });
  assert.equal(seen.meeting.people[0].name, 'Ben Ortiz');
  assert.equal(out.find((m) => m.type === 'meet.peek').id, PLAIN, 'only the code is asked of the introducer this Office is signed in to');
  assert.deepEqual(await svc.meetPeek({ link: LIVE }), { error: 'not-allowed', starter: { name: 'Dana Ruiz' }, id: LIVE });
  const before = out.length;
  assert.deepEqual(await svc.meetPeek({ link: 'https://example.com/whatever' }), { error: 'no-meeting' });
  assert.equal(out.length, before, 'what is not a meeting link is never asked');
});

test('joining and leaving: every window knows which meeting this Office is in', async () => {
  const { svc, say, heard } = service({ answer: (msg) => {
    if (msg.type === 'meet.join') return msg.id === PLAIN ? { type: 'meet.joined', ref: msg.ref, meeting: meetingOf(PLAIN, { people: [{ id: 'u-ben', name: 'Ben Ortiz', contact: true }] }) } : { type: 'error', for: 'meet.join', code: 'full', most: 4, ref: msg.ref };
    if (msg.type === 'meet.leave') return { type: 'meet.left', ref: msg.ref, id: msg.id };
    return null;
  } });
  await online(svc);
  assert.deepEqual(await svc.meetJoin({ id: LIVE }), { error: 'full', most: 4 });
  assert.equal(svc.state().meeting, null);
  assert.deepEqual(await svc.meetJoin({ id: 'not a code' }), { error: 'no-meeting' });

  const r = await svc.meetJoin({ id: PLAIN });
  assert.equal(r.meeting.people.length, 1, 'those already there, whom the newcomer offers to');
  assert.equal(svc.state().meeting, PLAIN);
  assert.deepEqual(await svc.meetLeave({ id: PLAIN }), { left: true });
  assert.equal(svc.state().meeting, null);

  await svc.meetJoin({ id: PLAIN });
  let told = heard('meet.removed');
  say({ type: 'meet.removed', id: PLAIN, why: 'removed', by: { id: 'u-ben', name: 'Ben Ortiz' } });
  await told;
  assert.equal(svc.state().meeting, null, 'taken out');

  await svc.meetJoin({ id: PLAIN });
  told = heard('meet.ended');
  say({ type: 'meet.ended', id: PLAIN, by: { id: 'u-ben', name: 'Ben Ortiz' } });
  await told;
  assert.equal(svc.state().meeting, null, 'ended for everyone');
});

test('what only the person who set a meeting up can do is asked of the introducer, and its refusal comes back', async () => {
  const { svc, out } = service({ answer: (msg) => {
    if (msg.type === 'meet.lock') return { type: 'meet.changed', ref: msg.ref, meeting: meetingOf(PLAIN, { locked: msg.locked }) };
    if (msg.type === 'meet.remove') return msg.who === 'u-ben' ? { type: 'meet.changed', ref: msg.ref, meeting: meetingOf(PLAIN) } : { type: 'error', for: 'meet.remove', code: 'nobody', ref: msg.ref };
    if (msg.type === 'meet.end') return msg.id === PLAIN ? { type: 'meet.ended', ref: msg.ref, id: msg.id } : { type: 'error', for: 'meet.end', code: 'not-yours', ref: msg.ref };
    return null;
  } });
  await online(svc);
  assert.equal((await svc.meetLock({ id: PLAIN, locked: true })).meeting.locked, true);
  assert.equal((await svc.meetLock({ id: PLAIN, locked: false })).meeting.locked, false);
  assert.ok((await svc.meetRemove({ id: PLAIN, who: 'u-ben' })).meeting);
  assert.deepEqual(await svc.meetRemove({ id: PLAIN, who: 'u-zed' }), { error: 'nobody' });
  assert.deepEqual(await svc.meetEnd({ id: PLAIN }), { ended: true });
  assert.deepEqual(await svc.meetEnd({ id: LIVE }), { error: 'not-yours' });
  assert.deepEqual(out.filter((m) => m.type.startsWith('meet.')).map((m) => m.type), ['meet.lock', 'meet.lock', 'meet.remove', 'meet.remove', 'meet.end', 'meet.end']);
});

test("someone else's change to a meeting never answers this person's own question", async () => {
  // The lock is refused, and the introducer says of the same meeting that it changed (another change, with no ref).
  const { svc } = service({ answer: (msg) => (msg.type === 'meet.lock' ? [{ type: 'meet.changed', meeting: meetingOf(PLAIN, { locked: false }) }, { type: 'error', for: 'meet.lock', code: 'not-yours', ref: msg.ref }] : null) });
  await online(svc);
  assert.deepEqual(await svc.meetLock({ id: PLAIN, locked: true }), { error: 'not-yours' });
});

test('a meeting this person is asked into is offered until they join it, say not now, or it ends', async () => {
  const { svc, say, heard, told } = service({ answer: (msg) => (msg.type === 'meet.join' ? { type: 'meet.joined', ref: msg.ref, meeting: meetingOf(msg.id) } : null) });
  await online(svc);
  const ask = async (id) => { const t = heard('meet.invited'); say({ type: 'meet.invited', meeting: meetingOf(id, { starter: { id: 'u-ben', name: 'Ben Ortiz' } }), by: { id: 'u-ben', name: 'Ben Ortiz' } }); await t; };

  await ask(PLAIN);
  await ask(PLAIN);
  assert.equal(svc.state().meetings.length, 1, 'asked twice, offered once');
  assert.equal(svc.state().meetings[0].by.name, 'Ben Ortiz');

  const before = told.filter((x) => x.event === 'talk:state').length;
  assert.deepEqual(svc.meetDismiss({ id: PLAIN }), { dismissed: true });
  assert.equal(svc.state().meetings.length, 0);
  assert.ok(told.filter((x) => x.event === 'talk:state').length > before, 'every window is told it is no longer offered');

  await ask(LIVE);
  assert.equal(svc.state().meetings[0].meeting.live, true, 'the offer says it will be streamed live');
  const ended = heard('meet.ended');
  say({ type: 'meet.ended', id: LIVE, by: { id: 'u-ben', name: 'Ben Ortiz' } });
  await ended;
  assert.equal(svc.state().meetings.length, 0);

  await ask(PLAIN);
  await svc.meetJoin({ id: PLAIN });
  assert.equal(svc.state().meetings.length, 0, 'joined, so no longer offered');
});

test("a meeting started in a group leaves one line in that group's conversation", async () => {
  const group = { id: 'g_1', name: 'Launch', starter: 'u-a', created: 5, members: [{ id: 'u-a', name: 'Amira Khan' }, { id: 'u-ben', name: 'Ben Ortiz', contact: true }] };
  const { svc, say, heard } = service({ answer: (msg) => (msg.type === 'meet.create' ? { type: 'meet.created', ref: msg.ref, meeting: meetingOf(LIVE, { groupId: msg.groupId, name: msg.name }) } : null) });
  await online(svc);
  const told = heard('groups');
  say({ type: 'groups', groups: [group] });
  await told;

  await svc.meetCreate({ name: 'Launch', groupId: 'g_1', live: true, now: true });
  let lines = svc.history({ contactId: groupKey('g_1') }).filter((m) => m.kind === 'meeting');
  assert.equal(lines.length, 1);
  assert.deepEqual(lines[0].meeting, { id: LIVE, name: 'Launch', live: true, recorded: false, link: `https://introducer.example/m/${LIVE}` });
  assert.equal(lines[0].by.id, 'u-a', 'started by this person');

  // Ben starts another there; told twice, it is still one line.
  for (let i = 0; i < 2; i++) {
    const t = heard('meet.invited');
    say({ type: 'meet.invited', meeting: meetingOf(PLAIN, { groupId: 'g_1', name: 'Launch' }), by: { id: 'u-ben', name: 'Ben Ortiz' } });
    await t;
  }
  lines = svc.history({ contactId: groupKey('g_1') }).filter((m) => m.kind === 'meeting');
  assert.equal(lines.length, 2);
  assert.equal(lines[1].by.name, 'Ben Ortiz');
});

test('a new connection starts clean: the introducer has this Office in no meeting', async () => {
  const { svc, say, heard } = service({ answer: (msg) => (msg.type === 'meet.join' ? { type: 'meet.joined', ref: msg.ref, meeting: meetingOf(msg.id) } : null) });
  await online(svc);
  await svc.meetJoin({ id: PLAIN });
  const asked = heard('meet.invited');
  say({ type: 'meet.invited', meeting: meetingOf(LIVE), by: { id: 'u-ben', name: 'Ben Ortiz' } });
  await asked;
  const again = heard('welcome');
  say({ type: 'welcome', me: AMIRA, contacts: [BEN], requests: { incoming: [], outgoing: [] } });
  await again;
  assert.equal(svc.state().meeting, null);
  assert.deepEqual(svc.state().meetings, []);
});

test('a meeting set up to be recorded says so to the introducer, and one that is not says nothing of it', async () => {
  const { svc, out } = service({ answer: (msg) => (msg.type === 'meet.create' ? { type: 'meet.created', ref: msg.ref, meeting: meetingOf(PLAIN, { name: msg.name, recorded: msg.recorded === true }) } : null) });
  await online(svc);
  const made = await svc.meetCreate({ name: 'An interview', recorded: true, invite: ['u-ben'], now: true });
  assert.equal(made.meeting.recorded, true);
  const [first] = out.filter((m) => m.type === 'meet.create');
  assert.deepEqual({ recorded: first.recorded, live: first.live, tell: first.tell }, { recorded: true, live: false, tell: true });
  await svc.meetCreate({ name: 'A word', recorded: 'yes' });
  assert.equal('recorded' in out.filter((m) => m.type === 'meet.create')[1], false, 'only true is true');
});

test('more contacts asked into a meeting that is on: how many were told, or why not', async () => {
  const { svc, out } = service({ answer: (msg) => {
    if (msg.type !== 'meet.invite') return null;
    return msg.id === LIVE ? { type: 'meet.asked', ref: msg.ref, id: msg.id, told: msg.invite.length } : { type: 'error', for: 'meet.invite', code: 'not-yours', ref: msg.ref };
  } });
  await online(svc);
  assert.deepEqual(await svc.meetInvite({ id: LIVE, invite: ['u-ben', 'u-cara'] }), { told: 2 });
  assert.deepEqual(out.find((m) => m.type === 'meet.invite').invite, ['u-ben', 'u-cara']);
  assert.deepEqual(await svc.meetInvite({ id: PLAIN, invite: 'u-ben' }), { error: 'not-yours' });
});

/* ── a meeting that is over: its line offers no Join (the screen review of 2026-11-22) ── */

test("a group's meeting line says the meeting has ended once it is known to be over", async () => {
  const group = { id: 'g_1', name: 'Launch', starter: 'u-a', created: 5, members: [{ id: 'u-a', name: 'Amira Khan' }, { id: 'u-ben', name: 'Ben Ortiz', contact: true }] };
  const { svc, say, heard } = service({ answer: (msg) => (msg.type === 'meet.peek' || msg.type === 'meet.join' ? { type: 'error', for: msg.type, code: 'no-meeting', ref: msg.ref } : null) });
  await online(svc);
  let told = heard('groups');
  say({ type: 'groups', groups: [group] });
  await told;
  const THIRD = 'meet-CCCCCCCCCCCCCCCCCCCC';
  for (const id of [PLAIN, LIVE, THIRD]) {
    told = heard('meet.invited');
    say({ type: 'meet.invited', meeting: meetingOf(id, { groupId: 'g_1', starter: { id: 'u-ben', name: 'Ben Ortiz' } }), by: { id: 'u-ben', name: 'Ben Ortiz' } });
    await told;
  }
  const lines = () => svc.history({ contactId: groupKey('g_1') }).filter((m) => m.kind === 'meeting').map((m) => [m.meeting.id, Boolean(m.ended)]);
  assert.deepEqual(lines(), [[PLAIN, false], [LIVE, false], [THIRD, false]]);
  assert.deepEqual([...svc.history({ groupId: 'g_1' }).filter((m) => m.kind === 'meeting').map((m) => 'ended' in m)], [false, false, false], 'nothing is said of a meeting not known to be over');

  // Ended by the person who set it up.
  told = heard('meet.ended');
  say({ type: 'meet.ended', id: PLAIN });
  await told;
  assert.deepEqual(lines(), [[PLAIN, true], [LIVE, false], [THIRD, false]]);
  // Looked at and found gone: it ran out, or ended while this person was not in it.
  assert.equal((await svc.meetPeek({ link: LIVE })).error, 'no-meeting');
  assert.deepEqual(lines(), [[PLAIN, true], [LIVE, true], [THIRD, false]]);
  // Or found gone on the way in.
  assert.equal((await svc.meetJoin({ id: THIRD })).error, 'no-meeting');
  assert.deepEqual(lines(), [[PLAIN, true], [LIVE, true], [THIRD, true]]);
  // Told twice, kept once; and what is not a meeting's code is never kept.
  told = heard('meet.ended');
  say({ type: 'meet.ended', id: PLAIN });
  await told;
  told = heard('meet.ended');
  say({ type: 'meet.ended', id: '../somewhere' });
  await told;
  assert.deepEqual(lines(), [[PLAIN, true], [LIVE, true], [THIRD, true]]);
});

test('any other refusal leaves the line as it was: a full or locked meeting is still on', async () => {
  const group = { id: 'g_1', name: 'Launch', starter: 'u-a', created: 5, members: [{ id: 'u-a', name: 'Amira Khan' }, { id: 'u-ben', name: 'Ben Ortiz', contact: true }] };
  const { svc, say, heard } = service({ answer: (msg) => (msg.type === 'meet.join' ? { type: 'error', for: 'meet.join', code: 'full', most: 4, ref: msg.ref } : null) });
  await online(svc);
  let told = heard('groups');
  say({ type: 'groups', groups: [group] });
  await told;
  told = heard('meet.invited');
  say({ type: 'meet.invited', meeting: meetingOf(PLAIN, { groupId: 'g_1', starter: { id: 'u-ben', name: 'Ben Ortiz' } }), by: { id: 'u-ben', name: 'Ben Ortiz' } });
  await told;
  assert.equal((await svc.meetJoin({ id: PLAIN })).error, 'full');
  assert.equal(svc.history({ contactId: groupKey('g_1') }).find((m) => m.kind === 'meeting').ended, undefined);
});

/* ── "Not now" does not lose the way in (2.27.0) ── */

test('a meeting a contact asks this person into keeps a line in the conversation with them, once, with its mark', async () => {
  const { svc, say, heard } = service();
  await online(svc);
  const ben = { id: 'u-ben', name: 'Ben Ortiz' };
  for (let i = 0; i < 2; i++) {
    const t = heard('meet.invited');
    say({ type: 'meet.invited', meeting: meetingOf(LIVE, { starter: ben, name: 'Town hall' }), by: ben });
    await t;
  }
  const lines = svc.history({ contactId: 'u-ben' }).filter((m) => m.kind === 'meeting');
  assert.equal(lines.length, 1, 'told twice, still one line');
  assert.deepEqual(lines[0].meeting, { id: LIVE, name: 'Town hall', live: true, recorded: false, link: `https://introducer.example/m/${LIVE}` });
  assert.deepEqual(lines[0].by, ben);

  // Said not now: the card goes, the line stays.
  svc.meetDismiss({ id: LIVE });
  assert.deepEqual(svc.state().meetings, []);
  assert.equal(svc.history({ contactId: 'u-ben' }).filter((m) => m.kind === 'meeting').length, 1);
});

test('a meeting started in a group keeps its line in the group only, and one this person set up keeps none with anyone', async () => {
  const group = { id: 'g_1', name: 'Launch', starter: 'u-a', created: 5, members: [{ id: 'u-a', name: 'Amira Khan' }, { id: 'u-ben', name: 'Ben Ortiz', contact: true }] };
  const { svc, say, heard } = service({ answer: (msg) => (msg.type === 'meet.create' ? { type: 'meet.created', ref: msg.ref, meeting: meetingOf(PLAIN, { name: msg.name }) } : null) });
  await online(svc);
  const told = heard('groups');
  say({ type: 'groups', groups: [group] });
  await told;
  const t = heard('meet.invited');
  say({ type: 'meet.invited', meeting: meetingOf(LIVE, { groupId: 'g_1', name: 'Launch' }), by: { id: 'u-ben', name: 'Ben Ortiz' } });
  await t;
  assert.equal(svc.history({ contactId: groupKey('g_1') }).filter((m) => m.kind === 'meeting').length, 1);
  assert.equal(svc.history({ contactId: 'u-ben' }).filter((m) => m.kind === 'meeting').length, 0, 'not said twice, in the group and with the person');

  await svc.meetCreate({ name: 'A word', invite: ['u-ben'], now: true });
  assert.equal(svc.history({ contactId: 'u-ben' }).filter((m) => m.kind === 'meeting').length, 0);
});
