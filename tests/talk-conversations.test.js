// Talk made better (Office 2.28): unread counts, missed calls, the mark on a
// message that has arrived, group conversations and reporting someone, as
// the Talk service keeps them, on a stand-in introducer.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTalkService } from '../apps/desktop/main/talk.js';
import { groupKey, CAN } from '../apps/desktop/main/talk-conversations.js';

const folders = [];
test.after(() => { for (const f of folders) fs.rmSync(f, { recursive: true, force: true }); });

const AMIRA = { id: 'u-a', name: 'Amira Khan', email: 'amira@example.com' };
const BEN = { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com', presence: 'online' };
const CARA = { id: 'u-cara', name: 'Cara Lund', email: 'cara@example.com', presence: 'offline' };
const GROUP = { id: 'g_1', name: 'Launch', starter: 'u-a', created: 5, members: [{ id: 'u-a', name: 'Amira Khan' }, { id: 'u-ben', name: 'Ben Ortiz', contact: true }, { id: 'u-cara', name: 'Cara Lund', contact: true }] };

/** The Talk service on a fake introducer that answers as `answer(msg)` says; `after` is said once the welcome has gone. */
function service({ answer = () => null, after = [], dataDir = null } = {}) {
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const out = [];
  const told = [];
  const dir = dataDir || fs.mkdtempSync(path.join(os.tmpdir(), 'talk-conversations-'));
  if (!dataDir) folders.push(dir);
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
        if (msg.type === 'hello') { say({ type: 'welcome', me: AMIRA, contacts: [BEN, CARA], requests: { incoming: [], outgoing: [] } }); for (const m of after) say(m); return; }
        for (const reply of [].concat(answer(msg) || [])) say(reply);
      }, close: () => {} };
      setImmediate(() => ws.onopen());
      return ws;
    },
  });
  return { svc, out, say, told, dir };
}
const online = async (svc) => {
  svc.resume();
  for (let i = 0; i < 50 && svc.state().phase !== 'online'; i++) await new Promise((r) => setTimeout(r, 5));
  assert.equal(svc.state().phase, 'online');
};
const tick = () => new Promise((r) => setTimeout(r, 15));

test('the hello says what this Office reads, so an introducer keeps notes and groups for it', async () => {
  const { svc, out } = service();
  await online(svc);
  assert.deepEqual(out[0].can, CAN);
  assert.deepEqual(CAN, ['groups', 'notices']);
});

test('unread: each message that comes in is counted on its person, in all, and cleared when the conversation is read', async () => {
  const { svc, say, out } = service();
  await online(svc);
  assert.deepEqual([svc.state().unread, svc.state().unreadTotal], [{}, 0]);
  say({ type: 'messages', messages: [{ id: 'm-00000001', from: 'u-ben', text: 'Hello', at: 10 }, { id: 'm-00000002', from: 'u-ben', text: 'Are you there?', at: 11 }] });
  say({ type: 'messages', messages: [{ id: 'm-00000003', from: 'u-cara', text: 'Lunch?', at: 12 }] });
  await tick();
  assert.deepEqual(svc.state().unread, { 'u-ben': 2, 'u-cara': 1 });
  assert.equal(svc.state().unreadTotal, 3);
  assert.deepEqual(out.filter((m) => m.type === 'msg.ack').map((m) => m.ids), [['m-00000001', 'm-00000002'], ['m-00000003']]);
  svc.read({ conversation: 'u-ben' });
  assert.deepEqual(svc.state().unread, { 'u-cara': 1 });
  assert.equal(svc.state().unreadTotal, 1);
  // A message of my own, sent from another computer of mine, is not unread.
  say({ type: 'msg.copy', message: { id: 'm-00000004', from: 'u-a', to: 'u-ben', text: 'On my way', at: 13 } });
  await tick();
  assert.deepEqual(svc.state().unread, { 'u-cara': 1 });
});

test('unread counts are kept between runs', async () => {
  const first = service();
  await online(first.svc);
  first.say({ type: 'messages', messages: [{ id: 'm-00000010', from: 'u-ben', text: 'Hello', at: 10 }] });
  await tick();
  const second = service({ dataDir: first.dir });
  await online(second.svc);
  assert.deepEqual(second.svc.state().unread, { 'u-ben': 1 });
});

test('a message of mine: sending, then sent once the introducer has it, then arrived once the other side has it', async () => {
  const { svc, say, out } = service({ answer: (m) => (m.type === 'msg.send' ? { type: 'msg.sent', id: m.id, at: 20, delivered: true } : null) });
  await online(svc);
  const entry = svc.message({ to: 'u-ben', text: 'Lunch at one?' });
  assert.equal(entry.state, 'sending');
  assert.equal(svc.history({ contactId: 'u-ben' })[0].state, 'sending');
  await tick();
  assert.equal(svc.history({ contactId: 'u-ben' })[0].state, 'sent');
  say({ type: 'notices', notices: [{ id: 'n_1', kind: 'delivered', ids: [entry.id], by: 'u-ben', at: 21 }] });
  await tick();
  assert.equal(svc.history({ contactId: 'u-ben' })[0].state, 'arrived');
  // The note is let go once it is kept here.
  assert.deepEqual(out.filter((m) => m.type === 'notice.ack').map((m) => m.ids), [['n_1']]);
  // A late "sent" does not take the mark back.
  say({ type: 'msg.sent', id: entry.id, at: 20 });
  await tick();
  assert.equal(svc.history({ contactId: 'u-ben' })[0].state, 'arrived');
});

test('a message that arrived while this computer was off is marked when the note comes at the next start', async () => {
  const first = service({ answer: (m) => (m.type === 'msg.send' ? { type: 'msg.sent', id: m.id, at: 20 } : null) });
  await online(first.svc);
  const entry = first.svc.message({ to: 'u-cara', text: 'See you tomorrow' });
  await tick();
  const second = service({ dataDir: first.dir, after: [{ type: 'notices', notices: [{ id: 'n_2', kind: 'delivered', ids: [entry.id], by: 'u-cara', at: 30 }] }] });
  await online(second.svc);
  await tick();
  assert.equal(second.svc.history({ contactId: 'u-cara' })[0].state, 'arrived');
});

test('missed calls: one that rang here unanswered, one the introducer kept a note of, each listed once, with a line in the conversation', async () => {
  const { svc, say } = service();
  await online(svc);
  // Rings, and the caller gives up.
  say({ type: 'call.incoming', callId: 'call-0001', from: BEN, sdp: 'v=0', video: true });
  say({ type: 'call.ended', callId: 'call-0001', reason: 'hung-up' });
  await tick();
  assert.equal(svc.state().missed, 1);
  // The introducer's note of the same call adds nothing; its note of one made while this computer was off does.
  say({ type: 'notices', notices: [
    { id: 'n_3', kind: 'missed-call', callId: 'call-0001', from: { id: 'u-ben', name: 'Ben Ortiz' }, video: true, why: 'unanswered', at: 40 },
    { id: 'n_4', kind: 'missed-call', callId: 'call-0002', from: { id: 'u-cara', name: 'Cara Lund' }, video: false, why: 'offline', at: 41 },
  ] });
  await tick();
  assert.equal(svc.state().missed, 2);
  const calls = svc.calls();
  assert.deepEqual(calls.map((c) => [c.id, c.with, c.name, c.video, c.why, c.seen]), [['call-0002', 'u-cara', 'Cara Lund', false, 'offline', false], ['call-0001', 'u-ben', 'Ben Ortiz', true, 'unanswered', false]]);
  // A line in each person's conversation, not counted as an unread message.
  assert.deepEqual(svc.history({ contactId: 'u-ben' }).map((e) => [e.kind, e.video]), [['missed-call', true]]);
  assert.deepEqual(svc.state().unread, {});
  // Looked at: the count goes, the list stays.
  svc.calls({ seen: true });
  assert.equal(svc.state().missed, 0);
  assert.equal(svc.calls().length, 2);
});

test('a call answered here, declined here, or answered on another computer is not missed; one that came while busy is', async () => {
  const { svc, say } = service();
  await online(svc);
  say({ type: 'call.incoming', callId: 'call-0011', from: BEN, sdp: 'v=0' });
  await tick();
  svc.send({ type: 'call.answer', callId: 'call-0011', sdp: 'v=0' });
  say({ type: 'call.ended', callId: 'call-0011' });
  say({ type: 'call.incoming', callId: 'call-0012', from: BEN, sdp: 'v=0' });
  await tick();
  svc.send({ type: 'call.decline', callId: 'call-0012' });
  say({ type: 'call.incoming', callId: 'call-0013', from: BEN, sdp: 'v=0' });
  say({ type: 'call.taken', callId: 'call-0013' });
  say({ type: 'call.ended', callId: 'call-0013' });
  await tick();
  assert.equal(svc.state().missed, 0);
  say({ type: 'call.incoming', callId: 'call-0014', from: CARA, sdp: 'v=0' });
  await tick();
  svc.send({ type: 'call.decline', callId: 'call-0014', reason: 'busy' });
  assert.equal(svc.state().missed, 1);
  assert.equal(svc.calls()[0].why, 'busy');
});

test('groups: listed at the welcome, made, changed, and gone from the list when left', async () => {
  const made = { ...GROUP, id: 'g_2', name: 'Weekend', created: 9 };
  const { svc, say, out } = service({
    after: [{ type: 'groups', groups: [GROUP] }],
    answer: (m) => {
      if (m.type === 'group.create') return { type: 'group.created', ref: m.ref, group: made };
      if (m.type === 'group.rename') return { type: 'group.changed', ref: m.ref, group: { ...made, name: m.name }, event: { kind: 'renamed', by: { id: 'u-a', name: 'Amira Khan' }, name: m.name, at: 50 } };
      if (m.type === 'group.leave') return { type: 'group.left', ref: m.ref, groupId: m.groupId };
      if (m.type === 'group.add') return { type: 'error', for: 'group.add', code: 'not-a-contact', ref: m.ref };
      return null;
    },
  });
  await online(svc);
  await tick();
  assert.deepEqual(svc.state().groups.map((g) => g.id), ['g_1']);
  const r = await svc.groupCreate({ name: 'Weekend', members: ['u-ben', 'u-cara'] });
  assert.equal(r.group.id, 'g_2');
  assert.deepEqual(out.find((m) => m.type === 'group.create').members, ['u-ben', 'u-cara']);
  assert.deepEqual(svc.state().groups.map((g) => g.id), ['g_1', 'g_2']);
  assert.equal((await svc.groupRename({ groupId: 'g_2', name: 'Saturday' })).group.name, 'Saturday');
  assert.equal(svc.state().groups[1].name, 'Saturday');
  // What changed is a line in the group's conversation.
  assert.deepEqual(svc.history({ groupId: 'g_2' }).map((e) => [e.kind, e.event.kind, e.event.name]), [['group-event', 'renamed', 'Saturday']]);
  assert.deepEqual(await svc.groupAdd({ groupId: 'g_2', members: ['u-x'] }), { error: 'not-a-contact' });
  assert.deepEqual(await svc.groupLeave({ groupId: 'g_2' }), { left: true });
  assert.deepEqual(svc.state().groups.map((g) => g.id), ['g_1']);
  // Put in a group by someone else: on the list, with who did it.
  say({ type: 'group.added', group: { ...GROUP, id: 'g_3', created: 12 }, by: { id: 'u-ben', name: 'Ben Ortiz' } });
  await tick();
  assert.deepEqual(svc.state().groups.map((g) => g.id), ['g_1', 'g_3']);
  assert.deepEqual(svc.history({ groupId: 'g_3' }).map((e) => [e.event.kind, e.event.by.name]), [['added-you', 'Ben Ortiz']]);
});

test('a group message: kept under its group with who wrote it, counted unread, and acknowledged', async () => {
  const { svc, say, out } = service({ after: [{ type: 'groups', groups: [GROUP] }] });
  await online(svc);
  say({ type: 'group.messages', messages: [{ id: 'm-00000020', group: 'g_1', from: 'u-ben', text: 'Shall we start Monday?', at: 60 }] });
  await tick();
  assert.deepEqual(svc.history({ groupId: 'g_1' }).map((e) => [e.from, e.text, e.mine]), [['u-ben', 'Shall we start Monday?', false]]);
  // Not mixed into Ben's own conversation.
  assert.deepEqual(svc.history({ contactId: 'u-ben' }), []);
  assert.deepEqual(svc.state().unread, { [groupKey('g_1')]: 1 });
  assert.deepEqual(out.filter((m) => m.type === 'group.ack').map((m) => m.ids), [['m-00000020']]);
  svc.read({ conversation: groupKey('g_1') });
  assert.equal(svc.state().unreadTotal, 0);
});

test('a group message of mine has arrived once everyone it went to has it', async () => {
  const { svc, say, out } = service({
    after: [{ type: 'groups', groups: [GROUP] }],
    answer: (m) => (m.type === 'group.send' ? { type: 'group.sent', id: m.id, groupId: m.groupId, at: 70, to: ['u-ben', 'u-cara'] } : null),
  });
  await online(svc);
  const entry = svc.message({ group: 'g_1', text: 'Monday it is' });
  assert.deepEqual([out.at(-1).type, out.at(-1).groupId, out.at(-1).text], ['group.send', 'g_1', 'Monday it is']);
  await tick();
  const one = () => svc.history({ groupId: 'g_1' })[0];
  assert.deepEqual([one().state, one().to], ['sent', ['u-ben', 'u-cara']]);
  say({ type: 'notices', notices: [{ id: 'n_5', kind: 'delivered', group: 'g_1', ids: [entry.id], by: 'u-ben', at: 71 }] });
  await tick();
  assert.deepEqual([one().state, one().arrived], ['sent', ['u-ben']]);
  say({ type: 'notices', notices: [{ id: 'n_6', kind: 'delivered', group: 'g_1', ids: [entry.id], by: 'u-cara', at: 72 }] });
  await tick();
  assert.deepEqual([one().state, one().arrived], ['arrived', ['u-ben', 'u-cara']]);
});

test('reporting someone: the reason and note go to the introducer, with their last messages only when the person chose so', async () => {
  const { svc, say, out } = service({ answer: (m) => (m.type === 'report' ? { type: 'report.taken', ref: m.ref, id: 'rep_1' } : null) });
  await online(svc);
  say({ type: 'messages', messages: [{ id: 'm-00000030', from: 'u-ben', text: 'Answer me', at: 80 }, { id: 'm-00000031', from: 'u-ben', text: 'Now', at: 81 }] });
  await tick();
  svc.message({ to: 'u-ben', text: 'Please stop' });
  assert.deepEqual(await svc.report({ about: 'u-ben', reason: 'rude' }), { error: 'reason' });
  assert.deepEqual(await svc.report({ about: 'u-ben', reason: 'harassment', note: 'Keeps writing' }), { reported: 'rep_1', messages: 0 });
  assert.deepEqual(out.filter((m) => m.type === 'report').at(-1).messages, []);
  assert.deepEqual(await svc.report({ about: 'u-ben', reason: 'harassment', note: 'Keeps writing', withMessages: true }), { reported: 'rep_1', messages: 2 });
  const sent = out.filter((m) => m.type === 'report').at(-1);
  // Theirs only: what I wrote back is not sent.
  assert.deepEqual(sent.messages.map((m) => m.text), ['Answer me', 'Now']);
  assert.deepEqual([sent.about, sent.reason, sent.note], ['u-ben', 'harassment', 'Keeps writing']);
  // Reporting is not blocking: nothing was blocked.
  assert.equal(out.some((m) => m.type === 'contact.block'), false);
});

test('signing out forgets the groups, and the counts are the next person\'s own', async () => {
  const { svc } = service({ after: [{ type: 'groups', groups: [GROUP] }, { type: 'messages', messages: [{ id: 'm-00000040', from: 'u-ben', text: 'Hi', at: 90 }] }] });
  await online(svc);
  await tick();
  assert.equal(svc.state().groups.length, 1);
  assert.equal(svc.state().unreadTotal, 1);
  svc.signOut();
  assert.deepEqual([svc.state().groups, svc.state().unreadTotal, svc.state().missed], [[], 0, 0]);
});

test('a message the introducer would not take is marked as not sent, to a person and to a group', async () => {
  const { svc, say } = service({ after: [{ type: 'groups', groups: [GROUP] }] });
  await online(svc);
  const one = svc.message({ to: 'u-ben', text: 'x'.repeat(10) });
  const two = svc.message({ group: 'g_1', text: 'Hello all' });
  assert.equal(svc.history({ contactId: 'u-ben' })[0].state, 'sending');
  say({ type: 'error', for: 'msg.send', code: 'too-many', id: one.id });
  say({ type: 'error', for: 'group.send', code: 'no-group', id: two.id, groupId: 'g_1' });
  await tick();
  assert.equal(svc.history({ contactId: 'u-ben' })[0].state, 'failed');
  assert.equal(svc.history({ groupId: 'g_1' }).find((m) => m.id === two.id).state, 'failed');
  // A refusal about something else changes no message.
  const three = svc.message({ to: 'u-ben', text: 'Again' });
  say({ type: 'error', for: 'contact.request', code: 'address' });
  await tick();
  assert.equal(svc.history({ contactId: 'u-ben' }).find((m) => m.id === three.id).state, 'sending');
});

test('every window is told when a conversation is read, the missed calls are looked at, or a call is turned away as busy', async () => {
  const { svc, say, told } = service();
  await online(svc);
  say({ type: 'messages', messages: [{ id: 'm-00000001', from: 'u-ben', text: 'Hello', at: 10 }] });
  say({ type: 'call.incoming', callId: 'c-busy', from: BEN, video: false, sdp: 'x' });
  await tick();
  const states = () => told.filter((t) => t.event === 'talk:state');
  let before = states().length;
  svc.read({ conversation: 'u-ben' });
  assert.equal(states().length, before + 1);
  assert.equal(states().at(-1).payload.unreadTotal, 0);
  before = states().length;
  svc.send({ type: 'call.decline', callId: 'c-busy', reason: 'busy' });
  assert.equal(states().length, before + 1);
  assert.equal(states().at(-1).payload.missed, 1);
  before = states().length;
  assert.equal(svc.calls().length, 1);
  assert.equal(states().length, before, 'only looking at the list changes its count');
  svc.calls({ seen: true });
  assert.equal(states().at(-1).payload.missed, 0);
});
