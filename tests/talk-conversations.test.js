// Talk made better (Office 2.28): unread counts, missed calls, the mark on a
// message that has arrived, group conversations and reporting someone, as
// the Talk service keeps them, on a stand-in introducer.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTalkService } from '../apps/desktop/main/talk.js';
import { groupKey, CAN, SENDING_KEEP_MS } from '../apps/desktop/main/talk-conversations.js';

const folders = [];
test.after(() => { for (const f of folders) fs.rmSync(f, { recursive: true, force: true }); });

const AMIRA = { id: 'u-a', name: 'Amira Khan', email: 'amira@example.com' };
const BEN = { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com', presence: 'online' };
const CARA = { id: 'u-cara', name: 'Cara Lund', email: 'cara@example.com', presence: 'offline' };
const GROUP = { id: 'g_1', name: 'Launch', starter: 'u-a', created: 5, members: [{ id: 'u-a', name: 'Amira Khan' }, { id: 'u-ben', name: 'Ben Ortiz', contact: true }, { id: 'u-cara', name: 'Cara Lund', contact: true }] };

/** Every stand-in made by a test, so `tick()` can wait for each to have said all it was going to. */
const live = [];

/** The Talk service on a fake introducer that answers as `answer(msg)` says; `after` is said once the welcome has gone. */
function service({ answer = () => null, after = [], dataDir = null, contacts = [BEN, CARA] } = {}) {
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const out = [];
  const told = [];
  const dir = dataDir || fs.mkdtempSync(path.join(os.tmpdir(), 'talk-conversations-'));
  if (!dataDir) folders.push(dir);
  let ws = null;
  const say = (m) => setImmediate(() => ws.onmessage({ data: JSON.stringify(m) }));
  /** Waiting for something the service tells its windows: each is asked at every telling, and goes once it says yes. */
  const watchers = new Set();
  const until = (seen) => new Promise((resolve) => {
    const w = (t) => { if (seen(t)) { watchers.delete(w); resolve(t); } };
    watchers.add(w);
  });
  let marks = 0;
  /**
   * Everything said so far has been heard and answered. The stand-in says a
   * mark of its own after whatever is already on its way, and the service
   * passes every message on to its windows in the order it came; when the
   * mark is heard and nothing more was sent meanwhile, nothing is left.
   */
  const settled = async () => {
    let sent;
    do {
      sent = out.length;
      const mark = ++marks;
      const heard = until((t) => t.event === 'talk:event' && t.payload?.type === 'test.mark' && t.payload.mark === mark);
      say({ type: 'test.mark', mark });
      await heard;
    } while (out.length !== sent);
  };
  const svc = createTalkService({
    stores: { settings: map(new Map()), secrets: map(new Map([['talk.token', 'amira']])) },
    dataDir: dir,
    broadcast: (event, payload) => { const t = { event, payload }; told.push(t); for (const w of [...watchers]) w(t); },
    connect: () => {
      ws = { readyState: 1, send: (text) => {
        const msg = JSON.parse(text);
        out.push(msg);
        if (msg.type === 'hello') { say({ type: 'welcome', me: AMIRA, contacts, requests: { incoming: [], outgoing: [] } }); for (const m of after) say(m); return; }
        for (const reply of [].concat(answer(msg) || [])) say(reply);
      }, close: () => {} };
      setImmediate(() => ws.onopen());
      return ws;
    },
  });
  /** The connection is lost, as when the network goes. */
  const drop = () => { const was = ws; was.readyState = 3; was.onclose(); };
  const it = { svc, out, say, told, dir, until, settled, drop };
  live.push(it);
  return it;
}
/** Connected and welcomed: waited for as the service tells its windows so, never for a length of time. */
const online = async (svc) => {
  const it = live.find((s) => s.svc === svc);
  const there = it.until(() => svc.state().phase === 'online');
  svc.resume();
  await there;
  assert.equal(svc.state().phase, 'online');
};
/** Every stand-in that is connected has said all it was going to, and been heard. */
const tick = async () => { for (const it of live) if (it.svc.state().phase === 'online') await it.settled(); };

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
  // A note that names one person (an introducer that still told who and when) marks nothing and keeps nothing of them.
  say({ type: 'notices', notices: [{ id: 'n_5', kind: 'delivered', group: 'g_1', ids: [entry.id], by: 'u-ben', at: 71 }] });
  await tick();
  assert.deepEqual([one().state, 'arrived' in one()], ['sent', false]);
  // The introducer says it once, when everyone has it.
  say({ type: 'notices', notices: [{ id: 'n_6', kind: 'delivered', group: 'g_1', ids: [entry.id], all: true }] });
  await tick();
  assert.deepEqual([one().state, 'arrived' in one()], ['arrived', false]);
  assert.deepEqual(out.filter((m) => m.type === 'notice.ack').map((m) => m.ids), [['n_5'], ['n_6']]);
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

/* ── what the review of phase 1 asked for ─────────────────────────── */

test("an id is one message's: someone else's words under it are kept as their own message, never in place of it", async () => {
  const { svc, say, out } = service({ after: [{ type: 'groups', groups: [GROUP] }] });
  await online(svc);
  say({ type: 'group.messages', messages: [{ id: 'm-00000100', group: 'g_1', from: 'u-ben', text: 'The real words', at: 60 }] });
  await tick();
  // Cara sends the same id with words of her own (an introducer that does not refuse it).
  say({ type: 'group.messages', messages: [{ id: 'm-00000100', group: 'g_1', from: 'u-cara', text: 'Forged words', at: 61 }] });
  await tick();
  const kept = () => svc.history({ groupId: 'g_1' }).map((e) => [e.id, e.from, e.text]);
  assert.deepEqual(kept(), [['m-00000100', 'u-ben', 'The real words'], ['m-00000100~u-cara', 'u-cara', 'Forged words']]);
  assert.equal(svc.state().unread[groupKey('g_1')], 2);
  // Ben's own id again with other words changes nothing either, and is not counted.
  say({ type: 'group.messages', messages: [{ id: 'm-00000100', group: 'g_1', from: 'u-ben', text: 'Changed afterwards', at: 62 }] });
  say({ type: 'group.messages', messages: [{ id: 'm-00000100', group: 'g_1', from: 'u-cara', text: 'Forged again', at: 63 }] });
  await tick();
  assert.deepEqual(kept(), [['m-00000100', 'u-ben', 'The real words'], ['m-00000100~u-cara', 'u-cara', 'Forged words']]);
  assert.equal(svc.state().unread[groupKey('g_1')], 2);
  // Each is still acknowledged, so the introducer lets it go.
  assert.equal(out.filter((m) => m.type === 'group.ack').length, 4);

  // Between two people: a contact cannot put their words in the place of a message of mine.
  const mine = svc.message({ to: 'u-ben', text: 'I agree to nothing' });
  say({ type: 'messages', messages: [{ id: mine.id, from: 'u-ben', text: 'I agree to everything', at: 70 }] });
  await tick();
  const two = svc.history({ contactId: 'u-ben' });
  assert.deepEqual(two.map((e) => [e.id, e.mine, e.text]), [[mine.id, true, 'I agree to nothing'], [`${mine.id}~u-ben`, false, 'I agree to everything']]);
});

test('unread does not drift: a message given again is counted once, and a conversation that is gone holds nothing up', async () => {
  const first = service({ after: [{ type: 'groups', groups: [GROUP] }] });
  const { svc, say, out } = first;
  await online(svc);
  const hello = { type: 'messages', messages: [{ id: 'm-00000110', from: 'u-ben', text: 'Hello', at: 10 }] };
  say(hello);
  // The introducer did not hear that Office has it, and gives it again.
  say(hello);
  say({ type: 'messages', messages: [{ id: 'm-00000111', from: 'u-cara', text: 'Lunch?', at: 11 }] });
  say({ type: 'group.messages', messages: [{ id: 'm-00000112', group: 'g_1', from: 'u-ben', text: 'All here?', at: 12 }] });
  await tick();
  assert.deepEqual(svc.state().unread, { 'u-ben': 1, 'u-cara': 1, [groupKey('g_1')]: 1 });
  assert.equal(svc.history({ contactId: 'u-ben' }).length, 1);
  assert.deepEqual(out.filter((m) => m.type === 'msg.ack').map((m) => m.ids), [['m-00000110'], ['m-00000110'], ['m-00000111']], 'given again, acknowledged again');

  // Ben is blocked, or removes this person: his count goes with him.
  say({ type: 'contact.blocked', id: 'u-ben' });
  await tick();
  assert.deepEqual(svc.state().unread, { 'u-cara': 1, [groupKey('g_1')]: 1 });
  assert.equal(svc.state().unreadTotal, 2);

  // Next start: Cara is no longer a contact and the person was taken out of the group while away.
  const second = service({ dataDir: first.dir, contacts: [BEN], after: [{ type: 'groups', groups: [] }] });
  await online(second.svc);
  await tick();
  assert.deepEqual([second.svc.state().unread, second.svc.state().unreadTotal], [{}, 0]);
  // Gone from the file too, so they never come back.
  const third = service({ dataDir: first.dir, after: [{ type: 'groups', groups: [GROUP] }] });
  await online(third.svc);
  await tick();
  assert.equal(third.svc.state().unreadTotal, 0);
});

test('a count for a group the person is no longer in is not shown, even by an introducer that does not say so', async () => {
  const first = service({ after: [{ type: 'groups', groups: [GROUP] }] });
  await online(first.svc);
  first.say({ type: 'group.messages', messages: [{ id: 'm-00000120', group: 'g_1', from: 'u-ben', text: 'Hi', at: 12 }] });
  await tick();
  assert.equal(first.svc.state().unreadTotal, 1);
  // An introducer from before says nothing of groups when there are none.
  const second = service({ dataDir: first.dir });
  await online(second.svc);
  await tick();
  assert.deepEqual([second.svc.state().unread, second.svc.state().unreadTotal], [{}, 0]);
});

test('a message on its way when the connection dropped is sent again under its own id, also after a restart, and never waits for ever', async () => {
  let answering = false;
  const answer = (m) => {
    if (!answering) return null;
    if (m.type === 'msg.send') return { type: 'msg.sent', id: m.id, at: 20 };
    if (m.type === 'group.send') return { type: 'group.sent', id: m.id, groupId: m.groupId, at: 21, to: ['u-ben', 'u-cara'] };
    return null;
  };
  const first = service({ answer, after: [{ type: 'groups', groups: [GROUP] }] });
  await online(first.svc);
  const one = first.svc.message({ to: 'u-ben', text: 'Did this go?' });
  const two = first.svc.message({ group: 'g_1', text: 'And this?' });
  // The connection goes before the introducer answers, and comes back.
  first.drop();
  assert.equal(first.svc.state().phase, 'offline');
  assert.equal(first.svc.history({ contactId: 'u-ben' })[0].state, 'sending');
  answering = true;
  await online(first.svc);
  await tick();
  const again = first.out.filter((m) => m.type === 'msg.send' || m.type === 'group.send');
  assert.deepEqual(again.map((m) => [m.type, m.id, m.text]), [['msg.send', one.id, 'Did this go?'], ['group.send', two.id, 'And this?'], ['msg.send', one.id, 'Did this go?'], ['group.send', two.id, 'And this?']]);
  assert.equal(first.svc.history({ contactId: 'u-ben' })[0].state, 'sent');
  assert.equal(first.svc.history({ groupId: 'g_1' })[0].state, 'sent');
  // Answered for: not sent a third time at the next connection.
  first.drop();
  await online(first.svc);
  await tick();
  assert.equal(first.out.filter((m) => m.type === 'msg.send').length, 2);

  // Office closed with one on its way: sent again at the next start.
  answering = false;
  const three = first.svc.message({ to: 'u-cara', text: 'Before closing' });
  answering = true;
  const second = service({ answer, dataDir: first.dir });
  await online(second.svc);
  await tick();
  assert.deepEqual(second.out.filter((m) => m.type === 'msg.send').map((m) => [m.id, m.to, m.text]), [[three.id, 'u-cara', 'Before closing']]);
  assert.equal(second.svc.history({ contactId: 'u-cara' })[0].state, 'sent');
});

test('a message that could not be sent for a day says "not sent" in place of waiting', async () => {
  const first = service();
  await online(first.svc);
  const one = first.svc.message({ to: 'u-ben', text: 'Too late now' });
  // A day and more passes with no connection: what the file says is made that old.
  const folder = path.join(first.dir, fs.readdirSync(first.dir).find((f) => fs.statSync(path.join(first.dir, f)).isDirectory()));
  const file = path.join(folder, 'marks.json');
  const marks = JSON.parse(fs.readFileSync(file, 'utf8'));
  assert.deepEqual(Object.keys(marks.sending), [one.id]);
  marks.sending[one.id].at = Date.now() - SENDING_KEEP_MS - 1000;
  fs.writeFileSync(file, JSON.stringify(marks));
  const second = service({ dataDir: first.dir });
  await online(second.svc);
  await tick();
  assert.equal(second.svc.history({ contactId: 'u-ben' })[0].state, 'failed');
  assert.equal(second.out.some((m) => m.type === 'msg.send'), false, 'not sent a day late');
  assert.deepEqual(JSON.parse(fs.readFileSync(file, 'utf8')).sending, {});
});

test("someone else's change to a group never answers this person's own question, and a refusal shows as a refusal", async () => {
  let first = null;
  const { svc } = service({
    after: [{ type: 'groups', groups: [GROUP] }],
    answer: (m) => {
      const elsewhere = { type: 'group.changed', group: { ...GROUP, id: 'g_9', name: 'Other', created: 20 }, event: { kind: 'renamed', by: { id: 'u-ben', name: 'Ben Ortiz' }, name: 'Other', at: 5 } };
      if (m.type === 'group.remove') return [elsewhere, { type: 'error', for: 'group.remove', code: 'not-yours', ref: m.ref }];
      if (m.type === 'group.add') return [elsewhere, { type: 'error', for: 'group.add', code: 'too-many-groups', ref: m.ref, who: [{ id: 'u-cara', name: 'Cara Lund' }] }];
      if (m.type === 'group.create') return [{ type: 'group.created', group: { ...GROUP, id: 'g_8' } }, { type: 'error', for: 'group.create', code: 'too-many', ref: m.ref, most: 50 }];
      if (m.type === 'report') return [{ type: 'report.taken', id: 'rep_other' }, { type: 'error', for: 'report', code: 'too-many', ref: m.ref }];
      if (m.type === 'group.leave') return [{ type: 'group.left', groupId: m.groupId, by: { id: 'u-ben', name: 'Ben Ortiz' } }, { type: 'error', for: 'group.leave', code: 'no-group', ref: m.ref }];
      if (m.type === 'group.rename' && m.name === 'First') { first = m.ref; return null; }
      // The second of two renames is refused before the first is answered.
      if (m.type === 'group.rename') return [{ type: 'error', for: 'group.rename', code: 'too-many', ref: m.ref }, { type: 'group.changed', ref: first, group: { ...GROUP, name: 'First' }, event: { kind: 'renamed', by: { id: 'u-a', name: 'Amira Khan' }, name: 'First', at: 6 } }];
      return null;
    },
  });
  await online(svc);
  assert.deepEqual(await svc.groupRemove({ groupId: 'g_1', member: 'u-ben' }), { error: 'not-yours' });
  // What the wording needs comes with the refusal.
  assert.deepEqual(await svc.groupAdd({ groupId: 'g_1', members: ['u-cara'] }), { error: 'too-many-groups', who: [{ id: 'u-cara', name: 'Cara Lund' }] });
  assert.deepEqual(await svc.groupCreate({ name: 'Big', members: ['u-ben', 'u-cara'] }), { error: 'too-many', most: 50 });
  assert.deepEqual(await svc.report({ about: 'u-ben', reason: 'spam' }), { error: 'too-many' });
  assert.deepEqual(await svc.groupLeave({ groupId: 'g_1' }), { error: 'no-group' });
  const [a, b] = await Promise.all([svc.groupRename({ groupId: 'g_1', name: 'First' }), svc.groupRename({ groupId: 'g_1', name: 'Second' })]);
  assert.equal(a.group?.name, 'First');
  assert.deepEqual(b, { error: 'too-many' });
});

test('missed calls fold: someone calling again and again is one entry and one line, each with a count', async () => {
  const { svc, say } = service();
  await online(svc);
  const from = { id: 'u-ben', name: 'Ben Ortiz' };
  say({ type: 'notices', notices: [{ id: 'n_a', kind: 'missed-call', callId: 'call-0101', from, video: false, why: 'offline', at: 40, count: 1 }] });
  // Another before Office had said it has the note: the same note, told anew.
  say({ type: 'notices', notices: [{ id: 'n_a', kind: 'missed-call', callId: 'call-0102', from, video: true, why: 'offline', at: 41, count: 2 }] });
  // Given again as it was: nothing more.
  say({ type: 'notices', notices: [{ id: 'n_a', kind: 'missed-call', callId: 'call-0102', from, video: true, why: 'offline', at: 41, count: 2 }] });
  await tick();
  assert.equal(svc.state().missed, 1);
  assert.deepEqual(svc.calls().map((c) => [c.id, c.with, c.count, c.video, c.at]), [['call-0102', 'u-ben', 2, true, 41]]);
  assert.deepEqual(svc.history({ contactId: 'u-ben' }).map((e) => [e.kind, e.count]), [['missed-call', 2]]);

  // Something is said, then a note that stands for three calls made while this computer was off.
  say({ type: 'messages', messages: [{ id: 'm-00000130', from: 'u-ben', text: 'Call me', at: 42 }] });
  say({ type: 'notices', notices: [{ id: 'n_b', kind: 'missed-call', callId: 'call-0105', from, video: false, why: 'offline', at: 50, count: 3 }] });
  await tick();
  // Still the one entry, not yet looked at; the line is a new one, since a message came between.
  assert.deepEqual(svc.calls().map((c) => [c.id, c.count]), [['call-0105', 5]]);
  assert.deepEqual(svc.history({ contactId: 'u-ben' }).map((e) => [e.kind || 'message', e.count ?? null]), [['missed-call', 2], ['message', null], ['missed-call', 3]]);

  // A hundred more fill neither the list nor the conversation.
  for (let i = 0; i < 100; i++) say({ type: 'notices', notices: [{ id: `n_c${i}`, kind: 'missed-call', callId: `call-02${String(i).padStart(2, '0')}`, from, video: false, why: 'offline', at: 60 + i, count: 1 }] });
  await tick();
  assert.deepEqual(svc.calls().map((c) => c.count), [105]);
  assert.equal(svc.history({ contactId: 'u-ben' }).length, 3);
  assert.equal(svc.history({ contactId: 'u-ben' }).at(-1).count, 103);

  // Looked at: the next call missed is a new entry. One that rang here and was also noted is the one call.
  svc.calls({ seen: true });
  say({ type: 'call.incoming', callId: 'call-0301', from: BEN, sdp: 'v=0', video: false });
  say({ type: 'call.ended', callId: 'call-0301' });
  say({ type: 'notices', notices: [{ id: 'n_d', kind: 'missed-call', callId: 'call-0301', from, video: false, why: 'unanswered', at: 200, count: 1 }] });
  await tick();
  assert.deepEqual(svc.calls().map((c) => [c.id, c.count, c.seen]), [['call-0301', 1, false], ['call-0299', 105, true]]);
  assert.equal(svc.state().missed, 1);
  // Someone else's call is their own entry.
  say({ type: 'notices', notices: [{ id: 'n_e', kind: 'missed-call', callId: 'call-0401', from: { id: 'u-cara', name: 'Cara Lund' }, video: false, why: 'busy', at: 201, count: 1 }] });
  await tick();
  assert.equal(svc.state().missed, 2);
});
