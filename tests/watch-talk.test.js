// Talk, as the notifier hears of it (apps/desktop/main/watch-talk.js,
// D-2026-11-22-23): a call coming in, a call missed, a message, a meeting
// someone asks the person into. What Talk says is worth a notification is
// handed in here as main/talk-notify.js tells it ('talk:notify'); nothing is
// shown on this computer.

import test from 'node:test';
import assert from 'node:assert/strict';
import { watchTalk, NOTIFY_EVENT } from '../apps/desktop/main/watch-talk.js';
import { NOTIFY_EVENT as TALKS_OWN, NOTIFY_KINDS } from '../apps/desktop/main/talk-notify.js';
import { EVENTS } from '../packages/office-shell/src/contract.js';
import { createNotifier } from '../packages/office-shell/src/electron/notifier.js';

const AMINA = { id: 'u1', name: 'Amina Yusuf' };
const BILAL = { id: 'u2', name: 'Bilal Khan' };
const MEET = 'meet-ABCDEFGHIJ0123456789';

function setup({ counts = true } = {}) {
  const kept = new Map();
  const shown = [];
  const opened = [];
  const notifier = createNotifier({
    settings: { get: (k, f) => (kept.has(k) ? kept.get(k) : f), set: (k, v) => kept.set(k, v) },
    show: (note) => {
      const entry = { ...note, closed: false };
      shown.push(entry);
      return { close: () => { entry.closed = true; } };
    },
    open: (target) => opened.push(target),
    setTimer: () => ({}),
    clearTimer: () => {},
  });
  const listeners = new Map();
  const on = (event, fn) => {
    if (!listeners.has(event)) listeners.set(event, new Set());
    listeners.get(event).add(fn);
    return () => listeners.get(event).delete(fn);
  };
  const emit = (event, payload) => { for (const fn of [...(listeners.get(event) || [])]) fn(payload); };
  const sent = [];
  const ctx = { front: false };
  const talk = watchTalk({ notifier, on, front: (app) => ctx.front && app === 'talk', send: (msg) => { sent.push(msg); return Promise.resolve(); } });
  if (counts) emit('talk:state', { phase: 'online', unreadTotal: 0, missed: 0 });
  const tell = (n) => emit(NOTIFY_EVENT, n);
  return { talk, notifier, emit, tell, shown, opened, sent, ctx, listeners };
}

test('it listens to what Talk itself tells, by the name Talk and the shell give it', () => {
  assert.equal(NOTIFY_EVENT, TALKS_OWN);
  assert.ok(EVENTS.includes(NOTIFY_EVENT));
  assert.deepEqual(NOTIFY_KINDS, ['call', 'call-over', 'missed', 'message', 'meeting', 'meeting-over'], 'each of these is handled here');
});

test('a message is shown with who wrote it and its words, and pressed opens Talk at the person', () => {
  const { tell, shown, opened } = setup();
  tell({ kind: 'message', id: 'm1', from: AMINA, group: null, text: 'Are you free at three?', at: 1 });
  assert.deepEqual(shown.map((n) => [n.kind, n.title, n.body]), [['chat', 'Amina Yusuf', 'Are you free at three?']]);
  shown[0].onClick();
  assert.deepEqual(opened, [{ app: 'talk', query: { contact: 'u1' } }]);
});

test('more from the same person is one notification: the newest words and how many have come, until Talk says all is read', () => {
  const { tell, emit, shown } = setup();
  tell({ kind: 'message', id: 'm1', from: AMINA, group: null, text: 'One', at: 1 });
  tell({ kind: 'message', id: 'm2', from: AMINA, group: null, text: 'Two', at: 2 });
  tell({ kind: 'message', id: 'm3', from: BILAL, group: null, text: 'Hello', at: 3 });
  assert.equal(shown[0].closed, true, 'the earlier one comes down');
  assert.deepEqual(shown.map((n) => [n.title, n.body]), [['Amina Yusuf', 'One'], ['Amina Yusuf', 'Two (2 new messages)'], ['Bilal Khan', 'Hello']]);
  emit('talk:state', { phase: 'online', unreadTotal: 0, missed: 0 });
  tell({ kind: 'message', id: 'm4', from: AMINA, group: null, text: 'Three', at: 4 });
  assert.equal(shown[3].body, 'Three');
});

test('a message in a group says who and in which group, and opens Talk at the group', () => {
  const { tell, shown, opened } = setup();
  tell({ kind: 'message', id: 'm1', from: AMINA, group: { id: 'g1', name: 'Launch' }, text: 'Ready', at: 1 });
  assert.deepEqual([shown[0].title, shown[0].body], ['Amina Yusuf in Launch', 'Ready']);
  shown[0].onClick();
  assert.deepEqual(opened, [{ app: 'talk', query: { group: 'g1' } }]);
});

test('with a Talk window in front a message is neither shown nor counted', () => {
  const { tell, shown, notifier, ctx } = setup({ counts: false });
  ctx.front = true;
  tell({ kind: 'message', id: 'm1', from: AMINA, group: null, text: 'Seen as it comes', at: 1 });
  assert.deepEqual(shown, []);
  assert.equal(notifier.state().waiting.chat, 0);
  ctx.front = false;
  tell({ kind: 'message', id: 'm2', from: AMINA, group: null, text: 'Now', at: 2 });
  assert.equal(shown[0].body, 'Now', 'the first that was not seen');
});

test('how many wait is what Talk says; a Talk that does not say is counted here', () => {
  const a = setup();
  a.tell({ kind: 'message', id: 'm1', from: AMINA, group: null, text: 'One', at: 1 });
  assert.equal(a.notifier.state().waiting.chat, 0, 'Talk has not said so yet');
  a.emit('talk:state', { phase: 'online', unreadTotal: 3, missed: 2 });
  assert.deepEqual([a.notifier.state().waiting.chat, a.notifier.state().waiting.missed], [3, 2]);
  const b = setup({ counts: false });
  b.tell({ kind: 'message', id: 'm1', from: AMINA, group: null, text: 'One', at: 1 });
  b.tell({ kind: 'missed', id: 'u1', callId: 'c1', from: AMINA, video: false, why: 'missed', at: 1, count: 1 });
  assert.deepEqual([b.notifier.state().waiting.chat, b.notifier.state().waiting.missed], [1, 1]);
});

test('a call coming in stays up with Answer and Decline until Talk says it no longer rings', () => {
  const { tell, shown, opened, talk } = setup();
  tell({ kind: 'call', id: 'c1', from: AMINA, video: true });
  assert.deepEqual([shown[0].kind, shown[0].title, shown[0].body, shown[0].sticky], ['call', 'Incoming video call', 'Amina Yusuf', true]);
  assert.deepEqual(shown[0].actions.map((a) => a.id), ['answer', 'decline']);
  assert.deepEqual(talk.ringing(), ['c1']);
  shown[0].onAction('answer');
  assert.deepEqual(opened, [{ app: 'talk', query: { contact: 'u1' } }], 'Talk opens at the caller; its window shows the call to answer');
  for (const why of ['answered', 'declined', 'elsewhere', 'missed', 'ended']) {
    tell({ kind: 'call', id: `c-${why}`, from: AMINA, video: false });
    tell({ kind: 'call-over', id: `c-${why}`, why });
    assert.equal(shown.at(-1).closed, true, why);
  }
  tell({ kind: 'call-over', id: 'c1', why: 'answered' });
  tell({ kind: 'call-over', id: 'c1', why: 'answered' });
  assert.equal(shown[0].closed, true);
  assert.deepEqual(talk.ringing(), []);
});

test('declined from its notification, the call is turned away through Talk, which then says it no longer rings', () => {
  const { tell, shown, sent, opened } = setup();
  tell({ kind: 'call', id: 'c1', from: AMINA, video: false });
  shown[0].onAction('decline');
  assert.deepEqual(sent, [{ type: 'call.decline', callId: 'c1' }]);
  assert.deepEqual(opened, [], 'no window opens to turn a call away');
  tell({ kind: 'call-over', id: 'c1', why: 'declined' });
  assert.equal(shown[0].closed, true);
});

test('a missed call is said so, one notification a caller with how many, and opens Talk at the missed calls', () => {
  const { tell, shown, opened } = setup();
  tell({ kind: 'missed', id: 'u1', callId: 'c1', from: AMINA, video: false, why: 'missed', at: 1, count: 1 });
  assert.deepEqual([shown[0].kind, shown[0].title, shown[0].body], ['missed', 'Missed call', 'Amina Yusuf']);
  tell({ kind: 'missed', id: 'u1', callId: 'c2', from: AMINA, video: false, why: 'off', at: 2, count: 2 });
  assert.equal(shown[0].closed, true);
  assert.deepEqual([shown[1].title, shown[1].body], ['2 missed calls', 'Amina Yusuf']);
  shown[1].onClick();
  assert.deepEqual(opened, [{ app: 'talk', query: { contact: 'u1', show: 'calls' } }]);
});

test('asked into a meeting: who asks, Join, and that it will be streamed or recorded when it will; no longer offered, it comes down', () => {
  const { tell, shown, opened } = setup();
  tell({ kind: 'meeting', id: MEET, from: AMINA, name: 'Launch', live: false, recorded: false, group: null });
  assert.deepEqual([shown[0].kind, shown[0].title, shown[0].body], ['meeting', 'Launch', 'Amina Yusuf asks you to join']);
  assert.deepEqual(shown[0].actions.map((a) => a.id), ['join']);
  shown[0].onAction('join');
  shown[0].onClick();
  assert.deepEqual(opened, [{ app: 'talk', query: { meet: MEET } }, { app: 'talk', query: { meet: MEET } }]);
  tell({ kind: 'meeting-over', id: MEET, why: 'joined' });
  assert.equal(shown[0].closed, true);

  tell({ kind: 'meeting', id: 'live-ABCDEFGHIJ0123456789', from: AMINA, name: '', live: true, recorded: true, group: { id: 'g1', name: 'Team' } });
  assert.equal(shown[1].title, 'Team', 'named by its group when it has no name of its own');
  assert.match(shown[1].body, /^Amina Yusuf asks you to join · /);
  assert.match(shown[1].body, /streamed live/i);
  assert.match(shown[1].body, /recorded/i);
  tell({ kind: 'meeting', id: 'meet-00000000000000000001', from: null, name: '', live: false, recorded: false, group: null });
  assert.deepEqual([shown[2].title, shown[2].body], ['Meeting', '']);
});

test('with a Talk window in front a meeting asked into is not shown here: the window shows it', () => {
  const { tell, shown, ctx } = setup();
  ctx.front = true;
  tell({ kind: 'meeting', id: MEET, from: AMINA, name: 'Launch', live: false, recorded: false, group: null });
  assert.deepEqual(shown, []);
});

test('signed out of Talk, nothing of it stays up; stopped, nothing is heard', () => {
  const { tell, emit, shown, talk, listeners } = setup();
  tell({ kind: 'call', id: 'c1', from: AMINA, video: false });
  tell({ kind: 'meeting', id: MEET, from: AMINA, name: 'Launch', live: false, recorded: false, group: null });
  emit('talk:state', { phase: 'signed-out', unreadTotal: 0, missed: 0 });
  assert.deepEqual(shown.map((n) => n.closed), [true, true]);
  tell({ kind: 'call', id: 'c2', from: AMINA, video: false });
  talk.stop();
  assert.equal(shown[2].closed, true);
  assert.equal(listeners.get(NOTIFY_EVENT).size, 0);
  assert.equal(listeners.get('talk:state').size, 0);
});

test('what is no notification is passed over, and a fault here never reaches Talk', () => {
  const { tell, emit, shown, notifier } = setup();
  for (const n of [null, 'call', {}, { kind: 'call' }, { kind: 'unknown', id: 1 }, { kind: 'call-over', id: 'never-rang' }, { kind: 'meeting-over', id: 'never-asked' }]) tell(n);
  emit('talk:state', null);
  assert.deepEqual(shown, []);
  notifier.report = () => { throw new Error('a notifier with a fault'); };
  assert.doesNotThrow(() => tell({ kind: 'message', id: 'm1', from: AMINA, group: null, text: 'x', at: 1 }));
});
