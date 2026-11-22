// Talk, as the notifier hears of it (apps/desktop/main/watch-talk.js,
// D-2026-11-22-23): a call coming in, a call missed, a message. Talk's
// events are handed in here as main/talk.js tells them to its windows;
// nothing is shown on this computer.

import test from 'node:test';
import assert from 'node:assert/strict';
import { watchTalk, RING_MS } from '../apps/desktop/main/watch-talk.js';
import { createNotifier } from '../packages/office-shell/src/electron/notifier.js';

const CONTACTS = [{ id: 'u1', name: 'Amina Yusuf', email: 'amina@example.org' }, { id: 'u2', name: '', email: 'bilal@example.org' }];

function setup({ windows = [] } = {}) {
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
  const timers = [];
  const sent = [];
  const told = [];
  const ctx = { front: false };
  const talk = watchTalk({
    notifier,
    on,
    front: (app) => ctx.front && app === 'talk',
    send: (msg) => { sent.push(msg); return Promise.resolve(); },
    tellWindow: (win, msg) => told.push([win, msg]),
    talkWindows: () => windows,
    setTimer: (fn, ms) => { const t = { fn, ms }; timers.push(t); return t; },
    clearTimer: (t) => { const i = timers.indexOf(t); if (i >= 0) timers.splice(i, 1); },
  });
  emit('talk:state', { phase: 'online', me: { id: 'me' }, contacts: CONTACTS });
  return { talk, notifier, emit, shown, opened, timers, sent, told, ctx, listeners };
}

const up = (shown) => shown.filter((n) => !n.closed);

test('a message is told by the name Talk knows the person by, and pressed opens Talk at them', () => {
  const { emit, shown, opened, notifier } = setup();
  emit('talk:event', { type: 'messages', messages: [{ id: 'm1', from: 'u1', text: 'Are you free?', at: 10 }] });
  assert.equal(shown.length, 1);
  assert.deepEqual([shown[0].title, shown[0].body, shown[0].kind], ['Amina Yusuf', 'Are you free?', 'chat']);
  shown[0].onClick();
  assert.deepEqual(opened, [{ app: 'talk', query: { contact: 'u1' } }]);
  assert.equal(notifier.state().waiting.chat, 1);
});

test('several messages at once: one notification a person, the newest words and how many, and none for my own', () => {
  const { emit, shown, notifier } = setup();
  emit('talk:event', { type: 'messages', messages: [
    { id: 'm1', from: 'u1', text: 'first', at: 10 },
    { id: 'm2', from: 'u2', text: 'hello', at: 11 },
    { id: 'm3', from: 'u1', text: 'second', at: 12 },
    { id: 'm4', from: 'me', text: 'sent from my other computer', at: 13 },
  ] });
  assert.equal(shown.length, 2);
  assert.deepEqual([shown[0].title, shown[0].body], ['Amina Yusuf', 'second (2 new messages)']);
  assert.deepEqual([shown[1].title, shown[1].body], ['bilal@example.org', 'hello'], 'a person with no name is told by their address');
  assert.equal(notifier.state().waiting.chat, 3);
});

test('with a Talk window in front a message is neither shown nor counted', () => {
  const { emit, shown, notifier, ctx } = setup();
  ctx.front = true;
  emit('talk:event', { type: 'messages', messages: [{ id: 'm1', from: 'u1', text: 'hi', at: 1 }] });
  assert.equal(shown.length, 0);
  assert.equal(notifier.state().waiting.chat, 0);
});

test('a message in a group says who and in which group, and opens Talk at the group', () => {
  const { emit, shown, opened } = setup();
  emit('talk:event', { type: 'group.messages', groupId: 'g7', group: { name: 'Accounts' }, messages: [{ id: 'm1', from: 'u1', text: 'Month end is Friday', at: 1 }] });
  assert.equal(shown.length, 1);
  assert.match(shown[0].title, /Amina Yusuf/);
  assert.match(shown[0].title, /Accounts/);
  shown[0].onClick();
  assert.deepEqual(opened, [{ app: 'talk', query: { group: 'g7' } }]);
});

test('where Talk says how many wait, its numbers are the ones by the clock', () => {
  const { emit, notifier, shown } = setup();
  emit('talk:state', { phase: 'online', me: { id: 'me' }, contacts: CONTACTS, unreadTotal: 4, missed: 1 });
  assert.deepEqual(notifier.state().waiting, { missed: 1, chat: 4, mail: 0, total: 5 });
  emit('talk:event', { type: 'messages', messages: [{ id: 'm1', from: 'u1', text: 'hi', at: 1 }] });
  assert.equal(shown.length, 1, 'still shown');
  assert.equal(notifier.state().waiting.chat, 4, 'and not counted twice');
  emit('talk:state', { phase: 'online', me: { id: 'me' }, contacts: CONTACTS, unreadTotal: 0, missed: 0 });
  assert.equal(notifier.state().waiting.total, 0);
});

test('a call coming in is shown until it stops; answered in a window, it comes down and was not missed', () => {
  const { emit, shown, talk, timers, opened } = setup();
  emit('talk:event', { type: 'call.incoming', callId: 'c1', from: { id: 'u1', name: 'Amina' }, video: false });
  assert.equal(shown.length, 1);
  assert.deepEqual([shown[0].kind, shown[0].title, shown[0].body, shown[0].sticky], ['call', 'Incoming call', 'Amina Yusuf', true]);
  assert.deepEqual(talk.ringing(), ['c1']);
  assert.equal(timers.length, 1);
  assert.equal(timers[0].ms, RING_MS);
  shown[0].onAction('answer');
  assert.deepEqual(opened, [{ app: 'talk', query: { contact: 'u1', call: 'c1' } }], 'Answer opens Talk at the call');
  assert.deepEqual(talk.ringing(), ['c1'], 'and the call rings on until Talk answers it');
  talk.sending({ type: 'call.answer', callId: 'c1', sdp: 'x' });
  assert.deepEqual(talk.ringing(), []);
  assert.equal(shown[0].closed, true);
  assert.equal(timers.length, 0);
  assert.equal(shown.length, 1, 'no missed call');
});

test('the same call told twice rings once', () => {
  const { emit, shown } = setup();
  emit('talk:event', { type: 'call.incoming', callId: 'c1', from: { id: 'u1' } });
  emit('talk:event', { type: 'call.incoming', callId: 'c1', from: { id: 'u1' } });
  assert.equal(shown.length, 1);
});

test('declined from its notification, the call is turned away through Talk and was not missed', () => {
  const { emit, shown, sent, talk } = setup();
  emit('talk:event', { type: 'call.incoming', callId: 'c1', from: { id: 'u1' } });
  shown[0].onAction('decline');
  assert.deepEqual(sent, [{ type: 'call.decline', callId: 'c1' }]);
  assert.deepEqual(talk.ringing(), []);
  assert.equal(shown[0].closed, true);
  assert.equal(shown.length, 1);
  shown[0].onAction('decline');
  assert.equal(sent.length, 1, 'once');
});

test('a call that stops ringing unanswered was missed: said so, counted, and opening Talk at the person', () => {
  const { emit, shown, opened, notifier } = setup();
  emit('talk:event', { type: 'call.incoming', callId: 'c1', from: { id: 'u1' }, video: true });
  assert.equal(shown[0].title, 'Incoming video call');
  emit('talk:event', { type: 'call.ended', callId: 'c1' });
  assert.equal(shown[0].closed, true);
  assert.equal(shown.length, 2);
  assert.deepEqual([shown[1].kind, shown[1].title, shown[1].body], ['missed', 'Missed call', 'Amina Yusuf']);
  assert.equal(notifier.state().waiting.missed, 1);
  shown[1].onClick();
  assert.deepEqual(opened, [{ app: 'talk', query: { contact: 'u1', show: 'calls' } }]);
  emit('talk:event', { type: 'call.ended', callId: 'c1' });
  assert.equal(shown.length, 2, 'the end of a call that was answered or already told is nothing');
});

test('a call nothing says has stopped is taken as missed after a while; one taken on another computer is not', () => {
  const a = setup();
  a.emit('talk:event', { type: 'call.incoming', callId: 'c1', from: { id: 'u2' } });
  a.timers[0].fn();
  assert.deepEqual(up(a.shown).map((n) => n.kind), ['missed']);
  assert.deepEqual(a.talk.ringing(), []);

  const b = setup();
  b.emit('talk:event', { type: 'call.incoming', callId: 'c2', from: { id: 'u2' } });
  b.emit('talk:event', { type: 'call.taken', callId: 'c2' });
  assert.equal(b.shown.length, 1);
  assert.equal(b.shown[0].closed, true);
  assert.equal(b.notifier.state().waiting.missed, 0);
});

test('turned away by a window because another call is on, the call was missed; turned away by hand, it was not', () => {
  const { emit, shown, talk } = setup();
  emit('talk:event', { type: 'call.incoming', callId: 'c1', from: { id: 'u1' } });
  talk.sending({ type: 'call.decline', callId: 'c1', reason: 'busy' });
  assert.deepEqual(up(shown).map((n) => n.kind), ['missed']);
  emit('talk:event', { type: 'call.incoming', callId: 'c2', from: { id: 'u1' } });
  talk.sending({ type: 'call.decline', callId: 'c2' });
  assert.deepEqual(shown.map((n) => n.kind), ['call', 'missed', 'call']);
  talk.sending({ type: 'message', to: 'u1', text: 'hi' });
  talk.sending(null);
});

test('a Talk window that starts while a call rings is told of it once; one that heard it ring is not told again', () => {
  const open = { id: 'already-open' };
  const { emit, talk, told } = setup({ windows: [open] });
  const call = { type: 'call.incoming', callId: 'c1', from: { id: 'u1' }, sdp: 'offer' };
  emit('talk:event', call);
  talk.windowReady(open);
  assert.equal(told.length, 0, 'the window that was open heard it as it came');
  const fresh = { id: 'new' };
  talk.windowReady(fresh);
  talk.windowReady(fresh);
  talk.windowReady(null);
  assert.deepEqual(told, [[fresh, call]]);
  emit('talk:event', { type: 'call.taken', callId: 'c1' });
  talk.windowReady({ id: 'later' });
  assert.equal(told.length, 1, 'nothing rings now');
});

test('signed out of Talk, nothing rings; stopped, nothing is heard', () => {
  const { emit, shown, talk, listeners } = setup();
  emit('talk:event', { type: 'call.incoming', callId: 'c1', from: { id: 'u1' } });
  emit('talk:state', { phase: 'signed-out', contacts: [] });
  assert.deepEqual(talk.ringing(), []);
  assert.equal(shown[0].closed, true);
  emit('talk:event', { type: 'call.incoming', callId: 'c2', from: { id: 'u1' } });
  talk.stop();
  assert.deepEqual(talk.ringing(), []);
  assert.equal([...listeners.values()].reduce((n, s) => n + s.size, 0), 0);
  emit('talk:event', { type: 'messages', messages: [{ id: 'm', from: 'u1', text: 'x' }] });
  assert.equal(shown.length, 2);
});

test('an event that is no event is passed over', () => {
  const { emit, shown } = setup();
  emit('talk:event', null);
  emit('talk:event', { type: 'presence', id: 'u1' });
  emit('talk:event', { type: 'messages' });
  emit('talk:event', { type: 'call.incoming' });
  emit('talk:state', null);
  assert.equal(shown.length, 0);
});
