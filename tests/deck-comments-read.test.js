// Review → Mark All as Read, the engine's half. A deck does not say who has
// read its comments, so what was read is kept on this computer as ids: a
// thread's own and one for each reply. Someone else's words are unread
// until read here; your own never are.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, commentIds, unreadThreads, markRead } from '@rutba/presentation';

function deckWithComments() {
  const deck = Deck.open(buildPptx({ title: 'Read', slides: [{ layout: 'title', title: 'One' }, { layout: 'title', title: 'Two' }] }));
  const a = deck.addComment(0, { text: 'Is this right?', author: 'Sara Khan' });
  deck.replyComment(0, a, { text: 'Yes.', author: 'Me Myself' });
  const b = deck.addComment(1, { text: 'Mine.', author: 'Me Myself' });
  return { deck, a, b };
}

test('every thread and reply has an id to keep, a reply under its thread', () => {
  const { deck, a, b } = deckWithComments();
  const ids = commentIds(deck.comments());
  assert.equal(ids.length, 3);
  assert.ok(ids.includes(a) && ids.includes(b));
  assert.ok(ids.some((id) => id.startsWith(`${a}/`)), 'the reply is kept under its thread');
});

test('someone else\'s thread is unread until it is marked, your own never is, and a new reply from them makes it unread again', () => {
  const { deck, a } = deckWithComments();
  const me = 'Me Myself';
  assert.deepEqual([...unreadThreads(deck.comments(), [], me)], [a], 'only Sara\'s thread');
  const read = markRead([], deck.comments());
  assert.equal(unreadThreads(deck.comments(), read, me).size, 0, 'Mark All as Read leaves nothing unread');
  deck.replyComment(0, a, { text: 'One more thing.', author: 'Sara Khan' });
  assert.deepEqual([...unreadThreads(deck.comments(), read, me)], [a], 'her new reply makes it unread');
  const thread = deck.comments().filter((c) => c.id === a);
  assert.equal(unreadThreads(deck.comments(), markRead(read, thread, deck.comments()), me).size, 0, 'opening the thread reads it');
});

test('ids of comments since deleted are not kept, and the list survives being stored as JSON', () => {
  const { deck, a } = deckWithComments();
  const read = markRead([], deck.comments());
  deck.removeComment(0, a);
  const kept = markRead(JSON.parse(JSON.stringify(read)), [], deck.comments());
  assert.equal(kept.length, 1, 'only the thread still in the deck');
  assert.ok(!kept.some((id) => id.startsWith(a)));
});

test('without a name for you, every comment of anyone\'s counts until read', () => {
  const { deck } = deckWithComments();
  assert.equal(unreadThreads(deck.comments(), [], '').size, 2);
});
