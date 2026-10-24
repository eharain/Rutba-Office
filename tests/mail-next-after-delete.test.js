// Deleting a message keeps the place in the list (owner, 2026-10-24: "the
// email deletion with delete button or delete from list reset the selection
// back to first email, while it should be selecting the next rather than
// first"): the message after the last one removed is opened, or at the foot
// of the list the one before, as Outlook does.
import test from 'node:test';
import assert from 'node:assert/strict';
import { nextAfterRemoved, rowKey } from '../apps/desktop/renderer/apps/mail/parts.js';

const row = (id, accountId = 'a', folder = 'Inbox') => ({ accountId, folder, id });
const list = ['1', '2', '3', '4'].map((id) => row(id));
const ids = (r) => r && r.id;

test('the middle of three removed: the one after it', () => {
  assert.equal(ids(nextAfterRemoved(list.slice(0, 3), [rowKey(list[1])])), '3');
});

test('the last removed: the one before it', () => {
  assert.equal(ids(nextAfterRemoved(list, [rowKey(list[3])])), '3');
  assert.equal(ids(nextAfterRemoved(list, [rowKey(list[2]), rowKey(list[3])])), '2');
});

test('several removed: the one after the last of them, wherever they were', () => {
  assert.equal(ids(nextAfterRemoved(list, [rowKey(list[0]), rowKey(list[2])])), '4');
  assert.equal(ids(nextAfterRemoved(list, [rowKey(list[0]), rowKey(list[3])])), '3');
});

test('everything removed, or nothing listed removed: nothing to open', () => {
  assert.equal(nextAfterRemoved(list, list.map(rowKey)), null);
  assert.equal(nextAfterRemoved(list, [rowKey(row('9'))]), null);
  assert.equal(nextAfterRemoved([], [rowKey(list[0])]), null);
});

test('All accounts: the same id in another account or folder is another message', () => {
  const mixed = [row('1', 'a'), row('1', 'b'), row('1', 'b', 'Work')];
  assert.deepEqual(nextAfterRemoved(mixed, [rowKey(mixed[0])]), { accountId: 'b', folder: 'Inbox', id: '1' });
  assert.deepEqual(nextAfterRemoved(mixed, [rowKey(mixed[2])]), { accountId: 'b', folder: 'Inbox', id: '1' });
});

test('conversations: a thread is removed whole, and the next thread opens on its own newest message', () => {
  const threads = [
    { ...row('3'), messages: [row('1'), row('3')] },
    { ...row('5'), messages: [row('4'), row('5')] },
  ];
  assert.equal(ids(nextAfterRemoved(threads, [rowKey(row('1')), rowKey(row('3'))], (t) => t.messages)), '5');
});
