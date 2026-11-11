// Delete beside each message in the list (Mail 2.25): what the control at a
// row's end takes, and where there is none.
import test from 'node:test';
import assert from 'node:assert/strict';
import { rowDeleteTargets } from '../apps/desktop/renderer/apps/mail/parts.js';

const row = (id, accountId = 'a', folder = 'Inbox') => ({ accountId, folder, id });
const conversation = { ...row('3'), messages: [row('1'), row('2'), row('3')] };
const readOnly = (accountId) => String(accountId).startsWith('archive:');

test('a row in a list of messages takes that message alone', () => {
  assert.deepEqual(rowDeleteTargets(row('7'), { threaded: false }), [row('7')]);
});

test('a row in a list of conversations takes the whole conversation', () => {
  assert.deepEqual(rowDeleteTargets(conversation, { threaded: true }).map((r) => r.id), ['1', '2', '3']);
});

test('one message picked out of an expanded conversation takes that one', () => {
  assert.deepEqual(rowDeleteTargets(conversation, { threaded: true, only: [row('2')] }), [row('2')]);
});

test('no control on a message from an archive opened in place, or in a message read in its own window', () => {
  assert.equal(rowDeleteTargets(row('1', 'archive:big'), { readOnly }), null);
  assert.equal(rowDeleteTargets(row('1'), { alone: true }), null);
  assert.equal(rowDeleteTargets(null), null);
});
