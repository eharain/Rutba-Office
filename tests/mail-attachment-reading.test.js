// Mail keeps what it read from an attachment (an invoice's details, read
// when the message is opened) with the message, so opening it again reads
// nothing: in the message's record, untouched by the folder's index, carried
// along when the message moves to another folder, and taken away again.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { MailStore } from '@rutba/mailbox/store';

const MESSAGE = {
  messageId: '<invoice-2675@northgate.example>',
  subject: 'Invoice INV-2675',
  from: [{ address: 'accounts@northgate.example', name: 'Northgate Supplies' }],
  date: '2026-08-03T09:00:00.000Z',
  text: 'Please find the invoice attached.',
  attachments: [
    { filename: 'invoice-2675.pdf', type: 'application/pdf', bytes: Buffer.from('%PDF-1.4 a stand-in') },
    { filename: 'terms.txt', type: 'text/plain', bytes: Buffer.from('Thirty days.') },
  ],
};

const READING = {
  version: 1, languages: ['eng'], kind: 'invoice', width: 1654, height: 2339,
  note: { kind: 'invoice', from: { value: 'Northgate Supplies', box: { x: 195, y: 165, width: 380, height: 80 }, sure: true }, total: { value: 1601.12, box: { x: 1330, y: 970, width: 130, height: 30 }, sure: true }, dueDate: null, number: null, currency: 'GBP', consistent: true },
};

test('an attachment\'s reading is kept in its message\'s record, not the folder\'s index, and moves with the message', () => {
  const store = new MailStore(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-reading-')));
  const id = store.put('acct', 'Inbox', MESSAGE);
  const indexBefore = fs.readFileSync(path.join(store.root, 'acct', 'Inbox', 'index.jsonl'), 'utf8');

  assert.equal(store.setAttachmentReading('acct', 'Inbox', id, 0, READING), true);
  assert.deepEqual(store.get('acct', 'Inbox', id).attachments[0].reading, READING);
  assert.equal(store.get('acct', 'Inbox', id).attachments[1].reading, undefined, 'the other attachment has none');
  assert.equal(fs.readFileSync(path.join(store.root, 'acct', 'Inbox', 'index.jsonl'), 'utf8'), indexBefore, 'the list is not touched');
  assert.equal(store.list('acct', 'Inbox').rows[0].subject, 'Invoice INV-2675');

  // Filed into another folder: the record goes with its attachments, the reading in it.
  const moved = store.put('acct', 'Archive', store.get('acct', 'Inbox', id), { force: true });
  assert.deepEqual(store.get('acct', 'Archive', moved).attachments[0].reading, READING);
  assert.ok(fs.existsSync(store.attachmentPath('acct', 'Archive', moved, store.get('acct', 'Archive', moved).attachments[0].stored)), 'and its file');

  // Taken away; a message or an attachment that is not there answers false.
  assert.equal(store.setAttachmentReading('acct', 'Inbox', id, 0, null), true);
  assert.equal(store.get('acct', 'Inbox', id).attachments[0].reading, undefined);
  assert.equal(store.setAttachmentReading('acct', 'Inbox', id, 5, READING), false);
  assert.equal(store.setAttachmentReading('acct', 'Inbox', 'no-such-message', 0, READING), false);
});
