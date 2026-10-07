// A message moved to another folder keeps its attachments.
//
// The store keeps an attachment's bytes in a file beside the message, and a
// move was a put of the record into the new folder and a remove from the
// old: the record kept the attachment's name, the remove took its file, and
// the message arrived in Archive, Junk or a rule's folder with an
// attachment that could no longer be opened.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { MailStore } from '@rutba/mailbox/store';
import { applyPlan } from '../apps/desktop/main/mail-rules.js';
import { createMailService } from '../apps/desktop/main/mail.js';

const PDF = Buffer.from('%PDF-1.4 the quarterly report');
const withFile = (messageId) => ({
  messageId, subject: 'The report', from: [{ address: 'ana@work.example' }], text: 'Attached.',
  attachments: [{ filename: 'report.pdf', type: 'application/pdf', bytes: PDF }],
});

/** The bytes of a message's first attachment in a folder, or null. */
function attachmentBytes(store, accountId, folder, id) {
  const a = store.get(accountId, folder, id)?.attachments?.[0];
  if (!a?.stored) return null;
  const file = store.attachmentPath(accountId, folder, id, a.stored);
  return fs.existsSync(file) ? fs.readFileSync(file) : null;
}

test('a message put into another folder and taken out of the first keeps its attachment', () => {
  const store = new MailStore(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-move-')));
  store.putMany('acc', 'INBOX', [withFile('m1@x')]);
  const id = store.list('acc', 'INBOX', { limit: 5 }).rows[0].id;
  store.put('acc', 'Archive', store.get('acc', 'INBOX', id), { force: true });
  store.remove('acc', 'INBOX', id);
  assert.deepEqual(attachmentBytes(store, 'acc', 'Archive', id), PDF);
  // Put again where it already is — a note added to it — it stays.
  store.put('acc', 'Archive', { ...store.get('acc', 'Archive', id), note: 1 }, { force: true });
  assert.deepEqual(attachmentBytes(store, 'acc', 'Archive', id), PDF);
});

test('a rule that moves or copies a message carries its attachment', () => {
  const store = new MailStore(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-move-')));
  store.putMany('acc', 'INBOX', [withFile('m2@x'), withFile('m3@x')]);
  const [a, b] = store.list('acc', 'INBOX', { limit: 5 }).rows;
  applyPlan([
    { row: { ...a, folder: 'INBOX' }, actions: [{ type: 'move', value: 'Reports' }] },
    { row: { ...b, folder: 'INBOX' }, actions: [{ type: 'copy', value: 'Kept' }] },
  ], { store, accountId: 'acc', folderFor: () => null });
  assert.deepEqual(attachmentBytes(store, 'acc', 'Reports', a.id), PDF, 'moved');
  assert.deepEqual(attachmentBytes(store, 'acc', 'Kept', b.id), PDF, 'copied');
  assert.deepEqual(attachmentBytes(store, 'acc', 'INBOX', b.id), PDF, 'and the original keeps its own');
});

test('Move to through the service keeps the attachment openable', () => {
  const settings = new Map();
  const stores = {
    settings: { get: (k, f) => (settings.has(k) ? settings.get(k) : f), set: (k, v) => settings.set(k, v), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: () => null, set: () => {}, delete: () => {}, keys: () => [] },
  };
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-move-'));
  const service = createMailService({ stores, holdBlob: (bytes) => ({ url: 'blob:x', bytes }), broadcast: () => {}, userData, smtp: { send: async () => ({ messageId: 'x' }) } });
  const account = service.addAccount({ account: { email: 'me@example.com', name: 'Me', imap: { host: 'h', port: 993 }, smtp: { host: 'h', port: 465 } }, password: 'secret' });
  const store = new MailStore(path.join(userData, 'mail'));
  store.putMany(account.id, 'INBOX', [withFile('m4@x')]);
  const id = store.list(account.id, 'INBOX', { limit: 5 }).rows[0].id;
  service.move({ accountId: account.id, folder: 'INBOX', ids: [id], to: 'Archive' });
  const opened = service.attachment({ accountId: account.id, folder: 'Archive', id, index: 0 });
  assert.ok(opened, 'the attachment is there to open');
  assert.deepEqual(Buffer.from(opened.bytes), PDF);
});
