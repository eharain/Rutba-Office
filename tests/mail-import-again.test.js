// A6.17: a file imported a second time goes into the mailbox made from it the
// first time, adding only the messages it did not hold, never a second
// mailbox with every message twice.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';

function service() {
  const settings = new Map();
  const stores = {
    settings: { get: (k, f) => (settings.has(k) ? settings.get(k) : f), set: (k, v) => settings.set(k, v), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: () => null, set: () => {}, delete: () => {}, keys: () => [] },
  };
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-import-again-'));
  return { mail: createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, smtp: { send: async () => ({ messageId: 'x' }) } }), userData };
}

const one = (n) => [
  `From MAILER-DAEMON Thu Oct  1 10:0${n}:00 2026`,
  `From: Sam <sam${n}@work.example>`,
  'To: <me@example.com>',
  `Subject: Imported ${n}`,
  `Date: ${new Date(Date.UTC(2026, 9, 1, 10, n)).toUTCString()}`,
  `Message-ID: <imported-${n}@test>`,
  'Content-Type: text/plain; charset=utf-8',
  '',
  `Message number ${n}.`,
  '',
].join('\n');

test('importing the same file again adds to the mailbox made from it, not a second one', () => {
  const { mail, userData } = service();
  const file = path.join(userData, 'Old mail.mbox');
  fs.writeFileSync(file, [one(1), one(2)].join('\n'));

  const first = mail.import({ path: file });
  assert.equal(first.messages, 2);
  assert.equal(first.again, false);
  assert.equal(mail.accounts().filter((a) => a.local).length, 1);

  // The same file named another way (a different case on Windows, a relative step elsewhere).
  const sameFile = process.platform === 'win32' ? file.toUpperCase() : path.join(path.dirname(file), '.', path.basename(file));
  const second = mail.import({ path: sameFile });
  assert.equal(second.accountId, first.accountId, 'the same mailbox');
  assert.equal(second.again, true);
  assert.equal(second.messages, 0, 'nothing held twice');
  assert.equal(mail.accounts().filter((a) => a.local).length, 1, 'still one mailbox');

  // The file grew: only the new message comes in, and the folders it had are kept.
  const foldersBefore = mail.folders({ accountId: first.accountId }).map((f) => f.path);
  fs.appendFileSync(file, '\n' + one(3));
  const third = mail.import({ path: file });
  assert.equal(third.messages, 1);
  const rows = mail.messages({ accountId: first.accountId, folder: foldersBefore[0], limit: 50 }).rows;
  assert.equal(rows.length, 3);
  for (const f of foldersBefore) assert.ok(mail.folders({ accountId: first.accountId }).some((g) => g.path === f), f);

  // A different file is still a mailbox of its own.
  const other = path.join(userData, 'Other.mbox');
  fs.writeFileSync(other, one(4));
  const fourth = mail.import({ path: other });
  assert.notEqual(fourth.accountId, first.accountId);
  assert.equal(fourth.again, false);
});
