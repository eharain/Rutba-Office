// Out of office on hosted mail (A6.6). Gmail, Outlook, Yahoo and iCloud keep
// their automatic replies in their own settings and offer no ManageSieve, so
// asking them waited ten seconds and then said the server "could not be
// reached". Now Mail says at once that the server cannot send the replies,
// and offers Rutba Office's own, without opening a connection.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';

function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: { get: (k, fallback) => (settings.has(k) ? structuredClone(settings.get(k)) : fallback), set: (k, v) => settings.set(k, structuredClone(v)), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k), keys: () => [...secrets.keys()] },
  };
}

const hosted = [
  { email: 'sam@gmail.com', host: 'imap.gmail.com', provider: 'google', label: 'Gmail' },
  { email: 'sam@outlook.com', host: 'outlook.office365.com', provider: 'microsoft', label: 'Outlook' },
  { email: 'sam@hotmail.com', host: 'imap-mail.outlook.com', provider: null, label: 'Outlook' },
  { email: 'sam@yahoo.com', host: 'imap.mail.yahoo.com', provider: 'yahoo', label: 'Yahoo' },
  { email: 'sam@icloud.com', host: 'imap.mail.me.com', provider: null, label: 'iCloud' },
];

test('hosted mail with no ManageSieve says at once that the server cannot send replies, and offers Rutba Office\'s own', async () => {
  const service = createMailService({
    stores: fakeStores(),
    holdBlob: () => ({ url: '' }),
    broadcast: () => {},
    userData: fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-ooo-hosted-')),
    smtp: { async send() { return { messageId: '<x@here>' }; } },
  });
  for (const h of hosted) {
    const { id } = service.addAccount({
      account: { email: h.email, name: h.email, provider: h.provider, imap: { host: h.host, port: 993, secure: true, user: h.email }, smtp: { host: h.host, port: 465, secure: true } },
      password: 'not-used',
    });
    const began = Date.now();
    const r = await service.serverAutoReply({ accountId: id, action: 'state' });
    assert.deepEqual(r, { supported: false, reason: 'provider', provider: h.label }, h.email);
    assert.ok(Date.now() - began < 500, `${h.email} answered at once, not after a wait`);
    // Asked to turn the server's on all the same: the same answer, nothing sent.
    assert.equal((await service.serverAutoReply({ accountId: id, action: 'on', settings: { message: 'Away' } })).reason, 'provider');
  }
});
