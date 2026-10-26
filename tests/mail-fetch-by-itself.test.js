// Mail fetches by itself (A6.2), against the test IMAP server: as the Mail
// window opens, and as soon as the server says mail came (IDLE), so an
// automatic reply goes out with nobody pressing Get mail; and one account
// is never fetched twice at once, however many ask.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';
import { imapServer } from '../packages/mailbox/test/imap-server.js';

function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: { get: (k, fallback) => (settings.has(k) ? structuredClone(settings.get(k)) : fallback), set: (k, v) => settings.set(k, structuredClone(v)), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k), keys: () => [...secrets.keys()] },
  };
}

const message = (n, from = 'Amina Yusuf <amina@work.example>') => [
  `From: ${from}`, 'To: Sam <sam@checks.example>', `Subject: Message ${n}`,
  `Date: ${new Date().toUTCString()}`, `Message-ID: <auto${n}-${Date.now()}@work.example>`, '', `Body of message ${n}.`, '',
].join('\r\n');

async function until(test, what, ms = 8000) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await test()) return true;
    await new Promise((r) => setTimeout(r, 50));
  }
  throw new Error(`waited for ${what}`);
}

test('Mail fetches by itself as it opens and as the server says mail came, and an automatic reply goes out with no Get mail', async () => {
  const server = await imapServer();
  server.deliver('INBOX', message(1));
  const sent = [];
  const service = createMailService({
    stores: fakeStores(),
    holdBlob: () => ({ url: '' }),
    broadcast: () => {},
    userData: fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mailauto-')),
    autoFetch: true,
    smtp: { async send({ draft }) { sent.push(draft); return { messageId: `<reply-${sent.length}@here>` }; } },
  });
  try {
    const { id: accountId } = service.addAccount({ account: server.account(), password: 'check' });
    service.updateAccount({ id: accountId, patch: { autoReply: { enabled: true, where: 'office', message: 'Away until Monday.', sentTo: [] } } });
    assert.deepEqual(service.fetchSettings(), { every: 5, push: true }, 'every five minutes, and as mail arrives, unless chosen otherwise');
    // The window opens.
    assert.equal(service.fetchNow().accounts, 1);
    const subjects = () => service.messages({ accountId, folder: 'INBOX', limit: 50 }).rows.map((r) => r.subject);
    await until(() => subjects().includes('Message 1'), 'the first fetch');
    await until(() => server.idling() === 1, 'a connection listening at the server');
    // Mail arrives: fetched without asking, and answered.
    sent.length = 0;
    // From someone not answered yet: one reply a sender.
    server.deliver('INBOX', message(2, 'Bilal Khan <bilal@work.example>'));
    await until(() => subjects().includes('Message 2'), 'the message that arrived');
    await until(() => sent.some((d) => /Message 2/.test(d.subject || '')), 'the automatic reply');
    // Many asking at once: one fetch at a time (beside the one listening).
    server.resetPeak();
    await Promise.all([service.sync({ accountId }), service.sync({ accountId }), service.sync({ accountId }), service.flushChanges({ accountId })]);
    assert.ok(server.peak() <= 2, `at most one fetch and the listener at once, not ${server.peak()}`);
    // Turned off: nothing listens.
    service.setFetchSettings({ every: 0, push: false });
    await until(() => server.idling() === 0, 'the listener closed');
    assert.deepEqual(service.fetchSettings(), { every: 0, push: false });
  } finally {
    await server.close();
  }
});
