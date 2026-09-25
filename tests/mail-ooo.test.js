// Out of office — automatic replies, and the RFC 3834 checks that keep one
// from mailing another one forever.
//
// `mail-ooo.js` is pure: a message, an account and a moment in, a plan (or a
// reply) out, nothing touched. Most of what the feature has to get right is
// provable there with no store and no service at all. The last two tests
// reach into `createMailService` (the same injected clock and fake `smtp`
// tests/mail-sendlater.test.js uses) to prove the setting itself is kept on
// the account, the way a signature is.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { skipReason, withinWindow, planAutoReplies, buildReply } from '../apps/desktop/main/mail-ooo.js';
import { createMailService } from '../apps/desktop/main/mail.js';

const ACCOUNT = { email: 'me@example.com', autoReply: { enabled: true, message: 'Back Monday.' } };

/** A message shaped the way the store hands one back — `headers` as the flat list `parseMessage` keeps. */
function message({ from = 'someone@example.org', subject = 'Hello', messageId = 'm1@example.org', references = [], headers = [] } = {}) {
  return {
    from: [{ name: '', address: from }],
    subject,
    messageId,
    references,
    headers: headers.map(([name, value]) => ({ name, key: name.toLowerCase(), value })),
  };
}

test('each RFC 3834 skip rule is honoured on its own', () => {
  assert.equal(skipReason(message({ from: 'me@example.com' }), ACCOUNT, {}), 'from this account itself');
  assert.equal(skipReason(message({ from: 'no-reply@service.example' }), ACCOUNT, {}), 'a no-reply address');
  assert.equal(skipReason(message({ from: 'mailer-daemon@service.example' }), ACCOUNT, {}), 'a no-reply address');
  assert.equal(skipReason(message({ headers: [['Auto-Submitted', 'auto-generated']] }), ACCOUNT, {}), 'already an automatic message');
  assert.equal(skipReason(message({ headers: [['Auto-Submitted', 'no']] }), ACCOUNT, {}), null, '"no" means an ordinary message');
  assert.equal(skipReason(message({ headers: [['Precedence', 'bulk']] }), ACCOUNT, {}), 'bulk mail');
  assert.equal(skipReason(message({ headers: [['Precedence', 'list']] }), ACCOUNT, {}), 'bulk mail');
  assert.equal(skipReason(message({ headers: [['List-Id', 'Announce <announce.example.org>']] }), ACCOUNT, {}), 'a mailing list');
  assert.equal(skipReason(message({ headers: [['List-Unsubscribe', '<mailto:off@example.org>']] }), ACCOUNT, {}), 'a mailing list');
  assert.equal(skipReason(message({ from: 'colleague@example.org' }), ACCOUNT, {}), null, 'an ordinary message is answered');
});

test('once a sender has been answered this period, a second message from them is skipped', () => {
  assert.equal(skipReason(message({ from: 'a@b.com' }), ACCOUNT, { sentTo: ['a@b.com'] }), 'already answered this period');
  assert.equal(skipReason(message({ from: 'a@b.com' }), ACCOUNT, { sentTo: ['c@d.com'] }), null);
});

test('a plan answers each sender once even across several messages in the same batch', () => {
  const plan = planAutoReplies(
    [message({ from: 'a@b.com', messageId: 'm1' }), message({ from: 'a@b.com', messageId: 'm2' }), message({ from: 'c@d.com', messageId: 'm3' })],
    ACCOUNT,
    Date.now(),
    {}
  );
  assert.deepEqual(plan.map((p) => p.to), ['a@b.com', 'c@d.com']);
});

test('the date window is honoured, including a moment before the start or after the end', () => {
  const start = Date.parse('2026-06-01T00:00:00Z');
  const end = Date.parse('2026-06-10T00:00:00Z');
  const settings = { start: '2026-06-01T00:00:00Z', end: '2026-06-10T00:00:00Z' };
  assert.equal(withinWindow(start - 1000, settings), false, 'before the start');
  assert.equal(withinWindow(start, settings), true, 'the start itself counts');
  assert.equal(withinWindow(end, settings), true, 'the end itself counts');
  assert.equal(withinWindow(end + 1000, settings), false, 'after the end');
  assert.equal(withinWindow(start + 3600_000, settings), true, 'inside the window');

  const account = { ...ACCOUNT, autoReply: { ...ACCOUNT.autoReply, start: settings.start, end: settings.end } };
  assert.deepEqual(planAutoReplies([message()], account, start - 1000, {}), []);
  assert.equal(planAutoReplies([message()], account, start + 1000, {}).length, 1);
  assert.deepEqual(planAutoReplies([message()], account, end + 1000, {}), []);
});

test('"only reply to people in my contacts" leaves everyone else unanswered', () => {
  const account = { ...ACCOUNT, autoReply: { ...ACCOUNT.autoReply, contactsOnly: true } };
  const isContact = (address) => address === 'friend@example.org';
  const plan = planAutoReplies([message({ from: 'friend@example.org' }), message({ from: 'stranger@example.org', messageId: 'm2' })], account, Date.now(), { isContact });
  assert.deepEqual(plan.map((p) => p.to), ['friend@example.org']);

  // The option is off by default: nobody is excluded just for being unknown.
  const open = planAutoReplies([message({ from: 'stranger@example.org' })], ACCOUNT, Date.now(), { isContact });
  assert.equal(open.length, 1);
});

test('the reply carries Auto-Submitted, threads under the original, and names it in the subject', () => {
  const original = message({ from: 'a@b.com', subject: 'Quarterly numbers', messageId: 'orig@b.com', references: ['earlier@b.com'] });
  const reply = buildReply(original, ACCOUNT, ACCOUNT.autoReply);
  assert.equal(reply.to, 'a@b.com');
  assert.equal(reply.headers['Auto-Submitted'], 'auto-replied');
  assert.equal(reply.inReplyTo, 'orig@b.com');
  assert.deepEqual(reply.references, ['earlier@b.com', 'orig@b.com']);
  assert.equal(reply.subject, 'Automatic reply: Quarterly numbers');
  assert.equal(reply.text, 'Back Monday.');

  // An account's own subject template still gets the original subject filled in.
  const custom = buildReply(original, ACCOUNT, { ...ACCOUNT.autoReply, subject: 'Away: <original subject>' });
  assert.equal(custom.subject, 'Away: Quarterly numbers');
});

test('turning automatic replies on again starts a fresh "answered" set', () => {
  const alreadyAnswered = { ...ACCOUNT, autoReply: { ...ACCOUNT.autoReply, sentTo: ['a@b.com'] } };
  assert.deepEqual(planAutoReplies([message({ from: 'a@b.com' })], alreadyAnswered, Date.now(), { sentTo: alreadyAnswered.autoReply.sentTo }), []);

  // Off, then on again: the caller resets `sentTo` to `[]` on that
  // transition (mail.js does this in updateAccount, not here) — once it
  // has, the same sender is owed a reply again.
  const turnedOnAgain = { ...ACCOUNT, autoReply: { ...ACCOUNT.autoReply, sentTo: [] } };
  assert.equal(planAutoReplies([message({ from: 'a@b.com' })], turnedOnAgain, Date.now(), { sentTo: [] }).length, 1);
});

/* ── through the service, so the setting is proved to live on the account ── */

function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: {
      get: (k, fallback) => (settings.has(k) ? settings.get(k) : fallback),
      set: (k, v) => settings.set(k, v),
      delete: (k) => settings.delete(k),
      all: () => Object.fromEntries(settings),
    },
    secrets: {
      available: () => true,
      get: (k) => secrets.get(k) ?? null,
      set: (k, v) => secrets.set(k, v),
      delete: (k) => secrets.delete(k),
      keys: () => [...secrets.keys()],
    },
  };
}

test("an account's automatic-reply settings persist the way its signature does", () => {
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-ooo-'));
  const stores = fakeStores();
  const service = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, smtp: { send: async () => ({ messageId: 'x' }) } });
  const account = service.addAccount({ account: { email: 'me@example.com', name: 'Me', imap: { host: 'h', port: 993 }, smtp: { host: 'h', port: 465 } }, password: 'secret' });

  const settings = { enabled: true, start: null, end: null, subject: '', message: 'Away this week.', contactsOnly: false, sentTo: [] };
  service.updateAccount({ id: account.id, patch: { autoReply: settings } });
  assert.deepEqual(service.accounts().find((a) => a.id === account.id).autoReply, settings);

  // A fresh service over the same settings — the restart every other Mail
  // setting has to survive — reads it back unchanged.
  const restarted = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, smtp: { send: async () => ({ messageId: 'x' }) } });
  assert.deepEqual(restarted.accounts().find((a) => a.id === account.id).autoReply, settings);
});
