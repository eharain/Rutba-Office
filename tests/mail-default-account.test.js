// The sending account (D-2026-11-11-02): "Send new mail from" in Mail,
// Accounts, and what Calendar and Talk call to send from it,
// mail.defaultAccount() and mail.sendFrom().
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';
import { parseMessage } from '../packages/mailbox/src/mime.js';

function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: {
      get: (k, fallback) => (settings.has(k) && settings.get(k) !== undefined ? structuredClone(settings.get(k)) : fallback),
      set: (k, v) => settings.set(k, structuredClone(v)),
      delete: (k) => settings.delete(k),
      all: () => Object.fromEntries(settings),
    },
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k), keys: () => [...secrets.keys()] },
  };
}

const dirs = [];
test.after(() => {
  for (const d of dirs) fs.rmSync(d, { recursive: true, force: true });
});

/** A service whose SMTP is a list: what it was handed, made into the bytes the server would have taken. */
function fixture() {
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-default-'));
  dirs.push(userData);
  const stores = fakeStores();
  const sent = [];
  const smtp = {
    async send({ account, draft }) {
      sent.push({ account: account.email, draft });
      return { messageId: `m-${sent.length}@example.com`, rejected: [] };
    },
  };
  // Local accounts only (no IMAP server to keep in step with), so nothing here opens a connection.
  const service = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, smtp, autoFetch: false });
  const add = (email, extra = {}) => service.addAccount({ account: { email, name: extra.name || email, imap: { host: '', port: 993 }, smtp: { host: 'smtp.example.com', port: 465, secure: true }, ...extra }, password: 'x' });
  return { service, stores, sent, add };
}

const INVITE = [
  'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Rutba//Calendar//EN', 'METHOD:REQUEST',
  'BEGIN:VEVENT', 'UID:planning-1@rutba', 'DTSTAMP:20261111T090000Z', 'DTSTART:20261112T100000Z', 'DTEND:20261112T110000Z',
  'SUMMARY:Planning', 'ORGANIZER:mailto:me@example.com', 'ATTENDEE;RSVP=TRUE:mailto:ann@example.org', 'END:VEVENT', 'END:VCALENDAR', '',
].join('\r\n');

test('no account: no default, and sendFrom says there is nothing to send from', async () => {
  const f = fixture();
  assert.equal(f.service.defaultAccount(), null);
  await assert.rejects(() => f.service.sendFrom({ to: 'ann@example.org', subject: 'Hi', text: 'Hi' }), (err) => err.code === 'no-account');
  assert.equal(f.sent.length, 0);
});

test('one account is the default without being chosen', () => {
  const f = fixture();
  const a = f.add('me@example.com', { name: 'Me' });
  assert.deepEqual(f.service.defaultAccount(), { id: a.id, address: 'me@example.com', name: 'Me' });
  assert.equal(f.service.accounts()[0].isDefault, true);
});

test('a mailbox brought in from a file is never the default, and cannot be chosen', () => {
  const f = fixture();
  const local = f.add('old@example.com', { local: true });
  assert.equal(f.service.defaultAccount(), null);
  assert.throws(() => f.service.setDefaultAccount({ id: local.id }), /cannot send/);
  const real = f.add('me@example.com');
  assert.equal(f.service.defaultAccount().id, real.id);
  assert.deepEqual(f.service.accounts().map((a) => [a.email, a.isDefault, a.canSend]), [['old@example.com', false, false], ['me@example.com', true, true]]);
});

test('Send new mail from: the chosen account is kept, and the first that can send takes over when it is removed', () => {
  const f = fixture();
  const first = f.add('first@example.com');
  const second = f.add('second@example.com');
  assert.equal(f.service.defaultAccount().id, first.id);
  assert.deepEqual(f.service.setDefaultAccount({ id: second.id }), { id: second.id, address: 'second@example.com', name: 'second@example.com' });
  assert.equal(f.service.defaultAccount().id, second.id);
  assert.equal(f.stores.settings.get('mail.defaultAccount'), second.id);
  assert.deepEqual(f.service.accounts().filter((a) => a.isDefault).map((a) => a.id), [second.id]);
  f.service.removeAccount({ id: second.id });
  assert.equal(f.service.defaultAccount().id, first.id);
  assert.throws(() => f.service.setDefaultAccount({ id: 'nobody' }), /not set up/);
});

test('sendFrom goes from the default account, and the message is kept in Sent as it went', async () => {
  const f = fixture();
  f.add('first@example.com');
  const me = f.add('me@example.com', { name: 'Me Myself' });
  f.service.setDefaultAccount({ id: me.id });
  const done = await f.service.sendFrom({
    to: [{ name: 'Ann Lee', address: 'ann@example.org' }, 'bob@example.org'],
    cc: 'Cy <cy@example.org>',
    subject: 'Lunch',
    text: 'Lunch on Thursday?',
    html: '<p>Lunch on Thursday?</p>',
    headers: { 'X-Rutba-Talk': 'invite' },
  });
  assert.equal(done.accountId, me.id);
  assert.equal(done.from, 'me@example.com');
  assert.deepEqual(done.refused, []);
  assert.equal(f.sent.length, 1);
  assert.equal(f.sent[0].account, 'me@example.com');
  assert.equal(f.sent[0].draft.to, '"Ann Lee" <ann@example.org>, bob@example.org');
  assert.deepEqual(f.sent[0].draft.envelope, { from: 'me@example.com', to: ['ann@example.org', 'bob@example.org', 'cy@example.org'] });
  const went = parseMessage(String(f.sent[0].draft.raw));
  assert.equal(went.subject, 'Lunch');
  assert.match(String(f.sent[0].draft.raw), /^X-Rutba-Talk: invite$/m);

  const sent = f.service.folders({ accountId: me.id }).find((x) => x.role === 'sent');
  assert.ok(sent, 'a Sent folder');
  const rows = f.service.messages({ accountId: me.id, folder: sent.path }).rows;
  assert.equal(rows.length, 1);
  assert.equal(rows[0].subject, 'Lunch');
  assert.equal(rows[0].unread, false);
});

test('sendFrom names its account by id or by address', async () => {
  const f = fixture();
  f.add('first@example.com');
  const other = f.add('Other@Example.com');
  assert.equal((await f.service.sendFrom({ account: 'other@example.com', to: 'ann@example.org', subject: 'a' })).accountId, other.id);
  assert.equal((await f.service.sendFrom({ account: other.id, to: 'ann@example.org', subject: 'b' })).accountId, other.id);
  await assert.rejects(() => f.service.sendFrom({ account: 'stranger@example.com', to: 'ann@example.org' }), (err) => err.code === 'no-account');
  const local = f.add('old@example.com', { local: true });
  await assert.rejects(() => f.service.sendFrom({ account: local.id, to: 'ann@example.org' }), (err) => err.code === 'cannot-send');
});

test('a text/calendar; method=REQUEST attachment goes as the invitation itself, and the Sent copy carries it', async () => {
  const f = fixture();
  const me = f.add('me@example.com');
  await f.service.sendFrom({
    to: 'ann@example.org',
    subject: 'Invitation: Planning',
    text: 'Planning, Thursday 12 November, 10:00.',
    attachments: [
      { filename: 'invite.ics', content: INVITE, contentType: 'text/calendar; method=REQUEST' },
      { filename: 'agenda.txt', content: new Uint8Array(Buffer.from('1. Budget')), contentType: 'text/plain' },
    ],
  });
  const raw = String(f.sent[0].draft.raw);
  // The calendar part beside the words, with its method, as iMIP has it (RFC 6047).
  assert.match(raw, /Content-Type: text\/calendar; charset=utf-8; method=REQUEST/i);
  assert.match(raw, /Content-Type: text\/plain; name=agenda\.txt/i);
  const parsed = parseMessage(raw);
  assert.ok(parsed.attachments.some((a) => /agenda\.txt/.test(a.filename || '')), 'the other attachment kept its name');

  const sent = f.service.folders({ accountId: me.id }).find((x) => x.role === 'sent');
  const row = f.service.messages({ accountId: me.id, folder: sent.path }).rows[0];
  const kept = f.service.message({ accountId: me.id, folder: sent.path, id: row.id });
  assert.ok((kept.attachments || []).some((a) => /calendar|ics/i.test(`${a.type} ${a.filename}`)), 'the invitation is in the Sent copy');
});

test('a calendar part with no method in its type takes its METHOD line', async () => {
  const f = fixture();
  f.add('me@example.com');
  await f.service.sendFrom({ to: 'ann@example.org', subject: 'Cancelled', attachments: [{ filename: 'cancel.ics', content: INVITE.replace('METHOD:REQUEST', 'METHOD:CANCEL'), contentType: 'text/calendar' }] });
  assert.match(String(f.sent[0].draft.raw), /method=CANCEL/i);
});

test('sendFrom refuses what it cannot send: no one to send to, an attachment by path, one with nothing in it', async () => {
  const f = fixture();
  f.add('me@example.com');
  await assert.rejects(() => f.service.sendFrom({ subject: 'To nobody' }), /someone to go to/);
  await assert.rejects(() => f.service.sendFrom({ to: 'ann@example.org', attachments: [{ filename: 'secrets.txt', path: 'C:/Windows/win.ini' }] }), /from its content/);
  await assert.rejects(() => f.service.sendFrom({ to: 'ann@example.org', attachments: [{ filename: 'empty.txt' }] }), /no content/);
  assert.equal(f.sent.length, 0);
});

test('a header cannot carry a line break into the message', async () => {
  const f = fixture();
  f.add('me@example.com');
  await f.service.sendFrom({ to: 'ann@example.org', subject: 'x', headers: { 'X-Note': 'one\r\nBcc: eve@example.org' } });
  const raw = String(f.sent[0].draft.raw);
  assert.doesNotMatch(raw, /^Bcc:/m);
  assert.deepEqual(f.sent[0].draft.envelope.to, ['ann@example.org']);
});

test('a failed send throws, and nothing is kept in Sent', async () => {
  const f = fixture();
  const me = f.add('me@example.com');
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-default-'));
  dirs.push(userData);
  const failing = createMailService({ stores: f.stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, smtp: { async send() { throw new Error('the server said no'); } }, autoFetch: false });
  await assert.rejects(() => failing.sendFrom({ to: 'ann@example.org', subject: 'x' }), /said no/);
  assert.equal(failing.folders({ accountId: me.id }).filter((x) => x.role === 'sent').length, 0);
});
