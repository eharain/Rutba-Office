// mail.onArrived (Mail 2.25): mail arrived, told to the other packs in the
// main process, so Calendar hears of invitations as they come instead of
// reading the Inboxes every 90 seconds. Against a fake IMAP server.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';
import { imapServer } from '../packages/mailbox/test/imap-server.js';
import { METHODS } from '@rutba/office-shell/contract';

const dirs = [];
test.after(() => {
  for (const d of dirs) fs.rmSync(d, { recursive: true, force: true });
});

const ics = ['BEGIN:VCALENDAR', 'METHOD:REQUEST', 'VERSION:2.0', 'BEGIN:VEVENT', 'UID:arrived-1@work.example', 'SUMMARY:Planning', 'DTSTART:20261120T100000Z', 'DTEND:20261120T110000Z', 'ORGANIZER:mailto:amina@work.example', 'END:VEVENT', 'END:VCALENDAR', ''].join('\r\n');
const plain = (n) => ['From: Amina Yusuf <amina@work.example>', 'To: Sam <sam@checks.example>', `Subject: Plain ${n}`, `Date: ${new Date(Date.UTC(2026, 10, 10, 9, n)).toUTCString()}`, `Message-ID: <plain${n}@work.example>`, '', `Body ${n}.`, ''].join('\r\n');
const outlookInvite = [
  'From: Amina Yusuf <amina@work.example>', 'To: Sam <sam@checks.example>', 'Subject: Planning', `Date: ${new Date(Date.UTC(2026, 10, 10, 10, 0)).toUTCString()}`,
  'Message-ID: <invite1@work.example>', 'MIME-Version: 1.0', 'Content-Type: multipart/alternative; boundary="a"', '',
  '--a', 'Content-Type: text/plain', '', 'Planning?', '', '--a', 'Content-Type: text/calendar; method=REQUEST; charset=utf-8', '', ics, '--a--', '',
].join('\r\n');

async function setup() {
  const server = await imapServer();
  const settings = new Map();
  const stores = {
    settings: { get: (k, f) => (settings.has(k) ? structuredClone(settings.get(k)) : f), set: (k, v) => settings.set(k, structuredClone(v)), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: () => 'check', set: () => {}, delete: () => {}, keys: () => [] },
  };
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-arrived-'));
  dirs.push(userData);
  const service = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, smtp: { send: async () => ({ messageId: 'x' }) }, autoFetch: false });
  const account = service.addAccount({ account: server.account(), password: 'check' });
  return { server, service, account };
}

test('each arrival is told once, with the ids, and each invitation with its calendar text', async () => {
  const s = await setup();
  try {
    await s.service.sync({ accountId: s.account.id });
    const heard = [];
    const stop = s.service.onArrived((e) => heard.push(e));
    // A listener that throws is passed over; the others still hear.
    s.service.onArrived(() => { throw new Error('a listener with a fault'); });
    s.server.deliver('INBOX', plain(1));
    s.server.deliver('INBOX', outlookInvite);
    await s.service.sync({ accountId: s.account.id });
    assert.equal(heard.length, 1, 'one arrival, told once');
    const e = heard[0];
    assert.equal(e.accountId, s.account.id);
    assert.equal(e.folder, 'INBOX');
    assert.equal(e.count, 2);
    assert.equal(e.ids.length, 2);
    assert.ok(!Number.isNaN(Date.parse(e.at)));
    assert.equal(e.filedByRules, false);
    assert.equal(e.invitations.length, 1, 'the Outlook invitation, its calendar an alternative with no file name');
    const inv = e.invitations[0];
    assert.equal(inv.subject, 'Planning');
    assert.equal(inv.messageId, 'invite1@work.example');
    assert.equal(inv.method, 'REQUEST');
    assert.deepEqual(inv.from, { name: 'Amina Yusuf', address: 'amina@work.example' });
    assert.match(inv.text, /UID:arrived-1@work\.example/);
    assert.ok(e.ids.includes(inv.id));
    // The message is where the event says, for a listener that wants the rest of it.
    assert.equal(s.service.message({ accountId: e.accountId, folder: e.folder, id: inv.id })?.subject, 'Planning');

    stop();
    s.server.deliver('INBOX', plain(2));
    await s.service.sync({ accountId: s.account.id });
    assert.equal(heard.length, 1, 'nothing told once stopped');
    // A sync with nothing new tells nothing.
  } finally {
    await s.server.close();
  }
});

test('onArrived is the main process\'s alone: no window has a channel for it', () => {
  assert.ok(!METHODS.mail.includes('onArrived'));
  const s = { service: createMailService({ stores: { settings: { get: (_k, f) => f, set: () => {}, delete: () => {}, all: () => ({}) }, secrets: { available: () => true, get: () => null, set: () => {}, delete: () => {}, keys: () => [] } }, holdBlob: () => ({}), broadcast: () => {}, userData: (() => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-arrived-')); dirs.push(d); return d; })(), autoFetch: false }) };
  assert.throws(() => s.service.onArrived('not a function'), /takes a function/);
});
