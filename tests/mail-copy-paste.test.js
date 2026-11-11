// Copy, cut and paste of messages between folders (Mail 2.25): within one
// account the server's own COPY or MOVE, between accounts an APPEND to the
// target, a cut then taking the source away; local folders the same way.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { MailStore } from '@rutba/mailbox/store';
import { createMailService } from '../apps/desktop/main/mail.js';
import { imapServer } from '../packages/mailbox/test/imap-server.js';

const PDF = Buffer.from('%PDF-1.4 the quarterly report');
const message = (messageId, extra = {}) => ({
  messageId, subject: `About ${messageId}`, from: [{ address: 'ana@work.example', name: 'Ana' }], to: [{ address: 'me@example.com' }],
  date: '2026-11-02T09:30:00.000Z', text: 'Attached.', unread: true,
  attachments: [{ filename: 'report.pdf', type: 'application/pdf', bytes: PDF }], ...extra,
});

const dirs = [];
test.after(() => {
  for (const d of dirs) fs.rmSync(d, { recursive: true, force: true });
});

function fixture() {
  const settings = new Map();
  const stores = {
    settings: { get: (k, f) => (settings.has(k) ? structuredClone(settings.get(k)) : f), set: (k, v) => settings.set(k, structuredClone(v)), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: () => 'x', set: () => {}, delete: () => {}, keys: () => [] },
  };
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-paste-'));
  dirs.push(userData);
  const service = createMailService({ stores, holdBlob: (bytes) => ({ url: 'blob:x', bytes }), broadcast: () => {}, userData, smtp: { send: async () => ({ messageId: 'x' }) }, autoFetch: false });
  const add = (email, { live = false } = {}) => service.addAccount({
    account: { email, name: email, local: !live, imap: { host: live ? 'imap.invalid' : '', port: 993, secure: true }, smtp: { host: live ? 'smtp.invalid' : '', port: 465 } },
    password: 'x',
  });
  const store = new MailStore(path.join(userData, 'mail'));
  const put = (accountId, folder, ...messages) => {
    store.putMany(accountId, folder, messages);
    store.upsertFolder(accountId, { path: folder, name: folder });
    return store.list(accountId, folder, { limit: 50 }).rows.map((r) => r.id);
  };
  const ids = (accountId, folder) => service.messages({ accountId, folder }).rows.map((r) => r.subject).sort();
  const bytesOf = (accountId, folder, id) => {
    const a = store.get(accountId, folder, id)?.attachments?.[0];
    return a?.stored ? store.attachmentBytes(accountId, folder, id, a.stored) : null;
  };
  return { service, store, add, put, ids, bytesOf };
}

test('copy within a local account: both folders have it, the attachment with it', async () => {
  const f = fixture();
  const a = f.add('old@example.com');
  const [id] = f.put(a.id, 'Inbox', message('one@x'));
  const r = await f.service.paste({ accountId: a.id, folder: 'Inbox', ids: [id], to: { accountId: a.id, folder: 'Kept' } });
  assert.deepEqual(r, { copied: 1, moved: 0 });
  assert.deepEqual(f.ids(a.id, 'Inbox'), ['About one@x']);
  assert.deepEqual(f.ids(a.id, 'Kept'), ['About one@x']);
  const kept = f.service.messages({ accountId: a.id, folder: 'Kept' }).rows[0];
  assert.deepEqual(Buffer.from(f.bytesOf(a.id, 'Kept', kept.id)), PDF);
  assert.ok(f.service.folders({ accountId: a.id }).some((x) => x.path === 'Kept'), 'the folder pasted into is listed');
});

test('cut within an account moves, and pasting where they are changes nothing', async () => {
  const f = fixture();
  const a = f.add('old@example.com');
  const [id] = f.put(a.id, 'Inbox', message('two@x'));
  assert.deepEqual(await f.service.paste({ accountId: a.id, folder: 'Inbox', ids: [id], to: { accountId: a.id, folder: 'Inbox' }, cut: true }), { copied: 0, moved: 0 });
  assert.deepEqual(await f.service.paste({ accountId: a.id, folder: 'Inbox', ids: [id], to: { accountId: a.id, folder: 'Done' }, cut: true }), { copied: 0, moved: 1 });
  assert.deepEqual(f.ids(a.id, 'Inbox'), []);
  assert.deepEqual(f.ids(a.id, 'Done'), ['About two@x']);
});

test('within a live account a copy is the server\'s COPY and a cut its MOVE', async () => {
  const f = fixture();
  const a = f.add('me@example.com', { live: true });
  const got = f.put(a.id, 'INBOX', message('three@x'), message('four@x'));
  await f.service.paste({ accountId: a.id, folder: 'INBOX', ids: [got[0]], to: { accountId: a.id, folder: 'Projects' } });
  await f.service.paste({ accountId: a.id, folder: 'INBOX', ids: [got[1]], to: { accountId: a.id, folder: 'Projects' }, cut: true });
  const kinds = f.store.changes(a.id).items.map((c) => `${c.kind} ${c.folder} -> ${c.to}`);
  assert.ok(kinds.includes('copy INBOX -> Projects'), kinds.join(' | '));
  assert.ok(kinds.includes('move INBOX -> Projects'), kinds.join(' | '));
  assert.ok(!f.store.changes(a.id).items.some((c) => c.kind === 'append'), 'nothing appended within one account');
});

test('between accounts: an APPEND to the target with its date and read state, the attachment carried, and a cut takes the source away', async () => {
  const f = fixture();
  const from = f.add('old@example.com');
  const to = f.add('me@example.com', { live: true });
  const [unread, read] = f.put(from.id, 'Inbox', message('five@x'), message('six@x', { unread: false, flagged: true }));
  const r = await f.service.paste({ accountId: from.id, folder: 'Inbox', ids: [unread, read], to: { accountId: to.id, folder: 'Imported' }, cut: true });
  assert.deepEqual(r, { copied: 2, moved: 2 });
  assert.deepEqual(f.ids(from.id, 'Inbox'), [], 'the cut took them from the source');
  const rows = f.service.messages({ accountId: to.id, folder: 'Imported' }).rows;
  assert.deepEqual(rows.map((x) => x.subject).sort(), ['About five@x', 'About six@x']);
  for (const row of rows) assert.deepEqual(Buffer.from(f.bytesOf(to.id, 'Imported', row.id)), PDF, 'the attachment came across');
  const appends = f.store.changes(to.id).items.filter((c) => c.kind === 'append' && c.folder === 'Imported');
  assert.equal(appends.length, 2);
  const byId = Object.fromEntries(appends.map((c) => [c.messageId, c]));
  assert.deepEqual(byId['five@x'].flags, []);
  assert.deepEqual(byId['six@x'].flags.sort(), ['\\Flagged', '\\Seen']);
  assert.equal(new Date(byId['five@x'].date).toISOString(), '2026-11-02T09:30:00.000Z');
  for (const c of appends) assert.ok(fs.existsSync(c.file), 'the message as the server is to take it, kept to send');
  assert.match(fs.readFileSync(appends[0].file, 'latin1'), /^Subject: About /m);
});

test('between accounts a copy leaves the source as it was', async () => {
  const f = fixture();
  const one = f.add('one@example.com');
  const two = f.add('two@example.com');
  const [id] = f.put(one.id, 'Inbox', message('seven@x'));
  assert.deepEqual(await f.service.paste({ accountId: one.id, folder: 'Inbox', ids: [id], to: { accountId: two.id, folder: 'Inbox' } }), { copied: 1, moved: 0 });
  assert.deepEqual(f.ids(one.id, 'Inbox'), ['About seven@x']);
  assert.deepEqual(f.ids(two.id, 'Inbox'), ['About seven@x']);
});

test('a message no longer there is passed over, and an archive opened in place is never cut from', async () => {
  const f = fixture();
  const a = f.add('old@example.com');
  assert.deepEqual(await f.service.paste({ accountId: a.id, folder: 'Inbox', ids: ['gone'], to: { accountId: a.id, folder: 'Kept' } }), { copied: 0, moved: 0 });
  await assert.rejects(() => f.service.paste({ accountId: 'archive:x', folder: 'Inbox', ids: ['m'], to: { accountId: a.id, folder: 'Kept' }, cut: true }), /read only/);
  await assert.rejects(() => f.service.paste({ accountId: a.id, folder: 'Inbox', ids: ['m'], to: { accountId: 'archive:x', folder: 'Kept' } }), /read only/);
  await assert.rejects(() => f.service.paste({ accountId: a.id, folder: 'Inbox', ids: ['m'], to: { accountId: 'nobody', folder: 'Kept' } }), /not set up/);
});

/* ── at the servers: two IMAP servers on this computer ─────────────────── */

const raw = (n) => [
  'From: Amina Yusuf <amina@work.example>', 'To: Sam <sam@checks.example>', `Subject: Server message ${n}`,
  `Date: ${new Date(Date.UTC(2026, 9, 1, 9, n)).toUTCString()}`, `Message-ID: <s${n}@work.example>`, '', `Body of message ${n}.`, '',
].join('\r\n');

async function twoServers() {
  const one = await imapServer();
  const two = await imapServer();
  for (let n = 1; n <= 3; n++) one.deliver('INBOX', raw(n));
  const settings = new Map();
  const stores = {
    settings: { get: (k, f) => (settings.has(k) ? structuredClone(settings.get(k)) : f), set: (k, v) => settings.set(k, structuredClone(v)), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: () => 'check', set: () => {}, delete: () => {}, keys: () => [] },
  };
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-paste-'));
  dirs.push(userData);
  const service = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, smtp: { send: async () => ({ messageId: 'x' }) }, autoFetch: false });
  const a = service.addAccount({ account: one.account(), password: 'check' });
  const b = service.addAccount({ account: two.account(), password: 'check' });
  await service.sync({ accountId: a.id });
  const row = (subject) => service.messages({ accountId: a.id, folder: 'INBOX' }).rows.find((r) => r.subject === subject);
  return { one, two, service, a, b, row, close: async () => { await one.close(); await two.close(); } };
}

test('at the server: Paste copies with COPY and a cut moves with MOVE, within one account', async () => {
  const s = await twoServers();
  try {
    await s.service.paste({ accountId: s.a.id, folder: 'INBOX', ids: [s.row('Server message 1').id], to: { accountId: s.a.id, folder: 'Projects' } });
    await s.service.paste({ accountId: s.a.id, folder: 'INBOX', ids: [s.row('Server message 2').id], to: { accountId: s.a.id, folder: 'Done' }, cut: true });
    await s.service.flushChanges({ accountId: s.a.id });
    assert.deepEqual(s.one.messages('Projects').map((m) => m.subject), ['Server message 1']);
    assert.deepEqual(s.one.messages('Done').map((m) => m.subject), ['Server message 2']);
    assert.deepEqual(s.one.messages('INBOX').map((m) => m.subject), ['Server message 1', 'Server message 3']);
    assert.ok(s.one.heard.some((l) => / UID COPY /i.test(l)) && s.one.heard.some((l) => / UID MOVE /i.test(l)), 'the server\'s own COPY and MOVE');
    assert.ok(!s.one.heard.some((l) => /APPEND/i.test(l)), 'nothing appended within one account');
  } finally {
    await s.close();
  }
});

test('at the servers: a cut pasted into another account is APPENDed there with its marks, then taken from the first', async () => {
  const s = await twoServers();
  try {
    const id = s.row('Server message 3').id;
    s.service.flag({ accountId: s.a.id, folder: 'INBOX', ids: [id], patch: { flagged: true } });
    const r = await s.service.paste({ accountId: s.a.id, folder: 'INBOX', ids: [id], to: { accountId: s.b.id, folder: 'Kept' }, cut: true });
    assert.deepEqual(r, { copied: 1, moved: 1 });
    await s.service.flushChanges({ accountId: s.b.id });
    await s.service.flushChanges({ accountId: s.a.id });
    const there = s.two.messages('Kept');
    assert.deepEqual(there.map((m) => m.subject), ['Server message 3']);
    assert.deepEqual(there[0].flags.sort(), ['\\Flagged'], 'unread still, and flagged');
    assert.match(there[0].raw, /Body of message 3\./, 'the message as its server had it');
    assert.doesNotMatch(there[0].raw, /X-Rutba-Rebuilt/, 'fetched from its server, not rebuilt');
    assert.deepEqual(s.one.messages('INBOX').map((m) => m.subject), ['Server message 1', 'Server message 2'], 'and gone from the first');
  } finally {
    await s.close();
  }
});
