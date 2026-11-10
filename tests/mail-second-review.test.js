// Faults the second review of Mail found, 10 Nov 2026, each reproduced on the code as it stood.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';
import { createMailSync } from '../apps/desktop/main/mail-sync.js';
import { MailStore } from '../packages/mailbox/src/store.js';
import { parseMessage } from '../packages/mailbox/src/mime.js';

function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: {
      get: (k, fallback) => (settings.has(k) ? structuredClone(settings.get(k)) : fallback),
      set: (k, v) => settings.set(k, structuredClone(v)),
      delete: (k) => settings.delete(k),
      all: () => Object.fromEntries(settings),
    },
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k) },
  };
}

function fixture(smtp) {
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-2nd-'));
  const stores = fakeStores();
  const clock = { at: Date.parse('2026-03-02T09:00:00Z') };
  const sent = [];
  const ship = smtp?.(sent) || { async send({ draft }) { sent.push(draft); return { messageId: `m-${sent.length}` }; } };
  const told = [];
  const service = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: (name, data) => told.push({ name, data }), userData, now: () => clock.at, smtp: ship });
  const account = service.addAccount({ account: { email: 'me@example.com', name: 'Me', imap: { host: 'h', port: 993 }, smtp: { host: 'h', port: 465 } }, password: 'x' });
  return { service, account, stores, sent, clock, userData, told };
}
const dueNow = (f) => new Date(f.clock.at).toISOString();
const later = (f) => new Date(f.clock.at + 3600_000).toISOString();

const raw = (n) => `From: a@b.com\r\nTo: me@example.com\r\nSubject: Message ${n}\r\nMessage-ID: <m${n}@x>\r\nDate: Mon, 2 Mar 2026 09:00:00 +0000\r\n\r\nBody ${n}\r\n`;

/** A server of the messages given, as much of IMAP as a fetch uses. */
function fakeServer(count) {
  const state = { count };
  const uids = () => Array.from({ length: state.count }, (_, i) => i + 1);
  const wanted = (range) => {
    const all = uids();
    if (/:\*$/.test(range)) return all.filter((u) => u >= Number(range.split(':')[0]));
    return String(range).split(',').flatMap((part) => {
      const [a, b] = part.split(':').map(Number);
      return all.filter((u) => u >= a && u <= (b ?? a));
    });
  };
  return {
    state,
    connect: async () => ({
      capabilities: new Map(),
      enabled: new Set(),
      get mailbox() { return { uidValidity: 7n, uidNext: state.count + 1, exists: state.count, messages: state.count }; },
      list: async () => [{ path: 'INBOX', name: 'INBOX', flags: new Set() }],
      status: async () => ({ messages: state.count, uidNext: state.count + 1, uidValidity: 7n }),
      getMailboxLock: async () => ({ release() {} }),
      search: async () => uids(),
      logout: async () => {},
      fetch: async function* fetch(range, query) {
        for (const uid of wanted(range)) {
          yield { uid, flags: new Set(), headers: Buffer.from(`Message-ID: <m${uid}@x>\r\n\r\n`), source: query.source ? Buffer.from(raw(uid)) : undefined };
        }
      },
    }),
  };
}
const syncOf = (store, server) => createMailSync({ store, account: () => ({ id: 'acc', email: 'me@example.com' }), connect: server.connect, describe: (f) => ({ path: f.path, name: f.name, role: 'inbox' }) });

test('a stop between a message file and its row: fetching again brings the message into the list', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-sync-2nd-'));
  const server = fakeServer(3);
  let store = new MailStore(root);
  await syncOf(store, server).fetch('acc');
  assert.equal(store.list('acc', 'INBOX', { limit: 100 }).rows.length, 3);

  // What the stop leaves: the files of the batch, no row for the last, and the folder not yet marked as read that far.
  const index = path.join(root, 'acc', 'INBOX', 'index.jsonl');
  const lines = fs.readFileSync(index, 'utf8').split('\n').filter(Boolean);
  assert.equal(lines.length, 3);
  fs.writeFileSync(index, `${lines.filter((l) => !l.includes('Message 3')).join('\n')}\n`);
  const prev = store.syncState('acc', 'INBOX');
  store = new MailStore(root);
  store.setSyncState('acc', 'INBOX', { ...prev, uidNext: 3, messages: 2, lastSeen: 2 });
  assert.equal(store.listedIds('acc', 'INBOX').size, 2);

  await syncOf(store, server).fetch('acc');
  assert.deepEqual(store.list('acc', 'INBOX', { limit: 100 }).rows.map((r) => r.subject).sort(), ['Message 1', 'Message 2', 'Message 3']);
  server.state.count = 4;
  await syncOf(store, server).fetch('acc');
  assert.equal(store.list('acc', 'INBOX', { limit: 100 }).rows.length, 4, 'and the next one arrives');
});

test('fetching twice makes no second rows', async () => {
  const server = fakeServer(2);
  const store = new MailStore(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-sync-2nd-')));
  const sync = syncOf(store, server);
  await sync.fetch('acc');
  await sync.fetch('acc');
  assert.equal(store.list('acc', 'INBOX', { limit: 100 }).rows.length, 2);
});

test('the same message twice in one batch makes one row', () => {
  const store = new MailStore(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-store-2nd-')));
  const m = () => parseMessage(raw(1));
  assert.equal(store.putMany('acc', 'INBOX', [m(), m(), m()]), 1);
  assert.equal(store.list('acc', 'INBOX', { limit: 100 }).rows.length, 1);
});

test('a spoiled list of folder directories does not stop the account: it is made again from the folders', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-store-2nd-'));
  let store = new MailStore(root);
  store.put('acc', 'INBOX', parseMessage(raw(1)));
  store.put('acc', 'Work/2026', parseMessage(raw(2)));
  store.setFolders('acc', [{ path: 'INBOX', name: 'INBOX' }, { path: 'Work/2026', name: '2026' }]);
  const list = path.join(root, 'acc', 'folder-dirs.json');
  assert.ok(fs.existsSync(list), 'the file is there to spoil');
  for (const spoiled of ['', '{"INBOX":']) {
    fs.writeFileSync(list, spoiled);
    store = new MailStore(root);
    assert.equal(store.list('acc', 'INBOX', { limit: 10 }).rows.length, 1);
    assert.equal(store.list('acc', 'Work/2026', { limit: 10 }).rows.length, 1, `the folder with a changed name, after ${JSON.stringify(spoiled)}`);
  }
});

test('Undo while the message is going to the server is refused, and the message is not also handed back', async () => {
  let release;
  const held = new Promise((r) => { release = r; });
  const f = fixture((sent) => ({ async send({ draft }) { await held; sent.push(draft); return { messageId: 'm-1' }; } }));
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Going', text: 'The words' }, at: dueNow(f) });
  const pass = f.service.pumpOutbox();
  await new Promise((r) => setTimeout(r, 30));
  assert.deepEqual(f.service.unsend({ id: item.id }), { sending: true });
  assert.equal(f.service.outbox().length, 1, 'still on the list while it goes');
  release();
  await pass;
  assert.equal(f.sent.length, 1);
  assert.equal(f.service.outbox().length, 0);
});

test('a message left marked as sending by a stopped program can be taken back', () => {
  const f = fixture();
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Left', text: 'The words' }, at: later(f) });
  f.stores.settings.set('mail.outbox', f.stores.settings.get('mail.outbox', []).map((o) => ({ ...o, sending: true })));
  assert.equal(f.service.unsend({ id: item.id }).draft.text, 'The words');
});

test('taking back that fails part way leaves the message in the Outbox with its words and attachments', () => {
  const f = fixture();
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-att-2nd-'));
  const file = path.join(dir, 'a.pdf');
  fs.writeFileSync(file, 'the attached bytes');
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Kept', text: 'The words', attachments: [{ filename: 'a.pdf', path: file }] }, at: later(f) });
  const real = fs.mkdtempSync;
  fs.mkdtempSync = () => { throw Object.assign(new Error('no space'), { code: 'ENOSPC' }); };
  try {
    assert.throws(() => f.service.unsend({ id: item.id }), /no space/);
  } finally {
    fs.mkdtempSync = real;
  }
  assert.equal(f.service.outbox().length, 1);
  const back = f.service.unsend({ id: item.id });
  assert.equal(back.draft.text, 'The words');
  assert.equal(fs.readFileSync(back.draft.attachments[0].path, 'utf8'), 'the attached bytes');
});

test('moved to Drafts: if the draft cannot be saved the message stays in the Outbox', async () => {
  const f = fixture();
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Kept', text: 'The words' }, at: later(f) });
  const real = MailStore.prototype.put;
  MailStore.prototype.put = () => { throw new Error('disk full'); };
  try {
    await assert.rejects(Promise.resolve().then(() => f.service.unsend({ id: item.id, toDrafts: true })), /disk full/);
  } finally {
    MailStore.prototype.put = real;
  }
  assert.equal(f.service.outbox().length, 1);
  await f.service.unsend({ id: item.id, toDrafts: true });
  assert.equal(f.service.outbox().length, 0);
});

test('a failure after the server took the message does not send it again', async () => {
  const f = fixture();
  f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Once', text: 'The words' }, at: dueNow(f) });
  const real = MailStore.prototype.put;
  MailStore.prototype.put = () => { throw new Error('disk full'); };
  try {
    await f.service.pumpOutbox();
    f.clock.at += 3600_000;
    await f.service.pumpOutbox();
  } finally {
    MailStore.prototype.put = real;
  }
  assert.equal(f.sent.length, 1, 'once');
  assert.equal(f.service.outbox().length, 0);
  assert.equal(f.told.find((x) => x.name === 'mail:sent').data.notKept, true, 'and the person is told the copy was not kept');
});

test('an mbox export keeps attachments and a body line starting From, and is under its name only when whole', async () => {
  const f = fixture();
  const bytes = Buffer.from('attached-bytes-here-'.repeat(50));
  const message = `From: a@b.com\r\nTo: me@example.com\r\nSubject: With file\r\nMessage-ID: <f1@x>\r\nDate: Mon, 2 Mar 2026 09:00:00 +0000\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary="B"\r\n\r\n--B\r\nContent-Type: text/plain\r\n\r\nFrom the desk of nobody\r\n\r\n--B\r\nContent-Type: application/octet-stream; name="a.bin"\r\nContent-Disposition: attachment; filename="a.bin"\r\nContent-Transfer-Encoding: base64\r\n\r\n${bytes.toString('base64')}\r\n--B--\r\n`;
  new MailStore(path.join(f.userData, 'mail')).put(f.account.id, 'INBOX', parseMessage(message));
  const out = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-out-2nd-')), 'backup.mbox');
  const done = await f.service.export({ accountId: f.account.id, folder: 'INBOX', path: out, format: 'mbox' });
  assert.equal(done.messages, 1);
  const text = fs.readFileSync(out, 'latin1');
  assert.match(text, /^From .*\n/, 'starts with its separator');
  assert.equal(/^From the desk/m.test(text), false, 'the body line cannot be read as the start of a message');
  assert.ok(text.replace(/[\r\n]/g, '').includes(bytes.toString('base64').replace(/[\r\n]/g, '')), 'the attachment is whole');
  assert.equal(fs.existsSync(`${out}.part`), false);
});

test('mail written out carries the mark of a file from the internet', { skip: process.platform !== 'win32' }, async () => {
  const f = fixture();
  new MailStore(path.join(f.userData, 'mail')).put(f.account.id, 'INBOX', parseMessage('From: a@b.com\r\nTo: me@example.com\r\nSubject: Marked\r\nMessage-ID: <m1@x>\r\nDate: Mon, 2 Mar 2026 09:00:00 +0000\r\n\r\nBody\r\n'));
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-out-mark-'));
  const mbox = path.join(dir, 'backup.mbox');
  await f.service.export({ accountId: f.account.id, folder: 'INBOX', path: mbox, format: 'mbox' });
  assert.match(fs.readFileSync(`${mbox}:Zone.Identifier`, 'utf8'), /ZoneId=3/);
  const folder = path.join(dir, 'each');
  await f.service.export({ accountId: f.account.id, folder: 'INBOX', path: folder, format: 'eml' });
  const [one] = fs.readdirSync(folder);
  assert.match(one, /\.eml$/);
  assert.match(fs.readFileSync(path.join(folder, `${one}:Zone.Identifier`), 'utf8'), /ZoneId=3/);
});
