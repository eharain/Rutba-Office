// What the review found could lose mail, each as the failure it would be.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';
import { MailStore } from '../packages/mailbox/src/store.js';
import { freeNameIn } from '../apps/desktop/renderer/apps/mail/parts.js';

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
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k), keys: () => [...secrets.keys()] },
  };
}

function fixture({ smtp } = {}) {
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-lose-'));
  const stores = fakeStores();
  const clock = { at: Date.parse('2026-03-02T09:00:00Z') };
  const sent = [];
  const ship = smtp || { async send({ draft }) { sent.push(draft); return { messageId: `m-${sent.length}` }; } };
  const make = (over = ship) => createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, now: () => clock.at, smtp: over });
  const service = make();
  const account = service.addAccount({ account: { email: 'me@example.com', name: 'Me', imap: { host: 'h', port: 993 }, smtp: { host: 'h', port: 465 } }, password: 'x' });
  return { service, make, account, stores, sent, clock, userData };
}
const dueNow = (f) => new Date(f.clock.at).toISOString();
const later = (f) => new Date(f.clock.at + 3600_000).toISOString();

test('Send now that fails leaves the message in the Outbox', async () => {
  const f = fixture({ smtp: { async send() { throw new Error('the server said no'); } } });
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Keep me' }, at: later(f) });
  const done = await f.service.sendNow({ id: item.id });
  assert.equal(done.sent, false);
  assert.match(done.error, /said no/);
  assert.equal(f.service.outbox().length, 1);
  assert.equal(f.service.outbox()[0].draft.subject, 'Keep me');
});

test('Send now that works takes it off, once', async () => {
  const f = fixture();
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Go' }, at: later(f) });
  assert.deepEqual(await f.service.sendNow({ id: item.id }), { sent: true });
  assert.equal(f.sent.length, 1);
  assert.equal(f.service.outbox().length, 0);
});

test('Send now of something no longer waiting says so', async () => {
  const f = fixture();
  assert.match((await f.service.sendNow({ id: 'gone' })).error, /no longer waiting/);
});

test('a stop in the middle of sending loses nothing: the rest are still due and the one going is held', async () => {
  let release;
  const stuck = new Promise((r) => { release = r; });
  const f = fixture({ smtp: { async send() { await stuck; return { messageId: 'x' }; } } });
  for (const subject of ['One', 'Two', 'Three']) f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject }, at: dueNow(f) });
  // Queue pumps by itself; give it the moment to reach the first send.
  for (let i = 0; i < 50 && !f.service.outbox().some((o) => o.sending); i++) await new Promise((r) => setTimeout(r, 5));
  const during = f.service.outbox();
  assert.equal(during.length, 3, 'all three still written down while the first goes');
  assert.deepEqual(during.filter((o) => o.sending).map((o) => o.draft.subject), ['One']);

  // The program stops here. Another run starts from what was written.
  const sent = [];
  const after = f.make({ async send({ draft }) { sent.push(draft.subject); return { messageId: 'y' }; } });
  await after.pumpOutbox();
  assert.deepEqual(sent, ['Two', 'Three'], 'those not reached go; the one in doubt does not go twice');
  const left = after.outbox();
  assert.equal(left.length, 1);
  assert.equal(left[0].draft.subject, 'One');
  assert.equal(left[0].held, true, 'held for the person to decide');
  await after.pumpOutbox();
  assert.deepEqual(sent, ['Two', 'Three'], 'and not sent by the next pass');

  // The person looked, and sends it.
  assert.deepEqual(await after.sendNow({ id: left[0].id }), { sent: true });
  assert.deepEqual(sent, ['Two', 'Three', 'One']);
  release();
});

test('a message taken back while the queue is working is not sent', async () => {
  const f = fixture();
  const a = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'A' }, at: later(f) });
  const b = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'B' }, at: later(f) });
  f.clock.at += 7200_000;
  f.service.unsend({ id: b.id });
  await f.service.pumpOutbox();
  assert.deepEqual(f.sent.map((d) => d.subject), ['A']);
  assert.ok(a.id);
});

test('a scheduled message keeps its attachments when the temporary folder is cleared', async () => {
  const f = fixture();
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-att-'));
  const file = path.join(dir, 'scratch.pdf');
  fs.writeFileSync(file, 'the attached bytes');
  f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'With file', attachments: [{ filename: 'report.pdf', path: file }] }, at: later(f) });
  fs.rmSync(dir, { recursive: true });
  assert.equal(f.service.outbox()[0].draft.attachments[0].path, undefined, 'no temporary path is relied on');

  await f.service.pumpOutbox();
  const restarted = f.make();
  f.clock.at += 7200_000;
  await restarted.pumpOutbox();
  assert.equal(f.sent.length, 1);
  assert.equal(f.sent[0].attachments[0].filename, 'report.pdf');
  assert.equal(Buffer.from(f.sent[0].attachments[0].content).toString(), 'the attached bytes');
});

test('a scheduled attachment that cannot be read stops the scheduling, out loud', () => {
  const f = fixture();
  assert.throws(() => f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', attachments: [{ filename: 'x', path: path.join(os.tmpdir(), 'not-there-at-all.pdf') }] }, at: later(f) }), /ENOENT/);
  assert.equal(f.service.outbox().length, 0);
});

test('taken back, a scheduled message has its attachments as files again', () => {
  const f = fixture();
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-att-')), 'a.txt');
  fs.writeFileSync(file, 'words');
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', attachments: [{ filename: 'a.txt', path: file }] }, at: later(f) });
  const kept = f.service.outbox()[0].draft.attachments[0].kept;
  assert.ok(fs.existsSync(kept));
  const back = f.service.unsend({ id: item.id });
  assert.equal(fs.readFileSync(back.draft.attachments[0].path, 'utf8'), 'words');
  assert.equal(fs.existsSync(kept), false, 'and the copy is let go');
});

// ---- the store

function store() {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-store-lose-'));
  return { root, s: new MailStore(root) };
}
const msg = (n, extra = {}) => ({ id: `m${n}`, subject: `Message ${n}`, from: { address: 'a@b.c' }, to: [], date: '2026-01-01T00:00:00Z', text: `body ${n}`, attachments: [], ...extra });
const subjects = (s, folder) => s.list('acc', folder, { limit: 100 }).rows.map((r) => r.subject).sort();

test('folders whose names come to one directory keep their own mail', () => {
  const { s } = store();
  ['Old Mail', 'Old-Mail', 'Old/Mail', 'INBOX', 'Inbox'].forEach((folder, n) => s.put('acc', folder, msg(n)));
  assert.deepEqual(subjects(s, 'Old Mail'), ['Message 0']);
  assert.deepEqual(subjects(s, 'Old-Mail'), ['Message 1']);
  assert.deepEqual(subjects(s, 'Old/Mail'), ['Message 2']);
  assert.deepEqual(subjects(s, 'INBOX'), ['Message 3']);
  assert.deepEqual(subjects(s, 'Inbox'), ['Message 4']);
});

test('removing one of those folders leaves the others', () => {
  const { s } = store();
  s.put('acc', 'Old Mail', msg(1));
  s.put('acc', 'Old/Mail', msg(2));
  s.setFolders('acc', [{ path: 'Old Mail', name: 'Old Mail' }, { path: 'Old/Mail', name: 'Mail' }]);
  s.removeFolder('acc', 'Old Mail');
  assert.deepEqual(subjects(s, 'Old/Mail'), ['Message 2']);
});

test('a folder called outgoing is not the queue of sent mail, and a folder keeps its directory across a restart', () => {
  const { s, root } = store();
  s.put('acc', 'outgoing', msg(1));
  s.put('acc', 'Old Mail', msg(2));
  s.put('acc', 'Old/Mail', msg(3));
  const again = new MailStore(root);
  assert.deepEqual(subjects(again, 'outgoing'), ['Message 1']);
  assert.deepEqual(subjects(again, 'Old/Mail'), ['Message 3']);
  assert.equal(fs.existsSync(path.join(root, 'acc', 'outgoing', 'index.jsonl')), false);
});

test('a directory from before is still found by its folder', () => {
  const { s, root } = store();
  s.put('acc', 'Old Mail', msg(1));
  fs.rmSync(path.join(root, 'acc', 'folder-dirs.json'), { force: true });
  const again = new MailStore(root);
  assert.deepEqual(subjects(again, 'Old Mail'), ['Message 1']);
});

test('a store that holds an older list reads the one written since, and writes nothing over it', () => {
  const { s, root } = store();
  s.put('acc', 'Old Mail', msg(1));
  const searching = new MailStore(root);
  assert.deepEqual(subjects(searching, 'Old Mail'), ['Message 1']);
  s.put('acc', 'Old/Mail', msg(2));
  assert.deepEqual(subjects(searching, 'Old/Mail'), ['Message 2']);
  assert.deepEqual(subjects(new MailStore(root), 'Old Mail'), ['Message 1']);
});

test('renaming a folder onto a name that would collide keeps both', () => {
  const { s } = store();
  s.put('acc', 'A B', msg(1));
  s.put('acc', 'C', msg(2));
  s.setFolders('acc', [{ path: 'A B', name: 'A B' }, { path: 'C', name: 'C' }]);
  s.renameFolder('acc', 'C', 'A/B');
  assert.deepEqual(subjects(s, 'A B'), ['Message 1']);
  assert.deepEqual(subjects(s, 'A/B'), ['Message 2']);
});

test('a crash while a folder index is being rewritten leaves the old one whole', (t) => {
  const { s, root } = store();
  s.put('acc', 'INBOX', msg(1));
  s.put('acc', 'INBOX', msg(2));
  const index = path.join(root, 'acc', 'INBOX', 'index.jsonl');
  const before = fs.readFileSync(index, 'utf8');
  t.mock.method(fs, 'renameSync', () => { const e = new Error('the machine stopped'); e.code = 'EIO'; throw e; });
  assert.throws(() => s.removeMany('acc', 'INBOX', ['m1']));
  t.mock.restoreAll();
  assert.equal(fs.readFileSync(index, 'utf8'), before);
  assert.deepEqual(fs.readdirSync(path.dirname(index)).filter((n) => n.endsWith('.tmp')), [], 'and no half file left about');
});

test('a message file is written whole or not at all', (t) => {
  const { s, root } = store();
  s.put('acc', 'INBOX', msg(1));
  const dir = path.join(root, 'acc', 'INBOX', 'm');
  const file = path.join(dir, fs.readdirSync(dir)[0]);
  const before = fs.readFileSync(file, 'utf8');
  t.mock.method(fs, 'renameSync', () => { const e = new Error('stopped'); e.code = 'EIO'; throw e; });
  assert.throws(() => s.put('acc', 'INBOX', msg(1, { subject: 'Changed' }), { force: true }));
  t.mock.restoreAll();
  assert.equal(fs.readFileSync(file, 'utf8'), before);
});

test('a row added after a torn last line is not lost with it', () => {
  const { s, root } = store();
  s.put('acc', 'INBOX', msg(1));
  fs.appendFileSync(path.join(root, 'acc', 'INBOX', 'index.jsonl'), '{"id":"half');
  const again = new MailStore(root);
  again.put('acc', 'INBOX', msg(2));
  assert.deepEqual(subjects(new MailStore(root), 'INBOX'), ['Message 1', 'Message 2']);
});

test('Save all: two attachments of one name, or a file already there, are both kept', async () => {
  const there = new Set(['C:/out/report.pdf', 'C:/out/report (2).pdf']);
  const exists = async (p) => there.has(p);
  assert.equal(await freeNameIn('C:/out', 'report.pdf', exists), 'C:/out/report (3).pdf');
  assert.equal(await freeNameIn('C:/out', 'new.pdf', exists), 'C:/out/new.pdf');
  assert.equal(await freeNameIn('C:/out', '..\\..\\evil.exe', exists), 'C:/out/.._.._evil.exe', 'a name cannot climb out of the folder');
  assert.equal(await freeNameIn('C:/out', 'no extension', exists), 'C:/out/no extension');
});
