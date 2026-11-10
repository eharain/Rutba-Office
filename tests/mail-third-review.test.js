// Faults the third review of Mail found, 10 Nov 2026, each reproduced on the code as it stood.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailSync } from '../apps/desktop/main/mail-sync.js';
import { createMailService } from '../apps/desktop/main/mail.js';
import { MailStore } from '../packages/mailbox/src/store.js';
import { isSeparator, openMbox, writeMbox, scanMbox } from '../packages/mailbox/src/mbox.js';

const fresh = () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-3rd-'));
  return { root, s: new MailStore(root) };
};
const msg = (n, extra = {}) => ({ messageId: `m${n}@x`, subject: `Message ${n}`, from: [{ address: 'a@b.com' }], date: '2026-03-02T09:00:00Z', text: `Body ${n}`, uid: n, ...extra });
const subjects = (s, folder) => s.list('acc', folder, { limit: 100 }).rows.map((r) => r.subject).sort();
const fail = () => { const e = new Error('the machine stopped'); e.code = 'EIO'; throw e; };

test('a rename stopped before the list changed is finished the next time the folders are asked for', (t) => {
  const { s, root } = fresh();
  s.put('acc', 'Old', msg(1));
  s.put('acc', 'Old/Inner', msg(2));
  s.setFolders('acc', [{ path: 'Old', name: 'Old' }, { path: 'Old/Inner', name: 'Inner' }, { path: 'Other', name: 'Other' }]);
  t.mock.method(s, 'setFolders', fail);
  assert.throws(() => s.renameFolder('acc', 'Old', 'New'));
  t.mock.restoreAll();
  const after = new MailStore(root);
  assert.deepEqual(after.folders('acc').map((f) => f.path).sort(), ['New', 'New/Inner', 'Other']);
  assert.deepEqual(subjects(after, 'New'), ['Message 1']);
  assert.deepEqual(subjects(after, 'New/Inner'), ['Message 2']);
  assert.equal(after.get('acc', 'New', after.list('acc', 'New').rows[0].id).folder, 'New', 'each message says where it now is');
  assert.equal(fs.existsSync(path.join(root, 'acc', 'rename.json')), false, 'nothing left to finish');
});

test('a rename stopped after some folders moved is finished, and nothing is lost', (t) => {
  const { s, root } = fresh();
  s.put('acc', 'Old', msg(1));
  s.put('acc', 'Old/Inner', msg(2));
  s.setFolders('acc', [{ path: 'Old', name: 'Old' }, { path: 'Old/Inner', name: 'Inner' }]);
  const realRename = fs.renameSync;
  let moves = 0;
  t.mock.method(fs, 'renameSync', (a, b) => {
    if (!String(a).endsWith('.tmp') && ++moves === 2) fail();
    return realRename(a, b);
  });
  assert.throws(() => s.renameFolder('acc', 'Old', 'New'));
  t.mock.restoreAll();
  const after = new MailStore(root);
  assert.deepEqual(after.folders('acc').map((f) => f.path).sort(), ['New', 'New/Inner']);
  assert.deepEqual(subjects(after, 'New'), ['Message 1']);
  assert.deepEqual(subjects(after, 'New/Inner'), ['Message 2']);
});

test('a rename that went through leaves no mark behind and asks nothing of the next look', () => {
  const { s, root } = fresh();
  s.put('acc', 'Old', msg(1));
  s.setFolders('acc', [{ path: 'Old', name: 'Old' }]);
  s.renameFolder('acc', 'Old', 'New');
  assert.equal(fs.existsSync(path.join(root, 'acc', 'rename.json')), false);
  assert.deepEqual(s.folders('acc').map((f) => f.path), ['New']);
});

const list = (paths) => paths.map((p) => ({ path: p, name: p, flags: new Set() }));
function serverOf(state) {
  return {
    state,
    connect: async () => ({
      capabilities: new Map(),
      enabled: new Set(),
      get mailbox() { return { uidValidity: 7n, uidNext: 1, exists: 0, messages: 0 }; },
      list: async () => list(state.paths),
      status: async () => ({ messages: 0, uidNext: 1, uidValidity: 7n }),
      getMailboxLock: async () => ({ release() {} }),
      search: async () => [],
      logout: async () => {},
      fetch: async function* fetch() {},
    }),
  };
}
const syncOf = (store, server) => createMailSync({ store, account: () => ({ id: 'acc', email: 'me@example.com' }), connect: server.connect, describe: (f) => ({ path: f.path, name: f.name, role: f.path === 'INBOX' ? 'inbox' : null }) });
function matched(paths) {
  const { s, root } = fresh();
  s.setFolders('acc', paths.map((p) => ({ path: p, name: p })));
  for (const [i, p] of paths.entries()) {
    s.put('acc', p, msg(i + 1));
    s.setSyncState('acc', p, { uidValidity: '7', uidNext: 2, lastSeen: 1 });
  }
  return { s, root };
}

test('a server that answers with no folders at all does not delete the ones here', async () => {
  const { s } = matched(['INBOX', 'Projects', 'Receipts']);
  await syncOf(s, serverOf({ paths: [] })).fetch('acc').catch(() => {});
  assert.deepEqual(s.folders('acc').map((f) => f.path).sort(), ['INBOX', 'Projects', 'Receipts']);
  assert.deepEqual(subjects(s, 'Projects'), ['Message 2']);
});

test('a server whose list is cut short does not delete the folders it left out', async () => {
  const { s } = matched(['INBOX', 'Projects', 'Receipts', 'Travel']);
  await syncOf(s, serverOf({ paths: ['INBOX'] })).fetch('acc').catch(() => {});
  assert.deepEqual(s.folders('acc').map((f) => f.path).sort(), ['INBOX', 'Projects', 'Receipts', 'Travel']);
});

test('one folder deleted at the server still goes from here', async () => {
  const { s } = matched(['INBOX', 'Projects', 'Receipts']);
  await syncOf(s, serverOf({ paths: ['INBOX', 'Receipts'] })).fetch('acc').catch(() => {});
  assert.deepEqual(s.folders('acc').map((f) => f.path).sort(), ['INBOX', 'Receipts']);
});

test('a folder deleted at the server that holds mail only this computer has is kept, and shown as local', async () => {
  const { s } = matched(['INBOX', 'Projects', 'Receipts']);
  s.put('acc', 'Projects', msg(9, { uid: null }));
  await syncOf(s, serverOf({ paths: ['INBOX', 'Receipts'] })).fetch('acc').catch(() => {});
  const kept = s.folders('acc').find((f) => f.path === 'Projects');
  assert.equal(kept?.local, true);
  assert.ok(subjects(s, 'Projects').includes('Message 9'));
});

test('a spoiled folder list is made again from the folders: the next write keeps them', () => {
  const { s, root } = fresh();
  s.put('acc', 'INBOX', msg(1));
  s.put('acc', 'Only Here', msg(2));
  s.setFolders('acc', [{ path: 'INBOX', name: 'INBOX' }, { path: 'Only Here', name: 'Only Here' }]);
  fs.writeFileSync(path.join(root, 'acc', 'folders.json'), '{"half');
  const after = new MailStore(root);
  assert.deepEqual(after.folders('acc').map((f) => f.path).sort(), ['INBOX', 'Only Here']);
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(root, 'acc', 'folders.json'), 'utf8')).map((f) => f.path).sort(), ['INBOX', 'Only Here'], 'written back whole');
});

test('both lists spoiled: the folders are found on the disk and the program does not go round in circles', () => {
  const { s, root } = fresh();
  s.put('acc', 'Saved', msg(1));
  s.setFolders('acc', [{ path: 'Saved', name: 'Saved' }]);
  fs.writeFileSync(path.join(root, 'acc', 'folders.json'), '{"half');
  fs.writeFileSync(path.join(root, 'acc', 'folder-dirs.json'), '{"half');
  const after = new MailStore(root);
  assert.deepEqual(after.folders('acc').map((f) => f.path), ['Saved']);
  assert.deepEqual(subjects(after, 'Saved'), ['Message 1']);
});

test('no folder list yet is simply no folders, and nothing is written', () => {
  const { s, root } = fresh();
  assert.deepEqual(s.folders('acc'), []);
  assert.equal(fs.existsSync(path.join(root, 'acc', 'folders.json')), false);
});

test('a spoiled line of waiting changes is put aside, not written over', () => {
  const { s, root } = fresh();
  s.addChange('acc', { kind: 'flags' });
  const file = path.join(root, 'acc', 'changes.json');
  fs.writeFileSync(file, '{"next":2,"items":[{"kind":"flags"');
  assert.deepEqual(s.changes('acc').items, []);
  s.addChange('acc', { kind: 'move' });
  const aside = fs.readdirSync(path.join(root, 'acc')).filter((n) => n.startsWith('changes.json.damaged-'));
  assert.equal(aside.length, 1);
  assert.match(fs.readFileSync(path.join(root, 'acc', aside[0]), 'utf8'), /"kind":"flags"/);
});

test('the folder list and the waiting changes reach the disk before the call returns', (t) => {
  const { s } = fresh();
  const flushes = t.mock.method(fs, 'fsyncSync');
  s.setFolders('acc', [{ path: 'INBOX', name: 'INBOX' }]);
  assert.equal(flushes.mock.callCount(), 1);
  s.addChange('acc', { kind: 'flags' });
  assert.equal(flushes.mock.callCount(), 2);
});

test('a message put with flush reaches the disk, and its row with it; without flush nothing waits', (t) => {
  const { s } = fresh();
  const flushes = t.mock.method(fs, 'fsyncSync');
  s.put('acc', 'INBOX', msg(1));
  const plain = flushes.mock.callCount();
  s.put('acc', 'INBOX', msg(2), { flush: true });
  assert.equal(flushes.mock.callCount() - plain, 2, 'the message and its row');
});

test('a move has its message on the disk before the copy it came from is taken away', (t) => {
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-3rd-'));
  const settings = new Map();
  const stores = { settings: { get: (k, d) => (settings.has(k) ? structuredClone(settings.get(k)) : d), set: (k, v) => settings.set(k, structuredClone(v)), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) }, secrets: { available: () => true, get: () => null, set: () => {}, delete: () => {} } };
  const service = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData });
  const account = service.addAccount({ account: { email: 'me@example.com', name: 'Me', local: true }, password: 'x' });
  const store = new MailStore(path.join(userData, 'mail'));
  store.put(account.id, 'INBOX', msg(1));
  store.put(account.id, 'Archive', msg(2));
  store.upsertFolder(account.id, { path: 'Archive', name: 'Archive' });
  const { id } = store.list(account.id, 'INBOX').rows[0];
  const order = [];
  const realFsync = fs.fsyncSync;
  const realRm = fs.rmSync;
  t.mock.method(fs, 'fsyncSync', (h) => { order.push('flush'); return realFsync(h); });
  t.mock.method(fs, 'rmSync', (p, o) => { if (String(p).endsWith(`${id}.json`)) order.push('remove'); return realRm(p, o); });
  service.move({ accountId: account.id, folder: 'INBOX', ids: [id], to: 'Archive' });
  t.mock.restoreAll();
  const removed = order.indexOf('remove');
  assert.ok(removed >= 2, `the message and its row flushed first: ${order.join(', ')}`);
  assert.deepEqual(order.slice(0, 2), ['flush', 'flush']);
});

test('a line spoiled in the middle of an index: the message comes back from its own file', () => {
  const { s, root } = fresh();
  for (const n of [1, 2, 3]) s.put('acc', 'INBOX', msg(n));
  const index = path.join(root, 'acc', 'INBOX', 'index.jsonl');
  const lines = fs.readFileSync(index, 'utf8').split('\n').filter(Boolean);
  lines[1] = '{"id":"spoi';
  fs.writeFileSync(index, `${lines.join('\n')}\n`);
  const after = new MailStore(root);
  assert.deepEqual(subjects(after, 'INBOX'), ['Message 1', 'Message 2', 'Message 3']);
  // The next rewrite of the list keeps it.
  after.setFlagsMany('acc', 'INBOX', [after.list('acc', 'INBOX', { limit: 3 }).rows[0].id], { unread: false });
  assert.deepEqual(subjects(new MailStore(root), 'INBOX'), ['Message 1', 'Message 2', 'Message 3']);
});

test('a torn last line is only that: a message removed earlier does not come back', () => {
  const { s, root } = fresh();
  for (const n of [1, 2]) s.put('acc', 'INBOX', msg(n));
  const index = path.join(root, 'acc', 'INBOX', 'index.jsonl');
  const [first] = fs.readFileSync(index, 'utf8').split('\n');
  // The file of message 2 is still here, its row gone, as a stop between the two steps of a remove leaves it.
  fs.writeFileSync(index, `${first}\n{"id":"half`);
  assert.deepEqual(subjects(new MailStore(root), 'INBOX'), ['Message 1']);
});

test('a folder called Outgoing from before the name was reserved keeps its mail', () => {
  const { s, root } = fresh();
  // As 2.24.0 left it: the folder's files in the directory the account now uses for itself.
  const old = path.join(root, 'acc', 'Outgoing');
  s.put('acc', 'Elsewhere', msg(1));
  const source = path.join(root, 'acc', 'Elsewhere');
  fs.cpSync(source, old, { recursive: true });
  fs.writeFileSync(path.join(old, 'waiting.eml'), 'a sent message');
  s.setFolders('acc', [{ path: 'Outgoing', name: 'Outgoing' }]);
  const after = new MailStore(root);
  assert.deepEqual(subjects(after, 'Outgoing'), ['Message 1']);
  assert.ok(fs.existsSync(path.join(old, 'waiting.eml')), 'the account\'s own file stays');
  assert.equal(fs.existsSync(path.join(old, 'index.jsonl')), false, 'the folder has moved out');
});

for (const line of [
  'From  Mon Jan  1 00:00:00 2024',
  'From a@b.com Mon Jan  1 00:00:00 2024 +0000',
  'From a@b.com Mon Jan  1 00:00:00 +0000 2024',
  'From a@b.com 2024-01-02T03:04:05Z',
  'From a@b.com 2024-01-02 03:04:05 +05:00',
  'From "John Smith"@x.com Tue, 02 Jan 2024 03:04:05 GMT',
  'From a@b.com Tue Jan  2 03:04:05 2024',
]) {
  test(`this starts a message: ${line}`, () => assert.equal(isSeparator(Buffer.from(`${line}\n\nno header follows\n`)), true));
}
for (const line of [
  'From here on we use the new office',
  'From noon on Jan 5 10:30 we met there',
  'From the 2024-01-02 meeting',
  'From Jan 5 10:30 onwards',
]) {
  test(`this is a line of a message: ${line}`, () => assert.equal(isSeparator(Buffer.from(`${line}\nand more words\n`)), false));
}

test('the archive this program writes for a quoted address reads back as its messages', () => {
  const raw = (n) => `Subject: ${n}\r\n\r\nBody ${n}\r\n`;
  const bytes = writeMbox([{ from: { address: '"John Smith"@x.com' }, date: '2024-01-02T03:04:05Z', raw: raw(1) }, { from: { address: '"Ana B"@x.com' }, date: '2024-01-03T03:04:05Z', raw: raw(2) }]);
  assert.equal(scanMbox(bytes).ranges.length, 2);
});

test('an archive of nothing but lines that begin From is read without a read of the file for each', (t) => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mbox-3rd-')), 'a.mbox');
  fs.writeFileSync(file, `From a@b.com Mon Jan  1 00:00:00 2024\nSubject: x\n\n${'From here on\n'.repeat(200000)}`);
  const reads = t.mock.method(fs, 'readSync');
  const box = openMbox(file, { block: 1048576 });
  const count = reads.mock.callCount();
  t.mock.restoreAll();
  box.close();
  assert.equal(box.length, 1);
  assert.ok(count < 200, `${count} reads for 200,000 lines`);
});

test('a separator that crosses the end of a block is still found', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mbox-3rd-')), 'b.mbox');
  const first = `From a@b.com Mon Jan  1 00:00:00 2024\nSubject: x\n\n${'y'.repeat(980)}\n`;
  fs.writeFileSync(file, `${first}From c@d.com Mon Jan  2 00:00:00 2024\nSubject: z\n\nbody\n`);
  const box = openMbox(file, { block: first.length + 8 });
  box.close();
  assert.equal(box.length, 2);
});
