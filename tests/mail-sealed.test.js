// Mail sealed with the profile's password (D-2026-10-26-02, 2.19).
//
// In a profile with a password every file of the mail store is sealed with
// the profile's key: no subject, address or attachment byte is on disk in
// the clear. Without one nothing changes. Setting, changing and taking away
// the password seals every file again, and a run stopped part way (a crash)
// carries on at the next open, its files readable all the while, never
// left sealed with a key the lock no longer makes. The cost is measured.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { MailStore } from '../packages/mailbox/src/store.js';
import { MAGIC } from '../packages/mailbox/src/sealed.js';

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-sealed-'));
const key = () => crypto.randomBytes(32);

const message = (n, secret = 'Harbour') => ({
  messageId: `<sealed-${n}@work.example>`,
  subject: `${secret} plans ${n}`,
  from: [{ name: 'Amina Yusuf', address: 'amina@work.example' }],
  to: [{ name: 'Sam', address: 'sam@checks.example' }],
  date: new Date(Date.UTC(2026, 9, 1, 9, n % 60)).toISOString(),
  text: `The ${secret} figures for message ${n}.`,
  unread: true,
  attachments: n % 4 === 0 ? [{ filename: `${secret}-${n}.txt`, type: 'text/plain', bytes: Buffer.from(`${secret} attachment ${n} `.repeat(40)) }] : [],
});

/** Every file the store keeps, and what each holds. */
function files(root) {
  const out = [];
  const walk = (dir) => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, e.name);
      if (e.isDirectory()) walk(p);
      else out.push({ file: path.relative(root, p), bytes: fs.readFileSync(p) });
    }
  };
  walk(root);
  return out;
}
const clearText = (root, word) => files(root).filter((f) => f.file !== '.reseal.json' && f.bytes.includes(Buffer.from(word)));

function fill(store, count, secret) {
  store.setFolders('acc', [{ path: 'INBOX', name: 'INBOX' }, { path: 'Archive', name: 'Archive' }]);
  store.putMany('acc', 'INBOX', Array.from({ length: count }, (_, i) => message(i, secret)));
  store.putMany('acc', 'Archive', Array.from({ length: Math.ceil(count / 10) }, (_, i) => message(100000 + i, secret)));
  store.setSyncState('acc', 'INBOX', { uidValidity: '7', lastSeen: count, secret });
  store.addChange('acc', { kind: 'flags', folder: 'INBOX', items: [{ id: 'x', messageId: secret }] });
  store.writeOutgoing('acc', 'sent-1.eml', Buffer.from(`Subject: ${secret} sent\r\n\r\nHello`));
}

test('without a password nothing is sealed, as before', () => {
  const root = tmp();
  const store = new MailStore(root);
  fill(store, 8, 'Plainword');
  assert.ok(clearText(root, 'Plainword').length > 0);
  assert.ok(!files(root).some((f) => f.bytes.subarray(0, MAGIC.length).equals(MAGIC)));
  assert.equal(store.list('acc', 'INBOX').total, 8);
});

test('with a password every mail file is sealed, and reads back whole', () => {
  const root = tmp();
  const k = key();
  const store = new MailStore(root, { key: k });
  fill(store, 12, 'Lighthouse');
  const id = store.list('acc', 'INBOX').rows.find((r) => r.subject === 'Lighthouse plans 4').id;
  store.setFlags('acc', 'INBOX', id, { unread: false });
  store.put('acc', 'Archive', { ...store.get('acc', 'INBOX', id) }, { force: true });
  assert.deepEqual(clearText(root, 'Lighthouse'), [], 'no subject, body, attachment, list, sync state, change or outgoing message in the clear');
  assert.deepEqual(clearText(root, 'amina@work.example'), [], 'nor an address');
  assert.equal(store.list('acc', 'INBOX').total, 12);
  assert.equal(store.get('acc', 'INBOX', id).unread, false);
  const a = store.get('acc', 'INBOX', id).attachments[0];
  assert.match(store.attachmentBytes('acc', 'INBOX', id, a.stored).toString(), /^Lighthouse attachment 4/);
  assert.match(store.attachmentBytes('acc', 'Archive', id, a.stored).toString(), /^Lighthouse attachment 4/, 'carried to another folder sealed');
  assert.equal(store.syncState('acc', 'INBOX').secret, 'Lighthouse');
  assert.equal(store.changes('acc').items[0].items[0].messageId, 'Lighthouse');
  assert.match(store.readOutgoing(store.outgoingPath('acc', 'sent-1.eml')).toString(), /Lighthouse sent/);
  // Each index line sealed alone: a message arriving is one line appended.
  const before = fs.readFileSync(path.join(root, 'acc', 'INBOX', 'index.jsonl'), 'utf8').split('\n').filter(Boolean).length;
  store.put('acc', 'INBOX', message(500, 'Lighthouse'));
  const lines = fs.readFileSync(path.join(root, 'acc', 'INBOX', 'index.jsonl'), 'utf8').split('\n').filter(Boolean);
  assert.equal(lines.length, before + 1);
  assert.ok(lines.every((l) => l.startsWith('~')));
  // Another key opens nothing; the store with none sees an empty folder, and the files are not touched.
  const wrong = new MailStore(root, { key: key() });
  assert.equal(wrong.get('acc', 'INBOX', id), null);
  assert.equal(wrong.list('acc', 'INBOX').total, 0);
  assert.equal(new MailStore(root, { key: k }).list('acc', 'INBOX').total, 13);
});

test('setting, changing and taking away the password seals every file again, and a crash part way carries on at the next open', async () => {
  const root = tmp();
  const plain = new MailStore(root);
  fill(plain, 40, 'Seagull');
  const k1 = key();
  const k2 = key();

  // Set: plain to sealed.
  await plain.prepareReseal(null, k1);
  const set = await plain.rekey(null, k1);
  assert.ok(set.sealed > 40 && set.unopened === 0, JSON.stringify(set));
  assert.deepEqual(clearText(root, 'Seagull'), []);
  assert.ok(!plain.resealPending());

  // Change, stopped part way through: the second folder never reached.
  await plain.prepareReseal(k1, k2);
  let calls = 0;
  await assert.rejects(plain.rekey(k1, k2, { progress: () => { if (++calls === 1) throw new Error('the power went'); } }), /power/);
  assert.ok(plain.resealPending(), 'the run is written down');
  const marker = fs.readFileSync(path.join(root, '.reseal.json'), 'utf8');
  assert.ok(!marker.includes(k1.toString('base64')), 'the old key is never on disk on its own');

  // The next open has the new key only: everything reads, then the run ends.
  const reopened = new MailStore(root, { key: k2 });
  assert.equal(reopened.list('acc', 'INBOX').total, 40, 'readable before the run carries on');
  assert.equal(reopened.list('acc', 'Archive').total, 4);
  const resumed = await reopened.resumeReseal();
  assert.equal(resumed.unopened, 0);
  assert.ok(!reopened.resealPending());
  const fresh = new MailStore(root, { key: k2 });
  assert.equal(fresh.list('acc', 'INBOX').total, 40);
  const any = fresh.list('acc', 'INBOX').rows[0];
  assert.match(fresh.get('acc', 'INBOX', any.id).text, /Seagull/);
  assert.equal(new MailStore(root, { key: k1 }).list('acc', 'INBOX').total, 0, 'nothing is left sealed with the key replaced');

  // Taken away: plain again, as before there was a password.
  await fresh.prepareReseal(k2, null);
  await fresh.rekey(k2, null);
  assert.ok(clearText(root, 'Seagull').length > 0);
  assert.ok(!files(root).some((f) => f.bytes.subarray(0, MAGIC.length).equals(MAGIC) || f.bytes.toString('latin1').startsWith('~')));
  assert.equal(new MailStore(root).list('acc', 'INBOX').total, 40);
});

test('what sealing costs, on a store of a few thousand messages', async (t) => {
  const count = 3000;
  const time = (fn) => { const t0 = process.hrtime.bigint(); fn(); return Number(process.hrtime.bigint() - t0) / 1e6; };
  const rootPlain = tmp();
  const rootSealed = tmp();
  const plain = new MailStore(rootPlain);
  const sealed = new MailStore(rootSealed, { key: key() });
  const putPlain = time(() => fill(plain, count, 'Cost'));
  const putSealed = time(() => fill(sealed, count, 'Cost'));
  const listPlain = time(() => new MailStore(rootPlain).list('acc', 'INBOX'));
  const listSealed = time(() => sealed.list('acc', 'INBOX', { limit: 100 }) && new MailStore(rootSealed, { key: sealed.keys()[0] }).list('acc', 'INBOX'));
  const ids = plain.list('acc', 'INBOX', { limit: 200 }).rows.map((r) => r.id);
  const openPlain = time(() => ids.forEach((id) => plain.get('acc', 'INBOX', id)));
  const openSealed = time(() => ids.forEach((id) => sealed.get('acc', 'INBOX', id)));
  const t0 = Date.now();
  const r = await plain.rekey(null, key());
  const rekeyMs = Date.now() - t0;
  const size = (root) => files(root).reduce((n, f) => n + f.bytes.length, 0);
  t.diagnostic(`${count + count / 10} messages, ${r.files} files: storing ${putPlain.toFixed(0)} ms plain, ${putSealed.toFixed(0)} ms sealed; a 3000-row list read cold ${listPlain.toFixed(0)} ms plain, ${listSealed.toFixed(0)} ms sealed; 200 messages opened ${openPlain.toFixed(0)} ms plain, ${openSealed.toFixed(0)} ms sealed; sealing the whole store ${rekeyMs} ms; on disk ${(size(rootPlain) / 1e6).toFixed(1)} MB sealed (was plain), ${(size(rootSealed) / 1e6).toFixed(1)} MB sealed from the start`);
  assert.equal(r.unopened, 0);
  assert.equal(r.files, files(rootPlain).length);
});

test('the profile vault seals Mail with the rest, and Mail reads and fetches as before', async () => {
  const { createVault, makeLock } = await import('../packages/office-shell/src/profiles.js');
  const { createMailService } = await import('../apps/desktop/main/mail.js');
  const { imapServer } = await import('../packages/mailbox/test/imap-server.js');
  const server = await imapServer();
  const raw = (n) => [`From: Amina Yusuf <amina@work.example>`, 'To: Sam <sam@checks.example>', `Subject: Kestrel report ${n}`, `Date: ${new Date().toUTCString()}`, `Message-ID: <kestrel-${n}@work.example>`, '', `Kestrel body ${n}.`, ''].join('\r\n');
  for (let n = 1; n <= 3; n++) server.deliver('INBOX', raw(n));
  const userData = tmp();
  const settings = new Map();
  const secrets = new Map();
  const storesWith = (vault) => ({
    vault,
    settings: { get: (k, f) => (settings.has(k) ? structuredClone(settings.get(k)) : f), set: (k, v) => settings.set(k, structuredClone(v)), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k), keys: () => [...secrets.keys()] },
  });
  const heard = [];
  const make = (vault) => createMailService({ stores: storesWith(vault), holdBlob: () => ({ url: '' }), broadcast: (e, p) => heard.push([e, p]), userData, autoFetch: false, smtp: { async send() { return { messageId: '<x@here>' }; } } });
  try {
    const vault = createVault();
    const service = make(vault);
    const { id: accountId } = service.addAccount({ account: server.account(), password: 'check' });
    await service.sync({ accountId });
    const mailRoot = path.join(userData, 'mail');
    assert.ok(clearText(mailRoot, 'Kestrel').length > 0, 'plain before a password');

    // A password set: the lock's key seals the mail store too, with progress said.
    const { key } = await makeLock('harbour-7');
    await vault.prepare(key);
    await vault.rekey(key);
    assert.deepEqual(clearText(mailRoot, 'Kestrel'), []);
    assert.ok(heard.some(([e, p]) => e === 'mail:sealing' && p.total > 0 && p.done === p.total));
    assert.equal(service.messages({ accountId, folder: 'INBOX', limit: 10 }).rows.length, 3);
    // Fetching and marking go on, sealed.
    server.deliver('INBOX', raw(4));
    await service.sync({ accountId });
    assert.equal(service.messages({ accountId, folder: 'INBOX', limit: 10 }).rows.length, 4);
    assert.deepEqual(clearText(mailRoot, 'Kestrel'), []);

    // The next start opens it with the key the password makes; a start without it sees nothing.
    const again = make(createVault({ key }));
    assert.equal(again.messages({ accountId, folder: 'INBOX', limit: 10 }).rows.length, 4);
    const locked = make(createVault());
    assert.equal(locked.messages({ accountId, folder: 'INBOX', limit: 10 }).rows.length, 0);

    // Taken away: plain again.
    await vault.prepare(null);
    await vault.rekey(null);
    assert.ok(clearText(mailRoot, 'Kestrel').length > 0);
    assert.equal(make(createVault()).messages({ accountId, folder: 'INBOX', limit: 10 }).rows.length, 4);
  } finally {
    await server.close();
  }
});
