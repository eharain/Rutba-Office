// The mail service with archives opened in place: an archive reads like an
// account, is searched with the accounts, cannot be changed, and exports.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { createMailService } from '../apps/desktop/main/mail.js';
import { writeArchiveFixtures } from './fixtures/mail-archives.js';

const sha = (f) => crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex');

function service() {
  const settings = new Map();
  const stores = {
    settings: { get: (k, f) => (settings.has(k) ? settings.get(k) : f), set: (k, v) => settings.set(k, v), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: () => null, set: () => {}, delete: () => {}, keys: () => [] },
  };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-desk-'));
  const events = [];
  const mail = createMailService({ stores, holdBlob: (bytes) => ({ url: 'blob:x', bytes }), broadcast: (name, data) => events.push([name, data]), userData: path.join(dir, 'data'), smtp: { send: async () => ({ messageId: 'x' }) } });
  return { mail, dir, events, settings, done: () => fs.rmSync(dir, { recursive: true, force: true }) };
}

const indexed = async (mail, id) => {
  for (let i = 0; i < 200; i++) {
    if (mail.archives().find((a) => a.id === id)?.index.state === 'ready') return;
    await new Promise((r) => setTimeout(r, 10));
  }
  throw new Error('the archive was never indexed');
};

test('an archive opened in place reads like an account and cannot be changed', async () => {
  const s = service();
  try {
    const files = writeArchiveFixtures(path.join(s.dir, 'archives'));
    const before = sha(files.pst);
    const pst = await s.mail.openArchive({ path: files.pst });
    assert.match(pst.id, /^archive:/);
    assert.equal(pst.kind, 'pst');
    assert.equal(pst.messages, 6);
    assert.equal((await s.mail.openArchive({ path: files.pst })).id, pst.id, 'the same file twice is one archive');
    assert.deepEqual((await s.mail.folders({ accountId: pst.id })).map((f) => f.path).sort(), ['Inbox', 'Inbox/Projects', 'Sent Items']);
    const list = await s.mail.messages({ accountId: pst.id, folder: 'Inbox', limit: 10 });
    assert.equal(list.total, 4);
    const signed = list.rows.find((r) => r.subject === 'Halvorsen contract: signed copy');
    const m = await s.mail.message({ accountId: pst.id, folder: 'Inbox', id: signed.id });
    assert.equal(m.readOnly, true);
    assert.match(m.text, /indemnity clause/);
    assert.equal(m.attachments[0].bytes, undefined, 'the bytes stay in the main process');
    assert.equal((await s.mail.attachment({ accountId: pst.id, folder: 'Inbox', id: signed.id, index: 0 })).size, 48_000);
    // Read, not written: marking read does nothing, a move or delete is refused.
    s.mail.flag({ accountId: pst.id, folder: 'Inbox', ids: [signed.id], patch: { unread: false } });
    assert.throws(() => s.mail.move({ accountId: pst.id, folder: 'Inbox', ids: [signed.id], to: 'Archive' }), /read only/);
    assert.throws(() => s.mail.delete({ accountId: pst.id, folder: 'Inbox', ids: [signed.id] }), /read only/);
    await indexed(s.mail, pst.id);
    assert.equal(sha(files.pst), before);
    await s.mail.closeArchive({ id: pst.id });
    assert.equal(s.mail.archives().length, 0);
  } finally {
    s.done();
  }
});

test('searchAll finds across the accounts and the archives, and exportHold writes the hold', async () => {
  const s = service();
  try {
    const files = writeArchiveFixtures(path.join(s.dir, 'archives'));
    s.mail.addAccount({ account: { email: 'ejaz@example.com', name: 'Ejaz', local: true } });
    const a = await s.mail.openArchive({ path: files.pst });
    const b = await s.mail.openArchive({ path: files.olm });
    await indexed(s.mail, a.id);
    await indexed(s.mail, b.id);
    const found = await s.mail.searchAll({ query: 'halvorsen after:2019-01-01 before:2021-01-01' });
    assert.equal(found.total, 4);
    assert.deepEqual([...new Set(found.rows.map((r) => r.sourceName))].sort(), ['Contoso archive', 'Northwind 2019']);
    assert.ok(found.rows.every((r) => r.readOnly && r.sourceFile));
    assert.deepEqual(found.indexing, []);

    const out = path.join(s.dir, 'out');
    fs.mkdirSync(out);
    const result = await s.mail.exportHold({ job: 'j1', query: 'halvorsen after:2019-01-01 before:2021-01-01', dir: out, mbox: true, pdf: 'combined', selection: 'halvorsen' });
    assert.equal(result.exported, 4);
    assert.equal(result.sources, 2);
    assert.ok(fs.existsSync(path.join(result.dir, 'messages.mbox')));
    assert.ok(s.events.some(([name, d]) => name === 'mail:exportProgress' && d.done === 4 && d.total === 4));

    const byFolder = await s.mail.exportHold({ folders: [{ accountId: a.id, folder: 'Inbox' }, { accountId: a.id, folder: 'Sent Items' }], dir: out, mbox: true, pdf: null });
    assert.equal(byFolder.exported, 5);
    assert.notEqual(byFolder.dir, result.dir, 'a second export goes to a folder of its own');
  } finally {
    s.done();
  }
});
