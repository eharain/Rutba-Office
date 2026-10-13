// Export for legal hold: a search's results across an account, a .pst and an
// .olm, written to an mbox, PDFs and a manifest; read back and checked.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { MailStore } from '@rutba/mailbox/store';
import { parseMessage } from '@rutba/mailbox/mime';
import { Mbox } from '@rutba/mailbox/mbox';
import { openArchive, indexArchive } from '@rutba/mailbox/archive';
import { exportMessages, storeLoader, archiveLoader, freshFolder } from '../packages/mailbox/src/legal-export.js';
import { toRfc822 } from '../packages/mailbox/src/rfc822.js';
import { writeArchiveFixtures, ACCOUNT_MESSAGES, PST_MESSAGES, rfc822 } from './fixtures/mail-archives.js';

const sha = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex');
const pdfPages = (bytes) => (Buffer.from(bytes).toString('latin1').match(/\/Type\s*\/Page(?!s)/g) || []).length;

function setup() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-hold-'));
  const files = writeArchiveFixtures(path.join(dir, 'archives'));
  const store = new MailStore(path.join(dir, 'store'));
  store.putMany('acct', 'Inbox', ACCOUNT_MESSAGES.map((m) => parseMessage(rfc822(m))));
  store.setFolders('acct', [{ path: 'Inbox', name: 'Inbox' }]);
  store.syncIndex('acct', { name: 'ejaz@example.com' });
  const sources = {
    acct: { kind: 'account', name: 'ejaz@example.com', load: storeLoader(store) },
    'archive:pst': { kind: 'pst', name: 'Northwind 2019', file: files.pst, archive: openArchive(files.pst) },
    'archive:olm': { kind: 'olm', name: 'Contoso archive', file: files.olm, archive: openArchive(files.olm) },
  };
  for (const [id, s] of Object.entries(sources)) {
    if (s.archive) {
      indexArchive(s.archive, store.index, { id, name: s.name });
      s.load = archiveLoader(s.archive);
    }
  }
  const load = (item) => sources[item.sourceId].load(item);
  const itemsOf = (rows) => rows.map((r) => ({ sourceId: r.accountId, sourceName: r.sourceName, sourceKind: r.sourceKind, sourceFile: sources[r.accountId].file || null, folder: r.folder, id: r.id, date: r.date }));
  return { dir, files, store, sources, load, itemsOf, done: () => fs.rmSync(dir, { recursive: true, force: true }) };
}

test("a search's results export to an mbox that reads back with the same messages and hashes", async () => {
  const s = setup();
  try {
    const before = Object.values(s.files).map((f) => sha(fs.readFileSync(f)));
    const { rows } = s.store.index.search('halvorsen after:2019-01-01', { sources: Object.keys(s.sources) });
    assert.equal(rows.length, 5);
    const progress = [];
    const out = freshFolder(s.dir, 'Export');
    const result = await exportMessages({
      items: s.itemsOf(rows),
      load: s.load,
      dir: out,
      mbox: true,
      pdf: 'both',
      attachments: true,
      selection: 'halvorsen after:2019-01-01',
      onProgress: (p) => progress.push(p.done),
    });
    assert.equal(result.exported, 5);
    assert.equal(result.cancelled, false);
    assert.equal(result.sources, 3);
    assert.deepEqual(progress, [1, 2, 3, 4, 5]);

    const manifest = JSON.parse(fs.readFileSync(path.join(out, 'manifest.json'), 'utf8'));
    assert.equal(manifest.complete, true);
    assert.equal(manifest.messages.length, 5);

    // The mbox: five messages, their Message-IDs, and each one's hash.
    const mbox = new Mbox(fs.readFileSync(path.join(out, 'messages.mbox')));
    assert.equal(mbox.length, 5);
    for (let i = 0; i < mbox.length; i++) {
      const raw = Buffer.from(mbox.raw(i));
      const exported = raw.subarray(0, raw.length - 1); // less the line feed that ends it
      assert.equal(sha(exported), manifest.messages[i].sha256, `message ${i + 1}'s hash`);
      assert.equal(parseMessage(exported).messageId, manifest.messages[i].messageId.replace(/^<|>$/g, ''));
    }
    assert.deepEqual(
      manifest.messages.map((m) => m.messageId).sort(),
      ['<halvorsen-1@northwind.example>', '<halvorsen-2@northwind.example>', '<halvorsen-next@northwind.example>', '<lease-renewal@contoso.example>', '<site-visit@contoso.example>']
    );
    // A .pst message comes back whole: its attachment's bytes are the ones written.
    const signed = parseMessage(mbox.raw(manifest.messages.findIndex((m) => m.messageId === '<halvorsen-1@northwind.example>')));
    assert.equal(signed.subject, 'Halvorsen contract: signed copy');
    assert.equal(signed.cc[0].address, 'bo@northwind.example');
    assert.deepEqual(Buffer.from(signed.attachments[0].bytes), Buffer.from(PST_MESSAGES.Inbox[0].attachments[0].bytes));

    // The PDFs: one a message, and one of all after a summary page.
    const combined = fs.readFileSync(path.join(out, 'messages.pdf'));
    assert.ok(pdfPages(combined) >= 6, `${pdfPages(combined)} pages`);
    assert.match(combined.toString('latin1'), /\(Halvorsen contract: signed copy\) Tj/);
    assert.match(combined.toString('latin1'), /\(Export summary\) Tj/);
    for (const m of manifest.messages) {
      const bytes = fs.readFileSync(path.join(out, m.pdf.file));
      assert.equal(sha(bytes), m.pdf.sha256);
      assert.equal(pdfPages(bytes), m.pdf.pages);
      assert.ok(bytes.toString('latin1').includes(`(${m.subject.replace(/[()\\]/g, '\\$&')}) Tj`), `${m.pdf.file} shows its subject`);
    }
    // Attachments beside them, with their hashes.
    const att = manifest.messages.find((m) => m.attachments.length).attachments[0];
    assert.equal(att.name, 'Halvorsen contract signed.pdf');
    assert.equal(sha(fs.readFileSync(path.join(out, att.file))), att.sha256);
    // Every file the manifest names has the hash it says.
    for (const f of manifest.files) assert.equal(sha(fs.readFileSync(path.join(out, f.name))), f.sha256, f.name);
    const csvBytes = fs.readFileSync(path.join(out, 'manifest.csv'));
    assert.deepEqual([...csvBytes.subarray(0, 3)], [0xef, 0xbb, 0xbf], 'a byte-order mark, for spreadsheets');
    const csv = csvBytes.subarray(3).toString('utf8').trim().split(String.fromCharCode(13, 10));
    assert.equal(csv.length, 6);
    assert.match(csv[0], /^n,date,from,to,cc,subject,message_id,source/);
    assert.ok(fs.existsSync(path.join(out, 'summary.pdf')));

    // The sources are as they were.
    assert.deepEqual(Object.values(s.files).map((f) => sha(fs.readFileSync(f))), before);
  } finally {
    s.done();
  }
});

test('an export stopped part way keeps what it wrote and says so', async () => {
  const s = setup();
  try {
    const { rows } = s.store.index.search('halvorsen', { sources: Object.keys(s.sources) });
    let done = 0;
    const out = freshFolder(s.dir, 'Export');
    const result = await exportMessages({ items: s.itemsOf(rows), load: s.load, dir: out, pdf: 'combined', onProgress: (p) => (done = p.done), cancelled: () => done >= 2 });
    assert.equal(result.cancelled, true);
    assert.equal(result.exported, 2);
    const manifest = JSON.parse(fs.readFileSync(path.join(out, 'manifest.json'), 'utf8'));
    assert.equal(manifest.complete, false);
    assert.equal(manifest.requested, 6);
    assert.equal(new Mbox(fs.readFileSync(path.join(out, 'messages.mbox'))).length, 2);
    // A second export beside it does not write over the first.
    assert.notEqual(freshFolder(s.dir, 'Export'), out);
  } finally {
    s.done();
  }
});

test('a whole folder of an archive exports, and an mbox message goes out as the archive holds it', async () => {
  const s = setup();
  try {
    const mbox = openArchive(s.files.mbox);
    const folder = mbox.folders()[0].path;
    const items = mbox.list(folder, { limit: 100 }).rows.map((r) => ({ sourceId: 'mbox', sourceName: 'Account export', folder, id: r.id, date: r.date }));
    const out = freshFolder(s.dir, 'Folder');
    const result = await exportMessages({ items, load: archiveLoader(mbox), dir: out, pdf: null });
    assert.equal(result.exported, 2);
    const back = new Mbox(fs.readFileSync(path.join(out, 'messages.mbox')));
    for (let i = 0; i < back.length; i++) {
      const original = Buffer.from(mbox.raw(folder, items[i].id));
      assert.deepEqual(Buffer.from(back.raw(i)).subarray(0, -1), original, 'byte for byte');
    }
    assert.ok(!fs.existsSync(path.join(out, 'messages.pdf')));
  } finally {
    s.done();
  }
});

test('a message is written back as RFC 5322 that parses to the same thing, words in any script included', () => {
  const m = {
    subject: 'رپورٹ: Q3 - نتائج',
    from: [{ name: 'Amina Rahman', address: 'amina@example.com' }],
    to: [{ name: 'عائشہ', address: 'aisha@example.com' }],
    cc: [],
    date: '2024-03-11T09:14:00.000Z',
    messageId: 'abc@example.com',
    text: 'سلام\nThe figures are attached.\nFrom the team',
    html: '<p>سلام</p>',
    attachments: [{ filename: 'نتائج.xlsx', type: 'application/octet-stream', bytes: Buffer.from([1, 2, 3, 250]) }],
  };
  const bytes = toRfc822(m);
  assert.ok(/^[\x00-\x7f]*$/.test(bytes.toString('latin1')), 'the bytes are plain ASCII');
  assert.deepEqual(toRfc822(m), bytes, 'the same message gives the same bytes');
  const back = parseMessage(bytes);
  assert.equal(back.subject, m.subject);
  assert.equal(back.to[0].name, 'عائشہ');
  assert.equal(back.messageId, 'abc@example.com');
  assert.equal(back.date, m.date);
  assert.match(back.text, /The figures are attached/);
  assert.match(back.html, /سلام/);
  assert.equal(back.attachments[0].filename, 'نتائج.xlsx');
  assert.deepEqual(Buffer.from(back.attachments[0].bytes), Buffer.from([1, 2, 3, 250]));
});
