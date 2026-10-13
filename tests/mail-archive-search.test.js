// One search across every account and every archive opened in place.
//
// The archives are written here by tests/fixtures (a .pst by the suite's own
// writer, an .olm as a zip of XML, an mbox); the account is a local store.
// Each search in EXPECT must find exactly the messages it names, wherever
// they live, with the place each one lives named in the result.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { MailStore } from '@rutba/mailbox/store';
import { parseMessage } from '@rutba/mailbox/mime';
import { PstFile } from '@rutba/mailbox/pst';
import { fileReader } from '@rutba/mailbox/import';
import { openArchive, indexArchive } from '@rutba/mailbox/archive';
import { parseSearch, parseSize, parseDay } from '@rutba/mailbox/query';
import { SearchIndex } from '@rutba/mailbox/search';
import { writeArchiveFixtures, ACCOUNT_MESSAGES, PST_MESSAGES, EXPECT, rfc822 } from './fixtures/mail-archives.js';

const sha = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');

function setup() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-archive-'));
  const files = writeArchiveFixtures(path.join(dir, 'archives'));
  const store = new MailStore(path.join(dir, 'store'));
  store.putMany('acct', 'Inbox', ACCOUNT_MESSAGES.map((m) => parseMessage(rfc822(m))));
  store.setFolders('acct', [{ path: 'Inbox', name: 'Inbox', role: 'inbox' }]);
  const pst = openArchive(files.pst);
  const olm = openArchive(files.olm);
  indexArchive(pst, store.index, { id: 'archive:pst', name: 'Northwind 2019' });
  indexArchive(olm, store.index, { id: 'archive:olm', name: 'Contoso archive' });
  store.syncIndex('acct', { name: 'ejaz@example.com' });
  const all = ['acct', 'archive:pst', 'archive:olm'];
  return { dir, files, store, pst, olm, all, done: () => fs.rmSync(dir, { recursive: true, force: true }) };
}

test('the .pst writer makes a file the reader reads back whole', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-pstw-'));
  try {
    const { pst: file } = writeArchiveFixtures(dir);
    const pst = new PstFile(fileReader(file));
    try {
      const summary = pst.summary();
      assert.equal(summary.messages, 6);
      assert.equal(summary.store, 'Northwind 2019');
      const inbox = summary.folders.find((f) => f.name === 'Inbox');
      assert.equal(inbox.messages, 4);
      const rows = pst.messages(inbox.nid, { limit: 10 });
      assert.deepEqual(rows.map((r) => r.subject).sort(), PST_MESSAGES.Inbox.map((m) => m.subject).sort());
      const first = rows.find((r) => r.subject === 'Halvorsen contract: signed copy');
      const m = pst.message(first.nid);
      assert.equal(m.cc[0].address, 'bo@northwind.example');
      assert.match(m.text, /indemnity clause on page four/);
      assert.equal(m.attachments[0].filename, 'Halvorsen contract signed.pdf');
      assert.deepEqual(Buffer.from(m.attachments[0].bytes), Buffer.from(PST_MESSAGES.Inbox[0].attachments[0].bytes));
      const big = pst.message(rows.find((r) => r.subject === 'Quarterly figures').nid);
      assert.equal(big.attachments[0].bytes.length, 1_300_000, 'an attachment over a thousand blocks comes back whole');
      const light = pst.message(first.nid, { withBytes: false });
      assert.equal(light.attachments[0].bytes, null, 'a light read leaves the file unread');
      assert.equal(light.attachments[0].size, 48_000);
    } finally {
      pst.close();
    }
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('an archive opens in place: its folders, a list newest first, a message', () => {
  const s = setup();
  try {
    assert.deepEqual(s.pst.folders().map((f) => f.path), ['Inbox', 'Inbox/Projects', 'Sent Items']);
    assert.equal(s.pst.kind, 'pst');
    const list = s.pst.list('Inbox', { limit: 2 });
    assert.equal(list.total, 4);
    assert.deepEqual(list.rows.map((r) => r.subject), ['Quarterly figures', 'Mailbox quota warning']);
    const m = s.pst.message('Inbox', list.rows[0].id);
    assert.equal(m.attachments[0].filename, 'Q3 figures.xlsx');

    assert.deepEqual(s.olm.folders().map((f) => [f.path, f.total]).sort(), [['Archive', 1], ['Inbox', 2]]);
    const olmList = s.olm.list('Inbox');
    assert.equal(olmList.rows[0].subject, 'Invoice 4471');
    assert.equal(s.olm.message('Inbox', olmList.rows[0].id).attachments[0].bytes.length, 21_000);

    const mbox = openArchive(s.files.mbox);
    assert.equal(mbox.folders()[0].total, 2);
    const raw = Buffer.from(mbox.raw(mbox.folders()[0].path, '0')).toString('utf8');
    assert.match(raw, /Message-ID: <halvorsen-next@northwind\.example>/, 'an mbox hands back the message as it holds it');
  } finally {
    s.done();
  }
});

test('one search finds every match across the account, the .pst and the .olm', () => {
  const s = setup();
  try {
    for (const [query, subjects] of Object.entries(EXPECT)) {
      const found = s.store.index.search(query, { sources: s.all });
      assert.deepEqual(found.rows.map((r) => r.subject).sort(), subjects.slice().sort(), `the search ${query}`);
      assert.equal(found.total, subjects.length, `the count for ${query}`);
      assert.ok(found.exact, `${query} is counted exactly`);
    }
  } finally {
    s.done();
  }
});

test('a result says where it lives, and opens there', () => {
  const s = setup();
  try {
    const { rows } = s.store.index.search('halvorsen', { sources: s.all });
    const where = Object.fromEntries(rows.map((r) => [r.subject, `${r.sourceName} | ${r.folder} | ${r.sourceKind}`]));
    assert.equal(where['Warehouse lease renewal'], 'Northwind 2019 | Inbox/Projects | pst');
    assert.equal(where['Old lease papers'], 'Contoso archive | Archive | olm');
    assert.equal(where['Halvorsen: next steps'], 'ejaz@example.com | Inbox | account');
    // Newest first, across all three.
    assert.deepEqual(rows.map((r) => r.date.slice(0, 4)), ['2021', '2020', '2019', '2019', '2019', '2018']);
    const lease = rows.find((r) => r.subject === 'Warehouse lease renewal');
    assert.match(s.pst.message(lease.folder, lease.id).text, /Leeds warehouse/);
    const acct = rows.find((r) => r.sourceKind === 'account');
    assert.match(s.store.get(acct.accountId, acct.folder, acct.id).text, /revised draft/);
  } finally {
    s.done();
  }
});

test('bodies are indexed: a word found only in a body is found', () => {
  const s = setup();
  try {
    const subjects = (q) => s.store.index.search(q, { sources: s.all }).rows.map((r) => r.subject).sort();
    assert.deepEqual(subjects('solicitor'), ['Halvorsen: next steps'], 'the account');
    assert.deepEqual(subjects('leeds'), ['Warehouse lease renewal'], 'the .pst');
    assert.deepEqual(subjects('"floor plans"'), ['Halvorsen site visit'], 'the .olm, as a phrase');
    assert.deepEqual(subjects('landlo'), ['Warehouse lease renewal'], 'a word is found by its beginning');
    // A phrase whose words are all in a message but not together is not found.
    assert.deepEqual(subjects('"clause indemnity"'), []);
  } finally {
    s.done();
  }
});

test("the account's index follows its folders: a new message, a moved one, a removed folder", () => {
  const s = setup();
  try {
    const before = s.store.index.search('greenhouse', { sources: ['acct'] });
    assert.equal(before.total, 0);
    s.store.put('acct', 'Inbox', parseMessage(rfc822({ ...ACCOUNT_MESSAGES[1], subject: 'Plants', body: 'The greenhouse quote came in.', messageId: 'plants@x.example' })));
    assert.deepEqual(s.store.search('acct', 'greenhouse').map((r) => r.subject), ['Plants']);

    // Only the new message is read again when its folder changes.
    s.store.put('acct', 'Inbox', parseMessage(rfc822({ ...ACCOUNT_MESSAGES[1], subject: 'More plants', body: 'Another greenhouse.', messageId: 'plants2@x.example' })));
    const { read } = s.store.syncIndex('acct');
    assert.equal(read, 1);

    s.store.putMany('acct', 'Reading', [parseMessage(rfc822({ ...ACCOUNT_MESSAGES[1], subject: 'Filed away', body: 'Pricing notes.', messageId: 'filed@x.example' }))]);
    s.store.setFolders('acct', [...s.store.folders('acct'), { path: 'Reading', name: 'Reading' }]);
    const filed = s.store.search('acct', 'subject:filed');
    assert.equal(filed[0].folder, 'Reading');

    s.store.setFolders('acct', s.store.folders('acct').filter((f) => f.path !== 'Reading'));
    assert.equal(s.store.search('acct', 'pricing').length, 0, 'a folder that is gone takes its messages out of the index');
    // The archives are untouched by any of it.
    assert.equal(s.store.index.search('halvorsen', { sources: ['archive:pst', 'archive:olm'] }).total, 5);
  } finally {
    s.done();
  }
});

test('opening, listing, indexing and reading an archive never changes it', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-archive-ro-'));
  try {
    const files = writeArchiveFixtures(dir);
    const before = Object.fromEntries(Object.entries(files).map(([k, f]) => [k, [sha(f), fs.statSync(f).mtimeMs]]));
    const index = new SearchIndex();
    for (const [k, f] of Object.entries(files)) {
      const a = openArchive(f);
      indexArchive(a, index, { id: k, name: k });
      for (const folder of a.folders()) for (const r of a.list(folder.path).rows) a.message(folder.path, r.id);
      a.close();
    }
    for (const [k, f] of Object.entries(files)) assert.deepEqual([sha(f), fs.statSync(f).mtimeMs], before[k], `${k} is as it was`);
    assert.deepEqual(fs.readdirSync(dir).sort(), Object.values(files).map((f) => path.basename(f)).sort(), 'nothing is written beside it');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('the search language: sizes, dates, ranges, OR and a minus', () => {
  assert.equal(parseSize('10mb'), 10 * 1024 * 1024);
  assert.equal(parseSize('500K'), 500 * 1024);
  assert.equal(parseSize('2.5 MB'), Math.round(2.5 * 1024 * 1024));
  assert.equal(parseSize('lots'), null);
  const day = parseDay('2024-03-11');
  assert.equal(day.to - day.from, 86_400_000);
  assert.equal(new Date(parseDay('11/03/2024').from).getMonth(), 2, 'day first');
  assert.equal(new Date(parseDay('2024').to).getFullYear(), 2025);

  const atoms = (q) => parseSearch(q, { now: new Date(2026, 9, 13).getTime() }).clauses.map((c) => ({ ...c.any[0], negate: c.negate, n: c.any.length }));
  const [after] = atoms('after:2024-01-01');
  assert.equal(after.kind, 'date');
  assert.equal(after.from, new Date(2024, 0, 1).getTime());
  assert.equal(after.to, null);
  const [range] = atoms('received:2024-01-01..2024-01-31');
  assert.equal(range.to, new Date(2024, 1, 1).getTime());
  const [cmp] = atoms('received:>=2024-02-01');
  assert.equal(cmp.from, new Date(2024, 1, 1).getTime());
  const [older] = atoms('older_than:1y');
  assert.equal(older.to, new Date(2025, 9, 13).getTime());
  assert.deepEqual(atoms('size>1mb')[0], { kind: 'size', min: 1024 * 1024 + 1, max: null, negate: false, n: 1 });
  assert.deepEqual(atoms('larger:5M')[0].min, 5 * 1024 * 1024);
  assert.deepEqual(atoms('smaller:1mb')[0].max, 1024 * 1024 - 1);
  assert.deepEqual(atoms('size:<=10k')[0].max, 10 * 1024);
  const [neg] = atoms('-from:bob');
  assert.equal(neg.negate, true);
  assert.equal(neg.field, 'from');
  const or = parseSearch('invoice OR receipt from:dana').clauses;
  assert.equal(or.length, 2);
  assert.equal(or[0].any.length, 2);
  assert.equal(atoms('hasattachments:yes')[0].kind, 'has');
  assert.equal(atoms('hasattachments:no')[0].negate, true);
  assert.equal(atoms('attachment:report')[0].field, 'filename');
  assert.equal(atoms('in:inbox')[0].field, 'folder');
  assert.equal(atoms('"quarterly report"')[0].kind, 'phrase');
  // Anything that does not parse is searched for as words.
  assert.equal(atoms('http://example.com/x')[0].kind, 'phrase');
  assert.equal(atoms('after:someday')[0], undefined);
  assert.ok(parseSearch('').empty);
});
