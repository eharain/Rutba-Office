// Mbox and Outlook for Mac archives read from their files, never whole: the
// same messages as reading the bytes, From_ lines found across the blocks
// the file is read in, zip64 directories read, the file never read whole,
// and an archive opened in place listing, reading and indexing from its file.
// The archives past 2 GB are the performance budget's (mail-big-archives.perf.js)
// and the window check's (mail-big-mbox): here they are small.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Mbox, MboxFile, openMbox } from '@rutba/mailbox/mbox';
import { OlmArchive } from '@rutba/mailbox/olm';
import { ZipFile } from '@rutba/mailbox/zip-file';
import { readZip } from '@rutba/ooxml/zip';
import { openArchive, indexArchive } from '@rutba/mailbox/archive';
import { identify } from '@rutba/mailbox/import';
import { SearchIndex } from '@rutba/mailbox/search';
import { writeArchiveFixtures } from './fixtures/mail-archives.js';
import { writeCorpusMbox, writeCorpusOlm, LAST } from './fixtures/big-archives.js';

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-big-archives-'));

test('an mbox read from its file finds the messages reading its bytes finds, whatever block it is read in', () => {
  const dir = tmp();
  try {
    const file = path.join(dir, 'corpus.mbox');
    writeCorpusMbox(file, { count: 40, attachBytes: 3000, attachEvery: 3 });
    const bytes = fs.readFileSync(file);
    const whole = new Mbox(bytes);
    // Blocks of a few bytes put From_ lines across every kind of join.
    for (const block of [5, 6, 7, 64, 4096, 8 * 1048576]) {
      const m = new MboxFile(file, { block });
      try {
        assert.equal(m.length, whole.length, `block ${block}`);
        for (const i of [0, 1, 17, m.length - 1]) assert.deepEqual(Buffer.from(m.raw(i)), Buffer.from(whole.raw(i)), `block ${block}, message ${i}`);
      } finally {
        m.close();
      }
    }
    const m = openMbox(file);
    try {
      assert.equal(m.length, 41);
      const last = m.message(m.length - 1);
      assert.equal(last.subject, LAST.subject);
      assert.match(last.text, /quartzberry/);
      assert.equal(last.attachments[0].filename, 'harbour-lease-signed.pdf');
      const head = m.headers(m.length - 1);
      assert.equal(head.subject, LAST.subject);
      assert.equal(head.from[0].address, LAST.from.address);
      assert.equal(head.hasAttachments, true);
      assert.equal(head.text, null, 'the headers alone, not the body');
      assert.equal(m.size(0), m.range(1).start);
    } finally {
      m.close();
    }
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('an mbox read from its file undoes the From_ quoting and keeps what came before the first message out', () => {
  const dir = tmp();
  try {
    const file = path.join(dir, 'quoted.mbox');
    const text = 'stray line before any message\nFrom a@example.com Mon Jan  1 00:00:00 2024\nSubject: one\n\n>From the start of a line\nplain\n\nFrom b@example.com Mon Jan  1 00:00:00 2024\nSubject: two\n\nbody\n';
    fs.writeFileSync(file, text);
    const m = new MboxFile(file, { block: 9 });
    try {
      assert.equal(m.length, 2);
      assert.equal(m.message(0).subject, 'one');
      assert.match(m.message(0).text, /^From the start of a line/m);
      assert.equal(m.message(1).subject, 'two');
      assert.deepEqual(Buffer.from(m.raw(1)), Buffer.from(new Mbox(fs.readFileSync(file)).raw(1)));
    } finally {
      m.close();
    }
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a zip read from its file reads the directory and entries the whole file gives, zip64 records included', () => {
  const dir = tmp();
  try {
    const files = writeArchiveFixtures(path.join(dir, 'a'));
    const zip = new ZipFile(files.olm);
    try {
      const whole = readZip(fs.readFileSync(files.olm)).entries;
      assert.deepEqual(zip.entries.map((e) => e.name), whole.map((e) => e.name));
      zip.entries.forEach((e, i) => assert.deepEqual(zip.data(e), whole[i].data, e.name));
    } finally {
      zip.close();
    }
    for (const zip64 of [false, true]) {
      const file = path.join(dir, `corpus-${zip64}.olm`);
      writeCorpusOlm(file, { count: 30, attachBytes: 2000, attachEvery: 2, zip64 });
      const z = new ZipFile(file);
      try {
        assert.equal(z.entries.length, 31);
        assert.ok(z.entries.some((e) => e.method === 0) && z.entries.some((e) => e.method === 8), 'stored and deflated entries both read');
        assert.match(z.data(z.entries[30]).toString('utf8'), /quartzberry/);
      } finally {
        z.close();
      }
      const olm = OlmArchive.open(file);
      const bytes = new OlmArchive(fs.readFileSync(file));
      try {
        assert.deepEqual([...olm.messages()].map((m) => m.messageId), [...bytes.messages()].map((m) => m.messageId), `zip64 ${zip64}`);
      } finally {
        olm.close();
      }
    }
    const cut = path.join(dir, 'cut.olm');
    fs.writeFileSync(cut, fs.readFileSync(path.join(dir, 'corpus-true.olm')).subarray(0, 4000));
    assert.throws(() => new ZipFile(cut), /not a zip archive|not a complete zip archive/);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('an mbox or .olm opened in place is never read whole: listed, read and indexed from its file', () => {
  const dir = tmp();
  const own = fs.readFileSync;
  try {
    const mboxFile = path.join(dir, 'Big export.mbox');
    const olmFile = path.join(dir, 'Big archive.olm');
    writeCorpusMbox(mboxFile, { count: 60, attachBytes: 4000 });
    writeCorpusOlm(olmFile, { count: 60, attachBytes: 4000 });
    // A file other than these may be read whole; these two never are.
    const read = [];
    fs.readFileSync = function (p, ...rest) {
      if (typeof p === 'string' && (path.resolve(p) === mboxFile || path.resolve(p) === olmFile)) read.push(p);
      return own.call(this, p, ...rest);
    };
    for (const [file, kind] of [[mboxFile, 'mbox'], [olmFile, 'olm']]) {
      assert.equal(identify(file).kind, kind);
      const a = openArchive(file);
      try {
        const folders = a.folders();
        assert.equal(folders.reduce((n, f) => n + f.total, 0), 61, kind);
        const inbox = folders.find((f) => f.path === 'Inbox') || folders[0];
        const page = a.list(inbox.path, { limit: 500 });
        const row = page.rows.find((r) => r.subject === LAST.subject);
        assert.ok(row, `${kind}: the last message is listed`);
        const m = a.message(inbox.path, row.id);
        assert.match(m.text, /quartzberry/);
        assert.equal(m.attachments[0].bytes.length, LAST.attachments[0].bytes.length);
        const index = new SearchIndex();
        indexArchive(a, index, { id: kind, name: kind });
        const found = index.search('quartzberry');
        assert.equal(found.total, 1, kind);
        assert.equal(found.rows[0].subject, LAST.subject);
        assert.equal(index.search('"quartzberry clause"').total, 1, `${kind}: a phrase is confirmed against the body read from the file`);
      } finally {
        a.close();
      }
    }
    // An .olm named otherwise is still known by its zip, read from its directory.
    const renamed = path.join(dir, 'archive.bin');
    fs.copyFileSync(olmFile, renamed);
    assert.equal(identify(renamed).kind, 'olm');
    assert.deepEqual(read, []);
  } finally {
    fs.readFileSync = own;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
