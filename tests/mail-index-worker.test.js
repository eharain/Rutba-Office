// Archives indexed and searched on threads of their own: the answers are the
// ones the index gives on the calling thread, a search while indexing goes on
// answers from what is indexed so far and says how far that is, indexing
// pauses and carries on, and closing an archive stops its thread.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createArchiveHost } from '@rutba/mailbox/archive-host';
import { openArchive, indexArchive, indexSteps } from '@rutba/mailbox/archive';
import { SearchIndex } from '@rutba/mailbox/search';
import { writeArchiveFixtures } from './fixtures/mail-archives.js';
import { writeCorpusPst } from './fixtures/pst-writer.js';

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-index-worker-'));
const settled = (events, id, state, ms = 20000) =>
  new Promise((resolve, reject) => {
    const t0 = Date.now();
    const look = () => {
      const last = events.filter(([k]) => k === id).pop()?.[1];
      if (last?.state === state) return resolve(last);
      if (Date.now() - t0 > ms) return reject(new Error(`never ${state}: ${JSON.stringify(last)}`));
      setTimeout(look, 10);
    };
    look();
  });

test('an archive indexed on its thread answers as the index on this thread does', async () => {
  const dir = tmp();
  const events = [];
  const host = createArchiveHost({ onProgress: (id, p) => events.push([id, p]) });
  try {
    const files = writeArchiveFixtures(path.join(dir, 'a'));
    const opened = await host.open('archive:pst', files.pst, 'Northwind 2019');
    assert.equal(opened.kind, 'pst');
    assert.deepEqual(opened.folders.map((f) => f.path).sort(), ['Inbox', 'Inbox/Projects', 'Sent Items']);
    await host.call('archive:pst', 'index');
    const ready = await settled(events, 'archive:pst', 'ready');
    assert.equal(ready.done, ready.total);

    const here = new SearchIndex();
    indexArchive(openArchive(files.pst), here, { id: 'archive:pst', name: 'Northwind 2019' });
    for (const q of ['halvorsen', '"indemnity clause"', 'from:northwind', 'has:attachment', 'halvorsen -signed']) {
      const there = await host.call('archive:pst', 'search', { query: q, limit: 50 });
      const own = here.search(q, { limit: 50 });
      assert.deepEqual(there.rows.map((r) => r.id), own.rows.map((r) => r.id), q);
      assert.equal(there.total, own.total, q);
      assert.deepEqual(there.indexed, { state: 'ready', done: ready.total, total: ready.total });
    }
    // A folder's list comes from its index once it is whole, a message from the file.
    const inbox = await host.call('archive:pst', 'list', { folder: 'Inbox', limit: 10 });
    assert.equal(inbox.total, 4);
    const message = await host.call('archive:pst', 'message', { folder: 'Inbox', id: inbox.rows[0].id });
    assert.ok(message.subject);
    const slim = await host.call('archive:pst', 'search', { query: 'halvorsen', limit: 50, slim: true });
    assert.deepEqual(Object.keys(slim.rows[0]).sort(), ['accountId', 'date', 'folder', 'id', 't']);
    await assert.rejects(host.call('archive:pst', 'nonsense'), /No such archive operation/);
  } finally {
    await host.stop();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a search while indexing goes on answers from what is indexed so far; indexing pauses and carries on', async () => {
  const dir = tmp();
  const file = path.join(dir, 'corpus.pst');
  await writeCorpusPst(file, { count: 3000, attachBytes: 2_000 });
  const events = [];
  // Small slices and parts, so the run is long enough to be caught midway.
  const host = createArchiveHost({ onProgress: (id, p) => events.push([id, p]), slice: 2, part: 150 });
  try {
    await host.open('archive:big', file, 'Corpus');
    await host.call('archive:big', 'index');
    let partial = null;
    for (let i = 0; i < 2000 && !partial; i++) {
      const found = await host.call('archive:big', 'search', { query: 'invoice', limit: 20 });
      if (found.indexed.state === 'indexing' && found.indexed.done > 0 && found.total > 0) partial = found;
      else await new Promise((r) => setTimeout(r, 2));
    }
    assert.ok(partial, 'a search was answered while indexing went on');
    assert.ok(partial.indexed.done < partial.indexed.total, JSON.stringify(partial.indexed));
    assert.equal(partial.indexed.total, 3000);

    const paused = await host.call('archive:big', 'pause');
    assert.equal(paused.state, 'paused');
    await new Promise((r) => setTimeout(r, 60));
    const still = await host.call('archive:big', 'status');
    assert.equal(still.done, paused.done, 'nothing more is indexed while paused');
    const whilePaused = await host.call('archive:big', 'search', { query: 'invoice', limit: 5000 });
    assert.ok(whilePaused.total >= partial.total, 'what was indexed before the pause is searchable');

    await host.call('archive:big', 'index');
    const ready = await settled(events, 'archive:big', 'ready', 60000);
    assert.equal(ready.done, 3000);
    const all = await host.call('archive:big', 'search', { query: 'invoice', limit: 5000 });
    // The same answer as one index built at once on this thread.
    const here = new SearchIndex();
    const source = openArchive(file);
    const steps = indexSteps(source, here, { id: 'archive:big', name: 'Corpus' });
    while (!steps.next().done);
    const own = here.search('invoice', { limit: 5000 });
    assert.equal(all.total, own.total);
    assert.deepEqual(new Set(all.rows.map((r) => `${r.folder}/${r.id}`)), new Set(own.rows.map((r) => `${r.folder}/${r.id}`)));
    source.close();
    // A folder indexed in parts lists every one of its messages once.
    const inbox = await host.call('archive:big', 'list', { folder: 'Inbox', limit: 1e9 });
    assert.equal(new Set(inbox.rows.map((r) => r.id)).size, inbox.total);
    assert.equal(inbox.total, source.folders().find((f) => f.path === 'Inbox').total);
  } finally {
    await host.stop();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('indexing in parts finds what indexing in one go finds', () => {
  const dir = tmp();
  try {
    const files = writeArchiveFixtures(path.join(dir, 'a'));
    const whole = new SearchIndex();
    indexArchive(openArchive(files.olm), whole, { id: 'o', name: 'o' });
    const parted = new SearchIndex();
    const hooks = {};
    const steps = indexSteps(openArchive(files.olm), parted, { id: 'o', name: 'o' }, { part: 1, hooks });
    let r = steps.next();
    while (!r.done) {
      hooks.flush();
      r = steps.next();
    }
    assert.ok(parted.segments.size > whole.segments.size, 'the folders were made searchable in parts');
    for (const q of ['halvorsen', '"indemnity clause"', 'contoso']) assert.deepEqual(parted.search(q).rows.map((x) => x.id).sort(), whole.search(q).rows.map((x) => x.id).sort(), q);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('closing an archive stops its thread; a file that will not open says so and leaves none', async () => {
  const dir = tmp();
  const host = createArchiveHost();
  try {
    const files = writeArchiveFixtures(path.join(dir, 'a'));
    await host.open('archive:x', files.mbox, 'Account export');
    assert.equal(host.stats.threads, 1);
    await host.close('archive:x');
    assert.equal(host.stats.threads, 0);
    await assert.rejects(host.call('archive:x', 'folders'), /not open/);
    // Asked again with its file, it starts a thread again.
    const folders = await host.call('archive:x', 'folders', {}, { file: files.mbox, name: 'Account export' });
    assert.equal(folders.length, 1);

    const junk = path.join(dir, 'notes.txt');
    fs.writeFileSync(junk, 'not mail at all');
    await assert.rejects(host.open('archive:junk', junk, 'notes'), /not a mail archive/);
    await new Promise((r) => setTimeout(r, 50));
    assert.equal(host.stats.threads, 1);
  } finally {
    await host.stop();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
