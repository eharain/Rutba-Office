/**
 * Performance budget: an mbox and an Outlook for Mac archive past 2 GB open
 * where they lie, read a message from past the 2 GB mark and are indexed
 * whole, and the memory they take is their index, not their size.
 *
 * Both are written here by tests/fixtures/big-archives.js from the seeded
 * corpus: 9,000 messages, each with an attachment of about 190 KB (base64 in
 * the file), about 2.3 GB apiece, the last one written after the 2 GB mark;
 * the .olm with zip64 records, as an archive past 4 GB has. A Node buffer
 * holds at most 2 GB, so reading either whole would fail outright; the
 * budget below is the process's resident memory, sampled throughout, which
 * must stay under a third of the file's size. Set
 * RUTBA_PERF_BIG_MESSAGES=20000 for archives of about 5 GB.
 *
 * Measured on the four-core laptop (2026-10-14), 9,001 messages each: the
 * mbox (2,217 MB) opened in 2.0 s and indexed in 40 s, the .olm (2,193 MB,
 * zip64) opened in 0.07 s and indexed in 8.5 s; resident memory grew 25 MB
 * and 31 MB, the index 5 MB of postings each.
 *
 * Run alone (`npm run test:perf`), not in `npm test`.
 */
import test, { before, after } from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { writeCorpusMbox, writeCorpusOlm, LAST } from './fixtures/big-archives.js';
import { openArchive, indexArchive } from '@rutba/mailbox/archive';
import { SearchIndex } from '@rutba/mailbox/search';

const COUNT = Number(process.env.RUTBA_PERF_BIG_MESSAGES) || 9_000;
const TWO_GB = 2 ** 31;
/** How long opening may take: an mbox is read through once to find its messages. */
const OPEN_MS = 60_000;
const INDEX_MS_PER_MESSAGE = 15;
/** The most of the file's size the process may take in memory beyond where it started. */
const MEMORY_SHARE = 1 / 3;

let dir;
const files = {};
before(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'perf-big-'));
  files.mbox = path.join(dir, 'Big export.mbox');
  files.olm = path.join(dir, 'Big archive.olm');
  const t0 = performance.now();
  writeCorpusMbox(files.mbox, { count: COUNT });
  writeCorpusOlm(files.olm, { count: COUNT, zip64: true });
  console.log(`# written: mbox ${(fs.statSync(files.mbox).size / 1048576).toFixed(0)} MB, olm ${(fs.statSync(files.olm).size / 1048576).toFixed(0)} MB in ${Math.round(performance.now() - t0)}ms`);
});
after(() => fs.rmSync(dir, { recursive: true, force: true }));

for (const kind of ['mbox', 'olm']) {
  test(`an ${kind} past 2 GB (${COUNT.toLocaleString('en')} messages) opens in place, reads past the 2 GB mark and is indexed, in a third of its size in memory`, () => {
    const file = files[kind];
    const bytes = fs.statSync(file).size;
    assert.ok(bytes > TWO_GB, `the ${kind} is past 2 GB (${bytes})`);
    globalThis.gc?.();
    const base = process.memoryUsage().rss;
    let peak = base;
    const sample = () => {
      peak = Math.max(peak, process.memoryUsage().rss);
    };
    // Sampled between steps too: the work below is synchronous, so a timer would not fire.
    const t0 = performance.now();
    const archive = openArchive(file);
    const folders = archive.folders();
    const openMs = Math.round(performance.now() - t0);
    sample();
    assert.equal(folders.reduce((n, f) => n + f.total, 0), COUNT + 1, 'every message is counted');
    assert.ok(openMs < OPEN_MS, `open took ${openMs} ms, budget ${OPEN_MS}`);

    // The last message, written past the 2 GB mark, read from the file.
    const folder = kind === 'mbox' ? folders[0].path : 'Inbox';
    const lastId = kind === 'mbox' ? String(COUNT) : `${archive.olm.entries.length - 1}.0`;
    const last = archive.message(folder, lastId);
    sample();
    assert.equal(last.subject, LAST.subject);
    assert.match(last.text, /quartzberry/);
    if (kind === 'mbox') assert.ok(archive.mbox.range(COUNT).start > TWO_GB, 'it lies past the 2 GB mark');

    const index = new SearchIndex();
    const t1 = performance.now();
    let n = 0;
    // The archive as indexing sees it, with the memory sampled as it goes.
    const built = indexArchive(
      {
        kind: archive.kind,
        folders: () => archive.folders(),
        text: (f, id) => archive.text(f, id),
        *messages(f) {
          for (const m of archive.messages(f)) {
            if (++n % 250 === 0) sample();
            yield m;
          }
        },
      },
      index,
      { id: kind, name: kind }
    );
    const indexMs = Math.round(performance.now() - t1);
    sample();
    assert.equal(built.done, COUNT + 1);
    assert.ok(indexMs < (COUNT + 1) * INDEX_MS_PER_MESSAGE, `indexing took ${indexMs} ms, budget ${(COUNT + 1) * INDEX_MS_PER_MESSAGE}`);
    const found = index.search('quartzberry');
    assert.equal(found.total, 1);
    assert.equal(found.rows[0].subject, LAST.subject);
    assert.equal(index.search('"quartzberry clause"').total, 1, 'a phrase is confirmed against the body read from the file');
    sample();
    const grew = peak - base;
    console.log(`# ${kind}: ${(bytes / 1048576).toFixed(0)} MB, ${COUNT + 1} messages; open ${openMs}ms; last message read from past 2 GB; index ${indexMs}ms (${(indexMs / (COUNT + 1)).toFixed(2)} ms a message, ${(index.postingBytes() / 1048576).toFixed(0)} MB of postings); resident memory grew at most ${(grew / 1048576).toFixed(0)} MB (budget ${((bytes * MEMORY_SHARE) / 1048576).toFixed(0)} MB)`);
    assert.ok(grew < bytes * MEMORY_SHARE, `memory grew ${(grew / 1048576).toFixed(0)} MB for a ${(bytes / 1048576).toFixed(0)} MB archive`);
    archive.close();
  });
}
