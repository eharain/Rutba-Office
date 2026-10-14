/**
 * Performance budget: a large Outlook archive opens, is indexed and searched
 * in seconds, and its index stays small.
 *
 * The .pst is written here, by tests/fixtures/pst-writer.js from the seeded
 * corpus in tests/fixtures/mail-corpus.js: 20,000 messages over ten years and
 * ten folders, a quarter of them carrying an attachment, about 1 GB on disk,
 * encoded with Outlook's permute cipher as its own files are. Set
 * RUTBA_PERF_PST_MESSAGES=100000 for the 5 GB class (4,927 MB; about a
 * minute to write and index on a four-core laptop).
 *
 * Measured on that laptop (2026-10-13), the 5 GB file: written in 31 s;
 * opened in 0.25 s; 100,000 messages indexed, bodies included, in 24 s
 * (27 s of CPU, 0.27 ms a message; in the app it runs in slices beside the
 * windows); searches 19 to 210 ms, a two-word phrase about 0.5 s; the index
 * took 173 MB (49 MB of postings, 160,000 words) and the process 397 MB in
 * all. The 1 GB default: open 0.12 s, index 11 s, 35 MB. Under load from
 * other work the same index took up to 80 s of CPU. The budgets below are
 * per message and per search, so they hold at either size, and are generous
 * on purpose (see perf-budget.perf.js): they exist to catch a step gone
 * quadratic.
 *
 * In the app the archive is indexed on a thread of its own
 * (@rutba/mailbox/archive-host), so the second budget is the main thread's:
 * while the same .pst is opened, indexed and searched over and over on its
 * thread, the longest this thread is held is under MAIN_STALL_MS. Measured
 * on the four-core laptop (2026-10-14), the 1 GB file: indexed on its thread
 * in 7.1 s beside 52 searches, this thread's longest gap 35 ms (event-loop
 * delay p99 25 ms; a Windows timer ticks every 15.6 ms, so that is the floor).
 *
 * Run alone (`npm run test:perf`), not in `npm test`.
 */
import test, { before, after } from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import v8 from 'node:v8';
import { monitorEventLoopDelay } from 'node:perf_hooks';
import vm from 'node:vm';
import { writeCorpusPst } from './fixtures/pst-writer.js';
import { openArchive, indexArchive } from '@rutba/mailbox/archive';
import { SearchIndex } from '@rutba/mailbox/search';
import { createArchiveHost } from '@rutba/mailbox/archive-host';

const COUNT = Number(process.env.RUTBA_PERF_PST_MESSAGES) || 20_000;
const OPEN_MS = 5_000;
const INDEX_MS_PER_MESSAGE = 2;
const SEARCH_MS = 1_500;
const INDEX_BYTES_PER_MESSAGE = 4_096;
/** The longest the main thread may be held while an archive indexes on its own. */
const MAIN_STALL_MS = 50;

const cpuOf = (fn) => {
  const c0 = process.cpuUsage();
  const t0 = performance.now();
  const out = fn();
  const used = process.cpuUsage(c0);
  return { out, cpu: Math.round((used.user + used.system) / 1000), ms: Math.round(performance.now() - t0) };
};
// A collection before each reading of memory, so garbage is not counted as index.
v8.setFlagsFromString('--expose-gc');
const gc = vm.runInNewContext('gc');
const settle = () => {
  gc();
  const m = process.memoryUsage();
  return m.heapUsed + m.arrayBuffers;
};

let dir;
let file;
let written;
let mb;
before(async () => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'perf-pst-'));
  file = path.join(dir, 'large.pst');
  // Written first; not part of the budget.
  const t0 = performance.now();
  const out = await writeCorpusPst(file, { count: COUNT, attachBytes: 190_000 });
  written = { out, ms: Math.round(performance.now() - t0) };
  mb = (out.bytes / 1048576).toFixed(0);
});
after(() => fs.rmSync(dir, { recursive: true, force: true }));

test(`a ${COUNT.toLocaleString('en')}-message .pst opens, is indexed and searched inside the budget`, () => {
  const opened = cpuOf(() => {
    const a = openArchive(file);
    return { archive: a, folders: a.folders() };
  });
  const { archive } = opened.out;
  assert.equal(opened.out.folders.reduce((n, f) => n + f.total, 0), COUNT, 'every message is counted when it opens');
  assert.ok(opened.cpu < OPEN_MS, `open took ${opened.cpu} ms of CPU, budget ${OPEN_MS}`);

  const afterOpen = settle();
  const index = new SearchIndex();
  const built = cpuOf(() => indexArchive(archive, index, { id: 'large', name: 'Large archive' }));
  assert.equal(built.out.done, COUNT);
  const indexBytes = settle() - afterOpen;
  assert.ok(built.cpu < COUNT * INDEX_MS_PER_MESSAGE, `indexing took ${built.cpu} ms of CPU, budget ${COUNT * INDEX_MS_PER_MESSAGE}`);
  assert.ok(indexBytes < COUNT * INDEX_BYTES_PER_MESSAGE, `the index took ${(indexBytes / 1048576).toFixed(0)} MB, budget ${((COUNT * INDEX_BYTES_PER_MESSAGE) / 1048576).toFixed(0)} MB`);

  const searches = [
    'invoice',
    'meeting report',
    '"thanks regards"',
    'from:northwind after:2018-01-01 before:2019-01-01',
    'has:attachment size>200kb',
    'filename:pdf',
    'contract -budget',
    'in:legal is:unread',
    'received:2020-01..2020-06 cc:contoso',
    'ka',
  ];
  const times = [];
  for (const q of searches) {
    const s = cpuOf(() => index.search(q, { limit: 200 }));
    times.push(`${q}=${s.out.total} in ${s.ms}ms (cpu ${s.cpu}ms)`);
    assert.ok(s.out.total > 0, `${q} finds something`);
    assert.ok(s.cpu < SEARCH_MS, `${q} took ${s.cpu} ms of CPU, budget ${SEARCH_MS}`);
  }
  // A phrase is confirmed against the bodies, read again from the file.
  const phrase = index.search('"thanks regards"', { limit: 5 });
  for (const row of phrase.rows) assert.match(archive.text(row.folder, row.id).toLowerCase(), /thanks\s+regards/);

  console.log(
    `# pst: ${COUNT} messages, ${mb} MB written in ${written.ms}ms; open ${opened.ms}ms (cpu ${opened.cpu}ms); ` +
      `index ${built.ms}ms (cpu ${built.cpu}ms, ${(built.cpu / COUNT).toFixed(2)} ms a message); ` +
      `index memory ${(indexBytes / 1048576).toFixed(0)} MB (postings ${(index.postingBytes() / 1048576).toFixed(0)} MB, ${index.vocabularySize} words); ` +
      `rss ${(process.memoryUsage().rss / 1048576).toFixed(0)} MB`
  );
  console.log(`# searches: ${times.join('; ')}`);
  archive.close();
});

test(`while a ${COUNT.toLocaleString('en')}-message .pst is indexed and searched on its thread, the main thread is never held ${MAIN_STALL_MS} ms`, async () => {
  // What the first budget left behind collected first, so its garbage is not taken for a stall.
  gc();
  // The longest gap between this thread's turns, by a timer that asks for one every millisecond.
  let last = performance.now();
  let worst = 0;
  const tick = setInterval(() => {
    const now = performance.now();
    worst = Math.max(worst, now - last);
    last = now;
  }, 1);
  const delay = monitorEventLoopDelay({ resolution: 5 });
  delay.enable();
  let ready = null;
  const host = createArchiveHost({ onProgress: (id, p) => p.state === 'ready' && (ready = p) });
  const t0 = performance.now();
  try {
    const opened = await host.open('archive:large', file, 'Large archive');
    assert.equal(opened.folders.reduce((n, f) => n + f.total, 0), COUNT);
    await host.call('archive:large', 'index');
    let searches = 0;
    let partial = null;
    while (!ready) {
      const found = await host.call('archive:large', 'search', { query: 'invoice', limit: 200 });
      searches++;
      if (!partial && found.indexed.state === 'indexing' && found.indexed.done > 0 && found.total > 0) partial = found;
      await new Promise((r) => setTimeout(r, 100));
    }
    const ms = Math.round(performance.now() - t0);
    clearInterval(tick);
    delay.disable();
    const p99 = delay.percentile(99) / 1e6;
    const max = delay.max / 1e6;
    console.log(`# main thread while ${COUNT} messages (${mb} MB) were indexed on their thread in ${ms}ms with ${searches} searches: longest gap ${worst.toFixed(1)}ms, event-loop delay max ${max.toFixed(1)}ms, p99 ${p99.toFixed(1)}ms`);
    assert.ok(partial, 'a search while indexing answered from what was indexed so far');
    assert.ok(partial.indexed.done < COUNT, JSON.stringify(partial.indexed));
    assert.equal(ready.done, COUNT);
    assert.ok(worst < MAIN_STALL_MS, `the main thread was held ${worst.toFixed(1)} ms, budget ${MAIN_STALL_MS}`);
    const all = await host.call('archive:large', 'search', { query: 'invoice', limit: 200 });
    assert.ok(all.total >= partial.total);
  } finally {
    clearInterval(tick);
    await host.stop();
  }
});
