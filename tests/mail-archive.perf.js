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
 * Run alone (`npm run test:perf`), not in `npm test`.
 */
import test from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import v8 from 'node:v8';
import vm from 'node:vm';
import { PstWriter } from './fixtures/pst-writer.js';
import { corpus } from './fixtures/mail-corpus.js';
import { openArchive, indexArchive } from '@rutba/mailbox/archive';
import { SearchIndex } from '@rutba/mailbox/search';

const COUNT = Number(process.env.RUTBA_PERF_PST_MESSAGES) || 20_000;
const OPEN_MS = 5_000;
const INDEX_MS_PER_MESSAGE = 2;
const SEARCH_MS = 1_500;
const INDEX_BYTES_PER_MESSAGE = 4_096;

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

test(`a ${COUNT.toLocaleString('en')}-message .pst opens, is indexed and searched inside the budget`, () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'perf-pst-'));
  const file = path.join(dir, 'large.pst');
  try {
    // Written first; not part of the budget.
    const writer = new PstWriter(file);
    const folders = new Map();
    const folderOf = (p) => {
      if (folders.has(p)) return folders.get(p);
      const parts = p.split('/');
      const handle = writer.folder(parts[parts.length - 1], parts.length > 1 ? folderOf(parts.slice(0, -1).join('/')) : undefined);
      folders.set(p, handle);
      return handle;
    };
    const written = cpuOf(() => {
      for (const { folder, message } of corpus({ count: COUNT, attachBytes: 190_000 })) writer.message(folderOf(folder), message);
      return writer.close();
    });
    const mb = (written.out.bytes / 1048576).toFixed(0);

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
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
