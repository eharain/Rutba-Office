// Thumbnails: the queue keeps only the newest jobs waiting (the tiles in
// view), letting the oldest go with nothing; and the kept JPEGs are trimmed
// to a size, the least lately made first.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createQueue, trimCache } from '../packages/office-shell/src/electron/thumbs.js';

test('only the newest jobs wait; the oldest are answered with nothing', async () => {
  let release;
  const gate = new Promise((r) => { release = r; });
  const ran = [];
  const q = createQueue({ concurrency: 1, keep: 3, run: async (job) => { ran.push(job.n); await gate; return `made ${job.n}`; } });
  const answers = Array.from({ length: 8 }, (_, n) => q.push({ n }));
  assert.equal(q.pending, 3, 'three wait');
  release();
  const got = await Promise.all(answers);
  assert.equal(got[0], 'made 0', 'the one already running finishes');
  assert.deepEqual(got.slice(1, 5), [null, null, null, null], 'the oldest waiting were let go');
  assert.deepEqual(got.slice(5).sort(), ['made 5', 'made 6', 'made 7']);
  assert.deepEqual(ran, [0, 7, 6, 5], 'newest first');
});

test('the cache is trimmed to its size, the oldest files first', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-thumbs-'));
  for (let i = 0; i < 10; i++) {
    const f = path.join(dir, `t${i}.jpg`);
    fs.writeFileSync(f, Buffer.alloc(1000));
    const at = new Date(Date.now() - (10 - i) * 60000);
    fs.utimesSync(f, at, at);
  }
  fs.writeFileSync(path.join(dir, 'notes.txt'), 'not a thumbnail');
  assert.equal(await trimCache(dir, { max: 20000, trim: 5000 }), 0, 'under the limit, nothing goes');
  assert.equal(await trimCache(dir, { max: 8000, trim: 5000 }), 5);
  assert.deepEqual(fs.readdirSync(dir).sort(), ['notes.txt', 't5.jpg', 't6.jpg', 't7.jpg', 't8.jpg', 't9.jpg']);
});
