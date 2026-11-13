// Locked files opened in memory (2.26.0, D-2026-11-13-20): sealed straight
// from bytes, read whole or by range, and held in the table of files that
// never reach the disk as plain copies.
import { test } from 'node:test';
import { tempDir } from './temp-dirs.js';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { memorySource, openSource, CHUNK } from '../apps/desktop/main/extras/locker/format.js';
import { createMemoryFiles, RANGED_WHOLE } from '../apps/desktop/main/memory-files.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const random = (n) => crypto.randomBytes(n);

// A small large-file threshold, so pieces are made without sealing 50 MB.
const LARGE_AT = 512 * 1024;

async function make(choices = {}) {
  const profile = tempDir('rl-mem-');
  const folder = tempDir('rl-mem-folder-');
  const locker = createLocker({ dir: profile, cost: CHEAP, large: LARGE_AT });
  await locker.setup('pw');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3, ...choices });
  return { locker, folder, id, profile };
}

/** Every file under a folder, with its bytes, for a check that no plain copy was written anywhere. */
function everyFile(dir) {
  const out = [];
  const walk = (d) => {
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, e.name);
      if (e.isDirectory()) walk(p);
      else out.push(p);
    }
  };
  walk(dir);
  return out;
}

const holds = (file, needle) => fs.readFileSync(file).includes(needle);

test('a memory source reads like a file: by position, and in order with no position', async () => {
  const bytes = random(200000);
  const fh = await openSource(memorySource(bytes, 1234));
  const a = Buffer.alloc(1000);
  assert.equal((await fh.read(a, 0, 1000, 5000)).bytesRead, 1000);
  assert.deepEqual(a, bytes.subarray(5000, 6000));
  const b = Buffer.alloc(150000);
  assert.equal((await fh.read(b, 0, 150000, null)).bytesRead, 150000);
  const c = Buffer.alloc(150000);
  assert.equal((await fh.read(c, 0, 150000, null)).bytesRead, 50000);
  assert.deepEqual(Buffer.concat([b, c.subarray(0, 50000)]), bytes);
  assert.equal((await fh.read(c, 0, 10, null)).bytesRead, 0);
  const st = await fh.stat();
  assert.equal(st.size, bytes.length);
  assert.equal(st.mtimeMs, 1234);
});

test('a save from memory becomes the next version, and nothing plain is written on the way', async () => {
  const { locker, folder, id, profile } = await make();
  const first = Buffer.from(`first ${crypto.randomUUID()} `.repeat(500));
  fs.writeFileSync(path.join(folder, 'notes.docx'), first);
  await locker.lockWaiting(id);
  const marker = Buffer.from(`edited ${crypto.randomUUID()}`);
  const edited = Buffer.concat([first, marker]);
  await locker.saveBytes(id, 'notes.docx', edited);
  const got = await locker.readBytes(id, 'notes.docx');
  assert.deepEqual(got.bytes, edited);
  assert.equal((await locker.versions(id, 'notes.docx')).length, 2);
  for (const f of [...everyFile(folder), ...everyFile(profile)]) assert.ok(!holds(f, marker), `plain text found in ${f}`);
});

test('a save from memory that grows past the large size moves the file into pieces, earlier versions too, with no plain copy', async () => {
  const { locker, folder, id, profile } = await make();
  const marker = Buffer.from(`old-version ${crypto.randomUUID()}`);
  const small = Buffer.concat([marker, random(1000)]);
  fs.writeFileSync(path.join(folder, 'film.mov'), small);
  await locker.lockWaiting(id);
  const big = Buffer.concat([random(LARGE_AT + 300000), Buffer.from('tail')]);
  await locker.saveBytes(id, 'film.mov', big);
  const e = (await locker.scan(id)).find((x) => x.rel === 'film.mov');
  assert.ok(e.pieces, 'kept as pieces');
  assert.deepEqual((await locker.readBytes(id, 'film.mov')).bytes, big);
  assert.equal((await locker.versions(id, 'film.mov')).length, 2);
  for (const f of [...everyFile(folder), ...everyFile(profile)]) assert.ok(!holds(f, marker), `plain text of the old version found in ${f}`);
});

test('a file read whole refuses one over the limit, and a conflict', async () => {
  const { locker, folder, id } = await make();
  fs.writeFileSync(path.join(folder, 'a.pdf'), random(5000));
  await locker.lockWaiting(id);
  await assert.rejects(locker.readBytes(id, 'a.pdf', { limit: 4000 }), (err) => err.code === 'too-large');
  assert.equal((await locker.readBytes(id, 'a.pdf', { limit: 5000 })).bytes.length, 5000);
  await assert.rejects(locker.readBytes(id, 'missing.pdf'), (err) => err.code === 'state');
});

async function checkRanges(r, plain) {
  const cases = [
    [0, 10], [0, plain.length], [CHUNK - 5, 10], [CHUNK, CHUNK], [CHUNK * 2 - 1, 2],
    [plain.length - 7, 100], [plain.length, 10], [plain.length + 50, 10], [12345, 0],
  ];
  for (let i = 0; i < 40; i += 1) {
    const s = crypto.randomInt(0, plain.length);
    cases.push([s, crypto.randomInt(1, 300000)]);
  }
  for (const [s, n] of cases) {
    const got = await r.readAt(s, n);
    assert.deepEqual(got, plain.subarray(s, Math.min(plain.length, s + n)), `range ${s}+${n}`);
  }
}

test('a locked file reads by range across its frames', async () => {
  const { locker, folder, id } = await make();
  const plain = random(CHUNK * 5 + 777);
  fs.writeFileSync(path.join(folder, 'song.mp3'), plain);
  await locker.lockWaiting(id);
  const r = await locker.openRange(id, 'song.mp3');
  try {
    assert.equal(Number(r.meta.size), plain.length);
    await checkRanges(r, plain);
  } finally {
    await r.close();
  }
});

test('a file in pieces reads by range across its pieces, after the folder call has ended', async () => {
  const { locker, folder, id } = await make();
  const plain = random(LARGE_AT * 3 + 4321);
  fs.writeFileSync(path.join(folder, 'clip.mp4'), plain);
  await locker.lockWaiting(id);
  assert.ok((await locker.scan(id)).find((x) => x.rel === 'clip.mp4').pieces);
  const r = await locker.openRange(id, 'clip.mp4');
  try {
    // Another folder call in between wipes that call's key; the reader has its own.
    await locker.scan(id);
    await checkRanges(r, plain);
  } finally {
    await r.close();
  }
});

test('a frame altered on the disk is refused by a range read', async () => {
  const { locker, folder, id } = await make();
  const plain = random(CHUNK * 3);
  fs.writeFileSync(path.join(folder, 'a.wav'), plain);
  await locker.lockWaiting(id);
  const e = (await locker.scan(id)).find((x) => x.rel === 'a.wav');
  const raw = fs.readFileSync(e.path);
  raw[raw.length - 100] ^= 1;
  fs.writeFileSync(e.path, raw);
  const r = await locker.openRange(id, 'a.wav');
  try {
    assert.deepEqual(await r.readAt(0, 100), plain.subarray(0, 100));
    await assert.rejects(r.readAt(CHUNK * 2 + 10, 10), (err) => err.code === 'tampered');
  } finally {
    await r.close();
  }
});

// The table.

test('the table answers for its paths, and a path under it with no entry reads as gone', async () => {
  const root = path.join(tempDir('rl-memt-'), 'mem');
  const t = createMemoryFiles({ root });
  const bytes = Buffer.from('hello plain');
  const p = t.add({ name: 'Report.docx', bytes, save: async () => ({}) });
  assert.equal(path.basename(p), 'Report.docx');
  assert.ok(t.has(p) && t.isUnder(p));
  assert.ok(t.has(p.toUpperCase()) === (process.platform === 'win32' || process.platform === 'darwin'));
  assert.deepEqual(await t.read(p), bytes);
  assert.deepEqual(t.readSync(p), bytes);
  assert.deepEqual(await t.readRange(p, 6, 100), Buffer.from('plain'));
  assert.equal(t.stat(p).size, bytes.length);
  assert.equal(t.stat(p).readonly, false);
  const other = path.join(path.dirname(p), 'Other.docx');
  assert.ok(!t.has(other) && t.isUnder(other));
  const stray = path.join(root, 'ffff', 'x.docx');
  assert.ok(t.isUnder(stray) && !t.has(stray));
  await assert.rejects(t.read(stray), (err) => err.code === 'ENOENT');
  assert.ok(!t.isUnder(path.join(path.dirname(root), 'mem2', 'x')));
});

test('a plain write that misses the table cannot land under it', async () => {
  const root = path.join(tempDir('rl-memt-'), 'mem');
  const t = createMemoryFiles({ root });
  const p = t.add({ name: 'Budget.xlsx', bytes: Buffer.from('x'), save: async () => ({}) });
  assert.throws(() => fs.writeFileSync(p, 'plain'), (err) => ['ENOTDIR', 'ENOENT', 'EEXIST'].includes(err.code));
  assert.throws(() => fs.mkdirSync(path.dirname(p), { recursive: true }), (err) => ['ENOTDIR', 'ENOENT', 'EEXIST'].includes(err.code));
  assert.throws(() => fs.writeFileSync(`${p}.tmp`, 'plain'), (err) => ['ENOTDIR', 'ENOENT', 'EEXIST'].includes(err.code));
  assert.equal(fs.statSync(path.dirname(p)).size, 0);
  assert.ok(fs.statSync(path.dirname(p)).isFile());
});

test('a write seals through save, saves of one file run in order, and a read-only file refuses', async () => {
  const root = path.join(tempDir('rl-memt-'), 'mem');
  const t = createMemoryFiles({ root });
  const sealed = [];
  const p = t.add({
    name: 'a.docx',
    bytes: Buffer.from('v0'),
    save: async (b) => {
      await new Promise((r) => setTimeout(r, b.toString() === 'v1' ? 30 : 0));
      sealed.push(b.toString());
      return { modified: 42 };
    },
  });
  await Promise.all([t.write(p, Buffer.from('v1')), t.write(p, Buffer.from('v2'))]);
  assert.deepEqual(sealed, ['v1', 'v2']);
  assert.equal((await t.read(p)).toString(), 'v2');
  assert.equal(t.stat(p).mtimeMs, 42);
  const ro = t.add({ name: 'b.pdf', bytes: Buffer.from('x') });
  assert.equal(t.stat(ro).readonly, true);
  await assert.rejects(t.write(ro, Buffer.from('y')), (err) => err.code === 'EACCES');
});

test('a failed save keeps the bytes the file had', async () => {
  const root = path.join(tempDir('rl-memt-'), 'mem');
  const t = createMemoryFiles({ root });
  const p = t.add({ name: 'a.docx', bytes: Buffer.from('v0'), save: async () => { throw Object.assign(new Error('away'), { code: 'away' }); } });
  await assert.rejects(t.write(p, Buffer.from('v1')), (err) => err.code === 'away');
  assert.equal((await t.read(p)).toString(), 'v0');
});

test('dropping a file zeroes what it held and removes its marker; clear drops them all', async () => {
  const root = path.join(tempDir('rl-memt-'), 'mem');
  const t = createMemoryFiles({ root });
  const p = t.add({ name: 'a.docx', bytes: Buffer.from('secret'), save: async () => ({}) });
  const held = t.entry(p).bytes;
  let closed = 0;
  const q = t.add({ name: 'b.mp4', readAt: async () => Buffer.alloc(0), size: 10, close: async () => { closed += 1; } });
  assert.ok(await t.drop(p));
  assert.ok(held.every((b) => b === 0));
  assert.ok(!t.has(p));
  assert.ok(!fs.existsSync(path.dirname(p)));
  await assert.rejects(t.read(p), (err) => err.code === 'ENOENT');
  await t.clear();
  assert.equal(closed, 1);
  assert.ok(!t.has(q));
  assert.equal(t.entries().length, 0);
});

test('a ranged file reads through its reader, and a save hands it a new one', async () => {
  const root = path.join(tempDir('rl-memt-'), 'mem');
  const t = createMemoryFiles({ root });
  const v1 = random(1000);
  const v2 = random(800);
  const reader = (b) => async (s, n) => b.subarray(s, s + n);
  let closed = 0;
  const p = t.add({
    name: 'a.jpg',
    readAt: reader(v1),
    size: v1.length,
    close: async () => { closed += 1; },
    save: async () => ({ readAt: reader(v2), size: v2.length, close: async () => {} }),
  });
  assert.deepEqual(await t.readRange(p, 10, 20), v1.subarray(10, 30));
  assert.deepEqual(await t.read(p), v1);
  assert.throws(() => t.readSync(p), (err) => err.code === 'EINVAL');
  await t.write(p, v2);
  assert.equal(closed, 1);
  assert.equal(t.stat(p).size, 800);
  assert.deepEqual(await t.read(p), v2);
});

test('markers left by an earlier run are swept; live ones stay', async () => {
  const root = path.join(tempDir('rl-memt-'), 'mem');
  fs.mkdirSync(root, { recursive: true });
  fs.writeFileSync(path.join(root, 'abcdef'), '');
  const t = createMemoryFiles({ root });
  const p = t.add({ name: 'a.docx', bytes: Buffer.from('x') });
  assert.equal(t.sweep(), 1);
  assert.ok(!fs.existsSync(path.join(root, 'abcdef')));
  assert.ok(fs.existsSync(path.dirname(p)));
});

test('a file read by range is never read whole past the limit, and a failed save zeroes what it was handed', async () => {
  const root = tempDir('rl-memcap-');
  const t = createMemoryFiles({ root: path.join(root, 'mem') });
  const big = t.add({ name: 'Film.mp4', readAt: async (_s, n) => Buffer.alloc(n, 1), size: RANGED_WHOLE + 1 });
  await assert.rejects(t.read(big), (err) => err.code === 'EFBIG');
  assert.equal((await t.readRange(big, 0, 10)).length, 10, 'by range it still reads');
  assert.equal(t.stat(big).ranged, true);
  let handed = null;
  const doc = t.add({ name: 'a.docx', bytes: Buffer.from('old'), save: async (b) => { handed = b; throw new Error('the disk is away'); } });
  assert.equal(t.stat(doc).ranged, false);
  await assert.rejects(t.write(doc, Buffer.from('SECRET-NEW')), /the disk is away/);
  assert.ok(handed.every((b) => b === 0), 'the new bytes are zeroed');
  assert.equal((await t.read(doc)).toString(), 'old');
});
