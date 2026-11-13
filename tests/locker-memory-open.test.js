// The Locker's half of opening in memory (extras/locker/memory.js): which
// kinds open in memory, a save sealed as the next version, a save made while
// locked held until the next unlock, and files let go when no window shows
// them (2.26.0, D-2026-11-13-20).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { tempDir } from './temp-dirs.js';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { createMemoryOpener, memoryWay } from '../apps/desktop/main/extras/locker/memory.js';
import { createMemoryFiles } from '../apps/desktop/main/memory-files.js';

const CHEAP = { N: 1024, r: 8, p: 1 };

async function make({ limit } = {}) {
  const profile = tempDir('rl-memo-');
  const folder = tempDir('rl-memo-folder-');
  const locker = createLocker({ dir: profile, cost: CHEAP, large: 512 * 1024 });
  await locker.setup('pw');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3 });
  const table = createMemoryFiles({ root: path.join(profile, 'mem') });
  const memory = createMemoryOpener({ locker, table, ...(limit ? { limit } : {}) });
  return { locker, folder, id, profile, table, memory };
}

function everyFile(dir) {
  const out = [];
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push(...everyFile(p));
    else out.push(p);
  }
  return out;
}

const plainAnywhere = (dirs, needle) => dirs.flatMap(everyFile).filter((f) => fs.readFileSync(f).includes(needle));

test('Office\'s own kinds open in memory; the rest, and the kinds decoded through a path, from a plain copy', () => {
  for (const rel of ['a.docx', 'b/c.XLSX', 'd.pptx', 'e.pdf', 'f.txt', 'g.png', 'h.JPG', 'i.rimg', 'j.odt', 'k.csv']) assert.equal(memoryWay(rel), 'whole', rel);
  for (const rel of ['a.mp4', 'b.mp3', 'c.webm', 'd.wav']) assert.equal(memoryWay(rel), 'ranged', rel);
  for (const rel of ['a.heic', 'b.tiff', 'c.psd', 'd.emf', 'e.mov', 'f.mkv', 'g.eml', 'h.ics', 'i.vcf', 'j.rbase', 'k.zip', 'l.exe', 'noext']) assert.equal(memoryWay(rel), null, rel);
});

test('a document opened in memory is the locked file\'s bytes, and a save becomes its next version, never plain', async () => {
  const { locker, folder, id, profile, table, memory } = await make();
  const first = Buffer.from(`one ${crypto.randomUUID()}`.repeat(100));
  fs.writeFileSync(path.join(folder, 'Tender.docx'), first);
  await locker.lockWaiting(id);
  const opened = await memory.open(id, 'Tender.docx');
  assert.equal(opened.way, 'whole');
  assert.equal(path.basename(opened.path), 'Tender.docx');
  assert.deepEqual(await table.read(opened.path), first);
  const marker = Buffer.from(`saved-${crypto.randomUUID()}`);
  await table.write(opened.path, Buffer.concat([first, marker]));
  assert.equal((await locker.versions(id, 'Tender.docx')).length, 2);
  assert.ok((await locker.readBytes(id, 'Tender.docx')).bytes.includes(marker));
  assert.deepEqual(plainAnywhere([folder, profile], marker), []);
  assert.deepEqual(plainAnywhere([folder, profile], first.subarray(0, 40)), []);
});

test('a document over the limit is not opened in memory, and says why', async () => {
  const { locker, folder, id, memory } = await make({ limit: 1000 });
  fs.writeFileSync(path.join(folder, 'Big.pdf'), crypto.randomBytes(5000));
  await locker.lockWaiting(id);
  assert.deepEqual(await memory.open(id, 'Big.pdf'), { path: null, way: null, why: 'too-large' });
  assert.equal(await memory.open(id, 'Photo.heic'), null);
});

test('a save made while the Locker is locked is held, and sealed at the next unlock', async () => {
  const { locker, folder, id, profile, table, memory } = await make();
  fs.writeFileSync(path.join(folder, 'Notes.txt'), 'before');
  await locker.lockWaiting(id);
  const { path: p } = await memory.open(id, 'Notes.txt');
  await locker.lock();
  // The window still shows it, so it stays open across the lock.
  await memory.afterLock({ shown: (x) => x === p });
  assert.ok(table.has(p));
  const word = `HELD${crypto.randomUUID().slice(0, 8)}`;
  const out = await table.write(p, Buffer.from(`after ${word}`));
  assert.equal(out.held, true);
  assert.equal(memory.heldCount(), 1);
  // Its window closed meanwhile: the held save keeps it until it is sealed.
  assert.deepEqual(await memory.reap({ shown: () => false }), []);
  assert.deepEqual(plainAnywhere([folder, profile], Buffer.from(word)), []);
  await locker.unlock({ password: 'pw' });
  assert.deepEqual(await memory.flushHeld(), []);
  assert.equal(memory.heldCount(), 0);
  assert.equal((await locker.readBytes(id, 'Notes.txt')).bytes.toString(), `after ${word}`);
  assert.equal((await memory.reap({ shown: () => false })).length, 1);
  assert.ok(!table.has(p));
});

test('a lock lets go of every file no window shows, its bytes zeroed', async () => {
  const { locker, folder, id, table, memory } = await make();
  fs.writeFileSync(path.join(folder, 'a.docx'), 'aaaa');
  fs.writeFileSync(path.join(folder, 'b.docx'), 'bbbb');
  await locker.lockWaiting(id);
  const a = (await memory.open(id, 'a.docx')).path;
  const b = (await memory.open(id, 'b.docx')).path;
  const bBytes = table.entry(b).bytes;
  await locker.lock();
  await memory.afterLock({ shown: (x) => x === a });
  assert.ok(table.has(a));
  assert.ok(!table.has(b));
  assert.ok(bBytes.every((x) => x === 0));
  await memory.closeAll();
  assert.ok(!memory.hasOpen());
});

test('a film plays by range from its locked file, and a save of it reads back from the new one', async () => {
  const { locker, folder, id, table, memory } = await make();
  const film = crypto.randomBytes(700 * 1024);
  fs.writeFileSync(path.join(folder, 'clip.mp4'), film);
  await locker.lockWaiting(id);
  const { path: p, way } = await memory.open(id, 'clip.mp4');
  assert.equal(way, 'ranged');
  assert.equal(table.entry(p).bytes, null, 'not held whole');
  assert.deepEqual(await table.readRange(p, 300000, 1000), film.subarray(300000, 301000));
  const cut = film.subarray(0, 200000);
  await table.write(p, cut);
  assert.equal(table.stat(p).size, cut.length);
  assert.deepEqual(await table.readRange(p, 199000, 5000), cut.subarray(199000));
  await memory.closeAll();
});
