// The Locker's half of opening in memory (extras/locker/memory.js): which
// kinds open in memory, a save sealed as the next version, and, by the
// owner's answers of 13 Nov (D-2026-11-13-22), a lock that seals what is
// open, closes it and offers it again, and a crash whose sealed copy is
// offered at the next unlock (2.26.0).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { tempDir } from './temp-dirs.js';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { createMemoryOpener, memoryWay, plainWhy, memoryNote, freeRel } from '../apps/desktop/main/extras/locker/memory.js';
import { createMemoryFiles } from '../apps/desktop/main/memory-files.js';

const CHEAP = { N: 1024, r: 8, p: 1 };

async function make({ limit, unsaved = new Map() } = {}) {
  const profile = tempDir('rl-memo-');
  const folder = tempDir('rl-memo-folder-');
  const locker = createLocker({ dir: profile, cost: CHEAP, large: 512 * 1024 });
  await locker.setup('pw');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: true, keepVersions: 3 });
  const table = createMemoryFiles({ root: path.join(profile, 'mem') });
  const recoveryDir = path.join(profile, 'recovery');
  const opener = () => createMemoryOpener({ locker, table, recoveryDir, unsavedOf: (p) => unsaved.get(p) || null, ...(limit ? { limit } : {}) });
  return { locker, folder, id, profile, table, memory: opener(), opener, unsaved, recoveryDir };
}

function everyFile(dir) {
  if (!fs.existsSync(dir)) return [];
  const out = [];
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push(...everyFile(p));
    else out.push(p);
  }
  return out;
}

const plainAnywhere = (dirs, needle) => dirs.flatMap(everyFile).filter((f) => fs.readFileSync(f).includes(needle));

test('Office\'s own kinds open in memory; the rest, and the kinds decoded through a path, from a plain copy that says why', () => {
  for (const rel of ['a.docx', 'b/c.XLSX', 'd.pptx', 'e.pdf', 'f.txt', 'g.png', 'h.JPG', 'i.rimg', 'j.odt', 'k.csv']) assert.equal(memoryWay(rel), 'whole', rel);
  for (const rel of ['a.mp4', 'b.mp3', 'c.webm', 'd.wav']) assert.equal(memoryWay(rel), 'ranged', rel);
  for (const rel of ['a.heic', 'b.tiff', 'c.psd', 'd.emf', 'e.mov', 'f.mkv']) {
    assert.equal(memoryWay(rel), null, rel);
    assert.equal(plainWhy(rel), 'by-path', rel);
  }
  for (const rel of ['g.eml', 'h.ics', 'i.vcf', 'j.rbase', 'k.zip', 'l.exe', 'noext']) {
    assert.equal(memoryWay(rel), null, rel);
    assert.equal(plainWhy(rel), 'kind', rel);
  }
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

test('a document over the limit, or a kind decoded through a path, is not opened in memory, and says why', async () => {
  const { locker, folder, id, memory } = await make({ limit: 1000 });
  fs.writeFileSync(path.join(folder, 'Big.pdf'), crypto.randomBytes(5000));
  await locker.lockWaiting(id);
  assert.deepEqual(await memory.open(id, 'Big.pdf'), { path: null, way: null, why: 'too-large' });
  assert.deepEqual(await memory.open(id, 'Photo.heic'), { path: null, way: null, why: 'by-path' });
});

test('a lock seals each open file\'s unsaved work, lets every byte go, and the next unlock offers it back unsaved', async () => {
  const { locker, folder, id, profile, table, memory, unsaved, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Notes.txt'), 'before');
  fs.writeFileSync(path.join(folder, 'Plan.docx'), 'plan as saved');
  await locker.lockWaiting(id);
  const notes = (await memory.open(id, 'Notes.txt')).path;
  const plan = (await memory.open(id, 'Plan.docx')).path;
  const held = table.entry(notes).bytes;
  const word = `UNSAVED${crypto.randomUUID().slice(0, 8)}`;
  unsaved.set(notes, Buffer.from(`before ${word}`));

  // The lock, as index.js does it: sealed while the keys are here, then let go.
  const closing = await memory.sealOpen();
  assert.deepEqual(closing.sort(), [notes, plan].sort(), 'both windows close');
  await locker.lock();
  await memory.closeAll();
  assert.ok(!memory.hasOpen());
  assert.ok(held.every((b) => b === 0), 'the bytes are zeroed');
  await assert.rejects(table.read(notes), (err) => err.code === 'ENOENT');
  // Nothing plain, not the unsaved word nor either name.
  assert.deepEqual(plainAnywhere([folder, profile], Buffer.from(word)), []);
  assert.deepEqual(plainAnywhere([recoveryDir], Buffer.from('Notes')), []);
  assert.deepEqual(plainAnywhere([recoveryDir], Buffer.from('Plan')), []);
  assert.deepEqual(await memory.offers(), [], 'nothing is offered while it is locked');

  await locker.unlock({ password: 'pw' });
  const offers = await memory.offers();
  assert.deepEqual(offers.map((o) => [o.rel, o.unsaved]).sort(), [['Notes.txt', true], ['Plan.docx', false]]);
  const again = await memory.reopen(offers.find((o) => o.rel === 'Notes.txt').key);
  assert.equal(again.unsaved, true);
  assert.equal((await table.read(again.path)).toString(), `before ${word}`, 'with its unsaved work');
  assert.equal(table.entry(again.path).info.unsaved, true, 'marked unsaved, so its window asks before it is thrown away');
  assert.equal((await locker.readBytes(id, 'Notes.txt')).bytes.toString(), 'before', 'the locked file is as it was saved');
  const plain = await memory.reopen(offers.find((o) => o.rel === 'Plan.docx').key);
  assert.equal(plain.unsaved, false);
  assert.equal((await table.read(plain.path)).toString(), 'plan as saved');
  assert.deepEqual((await memory.offers()).length, 0, 'an offer taken is not offered again while it is open');
  // Saved: its recovery copy goes, and nothing is offered after the next lock but what is open then.
  await table.write(again.path, Buffer.from(`before ${word} saved`));
  assert.equal((await locker.readBytes(id, 'Notes.txt')).bytes.toString(), `before ${word} saved`);
  await memory.reap({ shown: () => false });
  assert.deepEqual(await memory.offers(), []);
});

test('a save after the lock is refused in words', async () => {
  const { locker, folder, id, table, memory } = await make();
  fs.writeFileSync(path.join(folder, 'a.txt'), 'x');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'a.txt')).path;
  await locker.lock();
  await assert.rejects(table.write(p, Buffer.from('y')), /The Locker has locked, so this file cannot be saved/);
});

test('what a crash leaves is the copy sealed as the app autosaved, offered at the next unlock', async () => {
  const { locker, folder, id, profile, table, memory, opener } = await make();
  fs.writeFileSync(path.join(folder, 'Budget.xlsx'), 'saved');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'Budget.xlsx')).path;
  const word = `AUTOSAVED${crypto.randomUUID().slice(0, 8)}`;
  // documents.js's autosave hands the unsaved work to the entry's recover.
  await table.entry(p).recover(Buffer.from(word), { kind: 'sheet' });
  assert.deepEqual(plainAnywhere([folder, profile], Buffer.from(word)), []);
  // The crash: this run's table and opener are gone; the next run starts empty and is unlocked.
  await table.clear();
  const after = opener();
  await locker.lock();
  await locker.unlock({ password: 'pw' });
  const offers = await after.offers();
  assert.equal(offers.length, 1);
  assert.equal(offers[0].rel, 'Budget.xlsx');
  assert.equal(offers[0].unsaved, true);
  // Turned down: it goes.
  await after.dismiss(offers[0].key);
  assert.deepEqual(await after.offers(), []);
});

test('a window the person closes takes its recovery copy with it', async () => {
  const { locker, folder, id, table, memory } = await make();
  fs.writeFileSync(path.join(folder, 'a.docx'), 'aaaa');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'a.docx')).path;
  await table.entry(p).recover(Buffer.from('changed'), {});
  assert.equal((await memory.reap({ shown: (x) => x === p })).length, 0, 'still shown');
  assert.equal((await memory.reap({ shown: () => false })).length, 1);
  assert.deepEqual(await memory.offers(), []);
});

test('a recovery copy of a folder no longer in the Locker is left alone, not offered', async () => {
  const { locker, folder, id, memory, unsaved } = await make();
  fs.writeFileSync(path.join(folder, 'a.docx'), 'aaaa');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'a.docx')).path;
  unsaved.set(p, Buffer.from('changed'));
  await memory.sealOpen();
  await memory.closeAll();
  await locker.removeFolder(id).catch(() => {});
  assert.deepEqual(await memory.offers(), []);
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

test('a lock keeps the unsaved work a PDF or Image window sealed itself, and offers it back unsaved', async () => {
  const { locker, folder, id, profile, table, memory } = await make();
  fs.writeFileSync(path.join(folder, 'Lease.pdf'), '%PDF-as-saved');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'Lease.pdf')).path;
  // PDF and Image keep their work through doc.keepCopy, which hands it to recover; documents.js has no session for them.
  const markup = `%PDF-MARKUP${crypto.randomUUID().slice(0, 8)}`;
  await table.entry(p).recover(Buffer.from(markup), { kind: 'pdf' });
  await memory.sealOpen();
  await locker.lock();
  await memory.closeAll();
  assert.deepEqual(plainAnywhere([folder, profile], Buffer.from(markup)), []);
  await locker.unlock({ password: 'pw' });
  const [offer] = await memory.offers();
  assert.equal(offer.unsaved, true, 'offered with its unsaved work');
  const again = await memory.reopen(offer.key);
  assert.equal((await table.read(again.path)).toString(), markup, 'the markup is there');
});

test('a file whose window has not shown it yet is kept, with its recovery copy, until a window shows it and is closed', async () => {
  const { locker, folder, id, table, memory, unsaved, recoveryDir } = await make();
  for (const n of ['a.docx', 'b.docx']) fs.writeFileSync(path.join(folder, n), n);
  await locker.lockWaiting(id);
  for (const n of ['a.docx', 'b.docx']) unsaved.set((await memory.open(id, n)).path, Buffer.from(`${n} changed`));
  await memory.sealOpen();
  await locker.lock();
  await memory.closeAll();
  unsaved.clear();
  await locker.unlock({ password: 'pw' });
  const offers = await memory.offers();
  assert.equal(offers.length, 2);
  // Reopen opens them one after another; the reaper runs every 3 s and may come between.
  const first = await memory.reopen(offers[0].key);
  assert.deepEqual(await memory.reap({ shown: () => false }), [], 'not yet shown: kept');
  const second = await memory.reopen(offers[1].key);
  assert.deepEqual(await memory.reap({ shown: () => false }), []);
  assert.equal(fs.readdirSync(recoveryDir).length, 2, 'both recovery copies are still there');
  assert.ok(table.has(first.path) && table.has(second.path));
  // Shown, then closed by the person: that one goes, with its copy.
  await memory.reap({ shown: (x) => x === first.path });
  assert.equal((await memory.reap({ shown: () => false })).length, 1);
  assert.ok(!table.has(first.path) && table.has(second.path));
  assert.equal(fs.readdirSync(recoveryDir).length, 1);
});

test('the older formats Office reads but does not write open in memory, with a note that changes cannot be saved while locked', async () => {
  for (const rel of ['a.doc', 'b.HTM', 'c.xls', 'd.ppt']) assert.equal(memoryNote(rel), 'old-kind', rel);
  for (const rel of ['a.docx', 'b.html', 'c.xlsx', 'd.pptx', 'e.odt', 'f.csv', 'g.pdf']) assert.equal(memoryNote(rel), null, rel);
  const { locker, folder, id, memory } = await make();
  fs.writeFileSync(path.join(folder, 'Old.doc'), 'old');
  fs.writeFileSync(path.join(folder, 'New.docx'), 'new');
  await locker.lockWaiting(id);
  const old = await memory.open(id, 'Old.doc');
  assert.equal(old.way, 'whole');
  assert.equal(old.note, 'old-kind');
  assert.equal((await memory.open(id, 'New.docx')).note, null);
});

test('a file in the table that is not the Locker\'s (a print job) is not sealed, offered or reaped', async () => {
  const { locker, id, table, memory } = await make();
  const job = table.add({ name: 'job.pdf', bytes: Buffer.from('%PDF-job'), info: { job: true } });
  assert.equal(memory.hasOpen(), false);
  assert.deepEqual(memory.paths(), []);
  assert.deepEqual(await memory.sealOpen(), []);
  assert.deepEqual(await memory.reap({ shown: () => false }), []);
  assert.ok(table.has(job));
  assert.ok(id && locker.isOpen());
});

test('Save As from a file in an older format seals a new file beside it in the same locked folder, under a free name', async () => {
  const { locker, folder, id, profile, table, memory } = await make();
  fs.mkdirSync(path.join(folder, 'Accounts'));
  fs.writeFileSync(path.join(folder, 'Accounts', 'Budget.xls'), 'old workbook');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'Accounts/Budget.xls')).path;
  const word = `NEWBOOK${crypto.randomUUID().slice(0, 8)}`;
  const now = await table.saveAs(p, 'Budget.xlsx', Buffer.from(word));
  assert.equal(path.basename(now), 'Budget.xlsx');
  assert.equal((await locker.readBytes(id, 'Accounts/Budget.xlsx')).bytes.toString(), word, 'sealed beside it');
  assert.equal((await locker.readBytes(id, 'Accounts/Budget.xls')).bytes.toString(), 'old workbook', 'the old file is as it was');
  assert.equal(table.entry(now).info.rel, 'Accounts/Budget.xlsx', 'the window has the new one open');
  assert.equal(table.has(p), false, 'the old path reads as gone');
  assert.deepEqual(plainAnywhere([folder, profile], Buffer.from(word)), [], 'nothing plain');
  // A save after that is a save of the new file.
  await table.write(now, Buffer.from(`${word} again`));
  assert.equal((await locker.versions(id, 'Accounts/Budget.xlsx')).length, 2);
  // A name already in the folder is never written over.
  fs.writeFileSync(path.join(folder, 'Accounts', 'Plan.xls'), 'plan');
  fs.writeFileSync(path.join(folder, 'Accounts', 'Plan.xlsx'), 'someone else');
  await locker.lockWaiting(id);
  const plan = (await memory.open(id, 'Accounts/Plan.xls')).path;
  assert.equal(path.basename(await table.saveAs(plan, 'Plan.xlsx', Buffer.from('mine'))), 'Plan (2).xlsx');
  assert.equal((await locker.readBytes(id, 'Accounts/Plan.xlsx')).bytes.toString(), 'someone else');
  await locker.lock();
  await assert.rejects(table.saveAs(now, 'Budget.xlsx', Buffer.from('x')), /The Locker has locked/);
});

test('a free name is the one asked for, or the first numbered one not taken', () => {
  assert.equal(freeRel('.', 'a.xlsx', new Set()), 'a.xlsx');
  assert.equal(freeRel('d/e', 'a.xlsx', new Set(['d/e/a.xlsx', 'd/e/a (2).xlsx'])), 'd/e/a (3).xlsx');
  assert.equal(freeRel('.', 'A.XLSX', new Set(['a.xlsx'])), 'A (2).XLSX');
});

test('after Save As beside it, the window that opened the old name still counts as showing the file', async () => {
  const { locker, folder, id, table, memory } = await make();
  fs.writeFileSync(path.join(folder, 'Deck.ppt'), 'old deck');
  await locker.lockWaiting(id);
  const opened = (await memory.open(id, 'Deck.ppt')).path;
  // The window's address names Deck.ppt; the reaper has seen it shown.
  const shown = (x) => x === opened;
  await memory.reap({ shown });
  const now = await table.saveAs(opened, 'Deck.pptx', Buffer.from('new deck'));
  assert.deepEqual(await memory.reap({ shown }), [], 'not let go while that window is open');
  assert.ok(table.has(now));
  assert.ok(memory.paths().includes(opened), 'the idle lock and the lock close that window');
  assert.ok((await memory.sealOpen()).includes(opened));
});

test('a save that lands while an autosave copy is being sealed leaves no stale copy to be offered as unsaved', async () => {
  const { locker, folder, id, table, memory } = await make();
  fs.writeFileSync(path.join(folder, 'Ledger.xlsx'), 'v0');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'Ledger.xlsx')).path;
  // The autosave of version 1 starts, and is slow to seal; meanwhile the person saves version 3.
  const realSeal = locker.sealRecovery;
  let release;
  const gate = new Promise((r) => { release = r; });
  locker.sealRecovery = async (...args) => { await gate; return realSeal(...args); };
  const autosave = table.entry(p).recover(Buffer.from('v1 autosaved'), { kind: 'sheet' });
  // The save is in the locked file (counted) while the copy is still being sealed; what it does with the
  // copy waits for that seal, in the file's chain.
  const saving = table.write(p, Buffer.from('v3 saved'));
  while (!table.entry(p).info.saves) await new Promise((r) => setTimeout(r, 5));
  release();
  await saving;
  assert.equal(await autosave, false, 'the stale copy is not kept');
  locker.sealRecovery = realSeal;
  assert.equal(table.entry(p).info.unsaved, false);
  // The document is clean at the lock; the next unlock offers it as saved, never v1 as unsaved work.
  await memory.sealOpen();
  await locker.lock();
  await memory.closeAll();
  await locker.unlock({ password: 'pw' });
  const [offer] = await memory.offers();
  assert.equal(offer.unsaved, false);
  const again = await memory.reopen(offer.key);
  assert.equal((await table.read(again.path)).toString(), 'v3 saved');
});

test('two copies of one file sealed close together end with the later one', async () => {
  const { locker, folder, id, table, memory } = await make();
  fs.writeFileSync(path.join(folder, 'Map.pdf'), '%PDF-0');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'Map.pdf')).path;
  const e = table.entry(p);
  await Promise.all([e.recover(Buffer.from('%PDF-first')), e.recover(Buffer.from('%PDF-second'))]);
  await memory.sealOpen();
  await locker.lock();
  await memory.closeAll();
  await locker.unlock({ password: 'pw' });
  const [offer] = await memory.offers();
  assert.equal(offer.unsaved, true);
  assert.equal((await table.read((await memory.reopen(offer.key)).path)).toString(), '%PDF-second');
});

test('a lock does not wait past its time for an app that does not answer, and keeps the copy already sealed', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Hung.docx'), 'locked words');
  await locker.lockWaiting(id);
  // The document thread, busy or hung: it never answers for this file.
  const memory = createMemoryOpener({ locker, table, recoveryDir, unsavedOf: () => new Promise(() => {}), unsavedWaitMs: 150 });
  const p = (await memory.open(id, 'Hung.docx')).path;
  await table.entry(p).recover(Buffer.from('autosaved work'));
  const slot = table.entry(p).info.slot;
  const started = Date.now();
  await memory.sealOpen({ deadline: Date.now() + 5000 });
  assert.ok(Date.now() - started < 2000, `${Date.now() - started} ms`);
  const offer = (await memory.offers()).find((o) => o.key === slot);
  assert.equal(offer, undefined, 'still open, so not offered yet');
  await table.drop(p);
  const kept = (await memory.offers()).find((o) => o.key === slot);
  assert.equal(kept?.unsaved, true, 'the copy sealed before is the one offered');
  const back = await memory.reopen(slot);
  assert.equal((await table.read(back.path)).toString(), 'autosaved work');
});

test('a window that gives its work up leaves no copy to be offered, and a copy made before a save that lands first is stale', async () => {
  const { locker, folder, id, table, memory, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Given.pdf'), '%PDF-1.7 locked');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'Given.pdf')).path;
  const e = table.entry(p);
  await e.recover(Buffer.from('%PDF-1.7 unsaved'));
  assert.equal(e.info.unsaved, true);
  await e.forget();
  assert.equal(e.info.unsaved, false);
  assert.equal(fs.existsSync(path.join(recoveryDir, e.info.slot)), false, 'its copy gone');
  // Bytes taken at save count n, reaching the Locker after a save made it n + 1: stale, not offered.
  const saves = e.info.saves;
  await table.write(p, Buffer.from('%PDF-1.7 saved'));
  assert.equal(await e.recover(Buffer.from('%PDF-1.7 older'), { saves }), false);
  assert.equal(e.info.unsaved, false);
  assert.equal(fs.existsSync(path.join(recoveryDir, e.info.slot)), false);
});

/** A seal that waits for `open()` to go on: two copies of one file in flight, in the order a test sets. */
function gatedSeals(locker) {
  const real = locker.sealRecovery.bind(locker);
  let release;
  const gate = new Promise((r) => { release = r; });
  locker.sealRecovery = async (...a) => { await gate; return real(...a); };
  return { open: () => { locker.sealRecovery = real; release(); } };
}

test('an older save landing just after a newer copy is sealed leaves that copy, and the lock that follows keeps it', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Plan.docx'), 'v0');
  await locker.lockWaiting(id);
  // At the lock the window has typed nothing since version 5.
  const memory = createMemoryOpener({ locker, table, recoveryDir, unsavedOf: () => ({ bytes: Buffer.from('v5 typed'), version: 5, key: 'k' }) });
  const p = (await memory.open(id, 'Plan.docx')).path;
  const e = table.entry(p);
  // The copy of version 5 is sealed (renamed into place), and its seal has not finished when the save of version 3 lands.
  const real = locker.sealRecovery.bind(locker);
  let release;
  const gate = new Promise((r) => { release = r; });
  let sealed = false;
  locker.sealRecovery = async (...a) => {
    const out = await real(...a);
    sealed = true;
    await gate;
    return out;
  };
  const copy = e.recover(Buffer.from('v5 typed'), { version: 5, key: 'k' });
  while (!sealed) await new Promise((r) => setTimeout(r, 5));
  const saving = table.write(p, Buffer.from('v3 saved'), { version: 3, key: 'k' });
  while (!e.info.saves) await new Promise((r) => setTimeout(r, 5));
  locker.sealRecovery = real;
  release();
  assert.equal(await copy, true);
  await saving;
  assert.equal(e.info.unsaved, true, 'the copy holds newer work than the save');
  assert.ok(fs.existsSync(path.join(recoveryDir, e.info.slot)), 'and it is there');
  await memory.sealOpen();
  await table.drop(p);
  const offer = (await memory.offers()).find((o) => o.key === e.info.slot);
  assert.equal(offer?.unsaved, true);
  const back = await memory.reopen(offer.key);
  assert.equal((await table.read(back.path)).toString(), 'v5 typed');
  await table.drop(back.path);
  await memory.dismiss(offer.key);
});

test('a copy from PDF or Image still on its way when a lock reaches the file is sealed, though one from Word is refused', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Scan.pdf'), '%PDF-saved');
  await locker.lockWaiting(id);
  const memory = createMemoryOpener({ locker, table, recoveryDir });
  const p = (await memory.open(id, 'Scan.pdf')).path;
  const e = table.entry(p);
  await memory.sealOpen();
  assert.equal(e.info.final, true);
  // No version: PDF and Image hand the lock nothing, so their copy is the last word.
  assert.equal(await e.recover(Buffer.from('%PDF-late work'), { kind: 'pdf', saves: 0 }), true);
  assert.equal(await e.recover(Buffer.from('%PDF-word'), { version: 9, key: 'k' }), false, 'a versioned copy is refused');
  await table.drop(p);
  const offer = (await memory.offers()).find((o) => o.key === e.info.slot);
  assert.equal(offer?.unsaved, true);
  assert.equal((await table.read((await memory.reopen(offer.key)).path)).toString(), '%PDF-late work');
});

test('work set aside while its file is open is offered once the file is closed, never beside it while it is open', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Brief.docx'), 'saved');
  await locker.lockWaiting(id);
  const memory = createMemoryOpener({ locker, table, recoveryDir });
  const p = (await memory.open(id, 'Brief.docx')).path;
  const e = table.entry(p);
  assert.equal(await e.recover(Buffer.from('earlier work'), { version: 1, key: 'old' }), true);
  const old = e.info.slot;
  assert.equal(await e.setAside(), 'offered');
  assert.notEqual(e.info.slot, old);
  assert.equal(e.info.unsaved, false);
  assert.deepEqual((await memory.offers()).map((o) => o.key), [], 'not offered while Brief.docx is open');
  await table.drop(p);
  const offer = (await memory.offers()).find((o) => o.key === old);
  assert.equal(offer?.unsaved, true, 'offered once it is closed');
  const back = await memory.reopen(old);
  assert.equal((await table.read(back.path)).toString(), 'earlier work');
  await table.drop(back.path);
  await memory.dismiss(old);
});

test('two copies of one file\'s work reopened together open the file once, and the other copy\'s work as a new file beside it', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Brief.docx'), 'saved');
  await locker.lockWaiting(id);
  const memory = createMemoryOpener({ locker, table, recoveryDir });
  const p = (await memory.open(id, 'Brief.docx')).path;
  const e = table.entry(p);
  // Work set aside, then more typed and sealed into the file's new copy: two copies of Brief.docx.
  assert.equal(await e.recover(Buffer.from('earlier work'), { version: 1, key: 'old' }), true);
  const first = e.info.slot;
  assert.equal(await e.setAside(), 'offered');
  assert.equal(await e.recover(Buffer.from('later work'), { version: 1, key: 'new' }), true);
  const second = e.info.slot;
  await memory.sealOpen();
  await locker.lock();
  await memory.closeAll();
  await locker.unlock({ password: 'pw' });
  const offered = await memory.offers();
  assert.deepEqual(offered.map((o) => o.key).sort(), [first, second].sort());
  // "Reopen them": every key, one after another.
  const opened = [];
  for (const o of offered) opened.push(await memory.reopen(o.key));
  assert.equal(table.entries().filter((x) => x.info.rel === 'Brief.docx').length, 1, 'Brief.docx is open once');
  const beside = opened.find((o) => o.note === 'beside');
  assert.equal(beside?.beside, 'Brief (2).docx');
  assert.equal(beside.unsaved, true);
  const original = opened.find((o) => o !== beside);
  const words = [(await table.read(original.path)).toString(), (await table.read(beside.path)).toString()].sort();
  assert.deepEqual(words, ['earlier work', 'later work'], 'both copies\' work is open');
  // Saving the one beside makes a new locked file; Brief.docx is as it was.
  await table.write(beside.path, Buffer.from('beside saved'));
  assert.equal((await locker.readBytes(id, 'Brief (2).docx')).bytes.toString(), 'beside saved');
  assert.equal((await locker.readBytes(id, 'Brief.docx')).bytes.toString(), 'saved');
  // A copy with nothing unsaved for a file open already adds nothing, and goes.
  await table.drop(beside.path);
  await table.drop(original.path);
  for (const o of await memory.offers()) await memory.dismiss(o.key);
});

test('a copy with nothing unsaved, for a file that is open already, is let go rather than opened twice', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Note.docx'), 'saved');
  await locker.lockWaiting(id);
  const memory = createMemoryOpener({ locker, table, recoveryDir });
  const p = (await memory.open(id, 'Note.docx')).path;
  // A lock seals its clean copy; at the next unlock it is offered as saved.
  await memory.sealOpen();
  await locker.lock();
  await memory.closeAll();
  await locker.unlock({ password: 'pw' });
  const [offer] = await memory.offers();
  assert.equal(offer.unsaved, false);
  const q = (await memory.open(id, 'Note.docx')).path;
  const again = await memory.reopen(offer.key);
  assert.equal(again.path, null);
  assert.equal(again.already, q);
  assert.equal(table.entries().length, 1);
  assert.equal(fs.existsSync(path.join(recoveryDir, offer.key)), false, 'its copy goes');
  await table.drop(q);
  assert.ok(p);
});

test('work set aside whose copy does not open is said to be so, and a seal under way when it is set aside lands in the old copy', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Memo.docx'), 'saved');
  await locker.lockWaiting(id);
  const memory = createMemoryOpener({ locker, table, recoveryDir });
  const p = (await memory.open(id, 'Memo.docx')).path;
  const e = table.entry(p);
  assert.equal(await e.recover(Buffer.from('first'), { version: 1, key: 'old' }), true);
  // A newer copy is being sealed when the work is set aside: it lands in the old copy, and the new one is not marked unsaved.
  const gate = gatedSeals(locker);
  const copy = e.recover(Buffer.from('second'), { version: 2, key: 'old' });
  const old = e.info.slot;
  const aside = e.setAside();
  gate.open();
  assert.equal(await copy, true);
  assert.equal(await aside, 'offered');
  assert.notEqual(e.info.slot, old);
  assert.equal(e.info.unsaved, false, 'nothing marked unsaved against the new copy');
  assert.equal(fs.existsSync(path.join(recoveryDir, e.info.slot)), false);
  // A copy that will not open: kept, and the answer says it is not offered.
  assert.equal(await e.recover(Buffer.from('third'), { version: 3, key: 'old' }), true);
  const real = locker.openRecovery;
  locker.openRecovery = async () => { throw new Error('will not open'); };
  try {
    assert.equal(await e.setAside(), 'unread');
  } finally {
    locker.openRecovery = real;
  }
  assert.equal(await e.setAside(), 'none', 'nothing left to set aside');
  await table.drop(p);
  const offered = await memory.offers();
  const first = offered.find((o) => o.key === old);
  assert.equal((await table.read((await memory.reopen(first.key)).path)).toString(), 'second');
});

test('once a lock has reached a file, a copy that arrives later is refused, so it never seals over what the lock sealed', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Book.docx'), 'locked');
  await locker.lockWaiting(id);
  let answer;
  const asked = new Promise((r) => { answer = r; });
  const memory = createMemoryOpener({ locker, table, recoveryDir, unsavedOf: () => asked });
  const p = (await memory.open(id, 'Book.docx')).path;
  const e = table.entry(p);
  const lock = memory.sealOpen();
  // The autosave of an older version reaches the Locker while the lock is asking for the newest work.
  await new Promise((r) => setTimeout(r, 20));
  assert.equal(await e.recover(Buffer.from('older'), { version: 1, key: 'k' }), false, 'refused');
  answer({ bytes: Buffer.from('newest'), version: 2, key: 'k' });
  await lock;
  await table.drop(p);
  const offer = (await memory.offers()).find((o) => o.key === e.info.slot);
  assert.equal(offer?.unsaved, true);
  assert.equal((await table.read((await memory.reopen(e.info.slot)).path)).toString(), 'newest');
});

test('a copy a window sent just before a lock is waited for, and the lock does not seal an empty one over it', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Scan.pdf'), '%PDF-1.7 locked');
  await locker.lockWaiting(id);
  const memory = createMemoryOpener({ locker, table, recoveryDir, unsavedOf: () => null });
  const p = (await memory.open(id, 'Scan.pdf')).path;
  const e = table.entry(p);
  const gate = gatedSeals(locker);
  const copied = e.recover(Buffer.from('%PDF-1.7 window work'));
  const lock = memory.sealOpen();
  setTimeout(gate.open, 50);
  await lock;
  assert.equal(await copied, true);
  await table.drop(p);
  const offer = (await memory.offers()).find((o) => o.key === e.info.slot);
  assert.equal(offer?.unsaved, true, 'the window\'s copy is the one offered');
  assert.equal((await table.read((await memory.reopen(e.info.slot)).path)).toString(), '%PDF-1.7 window work');
});

test('typing during a long save, then a lock: the older save landing last does not take the newer work away', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Long.docx'), 'locked');
  await locker.lockWaiting(id);
  const memory = createMemoryOpener({ locker, table, recoveryDir, unsavedOf: () => ({ bytes: Buffer.from('typed after'), version: 2, key: 'k' }) });
  const p = (await memory.open(id, 'Long.docx')).path;
  const e = table.entry(p);
  // The save of version 1 was still being made when the lock asked; it reaches the Locker after the lock sealed version 2.
  await memory.sealOpen();
  await table.write(p, Buffer.from('saved before'), { version: 1, key: 'k' });
  assert.equal(e.info.unsaved, true, 'still unsaved: the copy is newer than the save');
  assert.equal((await locker.readBytes(id, 'Long.docx')).bytes.toString(), 'saved before');
  // The same version saved: the copy is in the file now, and goes.
  await table.write(p, Buffer.from('typed after'), { version: 2, key: 'k' });
  assert.equal(e.info.unsaved, false);
  assert.equal(fs.existsSync(path.join(recoveryDir, e.info.slot)), false);
});

test('a copy older than what a save put in the file is refused; one as new as a copy already sealed is in it', async () => {
  const { locker, folder, id, table, memory } = await make();
  fs.writeFileSync(path.join(folder, 'Ver.docx'), 'locked');
  await locker.lockWaiting(id);
  const p = (await memory.open(id, 'Ver.docx')).path;
  const e = table.entry(p);
  await table.write(p, Buffer.from('v5'), { version: 5, key: 'k' });
  assert.equal(await e.recover(Buffer.from('v4'), { version: 4, key: 'k' }), false, 'older than the save');
  assert.equal(await e.recover(Buffer.from('v6'), { version: 6, key: 'k' }), true);
  assert.equal(await e.recover(Buffer.from('v6 again'), { version: 6, key: 'k' }), true, 'as new as the copy: nothing to do');
  assert.equal((await e.recovered()).toString(), 'v6', 'and the copy is the first one');
  // A document opened again (another key) is newer than all an earlier one did.
  assert.equal(await e.recover(Buffer.from('reopened'), { version: 1, key: 'other' }), true);
  assert.equal((await e.recovered()).toString(), 'reopened');
});

test('the Locker does not let a file go under a save, a seal, or an app that says a save is on its way', async () => {
  const { locker, folder, id, table, recoveryDir } = await make();
  fs.writeFileSync(path.join(folder, 'Reap.docx'), 'locked');
  await locker.lockWaiting(id);
  let busy = true;
  const memory = createMemoryOpener({ locker, table, recoveryDir, busyOf: async () => busy });
  const p = (await memory.open(id, 'Reap.docx')).path;
  await table.read(p);
  assert.deepEqual(await memory.reap({ shown: () => false }), [], 'the app has a save on its way');
  busy = false;
  const gate = gatedSeals(locker);
  const copied = table.entry(p).recover(Buffer.from('work'));
  assert.deepEqual(await memory.reap({ shown: () => false }), [], 'a seal on its way');
  gate.open();
  await copied;
  assert.equal((await memory.reap({ shown: () => false })).length, 1, 'let go once nothing is on its way');
});
