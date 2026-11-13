// Files opened from the Locker in memory, reached by Word, Sheets and Slides
// where they really run: the document service on its own thread
// (doc-worker.js, started by doc-host.js), with the table of files opened in
// memory and the Locker on the main process. The table reaches the thread
// through a mirror and requests (memory-mirror.js), and a lock reaches the
// thread's documents the other way. Everything here goes through the real
// thread: a test of the service built in this process would not see a
// thread that has no table (2.26.0, the integrator's second review).
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { tempDir } from './temp-dirs.js';
import { buildDocx } from '@rutba/ooxml/build';
import { readZip } from '@rutba/ooxml/zip';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { createMemoryOpener } from '../apps/desktop/main/extras/locker/memory.js';
import { createMemoryFiles, setMemoryFiles, unsavedOf } from '../apps/desktop/main/memory-files.js';
import { createDocumentHost } from '../apps/desktop/main/doc-host.js';

const CHEAP = { N: 1024, r: 8, p: 1 };

/** The words of a Word file's body. */
const wordsOf = (bytes) => {
  const part = readZip(Buffer.from(bytes)).entries.find((e) => e.name.replace(/^\//, '') === 'word/document.xml');
  return part ? Buffer.from(part.data).toString('utf8').replace(/<[^>]+>/g, '') : '';
};

function everyFile(dir) {
  let list = [];
  try {
    list = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return [];
  }
  const out = [];
  for (const e of list) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push(...everyFile(p));
    else out.push(p);
  }
  return out;
}

/** Files under these folders with these words in them, as bytes or as a Word file's text. */
const plainWith = (dirs, word) => dirs.flatMap(everyFile).filter((f) => {
  let b;
  try {
    b = fs.readFileSync(f);
  } catch {
    return false; // gone while looking: another test's file
  }
  if (b.includes(Buffer.from(word))) return true;
  if (b.subarray(0, 2).toString('latin1') !== 'PK') return false;
  try {
    return wordsOf(b).includes(word);
  } catch {
    return false;
  }
});

/** What Office made in the temp folder since `before` (print jobs, embedded objects, scratch files); other tests share the folder. */
const newInTemp = (seen) => fs.readdirSync(os.tmpdir()).filter((n) => !seen.has(n) && /^rutba-(print|object|office)-/.test(n)).map((n) => path.join(os.tmpdir(), n));

const settle = async (fn, ms = 5000) => {
  const until = Date.now() + ms;
  for (;;) {
    if (await fn()) return true;
    if (Date.now() > until) return false;
    await new Promise((r) => setTimeout(r, 25));
  }
};

const insert = (text) => [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text }];

let L;
before(async () => {
  const profile = tempDir('rl-thread-');
  const folder = tempDir('rl-thread-folder-');
  const locker = createLocker({ dir: profile, cost: CHEAP, large: 512 * 1024 });
  await locker.setup('pw');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: true, keepVersions: 5 });
  const table = createMemoryFiles({ root: path.join(profile, 'mem') });
  setMemoryFiles(table);
  const recoveryDir = path.join(profile, 'recovery');
  const memory = createMemoryOpener({ locker, table, recoveryDir, unsavedOf, unsavedWaitMs: 4000 });
  // The document service's own recovery folder, where a document not from the Locker keeps its copies in the clear.
  const docRecovery = tempDir('rl-thread-recovery-');
  const tempBefore = new Set(fs.readdirSync(os.tmpdir()));
  const host = createDocumentHost({ holdBlob: (b, ty, n, o) => ({ id: o?.id || 'b', url: `rutba://blob/${o?.id || 'b'}` }), releaseBlob: () => {}, blobUrl: (i) => `rutba://blob/${i}` }).start();
  host.provide({ measureMath: async () => new Map(), setting: () => {}, teach: () => {} });
  host.init({ recoveryDir: docRecovery, locale: 'en-GB', settings: {} });
  const { doc, direct } = await host.ready();
  L = { profile, folder, locker, id, table, memory, recoveryDir, docRecovery, tempBefore, host, doc, direct };
});
after(async () => {
  await L?.host.stop();
});

/** A locked Word file with these words, opened in memory, and open in the service on its thread. */
async function openLocked(name, words) {
  fs.writeFileSync(path.join(L.folder, name), buildDocx({ paragraphs: [words] }));
  await L.locker.lockWaiting(L.id);
  const { path: p } = await L.memory.open(L.id, name);
  const s = await L.doc.open({ path: p });
  return { p, s };
}

const lockedWords = async (rel) => wordsOf((await L.locker.readBytes(L.id, rel)).bytes);
const refused = (err) => err?.code === 'LOCKER_ONLY';

test('the service runs on its own thread, and opens a locked document from the table, not from the disk', async () => {
  assert.equal(L.host.threaded, true);
  const word = `OPENED${crypto.randomUUID().slice(0, 6)}`;
  const { p, s } = await openLocked('Opened.docx', word);
  assert.equal(s.path, p);
  assert.match(JSON.stringify(await L.doc.model({ id: s.id })), new RegExp(word));
  await L.doc.close({ id: s.id });
  // A check's open, which waits for its answer: the thread waits on the main process for the bytes in turn.
  const again = L.direct.open({ path: p });
  assert.match(JSON.stringify(L.direct.model({ id: again.id })), new RegExp(word));
  L.direct.close({ id: again.id });
  await L.table.drop(p);
});

test('a save from the thread is sealed into the locked file, and nothing goes out of the Locker', async () => {
  const word = `SAVED${crypto.randomUUID().slice(0, 6)}`;
  const { p, s } = await openLocked('Tender.docx', 'Tender one.');
  await L.doc.apply({ id: s.id, ops: insert(`${word} `) });
  await L.doc.save({ id: s.id, path: p });
  assert.match(await lockedWords('Tender.docx'), new RegExp(word), 'the next version of the locked file');
  // Printing knows it is from the Locker, so its job stays in memory.
  assert.equal((await L.doc.printSource({ id: s.id })).locked, true);
  // Save As, Export as PDF's check, an embedded object, letters from a merge, a comparison: all refused.
  const outside = path.join(tempDir('rl-thread-out-'), 'Tender.docx');
  await assert.rejects(L.doc.save({ id: s.id, path: outside }), refused);
  await assert.rejects(Promise.resolve().then(() => L.doc.staysInLocker({ id: s.id, to: outside.replace(/docx$/, 'pdf') })), refused);
  await assert.rejects(L.doc.objectFile({ id: s.id, part: 'word/embeddings/oleObject1.xlsx' }), refused);
  await assert.rejects(L.doc.mailMerge({ id: s.id, action: 'finish' }), refused);
  await assert.rejects(L.doc.mailMerge({ id: s.id, action: 'messages' }), refused);
  const other = path.join(tempDir('rl-thread-other-'), 'Other.docx');
  fs.writeFileSync(other, buildDocx({ paragraphs: ['Other.'] }));
  await assert.rejects(L.doc.compare({ original: p, revised: other }), refused);
  assert.equal(fs.existsSync(outside), false);
  await L.doc.close({ id: s.id });
  await L.table.drop(p);
  assert.deepEqual(plainWith([L.profile, L.folder, L.docRecovery, ...newInTemp(L.tempBefore)], word), []);
});

test('autosave and a PDF or Image window\'s copy are sealed by the Locker, never written in the clear', async () => {
  const word = `AUTOSAVED${crypto.randomUUID().slice(0, 6)}`;
  const { p, s } = await openLocked('Draft.docx', 'Draft one.');
  const entry = () => L.table.entry(p);
  await L.doc.apply({ id: s.id, ops: insert(`${word} `) });
  await L.doc.autosave();
  assert.ok(await settle(() => entry()?.info.unsaved === true), 'sealed as its locked recovery copy');
  assert.ok(fs.existsSync(path.join(L.recoveryDir, entry().info.slot)));
  // A PDF window's work, handed over by its key: sealed, and given up again when the window gives it up.
  const pdfWord = `PDFWORK${crypto.randomUUID().slice(0, 6)}`;
  const kept = await L.doc.keepCopy({ key: 'pdfwin1', kind: 'pdf', name: 'Draft.docx', path: p, bytes: Buffer.from(`%PDF-1.7 ${pdfWord}`) });
  assert.equal(kept.locker, true);
  assert.equal(kept.file, null, 'no file of its own');
  await L.doc.dropCopy({ key: 'pdfwin1' });
  assert.ok(await settle(() => entry()?.info.unsaved === false), 'given up: not unsaved');
  assert.deepEqual(everyFile(L.docRecovery), [], 'nothing in the service\'s own recovery folder');
  await L.doc.close({ id: s.id });
  await L.table.drop(p);
  assert.deepEqual(plainWith([L.profile, L.folder, L.docRecovery, ...newInTemp(L.tempBefore)], word), []);
  assert.deepEqual(plainWith([L.profile, L.folder, L.docRecovery, ...newInTemp(L.tempBefore)], pdfWord), []);
});

test('a lock asks the thread for a document\'s unsaved work, and letting the file go closes it on the thread', async () => {
  const word = `UNSAVED${crypto.randomUUID().slice(0, 6)}`;
  const { p, s } = await openLocked('Letter.docx', 'Letter one.');
  await L.doc.apply({ id: s.id, ops: insert(`${word} `) });
  assert.match(wordsOf(await unsavedOf(p)), new RegExp(word), 'asked across the thread');
  const slot = L.table.entry(p).info.slot;
  await L.memory.sealOpen();
  assert.ok(fs.existsSync(path.join(L.recoveryDir, slot)), 'sealed before the lock');
  await L.table.drop(p);
  assert.ok(await settle(async () => !(await L.doc.sessions()).some((x) => x.id === s.id)), 'the session on the thread is closed');
  await L.memory.dismiss(slot);
});

test('a lock while a save is crossing the thread loses nothing: the save is in the locked file, or its work is offered again', async () => {
  const word = `INFLIGHT${crypto.randomUUID().slice(0, 6)}`;
  const { p, s } = await openLocked('Race.docx', 'Race one.');
  await L.doc.apply({ id: s.id, ops: insert(`${word} `) });
  const saving = L.doc.save({ id: s.id, path: p }).then(() => 'saved', (err) => err?.code || 'failed');
  await L.memory.sealOpen();
  await L.locker.lock();
  await L.memory.closeAll();
  await saving;
  await L.locker.unlock({ password: 'pw' });
  const inFile = (await lockedWords('Race.docx')).includes(word);
  const offered = (await L.memory.offers()).filter((o) => o.rel === 'Race.docx');
  let inOffer = false;
  for (const o of offered) {
    if (!o.unsaved) continue;
    const back = await L.memory.reopen(o.key);
    inOffer ||= wordsOf(await L.table.read(back.path)).includes(word);
    await L.table.drop(back.path);
    await L.memory.dismiss(o.key);
  }
  for (const o of offered) await L.memory.dismiss(o.key);
  assert.ok(inFile || inOffer, JSON.stringify({ inFile, offered }));
  assert.ok(await settle(async () => !(await L.doc.sessions()).some((x) => x.id === s.id)));
});

test('a document in an older format saves from the thread as a new file in the current format, beside it', async () => {
  fs.writeFileSync(path.join(L.folder, 'Notice.htm'), '<html><body><p>Old notice.</p></body></html>');
  await L.locker.lockWaiting(L.id);
  const { path: p } = await L.memory.open(L.id, 'Notice.htm');
  const s = await L.doc.open({ path: p });
  await L.doc.apply({ id: s.id, ops: insert('NEWNOTICE ') });
  const saved = await L.doc.save({ id: s.id, path: p });
  assert.equal(path.basename(saved.path), 'Notice.docx');
  assert.match(await lockedWords('Notice.docx'), /NEWNOTICE/);
  await L.doc.close({ id: s.id });
  await L.table.drop(saved.path);
});

test('a thread started again is sent the table before its first call, and a file whose session was lost still seals and closes', async () => {
  const word = `RESTART${crypto.randomUUID().slice(0, 6)}`;
  const { p, s } = await openLocked('Restart.docx', 'Restart one.');
  await L.doc.apply({ id: s.id, ops: insert(`${word} `) });
  await L.doc.autosave();
  assert.ok(await settle(() => L.table.entry(p)?.info.unsaved === true), 'its work sealed before the thread stops');
  const was = L.host.threadId;
  await L.host.crashForCheck();
  assert.ok(await settle(() => L.host.threadId !== null && L.host.threadId !== was, 15000), 'started again');
  // The new thread knows the file: the first call names it, and the mirror is there before it.
  const again = await L.doc.open({ path: p });
  assert.match(JSON.stringify(await L.doc.model({ id: again.id })), /Restart one\./);
  await L.doc.close({ id: again.id });
  // A lock with the session gone: it does not wait for an answer that cannot come, and keeps the sealed copy.
  const slot = L.table.entry(p).info.slot;
  const started = Date.now();
  await L.memory.sealOpen();
  assert.ok(Date.now() - started < 4000, `the lock did not wait (${Date.now() - started} ms)`);
  await L.table.drop(p);
  const offer = (await L.memory.offers()).find((o) => o.key === slot);
  assert.equal(offer?.unsaved, true, 'the copy sealed before the thread stopped is offered');
  const back = await L.memory.reopen(slot);
  assert.match(wordsOf(await L.table.read(back.path)), new RegExp(word));
  await L.table.drop(back.path);
  await L.memory.dismiss(slot);
});
