/**
 * Binary files built to hurt: a compound file whose tables loop or lie, and
 * a workbook script that reaches past the grid.
 *
 * A reader of Word, Excel and PowerPoint 97-2003 runs in the main process
 * on whatever a person double-clicks, so a chain that points back at itself,
 * or a size field of two billion on a four-kilobyte file, must end in a
 * sentence in a few milliseconds, not in a hang, a gigabyte or a stack. Each
 * test here builds the smallest file that did that and times the answer.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { CompoundFile } from '@rutba/office-formats/cfb';
import { entriesIn, readFib, readLists } from '../packages/office-formats/src/msdoc.js';
import { fileURLToPath } from 'node:url';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { runScript } from '@rutba/sheet-view/scripts';
import { encryptPackage, decryptPackage, dataSpacesStreams, EncryptedFileError } from '@rutba/office-formats/crypt';
import { writeCompoundFile } from '@rutba/office-formats/cfb-write';
import { buildDocx } from '@rutba/ooxml/build';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const END = 0xfffffffe;
const FREE = 0xffffffff;

/** A header, one FAT sector (sector 0) and the directory (sector 1) with a root and one stream. */
function compound({ fat = [], rootChild = FREE, streamLeft = FREE, streamStart = FREE, streamSize = 0, nameLen = 8 } = {}) {
  const b = Buffer.alloc(512 * 4);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(b);
  b.writeUInt16LE(3, 26);
  b.writeUInt16LE(9, 30);
  b.writeUInt16LE(6, 32);
  b.writeUInt32LE(1, 44);
  b.writeUInt32LE(1, 48);
  b.writeUInt32LE(4096, 56);
  b.writeUInt32LE(END, 60);
  b.writeUInt32LE(END, 68);
  for (let i = 0; i < 109; i++) b.writeUInt32LE(FREE, 76 + i * 4);
  b.writeUInt32LE(0, 76);
  const table = [0xfffffffd, END, ...fat];
  table.forEach((v, i) => b.writeUInt32LE(v >>> 0, 512 + i * 4));
  const root = 1024;
  b[root + 66] = 5;
  b.writeUInt32LE(FREE, root + 68);
  b.writeUInt32LE(FREE, root + 72);
  b.writeUInt32LE(rootChild, root + 76);
  const e = root + 128;
  Buffer.from('S\0t\0r\0', 'latin1').copy(b, e);
  b.writeUInt16LE(nameLen, e + 64);
  b[e + 66] = 2;
  b.writeUInt32LE(streamLeft, e + 68);
  b.writeUInt32LE(FREE, e + 72);
  b.writeUInt32LE(FREE, e + 76);
  b.writeUInt32LE(streamStart, e + 116);
  b.writeUInt32LE(streamSize, e + 120);
  return b;
}

const quick = (fn, limitMs = 1500) => {
  const t = Date.now();
  const out = fn();
  assert.ok(Date.now() - t < limitMs, `took ${Date.now() - t} ms`);
  return out;
};

test('a directory tree whose entry names itself as a sibling is read once, not recursed on', () => {
  const cfb = quick(() => new CompoundFile(compound({ rootChild: 1, streamLeft: 1 })));
  assert.deepEqual(cfb.childrenOf(cfb.root).map((c) => c.name), ['Str']);
});

test('a sector chain that loops, on a stream said to be two gigabytes, reads no more than the file holds', () => {
  const cfb = new CompoundFile(compound({ fat: [2], rootChild: 1, streamStart: 2, streamSize: 0x7fffffff }));
  const bytes = quick(() => cfb.read(cfb.find(['Str'])));
  assert.ok(bytes.length <= 2048, `${bytes.length} bytes came back from a 2 KB file`);
});

test('a DIFAT chain that points at itself ends at the size of the file', () => {
  const b = compound({ rootChild: 1 });
  b.writeUInt32LE(2, 68);
  b.writeUInt32LE(1, 72);
  for (let i = 0; i < 127; i++) b.writeUInt32LE(0, 1536 + i * 4);
  b.writeUInt32LE(2, 1536 + 127 * 4);
  const cfb = quick(() => new CompoundFile(b));
  assert.ok(cfb.fat.length < 1e6, 'the table is bounded by the file, not by a million turns of the loop');
});

test('a DIFAT sector named beyond the end of a cut file is passed over', () => {
  const b = compound({ rootChild: 1 });
  b.writeUInt32LE(900, 68);
  b.writeUInt32LE(1, 72);
  const cfb = new CompoundFile(b);
  assert.ok(cfb.entries.length >= 1);
});

test('a directory name length past the 64 bytes an entry has does not read off the sector', () => {
  const b = compound({ rootChild: 1, nameLen: 0xffff });
  // The last entry of the last sector is where the read ran out of bytes.
  const cfb = new CompoundFile(b.subarray(0, 1024 + 256));
  assert.ok(cfb.entries.length >= 1);
});

test('a table that claims billions of entries in a few bytes counts only those that fit', () => {
  assert.equal(entriesIn(0xfffffff0, 12, 0, 100), 8, '100 bytes hold 8 entries of 12 after the 4-byte lead');
  assert.equal(entriesIn(3, 8, 0, 100), 0, 'a length shorter than its lead holds none');
  assert.equal(entriesIn(400, 8, 200, 100), 0, 'a table that starts past the stream holds none');
});

test('a Word list table whose counts are as large as their fields go reads only what its bytes hold', () => {
  // The overrides: a count of 0xFFFFFFFF in front of one entry.
  const lfo = Buffer.alloc(4 + 16);
  lfo.writeUInt32LE(0xffffffff, 0);
  const t = Date.now();
  const { lfos } = readLists(lfo, { fc: 0, lcb: 0 }, { fc: 0, lcb: lfo.length });
  assert.equal(lfos.length, 1);
  assert.ok(Date.now() - t < 200);
});

test('a Word list table counting 32,767 lists in a few bytes, the text of a level said to be 65,535 characters, reads no more than its bytes hold', () => {
  const t = Date.now();
  const lst = Buffer.alloc(2 + 28 + 28 + 4);
  lst.writeInt16LE(0x7fff, 0);
  lst.writeUInt16LE(0xffff, 2 + 28 + 28);
  const { lists } = readLists(lst, { fc: 0, lcb: 2 + 28 }, { fc: 0, lcb: 0 });
  assert.ok(lists.length >= 1 && lists.length <= Math.floor((lst.length - 2) / 28), String(lists.length));
  assert.ok(lists.every((l) => l.levels.every((v) => v.codes.length <= 1)));
  assert.ok(Date.now() - t < 200);
});

test('a Word file whose list overrides are counted at four billion opens in a moment, as the 2.25.1 gate\'s fuzz found', () => {
  // paragraphs.doc with 0xFF over the count at the head of its PlfLfo: the reader filled an array until the process ran out of memory.
  const good = fs.readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'binary', 'paragraphs.doc'));
  const cf = new CompoundFile(good);
  const fib = readFib(cf.read(cf.find(['WordDocument'])));
  const table = cf.read(cf.find([fib.fWhichTblStm ? '1Table' : '0Table']));
  const { fc, lcb } = fib.pair(74);
  assert.ok(lcb > 4, 'the fixture has list overrides');
  const run = Buffer.from(table.subarray(fc, fc + lcb));
  const at = good.indexOf(run);
  assert.ok(at > 0 && good.indexOf(run, at + 1) === -1, 'the overrides are found once in the file');
  const bad = Buffer.from(good);
  bad.fill(0xff, at, at + 4);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-hostile-'));
  try {
    const file = path.join(dir, 'overrides.doc');
    fs.writeFileSync(file, bad);
    const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
    const words = (p) => { const o = service.open({ path: p }); const m = service.model({ id: o.id }); service.close({ id: o.id }); return JSON.stringify(m.blocks); };
    const t = Date.now();
    const damaged = words(file);
    assert.ok(Date.now() - t < 2000, `${Date.now() - t} ms`);
    // Only a count was damaged: the paragraphs are all there.
    fs.writeFileSync(path.join(dir, 'good.doc'), good);
    assert.equal(JSON.parse(damaged).length, JSON.parse(words(path.join(dir, 'good.doc'))).length);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a Word file whose compound container loops is refused in a sentence by the service', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-hostile-'));
  try {
    const file = path.join(dir, 'loop.doc');
    fs.writeFileSync(file, compound({ rootChild: 1, streamLeft: 1 }));
    const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
    const t = Date.now();
    assert.throws(() => service.open({ path: file }), (err) => {
      assert.doesNotMatch(err.message, /call stack|DataView|undefined|Cannot read|is not a function/);
      return /\.doc/.test(err.message) || /Word/.test(err.message) || /document/.test(err.message);
    });
    assert.ok(Date.now() - t < 2000);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

const book = () => new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: [['a', 'b'], [1, 2]] }] }));

test('a script clearing every column of the sheet works through the used cells, not sixteen thousand columns of a million rows', () => {
  const view = book();
  const out = runScript(`function main(w) { w.getActiveWorksheet().getRange('A:XFD').clear(); }`, view.scriptSnapshot());
  assert.equal(out.error, null);
  const cells = quick(() => view.applyScriptEdits(out.edits));
  assert.equal(cells, 4);
  assert.equal(view.displayValue(1, 1).text, '');
});

test('a script cannot write a cell Excel has no room for, and nothing is changed when it tries', () => {
  const view = book();
  for (const code of [
    `function main(w) { w.getActiveWorksheet().getCell(2000000, 0).setValue('far'); }`,
    `function main(w) { w.getActiveWorksheet().getCell(0, 99999).setValue('wide'); }`,
    `function main(w) { w.getActiveWorksheet().getRange('A1').setValue('kept?'); w.getActiveWorksheet().getCell(-1, 0).setValue('above'); }`,
  ]) {
    const out = runScript(code, view.scriptSnapshot());
    assert.throws(() => view.applyScriptEdits(out.edits), /outside the worksheet/);
  }
  assert.equal(view.displayValue(0, 0).text, 'a', 'the refused script left the sheet as it was');
});

test('an encrypted package that asks for four billion rounds of hashing is refused, not spun', () => {
  const sealed = encryptPackage(buildDocx({ paragraphs: ['Secret.'] }), 'pw', { spinCount: 1000 });
  const cfb = new CompoundFile(sealed);
  const info = Buffer.from(cfb.read(cfb.find(['EncryptionInfo']))).toString('latin1').replace('spinCount="1000"', 'spinCount="4000000000"');
  const forged = writeCompoundFile([
    { path: ['EncryptionInfo'], data: Buffer.from(info, 'latin1') },
    { path: ['EncryptedPackage'], data: cfb.read(cfb.find(['EncryptedPackage'])) },
    ...dataSpacesStreams(),
  ]);
  const t = Date.now();
  assert.throws(() => decryptPackage(forged, 'pw'), (err) => err instanceof EncryptedFileError && /spin count/.test(err.message));
  assert.ok(Date.now() - t < 1000);
});

test('an attachment named to hide that it is a program is still opened under its real extension', async () => {
  const { attachmentExtension } = await import('@rutba/mailbox/mime');
  assert.equal(attachmentExtension('invoice.PDF'), 'pdf');
  assert.equal(attachmentExtension('update.exe '), 'exe', 'Windows drops the trailing space and runs it');
  assert.equal(attachmentExtension('update.exe::$DATA'), 'exe', 'and the default stream name');
  assert.equal(attachmentExtension('update.scr\t'), 'scr');
  assert.equal(attachmentExtension('notes.x/../y'), '', 'a slash is not in an extension; it cannot name a folder');
  assert.equal(attachmentExtension('README'), 'readme');
  assert.equal(attachmentExtension(null), '');
});

test('a message that is nothing but unclosed tags is read in a moment, by the preview and by the junk filter', async () => {
  const { stripHtml } = await import('@rutba/mailbox/mime');
  const { tokensOf } = await import('../apps/desktop/main/mail-junk.js');
  // 200 KB of "<" took 40 seconds in the junk filter and as long again in the
  // preview, in the process that holds every window: quadratic, from one message.
  for (const html of ['<'.repeat(200000), '<script '.repeat(25000), '<style>x</style>kept<script>'.repeat(5000)]) {
    quick(() => stripHtml(html), 500);
    quick(() => tokensOf({ html, subject: 'hello' }), 500);
  }
  assert.equal(stripHtml('a<script>alert(1)</script>b<style>p{}</style>c<b>d</b>'), 'a b c d');
  assert.equal(stripHtml('x<script>never closed'), 'x');
  assert.ok(tokensOf({ html: '<p>Invoice due</p><script>x</script>', subject: 's' }).includes('invoice'));
});

test('the reading pane draws a message of unclosed tags and image openers without stalling the window', async () => {
  const { bodyDocument, stripTags } = await import('../apps/desktop/renderer/apps/mail/parts.js');
  for (const html of ['<img '.repeat(60000), '<'.repeat(100000)]) {
    quick(() => bodyDocument({ html }), 500);
    quick(() => stripTags(html), 500);
  }
  assert.match(bodyDocument({ html: '<img src="https://x.test/p.png">' }), /data-blocked-remote="https:\/\/x\.test\/p\.png"/, 'a remote picture is still held back');
});
