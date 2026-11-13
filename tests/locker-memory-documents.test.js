// Documents, workbooks and decks opened in memory from the Locker: read from
// the table, saved by sealing through it, and never written in the clear,
// not even as a recovery copy (2.26.0, D-2026-11-13-20).
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { tempDir } from './temp-dirs.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { createMemoryFiles, setMemoryFiles, setLockedCopyCheck, unsavedOf } from '../apps/desktop/main/memory-files.js';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';

after(() => setMemoryFiles(null));

function setup() {
  const home = tempDir('rl-memdoc-');
  const root = path.join(home, 'locker', 'mem');
  const table = createMemoryFiles({ root });
  setMemoryFiles(table);
  const recoveryDir = path.join(home, 'recovery');
  const service = createDocumentService({ holdBlob: () => ({}), recoveryDir });
  return { home, root, table, service, recoveryDir };
}

function everyFile(dir) {
  const out = [];
  if (!fs.existsSync(dir)) return out;
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push(...everyFile(p));
    else out.push(p);
  }
  return out;
}

const insert = (text) => [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text }];

test('a document opened in memory is read from the table and saved by sealing through it', async () => {
  const { home, table, service } = setup();
  const sealed = [];
  const p = table.add({ name: 'Tender.docx', bytes: buildDocx({ paragraphs: ['One.'] }), save: async (b) => { sealed.push(Buffer.from(b)); return { modified: 7 }; } });
  const session = service.open({ path: p });
  assert.equal(session.name, 'Tender.docx');
  const word = `SECRET${crypto.randomUUID().slice(0, 8)}`;
  service.apply({ id: session.id, ops: insert(`${word} `) });
  const saved = await service.saveAsync({ id: session.id });
  assert.equal(sealed.length, 1);
  assert.ok(OoxmlPackage.read(sealed[0]).text('word/document.xml').includes(word), 'the sealed bytes are the edited document');
  assert.equal(saved.dirty, false);
  assert.equal(saved.stat.size, sealed[0].length);
  // The sync Save, the other door, goes the same way.
  service.apply({ id: session.id, ops: insert('MORE ') });
  await service.save({ id: session.id });
  assert.equal(sealed.length, 2);
  assert.ok(!everyFile(home).some((f) => fs.readFileSync(f).includes(Buffer.from(word))), 'nothing in the profile holds the text');
  assert.deepEqual((await table.read(p)), sealed[1], 'the table holds what was sealed');
  service.close({ id: session.id });
});

test('a text file opened in memory saves through the export writer, captured, not written', async () => {
  const { home, table, service } = setup();
  const sealed = [];
  const p = table.add({ name: 'notes.txt', bytes: Buffer.from('first line\n'), save: async (b) => { sealed.push(Buffer.from(b)); } });
  const session = service.open({ path: p });
  service.apply({ id: session.id, ops: insert('PLAINWORD ') });
  await service.saveAsync({ id: session.id });
  assert.equal(sealed.length, 1);
  assert.match(sealed[0].toString('utf8'), /PLAINWORD first line/);
  assert.ok(!everyFile(home).some((f) => fs.readFileSync(f).includes(Buffer.from('PLAINWORD'))));
  service.close({ id: session.id });
});

test('a workbook opened in memory saves the same way', async () => {
  const { table, service } = setup();
  const sealed = [];
  const p = table.add({ name: 'Budget.xlsx', bytes: buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['a', 1]] }] }), save: async (b) => { sealed.push(Buffer.from(b)); } });
  const session = service.open({ path: p, kind: 'sheet' });
  await service.saveAsync({ id: session.id });
  assert.equal(sealed.length, 1);
  assert.ok(OoxmlPackage.read(sealed[0]).text('xl/workbook.xml'));
  service.close({ id: session.id });
});

test('a save after the Locker has locked is refused in words, and writes nothing', async () => {
  const { home, table, service } = setup();
  const p = table.add({ name: 'Tender.docx', bytes: buildDocx({ paragraphs: ['One.'] }), save: async () => {} });
  const session = service.open({ path: p });
  await table.clear();
  // Let go with its file: the session is gone, and nothing was written.
  await assert.rejects(service.saveAsync({ id: session.id }));
  assert.ok(!everyFile(home).some((f) => f.endsWith('.docx')));
});

test('a write to a path under the table that no save routed is refused', async () => {
  const { table, service } = setup();
  const p = table.add({ name: 'Tender.docx', bytes: buildDocx({ paragraphs: ['One.'] }), save: async () => {} });
  const session = service.open({ path: p });
  // An export to the same folder: not one of the routed saves.
  const beside = path.join(path.dirname(p), 'Tender.odt');
  assert.throws(() => service.export({ id: session.id, path: beside, format: 'odt' }), /is in the Locker, so it is saved only back into it/);
  service.close({ id: session.id });
});

test('a recovery copy of a document opened in memory goes to the Locker, never into the recovery folder', async () => {
  const { table, service, recoveryDir } = setup();
  const recovered = [];
  const p = table.add({
    name: 'Tender.docx',
    bytes: buildDocx({ paragraphs: ['One.'] }),
    save: async () => {},
    recover: async (bytes, how) => { recovered.push({ bytes: Buffer.from(bytes), how }); },
  });
  const session = service.open({ path: p });
  service.apply({ id: session.id, ops: insert('UNSAVEDWORD ') });
  assert.deepEqual(service.autosave(), [], 'nothing listed in the plain index');
  await new Promise((r) => setTimeout(r, 20));
  assert.equal(recovered.length, 1);
  assert.ok(OoxmlPackage.read(recovered[0].bytes).text('word/document.xml').includes('UNSAVEDWORD'));
  assert.ok(!everyFile(recoveryDir).some((f) => fs.readFileSync(f).includes(Buffer.from('UNSAVEDWORD'))));
  assert.ok(!everyFile(recoveryDir).some((f) => fs.readFileSync(f).includes(Buffer.from('Tender'))), 'nor its name');
  // A window's own copy (Video, Image, PDF) of a file opened in memory goes the same way.
  const kept = service.keepCopy({ key: 'w1', kind: 'pdf', name: 'Tender.pdf', path: p, bytes: Buffer.from('%PDF-WINDOWCOPY') });
  assert.equal(kept.locker, true);
  await new Promise((r) => setTimeout(r, 20));
  assert.equal(recovered.length, 2);
  assert.ok(!everyFile(recoveryDir).some((f) => fs.readFileSync(f).includes(Buffer.from('WINDOWCOPY'))));
  service.close({ id: session.id });
});

test('with no Locker to seal it, a document opened in memory has no recovery copy at all', async () => {
  const { table, service, recoveryDir } = setup();
  const p = table.add({ name: 'Tender.docx', bytes: buildDocx({ paragraphs: ['One.'] }), save: async () => {} });
  const session = service.open({ path: p });
  service.apply({ id: session.id, ops: insert('UNSAVEDWORD ') });
  assert.deepEqual(service.autosave(), []);
  assert.equal(service.keepCopy({ key: 'w2', kind: 'pdf', name: 'x.pdf', path: p, bytes: Buffer.from('x') }), null);
  assert.ok(!everyFile(recoveryDir).some((f) => fs.readFileSync(f).includes(Buffer.from('UNSAVEDWORD'))));
  service.close({ id: session.id });
});

test('a document from the Locker is saved only back into it: Save As, Export and an object opened elsewhere are refused', async () => {
  const { home, table, service } = setup();
  const p = table.add({ name: 'Tender.docx', bytes: buildDocx({ paragraphs: ['One.'] }), save: async () => {} });
  const session = service.open({ path: p });
  const elsewhere = path.join(home, 'out', 'Tender.docx');
  const said = /Tender\.docx is in the Locker, so it is saved only back into it\. To have a copy somewhere else, use Unlock or Send in the Locker\./;
  await assert.rejects(service.saveAsync({ id: session.id, path: elsewhere }), said);
  assert.throws(() => service.save({ id: session.id, path: elsewhere, warmed: true }), said);
  assert.throws(() => service.export({ id: session.id, path: path.join(home, 'out', 'Tender.pdf'), format: 'pdf' }), said);
  assert.throws(() => service.objectFile({ id: session.id, part: 'word/embeddings/x.bin' }), said);
  assert.ok(!fs.existsSync(path.join(home, 'out')));
  // Saving it back into itself is still a save.
  await service.saveAsync({ id: session.id, path: p });
  service.close({ id: session.id });
});

test('a document the Locker opened as a plain copy is kept to it the same way', async () => {
  const { home, service } = setup();
  const copy = path.join(home, 'locker', 'open', 'x', 'Big.docx');
  fs.mkdirSync(path.dirname(copy), { recursive: true });
  fs.writeFileSync(copy, buildDocx({ paragraphs: ['One.'] }));
  setLockedCopyCheck((f) => path.resolve(f) === path.resolve(copy));
  try {
    const session = service.open({ path: copy });
    await assert.rejects(service.saveAsync({ id: session.id, path: path.join(home, 'Big.docx') }), /in the Locker/);
    await service.saveAsync({ id: session.id });
    service.close({ id: session.id });
  } finally {
    setLockedCopyCheck(null);
  }
});

test('the Locker can ask for the unsaved work of a document, and letting its file go closes the session', async () => {
  const { table, service } = setup();
  const p = table.add({ name: 'Tender.docx', bytes: buildDocx({ paragraphs: ['One.'] }), save: async () => {} });
  const session = service.open({ path: p });
  assert.equal(await unsavedOf(p), null, 'nothing unsaved yet');
  service.apply({ id: session.id, ops: insert('UNSAVEDWORD ') });
  const got = await unsavedOf(p);
  const bytes = got.bytes;
  assert.ok(Number.isFinite(got.version) && typeof got.key === 'string', 'with the version it is of');
  assert.ok(OoxmlPackage.read(Buffer.from(service.unsavedBytes({ path: p }).bytes)).text('word/document.xml').includes('UNSAVEDWORD'), 'the same, asked as the main process asks the document thread');
  assert.ok(OoxmlPackage.read(bytes).text('word/document.xml').includes('UNSAVEDWORD'));
  await table.drop(p);
  assert.ok(!service.sessions().some((x) => x.id === session.id));
});

test('a document reopened with the work a lock left unsaved opens with that work, as unsaved', async () => {
  const { table, service } = setup();
  // A .doc's sealed copy is a .docx: it opens by what is in it.
  for (const name of ['Tender.docx', 'Tender.doc', 'Tender.htm']) {
    const p = table.add({ name, bytes: buildDocx({ paragraphs: ['Saved.'] }), save: async () => {}, recovered: async () => buildDocx({ paragraphs: ['Sealed work.'] }), info: { unsaved: true } });
    const session = await service.openAsync({ path: p });
    assert.equal(session.dirty, true, name);
    assert.match(JSON.stringify(session.model), /Sealed work\./, name);
    assert.equal(session.setAside, undefined);
    service.close({ id: session.id });
    await table.drop(p);
  }
  // An open that cannot wait for the Locker refuses, rather than open the saved file over that work.
  const p = table.add({ name: 'Tender.docx', bytes: buildDocx({ paragraphs: ['Saved.'] }), save: async () => {}, recovered: async () => null, info: { unsaved: true } });
  assert.throws(() => service.open({ path: p }), (err) => err.code === 'EUNSAVED');
  await table.drop(p);
});

test('work sealed in a copy that will not open is set aside for the next unlock, and the file opens as saved, clean', async () => {
  const { table, service } = setup();
  for (const recovered of [async () => null, async () => { throw new Error('will not open'); }, null]) {
    let asides = 0;
    const info = { unsaved: true };
    const p = table.add({
      name: 'Tender.docx',
      bytes: buildDocx({ paragraphs: ['Saved.'] }),
      save: async () => {},
      recovered,
      setAside: () => { asides += 1; info.unsaved = false; return true; },
      info,
    });
    const session = await service.openAsync({ path: p });
    assert.equal(asides, 1, 'set aside');
    assert.equal(session.dirty, false, 'nothing on screen is unsaved');
    assert.equal(session.setAside, true, 'and the window is told');
    assert.match(JSON.stringify(session.model), /Saved\./);
    service.close({ id: session.id });
    await table.drop(p);
  }
  // No way to set it aside: the open refuses, and the sealed work stays as it is.
  const p = table.add({ name: 'Tender.docx', bytes: buildDocx({ paragraphs: ['Saved.'] }), save: async () => {}, recovered: async () => null, info: { unsaved: true } });
  await assert.rejects(service.openAsync({ path: p }), (err) => err.code === 'EUNSAVED');
  assert.equal(table.entry(p).info.unsaved, true);
  await table.drop(p);
});

test('a document from the Locker says so to the print service, which keeps its PDF and its print job in the Locker', async () => {
  const { home, table, service } = setup();
  const p = table.add({ name: 'Tender.docx', bytes: buildDocx({ paragraphs: ['One.'] }), save: async () => {} });
  const session = service.open({ path: p });
  assert.equal(service.printSource({ id: session.id }).locked, true);
  assert.throws(() => service.staysInLocker({ id: session.id, to: path.join(home, 'Tender.pdf') }), (err) => err.code === 'LOCKER_ONLY');
  service.close({ id: session.id });
  const plain = path.join(home, 'Plain.docx');
  fs.writeFileSync(plain, buildDocx({ paragraphs: ['Two.'] }));
  const other = service.open({ path: plain });
  assert.equal(service.printSource({ id: other.id }).locked, false);
  assert.doesNotThrow(() => service.staysInLocker({ id: other.id, to: path.join(home, 'Plain.pdf') }));
  service.close({ id: other.id });
});

test('a document in an older format opened in memory saves as a new file in the current format beside it, and exports stay refused', async () => {
  const { table, service } = setup();
  const asked = [];
  const p = table.add({
    name: 'Notice.htm',
    bytes: Buffer.from('<html><body><p>Old page.</p></body></html>'),
    save: async () => { throw new Error('the .htm itself is never written'); },
    saveAs: async (name, bytes) => { asked.push({ name, bytes: Buffer.from(bytes) }); return { name }; },
  });
  const session = service.open({ path: p });
  service.apply({ id: session.id, ops: insert('NEWWORD ') });
  // Save, as Word and Slides do it: to the file's own path.
  const saved = await service.save({ id: session.id, path: p });
  assert.equal(path.basename(saved.path), 'Notice.docx');
  assert.equal(asked.length, 1);
  assert.equal(asked[0].name, 'Notice.docx');
  assert.ok(OoxmlPackage.read(asked[0].bytes).text('word/document.xml').includes('NEWWORD'), 'in the current format');
  assert.equal(path.dirname(saved.path), path.dirname(p), 'under the same token: the same locked folder');
  assert.equal(service.get?.({ id: session.id })?.dirty ?? false, false);
  // Save As with a name the app proposes beside it (Sheets proposes Name.xlsx; the shell answers that path).
  const again = table.add({ name: 'Old.htm', bytes: Buffer.from('<p>x</p>'), save: async () => {}, saveAs: async (name) => { asked.push({ name }); return { name }; } });
  const other = service.open({ path: again });
  await service.saveAsync({ id: other.id, path: path.join(path.dirname(again), 'Old copy.docx') });
  assert.equal(asked.at(-1).name, 'Old copy.docx');
  // Export as anything beside it is still not a save, and is refused.
  assert.throws(() => service.export({ id: other.id, path: path.join(path.dirname(again), 'Old.odt'), format: 'odt' }), (err) => err.code === 'LOCKER_ONLY');
  service.close({ id: session.id });
  service.close({ id: other.id });
});
