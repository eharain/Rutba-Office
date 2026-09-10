/**
 * The document service, in node.
 *
 * Everything a window does to a file goes through here: open, convert, save,
 * export, refuse. The engines have suites of their own and the windows have
 * the application checks; this is the layer between, and it had none — which
 * is where a PDF export that threw every time, an HTML export that wrote the
 * document's own characters as markup, and a refusal that read "EPERM:
 * operation not permitted" all lived unnoticed.
 *
 * It needs no Electron: the service takes a blob holder and does the rest in
 * Node.
 */
import test from 'node:test';
import { odfFlavour } from '../packages/office-formats/src/odf.js';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-doc-service-'));
const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures');
const write = (name, bytes) => {
  const p = path.join(dir, name);
  fs.writeFileSync(p, bytes);
  return p;
};
const textOf = (model) => (model.blocks || []).map((b) => (b.runs || []).map((r) => r.text).join('') || b.text || '');
const refusal = (fn) => {
  try {
    fn();
    return null;
  } catch (err) {
    return err.message;
  }
};

test('a file that is not the kind this window edits is refused, by name', () => {
  // The extension decides which window a double-click opens, so a document
  // saved under an .xlsx name opened Worksheets, which drew a document model
  // and threw on geometry it does not have: a blank window that never
  // resolved. The corpus run sat on it for thirty seconds.
  const asXlsx = write('docx-named.xlsx', buildDocx({ paragraphs: ['A document.'] }));
  const asPptx = write('xlsx-named.pptx', buildXlsx({ sheets: [{ name: 'S', rows: [[1, 2]] }] }));
  const asDocx = write('deck-named.docx', buildPptx({ title: 'D', slides: [{ layout: 'title', title: 'T' }] }));

  assert.match(refusal(() => doc.open({ path: asXlsx, kind: 'sheet' })), /is a document, not a workbook\. Open it in Rutba Word./);
  assert.match(refusal(() => doc.open({ path: asPptx, kind: 'deck' })), /is a workbook, not a presentation\. Open it in Worksheets\./);
  assert.match(refusal(() => doc.open({ path: asDocx, kind: 'doc' })), /is a presentation, not a document\. Open it in Presentation\./);

  // The right window still opens it, and a window that names no kind is not
  // second-guessed.
  const right = doc.open({ path: asXlsx, kind: 'doc' });
  assert.equal(right.kind, 'doc');
  doc.close({ id: right.id });
  const unasked = doc.open({ path: asXlsx });
  assert.equal(unasked.kind, 'doc');
  doc.close({ id: unasked.id });
});

test('a password-protected file is refused as one, whatever its extension', () => {
  // An encrypted .docx is a compound file holding the encrypted package —
  // not a 97-2003 document — and it used to open as an empty page, or as
  // "a document, not a presentation" under a .pptx name. The fixture is
  // the sample site's password-protected document (password 123), kept
  // small and never decrypted.
  const encrypted = path.join(FIXTURES, 'encrypted.docx');
  assert.match(refusal(() => doc.open({ path: encrypted, kind: 'doc' })), /encrypted\.docx is password-protected/);
  const asPptx = write('encrypted-named.pptx', fs.readFileSync(encrypted));
  assert.match(refusal(() => doc.open({ path: asPptx, kind: 'deck' })), /is password-protected/);
});

test('the disk failing is a sentence, not an error code', () => {
  const file = write('locked.docx', buildDocx({ paragraphs: ['One.'] }));
  const session = doc.open({ path: file });

  assert.match(refusal(() => doc.open({ path: path.join(dir, 'never.docx') })), /is not there any more/);
  assert.match(refusal(() => doc.open({ path: dir })), /is a folder, not a file/);
  assert.match(refusal(() => doc.save({ id: session.id, path: path.join(dir, 'no', 'such', 'a.docx') })), /the folder it was going into is not there/);
  assert.match(refusal(() => doc.export({ id: session.id, format: 'pdf', path: path.join(dir, 'no', 'a.pdf') })), /is not there/);

  // Read-only is the one a person meets: Word holds a file open, or the copy
  // came off a locked disk.
  fs.chmodSync(file, 0o444);
  const said = refusal(() => doc.save({ id: session.id, path: file }));
  fs.chmodSync(file, 0o666);
  // Windows honours the read-only bit here; a POSIX box running as root does
  // not, and there is nothing to assert when the write succeeds.
  if (said) assert.match(said, /read-only, or another program has it open/);
  assert.ok(!/^E[A-Z]+:/.test(said || ''), 'never an error code');
  doc.close({ id: session.id });
});

test('a converted file says what saving will actually do', () => {
  // "Saving will write a .docx" was said of every converted file, and it was
  // true of almost none of them: a .md saves as .md, an .rtf cannot be saved
  // at all until it is given a new name.
  const md = doc.open({ path: write('notes.md', '# Title\n\nA paragraph.\n') });
  assert.equal(md.converted.from, 'md');
  assert.equal(md.converted.writesBack, true, 'markdown is written back as markdown');

  // RTF used to be the example of a format that could not be written at all;
  // it is written now, and the flag says so. So is OpenDocument, which the
  // installer had registered this suite as the editor of for a year.
  const rtf = doc.open({ path: write('note.rtf', '{\\rtf1\\ansi A paragraph.\\par}') });
  assert.equal(rtf.converted.from, 'rtf');
  assert.equal(rtf.converted.writesBack, true, 'rich text is written back as rich text');
  doc.save({ id: rtf.id, path: path.join(dir, 'note.rtf') });

  const odt = doc.export({ id: rtf.id, format: 'odt', path: path.join(dir, 'note.odt') });
  assert.equal(odt.format, 'odt');
  assert.equal(odfFlavour(fs.readFileSync(path.join(dir, 'note.odt'))), 'odt', 'a real OpenDocument archive');
  const back = doc.open({ path: path.join(dir, 'note.odt') });
  assert.equal(back.converted.from, 'odt');
  assert.equal(back.converted.writesBack, true, 'OpenDocument is written back as OpenDocument');
  doc.close({ id: back.id });

  const csv = doc.open({ path: write('rows.csv', 'a,b\n1,2\n') });
  assert.equal(csv.converted.writesBack, true, 'a workbook writes a CSV back');
  for (const s of [md, rtf, csv]) doc.close({ id: s.id });
});

test('export writes what it says: a real PDF, and HTML whose text is text', () => {
  const file = write('agreement.docx', buildDocx({ paragraphs: ['5 < 6 & 7 > 4', 'A <script>alert(1)</script> line', 'Fish & Chips'] }));
  const session = doc.open({ path: file });

  // renderPdf answers { buffer, pages }; handing the object to Buffer.from
  // threw on every press of Export as PDF.
  const pdfPath = path.join(dir, 'agreement.pdf');
  const answer = doc.export({ id: session.id, format: 'pdf', path: pdfPath });
  const pdf = fs.readFileSync(pdfPath);
  assert.equal(pdf.subarray(0, 5).toString('latin1'), '%PDF-');
  assert.ok(pdf.toString('latin1').trimEnd().endsWith('%%EOF'), 'and it is closed off');
  assert.ok(answer.pages >= 1);

  // The characters of a document are never markup: the round trip used to
  // lose "< 6 & 7 >" entirely and write a page that ran the script tag.
  const htmlPath = path.join(dir, 'agreement.html');
  doc.export({ id: session.id, format: 'html', path: htmlPath });
  const html = fs.readFileSync(htmlPath, 'utf8');
  assert.match(html, /5 &lt; 6 &amp; 7 &gt; 4/);
  assert.doesNotMatch(html, /<script>/, 'nothing executable reaches the file');

  const again = doc.open({ path: htmlPath });
  assert.deepEqual(textOf(again.model), ['5 < 6 & 7 > 4', 'A <script>alert(1)</script> line', 'Fish & Chips'], 'and it comes back as it went in');
  doc.close({ id: again.id });
  doc.close({ id: session.id });
});

test('every format the suite writes survives an edit, a save and a reopen', () => {
  const cases = [
    ['docx', () => buildDocx({ paragraphs: ['One.'] }), (id) => doc.apply({ id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'MARK' }] }), (m) => textOf(m).join('').includes('MARK')],
    ['xlsx', () => buildXlsx({ sheets: [{ name: 'S', rows: [['a', 1]] }] }), (id) => doc.apply({ id, ops: [{ op: 'setCell', row: 4, col: 0, value: 'MARK' }] }), (m) => (m.cells || []).some((c) => String(c.text) === 'MARK')],
    ['pptx', () => buildPptx({ title: 'D', slides: [{ layout: 'title', title: 'Hi' }] }), (id) => doc.apply({ id, ops: [{ op: 'setNotes', slide: 0, text: 'MARK' }] }), (m) => String(m.slide?.notes || '').includes('MARK')],
    ['txt', () => 'One line.\n', (id) => doc.apply({ id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'MARK' }] }), (m) => textOf(m).join('').includes('MARK')],
    ['csv', () => 'a,b\n1,2\n', (id) => doc.apply({ id, ops: [{ op: 'setCell', row: 4, col: 0, value: 'MARK' }] }), (m) => (m.cells || []).some((c) => String(c.text) === 'MARK')],
  ];
  for (const [ext, make, edit, found] of cases) {
    const file = write(`round.${ext}`, make());
    const session = doc.open({ path: file });
    edit(session.id);
    doc.save({ id: session.id, path: file });
    const again = doc.open({ path: file });
    assert.ok(found(again.model), `the edit survives a .${ext}`);
    doc.close({ id: again.id });
    doc.close({ id: session.id });
  }
});

test('a window closing takes its documents with it', () => {
  // Nothing let go of a session before this: every document ever opened
  // stayed in memory for the life of the application.
  const before = doc.sessions().length;
  const a = doc.open({ path: write('held-a.docx', buildDocx({ paragraphs: ['A.'] })) }, { id: 42 });
  const b = doc.open({ path: write('held-b.docx', buildDocx({ paragraphs: ['B.'] })) }, { id: 42 });
  const other = doc.open({ path: write('held-c.docx', buildDocx({ paragraphs: ['C.'] })) }, { id: 43 });
  assert.equal(doc.sessions().length, before + 3);

  const gone = doc.closeWindow(42);
  assert.deepEqual(gone.sort(), [a.id, b.id].sort(), 'both documents that window held, and only those');
  assert.equal(doc.sessions().length, before + 1);
  doc.close({ id: other.id });
});

test('unsaved work survives a crash, and is offered back once', () => {
  // Nothing was written until Ctrl+S, so a crash, a power cut or a closed lid
  // took everything since the last save. A dirty document is copied into the
  // profile on a timer; what is left in that folder when the application
  // starts is exactly what a crash took.
  const home = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-recovery-'));
  const recoveryDir = path.join(home, 'recovery');
  const service = createDocumentService({ holdBlob: () => ({}), recoveryDir });

  const file = write('tender.docx', buildDocx({ paragraphs: ['One.'] }));
  const session = service.open({ path: file });
  assert.deepEqual(service.autosave(), [], 'a document nobody has touched is not copied');

  service.apply({ id: session.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'UNSAVED ' }] });
  const written = service.autosave();
  assert.equal(written.length, 1, 'an edited document is');
  assert.equal(written[0].path, file, 'and it remembers where it came from');
  assert.deepEqual(service.autosave(), [], 'and is not written again until it changes again');

  // Nor immediately after it does. A copy costs what serialising the document
  // costs — twenty seconds for a 46 MB workbook, with the main process not
  // answering its windows meanwhile — so a document is copied at most every
  // half minute, and a costly one less often than that.
  service.apply({ id: session.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'MORE ' }] });
  assert.deepEqual(service.autosave(), [], 'a second change within the interval waits for it');

  // The crash: this service goes away without ever saving or closing.
  const after = createDocumentService({ holdBlob: () => ({}), recoveryDir });
  const found = after.recoverable();
  assert.equal(found.length, 1);
  assert.equal(found[0].name, 'tender.docx');
  assert.ok(found[0].size > 0);

  const back = after.recover({ file: found[0].file });
  assert.ok(textOf(back.model).join('').includes('UNSAVED'), 'the edit is there');
  assert.equal(back.path, file, 'and it knows the file it belongs to');
  assert.equal(back.dirty, true, 'and it is not pretending to be saved');

  // Saving it clears the copy: offering somebody their own saved work back
  // after the next crash is worse than not offering anything.
  after.save({ id: back.id, path: path.join(home, 'tender.docx') });
  assert.deepEqual(after.recoverable(), []);
});

test('a document that is closed properly leaves nothing to recover', () => {
  const home = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-recovery2-'));
  const recoveryDir = path.join(home, 'recovery');
  const service = createDocumentService({ holdBlob: () => ({}), recoveryDir });
  const opened = service.open({ path: write('notes2.docx', buildDocx({ paragraphs: ['One.'] })) }, { id: 7 });
  service.apply({ id: opened.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'X' }] });
  assert.equal(service.autosave().length, 1);
  assert.equal(service.recoverable().length, 1);

  // A window that closes has already asked about its unsaved work.
  service.closeWindow(7);
  assert.deepEqual(service.recoverable(), [], 'so what it held is not a crash');

  // And a copy whose file has been removed under us is dropped rather than
  // offered: a recovery list that cannot recover is worse than an empty one.
  const third = createDocumentService({ holdBlob: () => ({}), recoveryDir });
  const s = third.open({ path: write('notes3.docx', buildDocx({ paragraphs: ['One.'] })) });
  third.apply({ id: s.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'Y' }] });
  const [entry] = third.autosave();
  fs.rmSync(path.join(recoveryDir, entry.file), { force: true });
  assert.deepEqual(third.recoverable(), []);
  assert.match(String(refusal(() => third.recover({ file: entry.file }))), /no longer there/);
});

test('an RTF opened here can be saved back as an RTF', () => {
  // The installer registers this suite as the editor of .rtf; until the writer
  // existed, opening one and pressing Ctrl+S refused. The window says which
  // it is, so the flag and the behaviour have to agree.
  const source = '{' + String.fromCharCode(92) + 'rtf1' + String.fromCharCode(92) + 'ansi A paragraph.' + String.fromCharCode(92) + 'par Second.' + String.fromCharCode(92) + 'par}';
  const file = write('round.rtf', source);
  const opened = doc.open({ path: file });
  assert.equal(opened.converted.from, 'rtf');
  assert.equal(opened.converted.writesBack, true, 'and it says so on the way in');

  doc.apply({ id: opened.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'EDITED ' }] });
  doc.save({ id: opened.id, path: file });

  const again = doc.open({ path: file });
  assert.ok(textOf(again.model).join(' ').includes('EDITED A paragraph.'), 'the edit is in the file');
  assert.ok(fs.readFileSync(file, 'utf8').startsWith('{'), 'and the file is an RTF, not a document under an RTF name');
  doc.close({ id: again.id });
  doc.close({ id: opened.id });
});

test('a selection move that scrolls nothing answers with the selection, not a frame', () => {
  // An arrow key on an eighteen-million-cell workbook cost a hundred
  // milliseconds and a quarter of a megabyte: the whole viewport, rebuilt and
  // sent, for a selection that moved one cell. The cells on screen are the
  // same cells; only the selection and the active cell's fields travel.
  const s = doc.new({ kind: 'sheet' });
  const moved = doc.apply({ id: s.id, ops: [{ op: 'move', direction: 'right' }] });
  assert.ok(moved.patch && !moved.model, 'the reply is a patch');
  assert.equal(moved.patch.selection.ref, 'B1');
  assert.equal(moved.patch.selection.active.ref, 'B1');
  assert.ok('format' in moved.patch && 'formulaBar' in moved.patch && 'total' in moved.patch, 'the active cell\'s fields come with it');

  // Ctrl+click: the second rectangle joins the first, the reply says so.
  const added = doc.apply({ id: s.id, ops: [{ op: 'select', row: 0, col: 0 }, { op: 'select', row: 2, col: 2, add: true }] });
  assert.equal(added.patch.selection.ref, 'A1,C3');
  assert.equal(added.patch.selection.ranges.length, 2);
  assert.equal(added.patch.selection.active.ref, 'C3', 'typing goes where the last click was');

  // A move that scrolls needs the cells, and gets the frame.
  const far = doc.apply({ id: s.id, ops: [{ op: 'select', row: 400, col: 0 }] });
  assert.ok(far.model && !far.patch, 'a scroll answers with a frame');
  assert.equal(far.model.selection.ref, 'A401');

  // An edit is never a patch, whatever came before it.
  const typed = doc.apply({ id: s.id, ops: [{ op: 'select', row: 400, col: 1 }, { op: 'setCell', row: 400, col: 1, value: '7' }] });
  assert.ok(typed.model, 'an edit answers with a frame');
  doc.close({ id: s.id });
});
