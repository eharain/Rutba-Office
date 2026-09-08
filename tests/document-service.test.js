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
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-doc-service-'));
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

  assert.match(refusal(() => doc.open({ path: asXlsx, kind: 'sheet' })), /is a document, not a workbook\. Open it in Word\./);
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

  const rtf = doc.open({ path: write('note.rtf', '{\\rtf1\\ansi A paragraph.\\par}') });
  assert.equal(rtf.converted.from, 'rtf');
  assert.equal(rtf.converted.writesBack, false, 'RTF is not written at all');
  const said = refusal(() => doc.save({ id: rtf.id, path: path.join(dir, 'note.rtf') }));
  assert.match(said, /cannot write RTF yet/);
  assert.match(said, /Save as to write \.docx/, 'and it says what it can write instead');

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
