// Things a keystroke or a ribbon button used to do to the structure of a document.
import test, { mock } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, Document } from '@rutba/ooxml';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { locate } from '@rutba/doc-view/positions';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const fixture = fs.readFileSync(new URL('./fixtures/word-edits.docx', import.meta.url));
const PNG = (shade) => Buffer.from(`iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLv${shade}AAAAAElFTkSuQmCC`, 'base64');

/** A document whose body is the XML given. */
function withBody(xml, base = buildDocx({ styles: true, paragraphs: ['x'] })) {
  const pkg = OoxmlPackage.read(base);
  const part = 'word/document.xml';
  const now = pkg.text(part);
  const open = now.indexOf('>', now.indexOf('<w:body')) + 1;
  pkg.write_(part, now.slice(0, open) + xml + now.slice(now.lastIndexOf('<w:sectPr') >= 0 ? now.lastIndexOf('<w:sectPr') : now.indexOf('</w:body>')));
  return pkg.write();
}
const bodyOf = (bytes) => OoxmlPackage.read(bytes).text('word/document.xml');
const count = (s, re) => (s.match(re) || []).length;
const balanced = (xml) => {
  assert.equal(count(xml, /<w:p[ >]/g), count(xml, /<\/w:p>/g), 'every paragraph closes');
  assert.equal(count(xml, /<w:pPr[ >]/g), count(xml, /<\/w:pPr>/g), 'every paragraph property block closes');
};
const blockWith = (view, words) => view.blocks.findIndex((b) => b.text.includes(words));

test('a table of contents with no content control, its paragraph styled, updates without cutting the file', () => {
  const toc = '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
    + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + '<w:r><w:t>Old entry</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
  const view = openDocx(withBody(`${toc}<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Real heading</w:t></w:r></w:p>`));
  view.updateTableOfContents({ pages: [1] });
  const saved = view.save();
  balanced(bodyOf(saved));
  assert.deepEqual(openDocx(saved).render().tableOfContents.entries.map((e) => e.text), ['Real heading']);
  assert.doesNotMatch(bodyOf(saved), /Old entry/);
});

const tableAbove = () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: ['Opening', 'Under the table', { text: 'A heading', style: 'Heading1' }, 'Last words'] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertTable({ rows: 2, cols: 2 });
  return view;
};

test('with a table above the caret, a table of contents goes where the caret is', () => {
  const view = tableAbove();
  view.collapseTo({ block: blockWith(view, 'Under the table'), offset: 0 });
  view.insertTableOfContents({ levels: 3, pages: [1] });
  const texts = view.render().blocks.map((b) => b.text);
  assert.ok(texts.indexOf('A heading') > texts.indexOf('Under the table'));
  const xml = bodyOf(view.save());
  assert.ok(xml.indexOf('Under the table') < xml.indexOf('<w:sdt>') && xml.indexOf('<w:sdt>') < xml.indexOf('Last words'));
});

test('with a table above the caret, a bookmark marks the paragraph the caret is in', () => {
  const view = tableAbove();
  const block = blockWith(view, 'Last words');
  view.collapseTo({ block, offset: 0 });
  view.addBookmark('Here');
  const para = /<w:p[ >](?:(?!<\/w:p>)[\s\S])*Last words(?:(?!<\/w:p>)[\s\S])*<\/w:p>/.exec(bodyOf(view.save()))[0];
  assert.match(para, /w:name="Here"/);
});

test('a building block taken with a table above it holds the words selected', () => {
  const view = tableAbove();
  const block = blockWith(view, 'Last words');
  assert.match(view.doc.portableParagraph(block), /Last words/);
});

test('a block put in with a table above the caret lands at the caret', () => {
  const view = tableAbove();
  const block = blockWith(view, 'Under the table');
  view.doc.insertParagraphsXml(block, ['<w:p><w:r><w:t>Put in</w:t></w:r></w:p>']);
  const xml = bodyOf(view.save());
  assert.ok(xml.indexOf('Under the table') < xml.indexOf('Put in') && xml.indexOf('Put in') < xml.indexOf('A heading'));
});

test('Edit in Image with a table above changes the picture asked for', () => {
  const view = tableAbove();
  for (const [words, name] of [['Under the table', 'One'], ['Last words', 'Two']]) {
    view.collapseTo({ block: blockWith(view, words), offset: 0 });
    view.insertImage({ name, contentType: 'image/png', data: PNG(''), widthPx: 10, heightPx: 10 });
  }
  const target = view.blocks.findIndex((b) => /descr="Two"|name="Two"/.test(b.xml ?? ''));
  assert.ok(target > 0, 'the second picture is found');
  view.replacePicture({ block: target, image: 0, data: PNG('AA'), contentType: 'image/png' });
  assert.equal(view.doc.pictureSource(target, 0).data.equals(PNG('AA')), true);
});

const fieldLine = '<w:p><w:r><w:t>Before </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple><w:r><w:t> after</w:t></w:r></w:p>';

for (const [where, offset] of [['before', 7], ['after', 8]]) {
  test(`Enter ${where} a field leaves one field`, () => {
    const view = openDocx(withBody(fieldLine));
    view.collapseTo({ block: 0, offset });
    view.splitParagraph();
    assert.equal(count(bodyOf(view.save()), /<w:fldSimple\b/g), 1);
  });
}

const deletedFirst = '<w:p><w:del w:id="1" w:author="R"><w:r><w:delText>gone</w:delText></w:r></w:del><w:r><w:t>kept</w:t></w:r></w:p>';

test('words typed ahead of a tracked deletion are kept', () => {
  const view = openDocx(withBody(deletedFirst));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('X');
  assert.equal(view.blocks[0].text, 'Xkept');
  assert.equal(count(bodyOf(view.save()), /<w:delText\b/g), 1, 'and the deletion stays one');
});

test('recording, words typed ahead of a deletion do not write the deletion twice, and Reject gives it back once', () => {
  const view = openDocx(withBody(deletedFirst));
  view.setTrackChanges(true, 'R');
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('X');
  assert.equal(count(bodyOf(view.save()), /<w:delText\b/g), 1);
  view.rejectChanges({ all: true });
  assert.equal(view.blocks[0].text, 'gonekept');
});

test('a paragraph of nothing but deletions takes new words', () => {
  const view = openDocx(withBody('<w:p><w:del w:id="1" w:author="R"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>'));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('New');
  assert.equal(view.blocks[0].text, 'New');
  assert.equal(count(bodyOf(view.save()), /<w:delText\b/g), 1);
});

test('the caret steps over deletions from either side', () => {
  const del = { text: '', del: { text: 'x' } };
  const a = { text: 'abc' };
  assert.equal(locate([del, a], 0, 'left').runIndex, 1);
  assert.equal(locate([a, del], 3, 'right').runIndex, 0);
});

test('a deleted tab and line break stay elements', () => {
  const view = openDocx(withBody('<w:p><w:r><w:t xml:space="preserve">a</w:t><w:tab/><w:t xml:space="preserve">b</w:t><w:br/><w:t>c</w:t></w:r></w:p>'));
  view.setTrackChanges(true, 'R');
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.deleteSelection();
  const xml = bodyOf(view.save());
  assert.doesNotMatch(xml, /<w:delText[^>]*>[^<]*[\t\n]/);
  assert.match(xml, /<w:del\b[\s\S]*<w:tab\/>[\s\S]*<w:br\/>[\s\S]*<\/w:del>/);
});

const german = () => {
  const pkg = OoxmlPackage.read(fixture);
  for (const part of ['word/document.xml', 'word/styles.xml']) pkg.write_(part, pkg.text(part).replace(/w:(val|styleId)="Heading([1-9])"/g, 'w:$1="berschrift$2"'));
  return pkg.write();
};

test('headings are known by what the style is called, not by its id', () => {
  const view = openDocx(german());
  view.collapseTo({ block: blockWith(view, 'Closing words'), offset: 0 });
  view.insertTableOfContents({ levels: 3, pages: [1, 1] });
  assert.deepEqual(view.render().tableOfContents.entries.map((e) => e.text), ['Chapter one', 'Chapter two']);
  const heading = view.render().blocks.find((b) => b.text === 'Chapter two');
  assert.equal(heading.headingLevel, 2);
});

test('a file whose contents styles have their own ids gets no second set', () => {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'A heading', style: 'Heading1' }] }));
  const part = 'word/styles.xml';
  pkg.write_(part, pkg.text(part).replace('</w:styles>', () => '<w:style w:type="paragraph" w:styleId="Verzeichnis1"><w:name w:val="toc 1"/></w:style></w:styles>'));
  const view = openDocx(pkg.write());
  view.insertTableOfContents({ levels: 1, pages: [1] });
  const saved = OoxmlPackage.read(view.save());
  assert.equal(count(saved.text(part), /w:val="toc 1"/gi), 1);
  assert.match(saved.text('word/document.xml'), /<w:pStyle w:val="Verzeichnis1"\/>/);
});

test('the paragraph that anchors a text box is a paragraph, and the box\'s words are not', () => {
  const texts = Document.open(fixture).paragraphs().map((p) => p.text);
  assert.ok(texts.includes('Box anchor words'));
  assert.ok(!texts.some((t) => t.includes('Words in the box')));
});

test('a new paragraph after a section\'s last does not start another section', () => {
  const doc = Document.open(fixture);
  const index = doc.editParagraphs().findIndex((p) => p.text.includes('Last of section'));
  doc.insertEditParagraphAfter(index, []);
  assert.equal(count(doc.xml, /<w:sectPr\b/g), 2);
});

test('a symbol reads as one character and comes back as the element', () => {
  const view = openDocx(fixture);
  const block = blockWith(view, 'Symbol here');
  const run = view.blocks[block].runs.find((r) => r.sym);
  assert.deepEqual(run.sym, { font: 'Wingdings', char: 'F0E0' });
  assert.equal(run.text.length, 1);
  assert.deepEqual(view.render().blocks.find((b) => b.text.includes('Symbol here')).runs.find((r) => r.sym).sym, run.sym);
});

const service = (recoveryDir = null) => createDocumentService({ holdBlob: () => ({ url: 'blob://held' }), recoveryDir });

function savedDocument(docs) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-word-'));
  const file = path.join(dir, 'a.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: ['Words'] }));
  const { id } = docs.open({ path: file });
  return { id, file };
}

test('Undo after a save leaves something to save', () => {
  const docs = service();
  const { id, file } = savedDocument(docs);
  docs.apply({ id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'X' }] });
  docs.save({ id, path: file });
  assert.equal(docs.undo({ id }).dirty, true);
});

test('Undo with nothing to undo leaves a clean document clean', () => {
  const docs = service();
  const { id } = savedDocument(docs);
  assert.equal(docs.undo({ id }).dirty, false);
});

test('a torn recovery list falls back to the one before it', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-recovery-'));
  const docs = service(dir);
  const type = (id) => docs.apply({ id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'X' }] });
  type(savedDocument(docs).id);
  type(savedDocument(docs).id);
  const copies = docs.autosave();
  assert.equal(copies.length, 2);
  // A second write of the list keeps the good one; the crash then tears the new one.
  docs.discardRecovery({ file: copies[1].file });
  const index = path.join(dir, 'index.json');
  assert.ok(fs.existsSync(`${index}.bak`));
  fs.writeFileSync(index, '[{"file": "half');
  assert.ok(service(dir).recoverable().length >= 1);
  assert.equal(fs.readdirSync(dir).some((n) => n.endsWith('.writing')), false, 'no half-written file is left');
});

test('a save reaches the disk before the old file is replaced', () => {
  const docs = service();
  const { id, file } = savedDocument(docs);
  docs.apply({ id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'X' }] });
  const order = [];
  const { fsyncSync, renameSync } = fs;
  const sync = mock.method(fs, 'fsyncSync', (fd) => { order.push('sync'); return fsyncSync(fd); });
  const rename = mock.method(fs, 'renameSync', (a, b) => { order.push('rename'); return renameSync(a, b); });
  try {
    docs.save({ id, path: file });
  } finally {
    sync.mock.restore();
    rename.mock.restore();
  }
  assert.deepEqual(order.slice(-2), ['sync', 'rename']);
});
