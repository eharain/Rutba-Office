// Bookmarks — the engine's half.
//
// Word's older anchor: a name on a span of paragraphs rather than a content
// control's tag on a spot. `w:bookmarkStart` and `w:bookmarkEnd` are written
// where Word writes them, a bookmarked paragraph stays editable (a change
// from before, when carrying a bookmark made a paragraph read-only), and the
// name follows Word's own rule.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const fixture = () => openDocx(buildDocx({
  styles: true,
  paragraphs: [
    { text: 'First paragraph text', style: 'Heading1' },
    { text: 'Second paragraph text' },
    { text: 'Third paragraph text' },
    { text: 'Fourth paragraph text' },
  ],
}));

const pXml = (view, i) => view.doc.doc.paragraph(i).xml;

test('addBookmark writes bookmarkStart right after pPr and bookmarkEnd right before the paragraph closes, and bookmarks() reads the span back', () => {
  const view = fixture();
  const engine = view.doc.doc;

  const id = engine.addBookmark('Summary', 1, 2);
  assert.equal(id, 0, 'the first bookmark in an empty body is id 0');

  assert.match(pXml(view, 1), /^<w:p\b[^>]*>(<w:pPr>[\s\S]*?<\/w:pPr>)?<w:bookmarkStart w:id="0" w:name="Summary"\/>/, 'the start, right after pPr');
  assert.match(pXml(view, 2), /<w:bookmarkEnd w:id="0"\/><\/w:p>$/, 'the end, right before the paragraph closes');

  assert.deepEqual(engine.bookmarks(), [{ id: 0, name: 'Summary', from: 1, to: 2 }]);

  const id2 = engine.addBookmark('Notes', 3, 3);
  assert.equal(id2, 1, 'the second bookmark is one past the highest id already in the body');
  assert.deepEqual(engine.bookmarks().find((b) => b.name === 'Notes'), { id: 1, name: 'Notes', from: 3, to: 3 });

  // Word's own Add replaces a bookmark of the same name rather than stacking
  // a second one on it.
  engine.addBookmark('Summary', 0, 0);
  const summaries = engine.bookmarks().filter((b) => b.name === 'Summary');
  assert.equal(summaries.length, 1, 'still one Summary');
  assert.deepEqual(summaries[0], { id: 2, name: 'Summary', from: 0, to: 0 }, 'moved to the new span, a fresh id');
  assert.doesNotMatch(pXml(view, 1), /w:name="Summary"/, 'the old span no longer carries it');
});

test('a bookmark name follows Word\'s own rule', () => {
  const view = fixture();
  const engine = view.doc.doc;
  const message = /a bookmark name is letters, digits and underscores, starting with a letter, up to 40 characters/;

  assert.throws(() => engine.addBookmark('1Start', 0), message, 'starting with a digit');
  assert.throws(() => engine.addBookmark('has space', 0), message, 'a space');
  assert.throws(() => engine.addBookmark('_GoBack', 0), message, 'starting with an underscore');
  assert.throws(() => engine.addBookmark('a'.repeat(41), 0), message, 'forty-one characters');
  assert.doesNotThrow(() => engine.addBookmark('a'.repeat(40), 0), 'forty characters is still fine');
});

test('removeBookmark takes both marks out, and answers false the second time', () => {
  const view = fixture();
  const engine = view.doc.doc;
  engine.addBookmark('Summary', 1, 2);

  assert.equal(engine.removeBookmark('Summary'), true);
  assert.doesNotMatch(pXml(view, 1), /w:bookmarkStart/);
  assert.doesNotMatch(pXml(view, 2), /w:bookmarkEnd/);
  assert.deepEqual(engine.bookmarks(), []);

  assert.equal(engine.removeBookmark('Summary'), false, 'nothing left to remove');
});

test('a bookmarked paragraph stays editable: typing keeps the mark on it', () => {
  const view = fixture();
  view.doc.doc.addBookmark('Note', 1, 1);

  view.setSelection({ block: 1, offset: 0 });
  view.insertText('X');

  const xml = pXml(view, 1);
  assert.match(xml, /^<w:p\b[^>]*>(<w:pPr>[\s\S]*?<\/w:pPr>)?<w:bookmarkStart w:id="0" w:name="Note"\/>/, 'the start stayed at the paragraph\'s own start');
  assert.match(xml, /<w:bookmarkEnd w:id="0"\/><\/w:p>$/, 'the end stayed at the paragraph\'s own close');
  assert.equal(view.blocks[1].text, 'XSecond paragraph text');
  assert.equal(view.blocks[1].structural, false, 'the paragraph is not read-only');
});

test('Enter mid-paragraph keeps the bookmark on the first half', () => {
  const view = fixture();
  view.doc.doc.addBookmark('Note', 1, 1);

  // Split "Second paragraph text" after "Second" (offset 6), in the one run.
  view.setSelection({ block: 1, offset: 6 });
  view.splitParagraph();

  assert.equal(view.blocks[1].text, 'Second');
  assert.equal(view.blocks[2].text, ' paragraph text');
  const first = pXml(view, 1);
  assert.match(first, /<w:bookmarkStart w:id="0" w:name="Note"\/>/, 'the start stayed on the first half');
  assert.match(first, /<w:bookmarkEnd w:id="0"\/><\/w:p>$/, 'the end stayed on the first half too');
  assert.doesNotMatch(pXml(view, 2), /w:bookmark/, 'the second half carries neither');
});

test('merging keeps the bookmark', () => {
  const view = fixture();
  view.doc.doc.addBookmark('Note', 1, 1);

  // Backspace at the start of the third paragraph merges it into the bookmarked one.
  view.setSelection({ block: 2, offset: 0 });
  view.deleteBackward();

  assert.equal(view.blocks[1].text, 'Second paragraph textThird paragraph text');
  const merged = pXml(view, 1);
  assert.match(merged, /<w:bookmarkStart w:id="0" w:name="Note"\/>/);
  assert.match(merged, /<w:bookmarkEnd w:id="0"\/><\/w:p>$/);
});

test('a paragraph carrying Word\'s own _GoBack lists no bookmarks', () => {
  const view = fixture();
  const engine = view.doc.doc;
  // _GoBack is Word's own housekeeping bookmark — never offered through
  // addBookmark (its name breaks the name rule), so it is spliced in by
  // hand here, the way Word itself writes it.
  const p0 = engine.paragraph(0);
  const withGoBack = p0.open + (p0.pPr ?? '') + '<w:bookmarkStart w:id="9" w:name="_GoBack"/>'
    + p0.xml.slice(p0.open.length + (p0.pPr ? p0.pPr.length : 0), p0.xml.length - '</w:p>'.length)
    + '<w:bookmarkEnd w:id="9"/></w:p>';
  engine._spliceBody(p0.start, p0.end, withGoBack);

  assert.deepEqual(engine.bookmarks(), []);
});

test('a bookmark Word wrote mid-paragraph reads with the right paragraph, and the paragraph is not structural', () => {
  const view = fixture();
  const engine = view.doc.doc;
  // Hand-written the way Word itself splits a run around a mid-paragraph mark.
  const before = pXml(view, 2);
  assert.equal(before, '<w:p><w:r><w:t xml:space="preserve">Third paragraph text</w:t></w:r></w:p>');
  const p2 = engine.paragraph(2);
  const handWritten = '<w:p><w:r><w:t xml:space="preserve">Third </w:t></w:r>'
    + '<w:bookmarkStart w:id="7" w:name="Mid"/>'
    + '<w:r><w:t xml:space="preserve">paragraph text</w:t></w:r>'
    + '<w:bookmarkEnd w:id="7"/></w:p>';
  engine._spliceBody(p2.start, p2.end, handWritten);

  assert.deepEqual(engine.bookmarks(), [{ id: 7, name: 'Mid', from: 2, to: 2 }]);
  assert.equal(engine.paragraph(2).structural, false);
});

test('gotoBookmark selects the whole span, and answers false for an unknown name', () => {
  const view = fixture();
  view.doc.doc.addBookmark('Summary', 1, 2);

  const found = view.gotoBookmark('Summary');
  assert.equal(found, true);
  assert.deepEqual(view.selection.from, { block: 1, offset: 0 });
  assert.deepEqual(view.selection.to, { block: 2, offset: view.blocks[2].text.length });

  view.setSelection({ block: 0, offset: 0 });
  const missing = view.gotoBookmark('Nope');
  assert.equal(missing, false, 'unknown name');
  assert.deepEqual(view.selection.from, { block: 0, offset: 0 }, 'no edit, selection untouched');
});

test('a bookmark survives a save and reopen', () => {
  const view = fixture();
  view.doc.doc.addBookmark('Summary', 1, 2);

  const reopened = openDocx(view.save());
  assert.deepEqual(reopened.doc.doc.bookmarks(), [{ id: 0, name: 'Summary', from: 1, to: 2 }]);
  assert.match(pXml(reopened, 1), /<w:bookmarkStart w:id="0" w:name="Summary"\/>/);
  assert.match(pXml(reopened, 2), /<w:bookmarkEnd w:id="0"\/><\/w:p>$/);
});

test('undo after addBookmark takes it off again', () => {
  const view = fixture();
  view.setSelection({ block: 1, offset: 0 });
  view.addBookmark('Summary');

  assert.deepEqual(view.doc.doc.bookmarks(), [{ id: 0, name: 'Summary', from: 1, to: 1 }]);
  view.undo();
  assert.deepEqual(view.doc.doc.bookmarks(), []);
});
