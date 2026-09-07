/**
 * Comments — margin voices, rendered at last, still read-only.
 *
 * Preserved from the first commit and never shown; now the frame carries
 * them, each anchored to the paragraph its range starts in, and the page can
 * mark them. The write side stays deliberately closed: paragraphs carrying
 * comment anchors are structural, so an edit cannot orphan a range.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

const COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:comment w:id="1" w:author="Ejaz Arain" w:date="2026-08-27T10:00:00Z">' +
  '<w:p><w:r><w:t>Please double-check this figure.</w:t></w:r></w:p></w:comment>' +
  '<w:comment w:id="2" w:author="A. Reviewer">' +
  '<w:p><w:r><w:t>Agreed — and cite the source.</w:t></w:r></w:p></w:comment>' +
  '</w:comments>';

/** A letter whose second paragraph carries two comment ranges. */
function commented() {
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: ['Dear customer', 'The balance is 4,200', 'Yours'] }));
  const xml = pkg.text('word/document.xml').replace(
    '<w:p><w:r><w:t xml:space="preserve">The balance is 4,200</w:t></w:r></w:p>',
    '<w:p><w:commentRangeStart w:id="1"/><w:commentRangeStart w:id="2"/>' +
    '<w:r><w:t xml:space="preserve">The balance is 4,200</w:t></w:r>' +
    '<w:commentRangeEnd w:id="1"/><w:commentRangeEnd w:id="2"/>' +
    '<w:r><w:commentReference w:id="1"/></w:r><w:r><w:commentReference w:id="2"/></w:r></w:p>',
  );
  pkg.write_('word/document.xml', xml);
  pkg.addPart('word/comments.xml', COMMENTS_XML,
    'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
  return openDocx(pkg.write());
}

test('the frame carries every comment, anchored to its paragraph', () => {
  const frame = commented().render();
  assert.equal(frame.comments.length, 2);
  const [first, second] = frame.comments;
  assert.equal(first.author, 'Ejaz Arain');
  assert.equal(first.text, 'Please double-check this figure.');
  assert.equal(first.blockIndex, 1, 'anchored where the range starts');
  assert.equal(second.author, 'A. Reviewer');
  assert.equal(second.blockIndex, 1, 'two voices on one paragraph');
  assert.ok(first.date && !second.date, 'a date rides when the file has one');
});

test('a document without a comments part reports none, cheaply', () => {
  const frame = openDocx(buildDocx({ paragraphs: ['plain'] })).render();
  assert.deepEqual(frame.comments, []);
});

test('the commented paragraph stays structural — the range cannot be orphaned', () => {
  const view = commented();
  assert.equal(view.block(1).structural, true);
  assert.ok(view.block(1).structuralTags.includes('w:commentRangeStart'));
  view.collapseTo({ block: 1, offset: 3 });
  assert.throws(() => view.insertText('x'), /commentRangeStart/);
});

// -------------------------------------------------------------- authoring --

test('a new comment wires the part, anchors as a point, and the paragraph stays editable', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Dear customer', 'Body'] }));
  view.collapseTo({ block: 0, offset: 4 });
  view.addComment('Should this be Dear client?', { author: 'Ejaz Arain' });

  const pkg = OoxmlPackage.read(view.save());
  assert.ok(pkg.has('word/comments.xml'), 'the part was created');
  assert.ok(pkg.text('[Content_Types].xml').includes('comments+xml'), 'with its content type');
  assert.ok(pkg.text('word/_rels/document.xml.rels').includes('Target="comments.xml"'), 'and its relationship');

  const frame = view.render();
  assert.equal(frame.comments.length, 1);
  assert.equal(frame.comments[0].author, 'Ejaz Arain');
  assert.equal(frame.comments[0].blockIndex, 0, 'anchored to the caret paragraph');

  assert.equal(view.block(0).structural, false, 'a POINT comment does not freeze the text it discusses');
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('X');
  assert.equal(view.render().comments[0].blockIndex, 0, 'and the anchor survives the edit');
});

test('a second comment gets the next id; undoing the first-ever comment empties the part', () => {
  const view = openDocx(buildDocx({ paragraphs: ['One', 'Two'] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.addComment('first voice', { author: 'A' });
  view.collapseTo({ block: 1, offset: 0 });
  view.addComment('second voice', { author: 'B' });
  const xml = OoxmlPackage.read(view.save()).text('word/comments.xml');
  assert.ok(xml.includes('w:id="1"') && xml.includes('w:id="2"'), 'ids count up');

  view.undo();
  assert.equal(view.render().comments.length, 1, 'the second voice went');
  view.undo();
  assert.equal(view.render().comments.length, 0, 'and the first — the part is empty, not orphaned');
});

test('a comment lands even on a structural paragraph, without unlocking it', () => {
  const view = commented();
  view.collapseTo({ block: 1, offset: 0 });
  view.addComment('third voice on the locked line', { author: 'C' });
  const frame = view.render();
  assert.equal(frame.comments.length, 3);
  assert.equal(frame.comments[2].blockIndex, 1);
  assert.equal(view.block(1).structural, true, 'still locked — the ranges are still there');
});

test('editing elsewhere leaves the comments and their anchors intact', () => {
  const view = commented();
  view.collapseTo({ block: 0, offset: 'Dear'.length });
  view.insertText(' valued');
  const frame = view.render();
  assert.equal(frame.comments.length, 2, 'still two comments');
  assert.equal(frame.comments[0].blockIndex, 1, 'still anchored');
  const pkg = OoxmlPackage.read(view.save());
  assert.ok(pkg.text('word/document.xml').includes('<w:commentRangeStart w:id="1"/>'),
    'the anchors round-trip untouched');
  assert.equal(pkg.text('word/comments.xml'), COMMENTS_XML, 'the part is byte-identical');
});
