// Review → Delete in Documents: a comment goes from every place it lives.
//
// A comment Word wrote lives in five places: its entry in the comments part,
// its range marks and reference run in the text, and three parts beside the
// comments — the thread and resolved state (commentsExtended), the durable
// id (commentsIds) and the date (commentsExtensible). A delete that missed
// one would leave Word an orphan to complain about, or a reply hanging on
// nothing. The comment in the showcase document is Word's own.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Document } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const SHOWCASE = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'rich', 'showcase.docx');
const texts = (doc) => doc.editParagraphs().map((p) => p.text);

test('a comment Word wrote goes from the text and from all four comment parts, and nothing else moves', () => {
  const doc = Document.open(fs.readFileSync(SHOWCASE));
  const before = texts(doc);
  assert.deepEqual(doc.comments().map((c) => c.id), ['9']);

  assert.equal(doc.deleteComments(['9']), 1);
  const back = Document.open(doc.save());
  assert.deepEqual(back.comments(), [], 'no comment is left');
  assert.deepEqual(texts(back), before, 'every paragraph reads as it did');
  const body = back.pkg.text('word/document.xml');
  assert.ok(!/<w:comment(?:RangeStart|RangeEnd|Reference)\b/.test(body), 'no range marks and no reference');
  assert.ok(!/<w15:commentEx\b/.test(back.pkg.text('word/commentsExtended.xml')), 'the thread entry is gone');
  assert.ok(!/<w16cid:commentId\b/.test(back.pkg.text('word/commentsIds.xml')), 'the durable id is gone');
  assert.ok(!/<w16cex:commentExtensible\b/.test(back.pkg.text('word/commentsExtensible.xml')), 'the date is gone');
});

test('a reply goes with the comment it answers', () => {
  const doc = Document.open(fs.readFileSync(SHOWCASE));
  // A reply as Word writes one: its own comment, whose paragraph
  // commentsExtended ties to the parent's paragraph.
  const reply = '<w:comment w:id="10" w:author="B. Colleague" w:date="2026-09-11T09:00:00Z"><w:p w14:paraId="0A0B0C0D"><w:r><w:t>Agreed.</w:t></w:r></w:p></w:comment>';
  doc.pkg.write_('word/comments.xml', doc.pkg.text('word/comments.xml').replace('</w:comments>', () => reply + '</w:comments>'));
  doc.pkg.write_('word/commentsExtended.xml', doc.pkg.text('word/commentsExtended.xml').replace('</w15:commentsEx>', () => '<w15:commentEx w15:paraId="0A0B0C0D" w15:paraIdParent="5E4178CE" w15:done="0"/></w15:commentsEx>'));
  assert.equal(doc.comments().length, 2);

  assert.equal(doc.deleteComments(['9']), 2, 'the comment and its reply');
  assert.deepEqual(doc.comments(), []);
});

test('deleting is one step, undone with everything put back, and a reference sharing words keeps them', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Alpha.', 'Beta.'] }));
  view.setSelection({ block: 0, offset: 0 });
  view.addComment('first', { author: 'A' });
  view.setSelection({ block: 1, offset: 0 });
  view.addComment('second', { author: 'A' });
  const ids = view.doc.doc.comments().map((c) => c.id);
  assert.equal(ids.length, 2);

  assert.equal(view.deleteComments([ids[0]]), 1);
  assert.deepEqual(view.doc.doc.comments().map((c) => c.text), ['second'], 'only the one asked for');
  assert.deepEqual(view.doc.doc.editParagraphs().map((p) => p.text), ['Alpha.', 'Beta.'], 'the words are untouched');

  view.undo();
  assert.deepEqual(view.doc.doc.comments().map((c) => c.text), ['first', 'second'], 'undo puts it back');

  // A reference inside a run that also holds words loses only the reference.
  const doc = Document.open(buildDocx({ paragraphs: ['Gamma.'] }));
  doc.addComment(0, { author: 'A', text: 'third' });
  const id = doc.comments()[0].id;
  doc.xml = doc.xml.replace('<w:r><w:commentReference w:id="' + id + '"/></w:r>', () => '<w:r><w:commentReference w:id="' + id + '"/><w:t>!</w:t></w:r>');
  doc.deleteComments([id]);
  assert.equal(doc.editParagraphs()[0].text, 'Gamma.!', 'the shared run keeps its words');
});
