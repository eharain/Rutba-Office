/**
 * Tracked changes — visible, and finally safe.
 *
 * The bug this closes had teeth: a `w:ins` wrapper's runs are modelled text,
 * and the kept-fragments rule preserved the wrapper WHOLE — so one keystroke
 * in a reviewed paragraph doubled the inserted words. Reviewed paragraphs
 * are structural now (a rebuild cannot re-attribute a change; resolving
 * belongs to Word), the kept rule is about content rather than tag names,
 * and the review is SHOWN: who, how much, and what a deletion removed.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

function reviewed() {
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: ['placeholder', 'Body'] }));
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace(
    '<w:p><w:r><w:t xml:space="preserve">placeholder</w:t></w:r></w:p>',
    '<w:p><w:r><w:t xml:space="preserve">Agreed </w:t></w:r>' +
    '<w:ins w:id="1" w:author="A. Reviewer"><w:r><w:t xml:space="preserve">strongly </w:t></w:r></w:ins>' +
    '<w:del w:id="2" w:author="A. Reviewer"><w:r><w:delText xml:space="preserve">reluctantly </w:delText></w:r></w:del>' +
    '<w:r><w:t xml:space="preserve">by all</w:t></w:r></w:p>',
  ));
  return openDocx(pkg.write());
}

test('a reviewed paragraph reads whole, refuses edits, and says why', () => {
  const view = reviewed();
  const b = view.block(0);
  assert.equal(b.text, 'Agreed strongly by all', 'insertions read, deletions hide — the AFTER text');
  assert.equal(b.structural, true, 'read-only until the review resolves in Word');
  assert.ok(b.structuralTags.includes('w:ins'));
  view.collapseTo({ block: 0, offset: 3 });
  assert.throws(() => view.insertText('x'), /w:ins/,
    'the refusal names the reason — one keystroke here used to DOUBLE the inserted words');
});

test('the frame summarises the review for the margin', () => {
  const frame = reviewed().render();
  const t = frame.blocks[0].tracked;
  assert.equal(t.inserted, 1);
  assert.equal(t.deleted, 1);
  assert.deepEqual(t.authors, ['A. Reviewer']);
  assert.equal(t.deletedText, 'reluctantly ', 'what the deletion removed is not lost to the reader');
  assert.ok(!('tracked' in frame.blocks[1]), 'plain paragraphs carry no ballast');
});

test('a field paragraph is not mistaken for a reviewed one', () => {
  // '<w:ins' is a prefix of '<w:instrText' — the boundary matters.
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: ['placeholder'] }));
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace(
    '<w:p><w:r><w:t xml:space="preserve">placeholder</w:t></w:r></w:p>',
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r><w:t xml:space="preserve">1</w:t></w:r></w:p>',
  ));
  const b = openDocx(pkg.write()).render().blocks[0];
  assert.equal(b.structural, true, 'a field is structural on its own account');
  assert.ok(!b.structuralTags.includes('w:ins'), 'but not called a tracked change');
  assert.ok(!b.tracked, 'and no review is invented for it');
});

test('everything tracked survives an edit ELSEWHERE, byte-identical', () => {
  const view = reviewed();
  const before = OoxmlPackage.read(view.save()).text('word/document.xml');
  const reviewedSpan = /<w:p>[\s\S]*?<\/w:p>/.exec(before)[0];

  view.collapseTo({ block: 1, offset: 0 });
  view.insertText('New ');
  const after = OoxmlPackage.read(view.save()).text('word/document.xml');
  assert.ok(after.includes(reviewedSpan), 'the reviewed paragraph kept its bytes');
  assert.equal((after.match(/strongly/g) ?? []).length, 1, 'and nothing doubled');
});
