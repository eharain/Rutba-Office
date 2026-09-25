/**
 * Tracked changes — visible, and finally safe to sit beside.
 *
 * The bug this once closed had teeth: a `w:ins` wrapper's runs are modelled
 * text, and the kept-fragments rule preserved the wrapper WHOLE — so one
 * keystroke in a reviewed paragraph doubled the inserted words. Reviewed
 * paragraphs were made structural to stop that. Recording tracked changes
 * (see word-track.test.js) needed them editable again — typing beside a
 * change, your own pending one or somebody else's still open, is ordinary
 * work, not something a review should block — so `parseRuns`/`renderRuns`
 * now carry `w:ins`/`w:del` as a run property of their own (the same idea
 * as a hyperlink's group wrapper), and a rebuild puts the wrapper back
 * correctly instead of never touching the paragraph at all. The review is
 * still SHOWN: who, how much, and what a deletion removed.
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

test('a reviewed paragraph reads whole, stays editable, and a keystroke inside the insertion does not double it', () => {
  const view = reviewed();
  const b = view.block(0);
  assert.equal(b.text, 'Agreed strongly by all', 'insertions read, deletions hide — the AFTER text');
  assert.equal(b.structural, false, 'typing beside a tracked change is ordinary editing now');
  assert.ok(!b.structuralTags.includes('w:ins'));

  // The caret lands INSIDE "strongly" — the w:ins run itself — exactly the
  // spot the old bug doubled: rebuilding used to keep the whole wrapper
  // (its ORIGINAL words) beside the newly spliced run.
  view.collapseTo({ block: 0, offset: 10 });
  view.insertText('X');
  assert.equal(view.block(0).text, 'Agreed strXongly by all');
  assert.equal((view.block(0).text.match(/strongly/g) ?? []).length, 0, 'nothing doubled — the run was spliced, not duplicated');
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
    // A SEQ, not a PAGE: a PAGE field in the body is a run of its own now
    // (runs.js REFERENCE_KINDS) and no longer holds the paragraph.
    '<w:r><w:instrText xml:space="preserve"> SEQ Figure </w:instrText></w:r>' +
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
