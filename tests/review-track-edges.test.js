// Tracked changes — edges the first tests leave alone: a field taken out and
// put back, revision ids across every kind of revision, a Word file's own
// revisions through an edit, and undo after Accept.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';

const WHO = 'w:author="Ann" w:date="2026-01-01T00:00:00Z"';

const withParagraph = (para, mutate) => {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'PLACEHOLDER' }, { text: 'Second paragraph' }] }));
  const xml = pkg.text('word/document.xml').replace(/<w:p\b(?:(?!<\/w:p>)[\s\S])*?PLACEHOLDER[\s\S]*?<\/w:p>/, () => para);
  pkg.write_('word/document.xml', xml);
  if (mutate) mutate(pkg);
  return openDocx(pkg.write());
};
const bodyOf = (view) => OoxmlPackage.read(view.save()).text('word/document.xml');

test('Reject All on a field Word tracked as deleted brings its code back as instrText', () => {
  const view = withParagraph(`<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:del w:id="1" ${WHO}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:delText>3</w:delText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del></w:p>`);
  view.rejectChanges({ all: true });
  const xml = bodyOf(view);
  assert.doesNotMatch(xml, /delInstrText|delText/, 'nothing deleted is left once the deletion is rejected');
  assert.match(xml, /<w:instrText[^>]*> PAGE <\/w:instrText>/, 'the field keeps its instruction');
});

test('a new change gets an id past every kind of revision, including ones in notes and headers', () => {
  const view = withParagraph(
    `<w:p><w:r><w:t>Body text</w:t></w:r><w:ins w:id="2" ${WHO}><w:r><w:t> more</w:t></w:r></w:ins></w:p>` +
    `<w:tbl><w:tblPr><w:tblPrChange w:id="40" ${WHO}><w:tblPr/></w:tblPrChange></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    (pkg) => {
      pkg.addPart('word/footnotes.xml', Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:ins w:id="90" ${WHO}><w:r><w:t>note</w:t></w:r></w:ins></w:p></w:footnote></w:footnotes>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
    },
  );
  assert.ok(view.doc.doc.nextTrackChangeId() > 90, 'counted past the footnote and the table property change');
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: 0 });
  view.insertText('Z');
  const ids = [...bodyOf(view).matchAll(/<w:ins\b[^>]*w:author="Kim"[^>]*w:id="(\d+)"|<w:ins\b[^>]*w:id="(\d+)"[^>]*w:author="Kim"/g)].map((m) => Number(m[1] ?? m[2]));
  assert.equal(ids.length, 1);
  assert.ok(ids[0] > 90, 'id ' + ids[0] + ' must not collide with 2, 40 or 90');
});

test('a Word file\'s own revisions keep their ids and authors through an edit elsewhere, a save and a reopen', () => {
  const para = `<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:ins w:id="5" ${WHO}><w:r><w:t xml:space="preserve">big </w:t></w:r></w:ins><w:del w:id="6" ${WHO}><w:r><w:delText xml:space="preserve">old </w:delText></w:r></w:del><w:r><w:t>world</w:t></w:r></w:p>`;
  const view = withParagraph(para);
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 1, offset: 0 });
  view.insertText('>');
  const again = openDocx(view.save());
  const xml = bodyOf(again);
  assert.match(xml, /<w:ins w:id="5" w:author="Ann"/);
  assert.match(xml, /<w:del w:id="6" w:author="Ann"/);
  const mine = [...xml.matchAll(/<w:ins\b[^>]*w:id="(\d+)"[^>]*w:author="Kim"/g)].map((m) => m[1]);
  assert.equal(mine.length, 1);
  assert.ok(!['5', '6'].includes(mine[0]), 'the new change has an id of its own');
  assert.equal(again.blocks[0].text, 'Hello big world');
  assert.equal(again.blocks[1].text, '>Second paragraph');
});

test('undo after Accept All brings the revisions back, and a deletion across two paragraphs survives save and Reject All', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'First paragraph text' }, { text: 'Second paragraph text' }, { text: 'Third' }] }));
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: 6 }, { block: 1, offset: 6 });
  view.deleteSelection();

  const reopened = openDocx(view.save());
  assert.equal(reopened.blocks.length, 3, 'the deleted mark keeps its paragraph until accepted');
  reopened.acceptChanges({ all: true });
  assert.deepEqual(reopened.blocks.map((b) => b.text), ['First  paragraph text', 'Third']);
  reopened.undo();
  assert.deepEqual(reopened.blocks.map((b) => b.text), ['First ', ' paragraph text', 'Third']);
  assert.ok(reopened.blocks[0].runs.some((r) => r.del), 'the strike-through is back');
  reopened.rejectChanges({ all: true });
  assert.deepEqual(reopened.blocks.map((b) => b.text), ['First paragraph text', 'Second paragraph text', 'Third']);
});
