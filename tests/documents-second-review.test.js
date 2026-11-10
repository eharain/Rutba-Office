// A tracked deletion keeps what it took: Reject brings it back, and the file
// goes on saying so after it is closed and opened again.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';
import { coalesce, trackedRemoveRange } from '../packages/doc-view/src/positions.js';

const meta = { id: '7', author: 'Ann', date: '2026-01-01T00:00:00Z' };

const fromRuns = (runsXml) => {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'PLACEHOLDER' }] }));
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace(/<w:p\b(?:(?!<\/w:p>)[\s\S])*?PLACEHOLDER[\s\S]*?<\/w:p>/, () => '<w:p>' + runsXml + '</w:p>'));
  return pkg.write();
};
const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const bodyOf = (bytes) => OoxmlPackage.read(bytes).text('word/document.xml');


const symbolRun = { rPr: null, text: 'x', sym: { font: 'Wingdings', char: 'F04A' } };

test('a deleted symbol between deleted words keeps its own place in the deletion', () => {
  const runs = [{ rPr: null, text: 'ab' }, symbolRun, { rPr: null, text: 'cd' }];
  const out = coalesce(trackedRemoveRange(runs, 1, 4, true, meta));
  const dels = out.filter((r) => r.del);
  assert.deepEqual(dels.map((r) => r.del.text), ['b', 'x', 'c']);
  assert.equal(dels.filter((r) => r.del.xml).length, 1, 'only the symbol carries its XML');
});

test('deleting from the symbol on does not lose the word after it', () => {
  const runs = [{ rPr: null, text: 'ab' }, symbolRun, { rPr: null, text: 'cd' }];
  const out = coalesce(trackedRemoveRange(runs, 2, 4, true, meta));
  const dels = out.filter((r) => r.del);
  assert.deepEqual(dels.map((r) => r.del.text), ['x', 'c'], 'the c would otherwise ride behind the symbol and be written nowhere');
});

for (const [what, run, expected] of [
  ['a note reference', { rPr: null, text: '￼', noteRef: { kind: 'footnote', id: '2' } }, /<w:footnoteReference w:id="2"\/>/],
  ['an equation', { rPr: null, text: '￼', math: { xml: '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>', display: false } }, /<m:oMath>/],
  ['a field', { rPr: null, text: '12', field: { instr: ' PAGE ', kind: 'page', name: null } }, /<w:delInstrText xml:space="preserve"> PAGE <\/w:delInstrText>/],
  ['a content control', { rPr: null, text: '￼', control: { xml: '<w:sdt><w:sdtContent><w:r><w:t>y</w:t></w:r></w:sdtContent></w:sdt>' } }, /<w:delText>y<\/w:delText>/],
]) {
  test(`a tracked deletion records ${what}`, () => {
    const runs = [{ rPr: null, text: 'a' }, run, { rPr: null, text: 'b' }];
    const out = trackedRemoveRange(runs, 1, 1 + run.text.length, true, meta);
    assert.equal(out.length, 3);
    assert.ok(out[1].del, 'it stands in the paragraph as a deletion');
    assert.equal(out[1].text, '', 'and is no longer in the words the caret walks');
    assert.match(out[1].del.xml, expected);
  });
}

test('a note\'s own mark is still not the author\'s to take', () => {
  const out = trackedRemoveRange([{ rPr: null, text: '￼', noteMark: 'footnote' }, { rPr: null, text: 'b' }], 0, 1, true, meta);
  assert.deepEqual(out.map((r) => r.text), ['b']);
  assert.equal(out.some((r) => r.del), false);
});

test('half of a field is left whole', () => {
  const out = trackedRemoveRange([{ rPr: null, text: '123', field: { instr: ' PAGE ', kind: 'page', name: null } }], 1, 2, true, meta);
  assert.equal(out[0].text, '123');
  assert.equal(out[0].del, undefined);
});

for (const [what, inner, comesBack] of [
  ['a note reference', '<w:r><w:footnoteReference w:id="2"/></w:r>', /<w:footnoteReference w:id="2"\/>/],
  ['an equation', '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>', /<m:oMath>/],
  ['a field', '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:delText xml:space="preserve">12</w:delText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', /PAGE/],
]) {
  const wrapped = (extra = '') => t('before ') + `<w:del w:id="9" w:author="Ann" w:date="2026-01-01T00:00:00Z">${inner}</w:del>` + t(' after') + extra;

  test(`${what} deleted in a file stays deleted when the paragraph is edited`, () => {
    const view = openDocx(fromRuns(wrapped()));
    view.setSelection({ block: 0, offset: 0 });
    view.insertText('>');
    const xml = bodyOf(view.save());
    const del = /<w:del\b[^>]*>([\s\S]*?)<\/w:del>/.exec(xml);
    assert.ok(del, 'the deletion is still there');
    assert.match(del[1], comesBack);
  });

  test(`Reject brings ${what} back after an edit and a save`, () => {
    const view = openDocx(fromRuns(wrapped()));
    view.setSelection({ block: 0, offset: 0 });
    view.insertText('>');
    const again = openDocx(view.save());
    again.rejectChanges({ all: true });
    const xml = bodyOf(again.save());
    assert.doesNotMatch(xml, /<w:del\b/);
    assert.match(xml, comesBack);
    assert.doesNotMatch(xml, /delInstrText|delText/);
  });
}

test('words deleted in a file beside nothing special still read as they did', () => {
  const view = openDocx(fromRuns(t('a') + '<w:del w:id="9" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>' + t('b')));
  view.setSelection({ block: 0, offset: 0 });
  view.insertText('>');
  assert.match(bodyOf(view.save()), /<w:delText xml:space="preserve">gone<\/w:delText>/);
});

for (const [what, runXml, comesBack] of [
  ['a page field', '<w:fldSimple w:instr=" PAGE "><w:r><w:t>12</w:t></w:r></w:fldSimple>', /PAGE/],
  ['an equation', '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>', /<m:oMath>/],
  ['a note reference', '<w:r><w:footnoteReference w:id="2"/></w:r>', /<w:footnoteReference w:id="2"\/>/],
]) {
  test(`with Track Changes on, ${what} deleted and saved comes back on Reject`, () => {
    const view = openDocx(fromRuns(t('ab') + runXml + t('cd')));
    view.setTrackChanges(true, 'Ann');
    const end = view.block(0).text.length - 1;
    view.setSelection({ block: 0, offset: 1 }, { block: 0, offset: end });
    view.deleteSelection();
    const saved = view.save();
    assert.match(/<w:del\b[\s\S]*<\/w:del>/.exec(bodyOf(saved))?.[0] ?? '', comesBack, 'the file records it');
    const again = openDocx(saved);
    again.rejectChanges({ all: true });
    const xml = bodyOf(again.save());
    assert.match(xml, comesBack);
    assert.doesNotMatch(xml, /<w:del\b|delInstrText|delText/);
    assert.equal(again.block(0).text.startsWith('ab'), true);
  });
}

test('text put in at the start of a table cell stays in that cell', async () => {
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: ['Above'] }));
  const body = pkg.text('word/document.xml').replace('</w:p>', () => '</w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>In the cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  pkg.write_('word/document.xml', Buffer.from(body, 'utf8'));
  const backend = openDocx(pkg.write()).doc;
  const doc = backend.doc;
  assert.equal(doc.editParagraphCount(), 2, 'the paragraph above, and the one in the cell');
  backend.insertParagraphsXml(1, ['<w:p><w:r><w:t>Put in</w:t></w:r></w:p>'], { before: true });
  assert.match(doc.xml, /<w:tc>(?:(?!<\/w:tc>).)*Put in(?:(?!<\/w:tc>).)*In the cell/s);
});
