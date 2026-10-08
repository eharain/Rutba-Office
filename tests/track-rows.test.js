// Tracked table rows: a row inserted or deleted with Track Changes on carries
// a w:ins or w:del in its own w:trPr. Accept All keeps an inserted row (its
// mark gone) and takes a deleted one out; Reject All does the other; a table
// left with no rows goes with them.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';

const WHO = 'w:author="Ann" w:date="2026-01-01T00:00:00Z"';
const row = (text) => `<w:tr><w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr>`;
const inserted = (text, id) => `<w:tr><w:trPr><w:ins w:id="${id}" ${WHO}/></w:trPr><w:tc><w:p><w:ins w:id="${id + 1}" ${WHO}><w:r><w:t>${text}</w:t></w:r></w:ins></w:p></w:tc></w:tr>`;
const deleted = (text, id) => `<w:tr><w:trPr><w:del w:id="${id}" ${WHO}/></w:trPr><w:tc><w:p><w:del w:id="${id + 1}" ${WHO}><w:r><w:delText>${text}</w:delText></w:r></w:del></w:p></w:tc></w:tr>`;
const table = (...rows) => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${rows.join('')}</w:tbl>`;

const withTable = (tbl) => {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'PLACEHOLDER' }, { text: 'After the table' }] }));
  const xml = pkg.text('word/document.xml').replace(/<w:p\b(?:(?!<\/w:p>)[\s\S])*?PLACEHOLDER[\s\S]*?<\/w:p>/, () => tbl);
  pkg.write_('word/document.xml', xml);
  return openDocx(pkg.write());
};
const bodyOf = (view) => OoxmlPackage.read(view.save()).text('word/document.xml');
const rowTexts = (xml) => [...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((m) => [...m[0].matchAll(/<w:(?:t|delText)\b[^>]*>([^<]*)</g)].map((t) => t[1]).join(''));

test('Accept All keeps an inserted row without its mark and takes a deleted row out', () => {
  const view = withTable(table(row('Kept'), inserted('Added', 5), deleted('Removed', 7)));
  view.acceptChanges({ all: true });
  const xml = bodyOf(view);
  assert.deepEqual(rowTexts(xml), ['Kept', 'Added']);
  assert.doesNotMatch(xml, /<w:trPr\b[^>]*>[\s\S]*?<w:(?:ins|del)\b/, 'no row is left marked');
  assert.doesNotMatch(xml, /<w:(?:ins|del)\b/, 'nothing tracked is left');
});

test('Reject All takes an inserted row out and keeps a deleted row, its words back', () => {
  const view = withTable(table(row('Kept'), inserted('Added', 5), deleted('Removed', 7)));
  view.rejectChanges({ all: true });
  const xml = bodyOf(view);
  assert.deepEqual(rowTexts(xml), ['Kept', 'Removed']);
  assert.match(xml, /<w:t>Removed<\/w:t>/);
  assert.doesNotMatch(xml, /<w:(?:ins|del)\b/);
});

test('a table whose every row is rejected goes, and the words after it stay', () => {
  const view = withTable(table(inserted('Only', 5)));
  view.rejectChanges({ all: true });
  const xml = bodyOf(view);
  assert.doesNotMatch(xml, /<w:tbl\b/);
  assert.match(xml, /After the table/);
});

test('a row inside a row that goes goes with it', () => {
  const nested = `<w:tr><w:trPr><w:del w:id="9" ${WHO}/></w:trPr><w:tc>${table(inserted('Inner', 11))}<w:p/></w:tc></w:tr>`;
  const view = withTable(table(row('Kept'), nested));
  view.acceptChanges({ all: true });
  assert.deepEqual(rowTexts(bodyOf(view)), ['Kept']);
});

test('a deletion across runs and paragraphs is saved with an id for each piece, and still rejected whole', () => {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'First paragraph here' }, { text: 'Second paragraph here' }, { text: 'Third' }] }));
  const view = openDocx(pkg.write());
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: 6 }, { block: 1, offset: 6 });
  view.deleteSelection?.() ?? view.deleteBackward();
  const xml = bodyOf(view);
  const ids = [...xml.matchAll(/<w:(?:ins|del)\b[^>]*\bw:id="(\d+)"/g)].map((m) => m[1]);
  assert.ok(ids.length >= 2, `pieces: ${ids.length}`);
  assert.equal(new Set(ids).size, ids.length, `ids ${ids.join(', ')}`);
  const again = openDocx(view.save());
  again.rejectChanges({ all: true });
  assert.deepEqual(again.blocks.slice(0, 2).map((b) => b.text), ['First paragraph here', 'Second paragraph here']);
});

test('Accept All and Reject All reach the headers too, and undo puts a header back', () => {
  const HDR = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t xml:space="preserve">Draft </w:t></w:r><w:ins w:id="3" ${WHO}><w:r><w:t>two</w:t></w:r></w:ins><w:del w:id="4" ${WHO}><w:r><w:delText>one</w:delText></w:r></w:del></w:p></w:hdr>`;
  const make = () => {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Body' }] }));
    view.doc.doc.pkg.addPart('word/header1.xml', HDR, 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
    return view;
  };
  const header = (view) => view.doc.doc.pkg.text('word/header1.xml');
  const words = (xml) => [...xml.matchAll(/<w:(?:t|delText)\b[^>]*>([^<]*)</g)].map((m) => m[1]).join('');

  const accepted = make();
  accepted.acceptChanges({ all: true });
  assert.equal(words(header(accepted)), 'Draft two');
  assert.doesNotMatch(header(accepted), /<w:(?:ins|del)\b/);
  accepted.undo();
  assert.match(header(accepted), /<w:ins w:id="3"/, 'undo puts the header back as it was');

  const rejected = make();
  rejected.rejectChanges({ all: true });
  assert.equal(words(header(rejected)), 'Draft one');
  assert.doesNotMatch(header(rejected), /<w:(?:ins|del)\b/);
});
