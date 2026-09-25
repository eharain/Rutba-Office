// Review → Check Accessibility in Rutba Word: what it finds in a document,
// and the fixes it writes the way Word writes them — a picture's description
// on `wp:docPr descr` (and on the picture's own `pic:cNvPr`), "decorative" as
// the adec extension, a header row as `w:tblHeader`, the title in the core
// properties — each one undo step.

import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { checkAccessibility, describeWord, setWordAltText, setWordTableHeader, removeWordParagraphs, setWordTitle, readTitle } from '@rutba/proofing';
import { gradientPng } from '../apps/desktop/main/sample-picture.js';

function fixture() {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Annual report', style: 'Heading1' },
    { text: 'Detail', style: 'Heading3' },
    { text: '', style: 'Heading2' },
    { text: 'These words are pale grey.', colour: 'C8C8C8' },
    { runs: [{ text: 'For prices ' }, { text: 'click here', link: 'https://example.com/prices' }] },
    { text: '' }, { text: '' }, { text: '' },
    { table: { rows: [['Item', 'Price'], ['Pens', '12']] } },
    { text: 'The end.' },
  ] }));
  view.setSelection({ block: 0, offset: 0 });
  view.insertImage({ name: 'Chart', contentType: 'image/png', data: gradientPng(30, 20, [40, 90, 200], [200, 60, 40]), widthPx: 120, heightPx: 80 });
  return view;
}

const rulesOf = (view) => checkAccessibility(describeWord(view)).issues.map((i) => i.rule);

test('a document with every kind of trouble is described and flagged, each finding where it is', () => {
  const view = fixture();
  const { issues, verdict } = checkAccessibility(describeWord(view));
  const by = (rule) => issues.filter((i) => i.rule === rule);
  assert.equal(verdict, 'investigate');
  assert.equal(by('altText').length, 1, 'the picture has no description');
  assert.deepEqual(by('altText')[0].where, { block: 1, image: 0 }, 'found in the paragraph the page picks it in');
  assert.equal(by('headingOrder').length, 1, 'Heading 3 straight after Heading 1');
  assert.equal(by('headingOrder')[0].fixes[0].level, 2);
  assert.equal(by('emptyHeading').length, 1);
  assert.equal(by('contrast').length, 1, 'C8C8C8 on white is 1.7:1');
  assert.match(by('contrast')[0].detail, /^1\.\d:1/);
  assert.equal(by('linkText').length, 1, '"click here"');
  assert.equal(by('blankLines').length, 1, 'three empty paragraphs in a row');
  assert.deepEqual(by('blankLines')[0].fixes[0].target.blocks, [7, 8], 'all but the first');
  assert.equal(by('tableHeader').length, 1, 'the table repeats no header row');
  assert.ok(Number.isInteger(by('tableHeader')[0].where.block), 'the table is found by its first cell');
});

test('Edit Alt Text writes wp:docPr descr and the picture\'s own, Mark as decorative writes the adec extension, and undo takes each back', () => {
  const view = fixture();
  const [picture] = describeWord(view).objects;
  setWordAltText(view, { drawing: picture.target.drawing, descr: 'Sales by quarter, rising' });
  const doc = view.doc.doc;
  assert.match(doc.xml, /<wp:docPr [^>]*descr="Sales by quarter, rising"/);
  assert.match(doc.xml, /<pic:cNvPr [^>]*descr="Sales by quarter, rising"/);
  assert.ok(!rulesOf(view).includes('altText'));

  setWordAltText(view, { drawing: picture.target.drawing, decorative: true });
  assert.match(doc.xml, /<wp:docPr [^>]*>(?:(?!<\/wp:docPr>)[\s\S])*<adec:decorative [^>]*val="1"/, 'the extension inside docPr');
  assert.ok(!/<wp:docPr [^>]*descr=/.test(doc.xml), 'a decorative picture has no description');
  assert.ok(!rulesOf(view).includes('altText'));

  const reopened = openDocx(view.save());
  assert.equal(describeWord(reopened).objects[0].decorative, true, 'the saved file keeps it');

  view.undo();
  assert.match(doc.xml, /descr="Sales by quarter, rising"/, 'one undo: the description back');
  view.undo();
  assert.ok(rulesOf(view).includes('altText'), 'two: as it was');
});

test('the table fix repeats the first row as a header row; empty paragraphs go; the title is set', () => {
  const view = fixture();
  setWordTableHeader(view, { table: 0 });
  assert.match(view.doc.doc.xml, /<w:tbl>[\s\S]*?<w:tr><w:trPr><w:tblHeader\/><\/w:trPr>/);
  assert.ok(!rulesOf(view).includes('tableHeader'));

  const blank = checkAccessibility(describeWord(view)).issues.find((i) => i.rule === 'blankLines');
  const before = view.blocks.length;
  removeWordParagraphs(view, blank.fixes[0].target);
  assert.equal(view.blocks.length, before - 2);
  assert.ok(!rulesOf(view).includes('blankLines'));
  view.undo();
  assert.equal(view.blocks.length, before, 'one undo puts both back');

  setWordTitle(view, { title: 'Annual report 2026' });
  assert.equal(readTitle(OoxmlPackage.read(view.save())), 'Annual report 2026');
  assert.ok(!rulesOf(view).includes('titleMissing'));
});

test('heading and colour fixes are ordinary edits: a style, then an automatic colour', () => {
  const view = fixture();
  const issues = checkAccessibility(describeWord(view)).issues;
  const skip = issues.find((i) => i.rule === 'headingOrder');
  view.setSelection({ block: skip.where.block, offset: 0 });
  view.setParagraphFormat({ styleId: `Heading${skip.fixes[0].level}` });
  const pale = issues.find((i) => i.rule === 'contrast');
  view.setSelection({ block: pale.where.block, offset: 0 }, { block: pale.where.block, offset: pale.fixes[0].target.length });
  view.setRunFormat({ fontColour: pale.fixes[0].colour.slice(1) });
  const after = rulesOf(view);
  assert.ok(!after.includes('headingOrder') && !after.includes('contrast'), after.join(', '));
});
