// Home → Multilevel List's library: 1) a) i), I. A. 1., and Article I. /
// Section 1.01 / (a) — each written once, named, and labelled as Word labels
// it, legal numbering showing the article's number in figures.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { formatCounter } from '@rutba/doc-view/lists';

const doc = () => openDocx(buildDocx({ styles: true, paragraphs: ['One', 'One point one', 'One point two', 'Two', 'Two point one', 'Two point one a'].map((text) => ({ text })) }));
const labels = (view) => {
  const frame = view.render({ pages: false });
  return frame.blocks.map((b) => frame.listLabels?.[b.index]?.label ?? null);
};
const levels = [0, 1, 1, 0, 1, 2];
const listed = (style) => {
  const view = doc();
  for (let i = 0; i < 6; i++) {
    view.setSelection({ block: i, offset: 0 });
    view.setParagraphFormat({ list: style });
    if (levels[i]) view.setParagraphFormat({ indentDelta: levels[i] });
  }
  return view;
};

test('1) a) i): parentheses at each level', () => {
  assert.deepEqual(labels(listed('outlineParen')), ['1)', 'a)', 'b)', '2)', 'a)', 'i)']);
});

test('I. A. 1.: roman, letter and figure', () => {
  assert.deepEqual(labels(listed('outlineRoman')), ['I.', 'A.', 'B.', 'II.', 'A.', '1.']);
});

test('Article I. / Section 1.01: the section legal, the article\'s number in figures', () => {
  const view = listed('legal');
  assert.deepEqual(labels(view), ['Article I.', 'Section 1.01', 'Section 1.02', 'Article II.', 'Section 2.01', '(a)']);
  const xml = view.doc.doc.pkg.text('word/numbering.xml');
  assert.match(xml, /<w:name w:val="Rutba Article I\. Section 1\.01"\/>/);
  assert.match(xml, /<w:lvl w:ilvl="1"><w:start w:val="1"\/><w:numFmt w:val="decimalZero"\/><w:isLgl\/><w:lvlText w:val="Section %1\.%2"\/>/);
  assert.equal((xml.match(/Rutba Article I\. Section 1\.01/g) || []).length, 1, 'one definition, however many paragraphs take it');
  // Read back from the saved file: the same labels.
  assert.deepEqual(labels(openDocx(view.save())), ['Article I.', 'Section 1.01', 'Section 1.02', 'Article II.', 'Section 2.01', '(a)']);
});

test('the three Word gives every document are unchanged, and a zero-padded number reads 01', () => {
  assert.deepEqual(labels(listed('outline')), ['1.', '1.1.', '1.2.', '2.', '2.1.', '2.1.1.']);
  assert.equal(formatCounter('decimalZero', 7), '07');
  assert.equal(formatCounter('decimalZero', 12), '12');
});
