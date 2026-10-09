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

test('Define New Multilevel List: a list of one\'s own, level by level, written named, started where it says', () => {
  const view = doc();
  const spec = {
    name: 'Chapters',
    levels: [
      { format: 'upperLetter', text: 'Part %1:', start: 3, indentTw: 360, hangingTw: 360 },
      { format: 'decimal', text: '%1-%2', legal: true },
      { format: 'lowerRoman', text: '[%3]' },
    ],
  };
  view.setSelection({ block: 0, offset: 0 });
  view.setParagraphFormat({ list: spec });
  for (let i = 1; i < 6; i++) {
    view.setSelection({ block: i, offset: 0 });
    view.setParagraphFormat({ list: { numId: view.doc.doc.paragraph(0).numbering.numId } });
    if (levels[i]) view.setParagraphFormat({ indentDelta: levels[i] });
  }
  assert.deepEqual(labels(view), ['Part C:', '3-1', '3-2', 'Part D:', '4-1', '[i]'], 'the legal level shows the part in figures');
  const xml = view.doc.doc.pkg.text('word/numbering.xml');
  assert.match(xml, /<w:name w:val="Chapters"\/>/);
  assert.match(xml, /<w:lvl w:ilvl="0"><w:start w:val="3"\/><w:numFmt w:val="upperLetter"\/><w:lvlText w:val="Part %1:"\/><w:lvlJc w:val="left"\/><w:pPr><w:ind w:left="360" w:hanging="360"\/><\/w:pPr><\/w:lvl>/);
  assert.equal((xml.match(/<w:lvl w:ilvl=/g) || []).length >= 9, true, 'nine levels, the rest filled as Word fills them');
  assert.deepEqual(labels(openDocx(view.save())), ['Part C:', '3-1', '3-2', 'Part D:', '4-1', '[i]'], 'read back from the saved file');
  assert.throws(() => view.doc.defineList({ levels: [{ format: 'emoji' }] }), /number style/);
  assert.throws(() => view.doc.defineList({ levels: [{ text: '%1"/><x' }, { style: 'Nope' }] }), /no paragraph style "Nope"/);
});

test('a level linked to a heading style numbers every heading of that style, and a heading can be taken out', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Intro', style: 'Heading1' }, { text: 'Body' }, { text: 'Scope', style: 'Heading2' }, { text: 'Terms', style: 'Heading2' }, { text: 'Method', style: 'Heading1' }, { text: 'Steps', style: 'Heading2' },
  ] }));
  view.setSelection({ block: 0, offset: 0 });
  view.setParagraphFormat({ list: { name: 'Headings', levels: [{ format: 'decimal', text: '%1', style: 'Heading1' }, { format: 'decimal', text: '%1.%2', style: 'Heading2' }] } });
  assert.deepEqual(labels(view), ['1', null, '1.1', '1.2', '2', '2.1'], 'the headings numbered by their styles; the body not');
  const styles = view.doc.doc.pkg.text('word/styles.xml');
  assert.match(styles, /w:styleId="Heading2"[\s\S]*?<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="\d+"\/><\/w:numPr>/, 'the style carries the list, as Word links it');
  assert.doesNotMatch(view.doc.doc.paragraph(0).pPr || '', /<w:numPr>/, 'the heading is numbered by its style, not its own numPr');
  // Out of the list: numId 0, the style's number turned off for this one.
  view.setSelection({ block: 3, offset: 0 });
  view.setParagraphFormat({ list: null });
  assert.deepEqual(labels(view), ['1', null, '1.1', null, '2', '2.1']);
  assert.match(view.doc.doc.paragraph(3).pPr, /<w:numId w:val="0"\/>/);
  // A new heading of the style is numbered as it is typed.
  view.setSelection({ block: 1, offset: 0 });
  view.setParagraphFormat({ styleId: 'Heading2' });
  assert.deepEqual(labels(view), ['1', '1.1', '1.2', null, '2', '2.1']);
  assert.deepEqual(labels(openDocx(view.save())), ['1', '1.1', '1.2', null, '2', '2.1'], 'read back from the saved file');
});

test('the three Word gives every document are unchanged, and a zero-padded number reads 01', () => {
  assert.deepEqual(labels(listed('outline')), ['1.', '1.1.', '1.2.', '2.', '2.1.', '2.1.1.']);
  assert.equal(formatCounter('decimalZero', 7), '07');
  assert.equal(formatCounter('decimalZero', 12), '12');
});
