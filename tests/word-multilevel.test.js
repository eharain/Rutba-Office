// Multilevel lists — the engine's half.
//
// Home → Multilevel list points paragraphs at a nine-level definition whose
// numbers carry the ones above them (1. 1.1. 1.1.1.); Increase and Decrease
// indent move a list paragraph a level rather than a step, in a numbered
// list (1. a. i.) and a bulleted one too, since new definitions now have
// Word's nine levels; a paragraph in no list still moves its own indent.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';

const labels = (view) => Object.fromEntries(Object.entries(view.render().listLabels || {}).map(([k, v]) => [k, v.label]));
const listAll = (view, kind, count) => { for (let i = 0; i < count; i++) { view.collapseTo({ block: i, offset: 0 }); view.setParagraphFormat({ list: kind }); } };

test('a multilevel list numbers 1. 2. 3., Increase indent nests 2.1. and 2.1.1. under them, and the file says so in a nine-level definition', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: ['Scope', 'Parties', 'Buyer', 'Seller', 'Term'].map((text) => ({ text })) }));
  listAll(view, 'outline', 5);
  assert.deepEqual(labels(view), { 0: '1.', 1: '2.', 2: '3.', 3: '4.', 4: '5.' });

  view.collapseTo({ block: 2, offset: 0 });
  view.setParagraphFormat({ indentDelta: 1 });
  view.collapseTo({ block: 3, offset: 0 });
  view.setParagraphFormat({ indentDelta: 1 });
  view.setParagraphFormat({ indentDelta: 1 });
  assert.deepEqual(labels(view), { 0: '1.', 1: '2.', 2: '2.1.', 3: '2.1.1.', 4: '3.' }, 'each level carries the numbers above it');
  assert.equal(view.formatAtCaret().listLevel, 2);
  assert.equal(view.formatAtCaret().indentLevel, 2, 'the ribbon\'s indent buttons see the level');
  view.setParagraphFormat({ indentDelta: -1 });
  assert.equal(labels(view)[3], '2.2.', 'out one level');
  view.setParagraphFormat({ indentDelta: -1 });
  view.setParagraphFormat({ indentDelta: -1 });
  assert.equal(view.formatAtCaret().listLevel, 0, 'never above the top');

  const pkg = OoxmlPackage.read(view.save());
  const numbering = pkg.text('word/numbering.xml');
  assert.match(numbering, /<w:multiLevelType w:val="multilevel"\/>/);
  assert.match(numbering, /<w:lvl w:ilvl="1"><w:start w:val="1"\/><w:numFmt w:val="decimal"\/><w:lvlText w:val="%1\.%2\."\/><w:lvlJc w:val="left"\/><w:pPr><w:ind w:left="1440" w:hanging="360"\/><\/w:pPr><\/w:lvl>/);
  assert.match(numbering, /<w:lvl w:ilvl="8">[\s\S]*?<w:lvlText w:val="%1\.%2\.%3\.%4\.%5\.%6\.%7\.%8\.%9\."\/>/, 'nine levels');
  assert.match(view.doc.doc.editParagraph(2).pPr, /<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="\d+"\/><\/w:numPr>/);
  assert.deepEqual(labels(openDocx(view.save())), { 0: '1.', 1: '2.', 2: '2.1.', 3: '3.', 4: '4.' }, 'through a save');
});

test('a numbered list nests 1. a. i., a bulleted one changes its bullet, and a paragraph in no list moves its own indent', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: ['a', 'b', 'c'].map((text) => ({ text })) }));
  listAll(view, 'number', 3);
  view.collapseTo({ block: 1, offset: 0 });
  view.setParagraphFormat({ indentDelta: 1 });
  view.collapseTo({ block: 2, offset: 0 });
  view.setParagraphFormat({ indentDelta: 1 });
  view.setParagraphFormat({ indentDelta: 1 });
  assert.deepEqual(labels(view), { 0: '1.', 1: 'a.', 2: 'i.' }, 'Word\'s numbered list levels');
  const numbering = OoxmlPackage.read(view.save()).text('word/numbering.xml');
  assert.match(numbering, /<w:multiLevelType w:val="hybridMultilevel"\/><w:lvl w:ilvl="0"><w:start w:val="1"\/><w:numFmt w:val="decimal"\/><w:lvlText w:val="%1\."\/>/);
  assert.match(numbering, /<w:lvl w:ilvl="2"><w:start w:val="1"\/><w:numFmt w:val="lowerRoman"\/><w:lvlText w:val="%3\."\/>/);

  const bullets = openDocx(buildDocx({ styles: true, paragraphs: ['a', 'b'].map((text) => ({ text })) }));
  listAll(bullets, 'bullet', 2);
  bullets.collapseTo({ block: 1, offset: 0 });
  bullets.setParagraphFormat({ indentDelta: 1 });
  assert.equal(bullets.formatAtCaret().listLevel, 1);
  assert.equal(bullets.formatAtCaret().listType, 'bullet', 'still a bullet, one level in');
  assert.match(OoxmlPackage.read(bullets.save()).text('word/numbering.xml'), /<w:lvl w:ilvl="1"><w:start w:val="1"\/><w:numFmt w:val="bullet"\/><w:lvlText w:val="o"\/>/, 'the hollow bullet Word draws there');

  const plain = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'x' }] }));
  plain.setParagraphFormat({ indentDelta: 1 });
  assert.equal(plain.formatAtCaret().indentLevel, 1);
  assert.equal(plain.formatAtCaret().listLevel, null);
  assert.match(plain.doc.doc.editParagraph(0).pPr, /<w:ind w:left="720"\/>/, 'the paragraph\'s own indent, as before');
});
