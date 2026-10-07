// Documents: building blocks — Insert → Quick Parts → Save Selection to
// Quick Part Gallery, and the block put in again, in this document or
// another. What travels is the words and their look; what only made sense
// in the document it came from stays behind.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { portableParagraph, fitParagraph, blockText } from '@rutba/ooxml/blocks';

const source = () => openDocx(buildDocx({ styles: true, paragraphs: [
  { text: 'Terms of payment', style: 'Heading1' },
  { runs: [{ text: 'Invoices are due within ' }, { text: 'thirty days', bold: true }, { text: ' of their date.' }] },
  { runs: [{ text: 'Questions: ' }, { text: 'our site', link: 'https://example.com/help' }] },
] }));
const target = () => openDocx(buildDocx({ styles: true, paragraphs: ['First paragraph.', 'Second paragraph.'] }));
const textOf = (view) => view.render({ pages: false }).blocks.map((b) => b.runs.map((r) => r.text).join(''));

test('a paragraph travels as words and their look: tracked changes taken as accepted, the rest of its document left behind', () => {
  const xml = '<w:p w:rsidR="00A1"><w:pPr><w:pStyle w:val="Heading1"/><w:pPrChange w:id="1"><w:pPr/></w:pPrChange></w:pPr>'
    + '<w:bookmarkStart w:id="0" w:name="here"/><w:r><w:t>Kept </w:t></w:r>'
    + '<w:del w:id="2" w:author="A"><w:r><w:delText>gone </w:delText></w:r></w:del>'
    + '<w:ins w:id="3" w:author="A"><w:r><w:t>added </w:t></w:r></w:ins>'
    + '<w:r><w:drawing><wp:inline/></w:drawing></w:r>'
    + '<w:hyperlink r:id="rId9"><w:r><w:t>a link</w:t></w:r></w:hyperlink>'
    + '<w:hyperlink w:anchor="top"><w:r><w:t> and a place</w:t></w:r></w:hyperlink>'
    + '<w:bookmarkEnd w:id="0"/></w:p>';
  const out = portableParagraph(xml, { linkTarget: (id) => (id === 'rId9' ? 'https://example.com' : null) });
  assert.doesNotMatch(out, /gone|<w:del\b|<w:ins\b|pPrChange|bookmark|drawing|rsid/);
  assert.match(out, /<w:r><w:t>Kept <\/w:t><\/w:r><w:r><w:t>added <\/w:t><\/w:r>/);
  assert.match(out, /<w:fldSimple w:instr=" HYPERLINK &quot;https:\/\/example\.com&quot; "><w:r><w:t>a link<\/w:t><\/w:r><\/w:fldSimple>/, 'a link to the web as a field, needing no relationship');
  assert.match(out, /<w:r><w:t> and a place<\/w:t><\/w:r>/, 'a link to a place in its own document: its words');
  assert.equal(portableParagraph('<w:tbl/>'), null);
  assert.equal(blockText([out]), 'Kept added a link and a place');
});

test('arriving, a block loses the styles and lists its new document cannot resolve', () => {
  const xml = '<w:p><w:pPr><w:pStyle w:val="Fancy"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:rStyle w:val="Strong"/><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>';
  assert.equal(fitParagraph(xml, { styles: new Set(['Normal']), charStyles: new Set(), numIds: new Set() }), '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>');
  assert.equal(fitParagraph(xml, { styles: new Set(['Fancy']), charStyles: new Set(['Strong']), numIds: new Set(['7']) }), xml, 'all of it resolved: all of it kept');
});

test('whole paragraphs saved, and put in after the caret\'s paragraph of another document, looking as they did', () => {
  const from = source();
  from.setSelection({ block: 0, offset: 0 }, { block: 2, offset: 0 });
  const block = from.buildingBlock();
  assert.equal(block.inline, false);
  assert.equal(block.paragraphs.length, 2, 'a selection ending at a paragraph\'s start does not take it');
  const to = target();
  to.setSelection({ block: 0, offset: 'First paragraph.'.length });
  to.insertBuildingBlock(block);
  assert.deepEqual(textOf(to), ['First paragraph.', 'Terms of payment', 'Invoices are due within thirty days of their date.', 'Second paragraph.']);
  const blocks = to.render({ pages: false }).blocks;
  assert.equal(blocks[1].style, 'Heading1', 'the heading still a heading here');
  assert.ok(blocks[2].runs.find((r) => r.text === 'thirty days').bold, 'the bold still bold');
  assert.deepEqual([to.focus.block, to.focus.offset], [2, 'Invoices are due within thirty days of their date.'.length], 'the caret after what went in');
  to.undo();
  assert.deepEqual(textOf(to), ['First paragraph.', 'Second paragraph.'], 'one step of undo');
});

test('the caret inside a paragraph splits it round the block; at its start the block goes before it', () => {
  const from = source();
  from.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 'Terms of payment'.length });
  const heading = from.buildingBlock();
  assert.equal(heading.inline, false, 'a whole paragraph selected is a paragraph');
  const mid = target();
  mid.setSelection({ block: 0, offset: 5 });
  mid.insertBuildingBlock(heading);
  assert.deepEqual(textOf(mid), ['First', 'Terms of payment', ' paragraph.', 'Second paragraph.']);
  const start = target();
  start.setSelection({ block: 1, offset: 0 });
  start.insertBuildingBlock(heading);
  assert.deepEqual(textOf(start), ['First paragraph.', 'Terms of payment', 'Second paragraph.']);
  const top = target();
  top.setSelection({ block: 0, offset: 0 });
  top.insertBuildingBlock(heading);
  assert.deepEqual(textOf(top), ['Terms of payment', 'First paragraph.', 'Second paragraph.'], 'before the first paragraph of all');
});

test('words inside one paragraph are kept as words, and go in at the caret in their own look', () => {
  const from = source();
  const at = 'Invoices are due within '.length;
  from.setSelection({ block: 1, offset: at }, { block: 1, offset: at + 'thirty days'.length });
  const words = from.buildingBlock();
  assert.equal(words.inline, true);
  const to = target();
  to.setSelection({ block: 1, offset: 'Second '.length });
  to.insertBuildingBlock(words);
  assert.deepEqual(textOf(to), ['First paragraph.', 'Second thirty daysparagraph.']);
  assert.ok(to.render({ pages: false }).blocks[1].runs.find((r) => r.text === 'thirty days').bold);
  assert.deepEqual([to.focus.block, to.focus.offset], [1, 'Second thirty days'.length]);
});

test('with Track Changes on, a block that goes in is marked as put in', () => {
  const from = source();
  from.setSelection({ block: 0, offset: 0 }, { block: 1, offset: 0 });
  const block = from.buildingBlock();
  const to = target();
  to.setTrackChanges(true);
  assert.equal(to.recording, true);
  to.setSelection({ block: 0, offset: 'First paragraph.'.length });
  to.insertBuildingBlock(block);
  assert.match(to.doc.doc.paragraph(1).xml, /<w:ins w:id="\d+" w:author="[^"]+" w:date="[^"]+"><w:r>/);
  assert.equal(to.buildingBlock(), null, 'nothing selected after: nothing to keep');
});
