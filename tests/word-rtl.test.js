// Right to left in Documents: a paragraph that runs from the right margin,
// as Arabic, Hebrew, Persian and Urdu do — w:bidi read from the file and
// from its style, written by the ribbon's direction buttons, and the
// paragraph's alignment kept mirrored as Word keeps it; a table whose
// columns run from the right (w:bidiVisual) and a section whose columns do
// (w:bidi in the section), read and written.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage, readBidi } from '@rutba/ooxml';

const ARABIC = 'مرحبا بالعالم، هذه فقرة تبدأ من اليمين.';

/** A document whose paragraphs are given as their pPr and words, with an optional extra style. */
function docWith(paragraphs, extraStyle = '') {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: paragraphs.map(() => ({ text: 'x' })) }));
  const body = paragraphs.map(([pPr, text]) => `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('');
  const doc = pkg.text('word/document.xml').replace(/(<w:body>)[\s\S]*?(<w:sectPr\b|<\/w:body>)/, (_, open, close) => open + body + close);
  pkg.write_('word/document.xml', Buffer.from(doc, 'utf8'));
  if (extraStyle) pkg.write_('word/styles.xml', Buffer.from(pkg.text('word/styles.xml').replace('</w:styles>', () => extraStyle + '</w:styles>'), 'utf8'));
  return openDocx(pkg.write());
}
const pPrOf = (view, i) => /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(view.doc.doc.editParagraph(i).xml)?.[0] ?? '';

test('w:bidi reads as right to left, w:val="0" as a right-to-left style turned back, nothing as nothing', () => {
  assert.equal(readBidi('<w:pPr><w:bidi/></w:pPr>'), true);
  assert.equal(readBidi('<w:pPr><w:bidi w:val="1"/></w:pPr>'), true);
  assert.equal(readBidi('<w:pPr><w:bidi w:val="0"/></w:pPr>'), false);
  assert.equal(readBidi('<w:pPr><w:bidi w:val="false"/></w:pPr>'), false);
  assert.equal(readBidi('<w:pPr><w:jc w:val="left"/></w:pPr>'), null);
  assert.equal(readBidi('<w:pPr><w:bidiVisual/></w:pPr>'), null, 'a table\'s direction is not a paragraph\'s');
  assert.equal(readBidi('<w:pPr><w:pPrChange w:id="1"><w:pPr><w:bidi/></w:pPr></w:pPrChange></w:pPr>'), null, 'a recorded change\'s old properties are not the paragraph\'s own');
  assert.equal(readBidi(null), null);
});

test('a right-to-left paragraph is drawn from the right, its alignment read back as the page shows it', () => {
  const view = docWith([['<w:pPr><w:bidi/></w:pPr>', ARABIC], ['<w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>', ARABIC], ['', 'Left to right.']]);
  const blocks = view.render({ pages: false }).blocks;
  assert.equal(blocks[0].rtl, true);
  assert.equal(blocks[2].rtl, undefined, 'a paragraph that says nothing leaves it to its style');
  view.setSelection({ block: 0, offset: 2 });
  assert.equal(view.formatAtCaret().rtl, true);
  assert.equal(view.formatAtCaret().paragraphAlign, 'right', 'with no alignment of its own it sits at the right margin');
  view.setSelection({ block: 1, offset: 2 });
  assert.equal(view.formatAtCaret().paragraphAlign, 'left', 'Word\'s "right" in a right-to-left paragraph is the left margin');
  view.setSelection({ block: 2, offset: 2 });
  assert.equal(view.formatAtCaret().rtl, false);
  assert.equal(view.formatAtCaret().paragraphAlign, 'left');
});

test('the direction buttons write w:bidi in its place, and take it off again', () => {
  const view = docWith([['<w:pPr><w:spacing w:after="120"/><w:jc w:val="center"/></w:pPr>', ARABIC]]);
  view.setSelection({ block: 0, offset: 0 });
  view.setParagraphFormat({ rtl: true });
  assert.match(pPrOf(view, 0), /<w:bidi\/><w:spacing w:after="120"\/><w:jc w:val="center"\/>/, 'before the spacing and the alignment, as the schema orders them');
  assert.equal(view.formatAtCaret().rtl, true);
  assert.equal(view.formatAtCaret().paragraphAlign, 'center', 'centred stays centred');
  assert.equal(view.render({ pages: false }).blocks[0].rtl, true);

  const reopened = openDocx(view.save());
  assert.equal(reopened.render({ pages: false }).blocks[0].rtl, true, 'in the file');

  view.setParagraphFormat({ rtl: false });
  assert.doesNotMatch(pPrOf(view, 0), /bidi/);
  assert.equal(view.formatAtCaret().rtl, false);
  view.undo();
  assert.match(pPrOf(view, 0), /<w:bidi\/>/, 'one undo step');
});

test('aligning a right-to-left paragraph writes its alignment mirrored, as Word does', () => {
  const view = docWith([['<w:pPr><w:bidi/></w:pPr>', ARABIC]]);
  view.setSelection({ block: 0, offset: 0 });
  view.setParagraphFormat({ align: 'left' });
  assert.match(pPrOf(view, 0), /<w:jc w:val="right"\/>/, 'the page\'s left is the end of the line');
  assert.equal(view.formatAtCaret().paragraphAlign, 'left');
  view.setParagraphFormat({ align: 'right' });
  assert.match(pPrOf(view, 0), /<w:jc w:val="left"\/>/);
  assert.equal(view.formatAtCaret().paragraphAlign, 'right');
  view.setParagraphFormat({ align: 'both' });
  assert.match(pPrOf(view, 0), /<w:jc w:val="both"\/>/, 'justified is justified either way');
  // Direction and alignment together: the alignment is the new direction's.
  const other = docWith([['', 'Left to right.']]);
  other.setSelection({ block: 0, offset: 0 });
  other.setParagraphFormat({ rtl: true, align: 'right' });
  assert.match(pPrOf(other, 0), /<w:bidi\/>[\s\S]*<w:jc w:val="left"\/>/);
});

test('a paragraph whose style runs right to left follows it, and is turned back with w:val="0"', () => {
  const style = '<w:style w:type="paragraph" w:customStyle="1" w:styleId="ArabicBody"><w:name w:val="Arabic Body"/><w:basedOn w:val="Normal"/><w:pPr><w:bidi/></w:pPr></w:style>';
  const view = docWith([['<w:pPr><w:pStyle w:val="ArabicBody"/></w:pPr>', ARABIC]], style);
  view.setSelection({ block: 0, offset: 0 });
  assert.equal(view.docStyles.ArabicBody.rtl, true, 'the style says so');
  assert.equal(view.formatAtCaret().rtl, true, 'and the paragraph follows it');
  view.setParagraphFormat({ rtl: false });
  assert.match(pPrOf(view, 0), /<w:bidi w:val="0"\/>/, 'turned back against its style');
  assert.equal(view.formatAtCaret().rtl, false);
  assert.equal(view.render({ pages: false }).blocks[0].rtl, false);
  view.setParagraphFormat({ rtl: true });
  assert.doesNotMatch(pPrOf(view, 0), /bidi/, 'and back to its style\'s way: nothing of its own');
  assert.equal(view.formatAtCaret().rtl, true);
});

test('a right-to-left table is read from w:bidiVisual, and the caret\'s table turned either way writes it in its place', () => {
  const view = docWith([['', 'Before the table.']]);
  view.setSelection({ block: 0, offset: 0 });
  view.insertTable({ rows: 2, cols: 3 });
  const blocks = () => view.render({ pages: false }).blocks;
  const cell = blocks().findIndex((b) => /^t\d+:r0:c0$/.test(b.container || ''));
  assert.ok(cell >= 0, 'a cell to put the caret in');
  view.setSelection({ block: cell, offset: 0 });
  view.tableOp('direction', { rtl: true });
  const xml = view.doc.doc.xml;
  assert.match(xml, /<w:tblPr>(?:<w:tblStyle\b[^>]*\/>)?<w:bidiVisual\/>/, 'after the style, before the rest of the table\'s properties');
  const look = blocks().find((b) => b.tableLook)?.tableLook;
  assert.equal(look?.rtl, true);
  view.tableOp('direction', { rtl: false });
  assert.doesNotMatch(view.doc.doc.xml, /<w:bidiVisual/);
  view.undo();
  assert.match(view.doc.doc.xml, /<w:bidiVisual\/>/, 'one undo step each way');
});

test('a right-to-left section lays its columns from the right, and Layout writes w:bidi where the schema has it', () => {
  const view = docWith([['', ARABIC]]);
  view.setPageSetup({ columns: { count: 2, spaceTwips: 720 } });
  const ltr = view.section.columnBoxes;
  assert.ok(ltr[0].xPx < ltr[1].xPx, 'left to right, the first column at the left');
  view.setPageSetup({ rtl: true });
  assert.equal(view.section.rtl, true);
  const rtl = view.section.columnBoxes;
  assert.ok(rtl[0].xPx > rtl[1].xPx, 'right to left, the first column at the right');
  assert.equal(Math.round(rtl[0].xPx + rtl[0].widthPx), Math.round(view.section.contentWidthPx), 'flush with the right margin');
  const sectPr = /<w:sectPr\b[\s\S]*<\/w:sectPr>/.exec(view.doc.doc.xml)[0];
  assert.match(sectPr, /<w:cols\b[\s\S]*<w:bidi\/>/, 'after the columns');
  assert.ok(!/<w:docGrid\b/.test(sectPr) || sectPr.indexOf('<w:bidi/>') < sectPr.indexOf('<w:docGrid'), 'and before the grid');
  view.setPageSetup({ rtl: false });
  assert.equal(view.section.rtl, false);
  assert.doesNotMatch(view.doc.doc.xml, /<w:bidi\/>/);
});
