// The ruler and the grips on a table — the engine's half.
//
// The ruler moves four indents and the tab stops of a paragraph, the margins
// of the page and the columns of a table; the grips on the page move a
// table's columns and rows. Each is one paragraph or table property written
// where the schema wants it, read back the same, and kept through a save.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn. ';
const doc = (paragraphs) => openDocx(buildDocx({ styles: true, paragraphs }));
const xmlOf = (view) => view.doc.doc.xml;
const pPrOf = (view, i) => view.doc.doc.editParagraph(i).pPr;

test('the ruler\'s indents: first line and hanging are one setting, left is explicit, right clears', () => {
  const view = doc([{ text: 'Heading', style: 'Heading1' }, { text: lorem }, { text: 'after' }]);
  view.setSelection({ block: 1, offset: 0 });

  view.setParagraphFormat({ firstLineTwips: 720 });
  let props = view.doc.getParagraphProps(1);
  assert.equal(props.firstLineTwips, 720);
  assert.equal(props.hangingTwips, null);
  assert.equal(view.render({ pages: false }).blocks[1].firstLinePx, 48);

  // A hanging indent replaces the first-line one, as Word's ruler does.
  view.setParagraphFormat({ hangingTwips: 360 });
  props = view.doc.getParagraphProps(1);
  assert.equal(props.hangingTwips, 360);
  assert.equal(props.firstLineTwips, null);
  assert.match(pPrOf(view, 1), /<w:ind w:hanging="360"\/>/);

  // An explicit zero left indent is written, not dropped: it beats a style's indent.
  view.setParagraphFormat({ leftTwips: 0 });
  assert.match(pPrOf(view, 1), /<w:ind w:hanging="360" w:left="0"\/>/);
  assert.equal(view.doc.getParagraphProps(1).indentTwips, 0);
  assert.equal(view.render({ pages: false }).blocks[1].indentPx, 0);

  view.setParagraphFormat({ leftTwips: 1440, rightTwips: 720 });
  props = view.doc.getParagraphProps(1);
  assert.equal(props.indentTwips, 1440);
  assert.equal(props.rightTwips, 720);
  assert.equal(view.render({ pages: false }).blocks[1].rightPx, 48);

  // Clearing every indent takes the element away rather than leaving an empty one.
  view.setParagraphFormat({ leftTwips: null, rightTwips: null, hangingTwips: null });
  assert.doesNotMatch(pPrOf(view, 1) ?? '', /<w:ind\b/);

  // A survivor of the save.
  view.setParagraphFormat({ firstLineTwips: 1170, rightTwips: 300 });
  const again = openDocx(view.save());
  assert.equal(again.doc.getParagraphProps(1).firstLineTwips, 1170);
  assert.equal(again.doc.getParagraphProps(1).rightTwips, 300);
  assert.equal(again.render({ pages: false }).blocks[1].firstLinePx, 78);
});

test('tab stops: written sorted in the schema\'s slot, read back, cleared, and a clear-only list is empty', () => {
  const view = doc([{ text: 'a\tb\tc' }, { text: 'plain' }]);
  view.setSelection({ block: 0, offset: 0 });
  view.setParagraphFormat({ align: 'left', firstLineTwips: 360 });
  view.setParagraphFormat({ tabs: [{ align: 'right', posTwips: 5760, leader: 'dot' }, { align: 'left', posTwips: 1440 }] });

  const pPr = pPrOf(view, 0);
  // tabs sit before ind and jc, whatever order they were set in.
  assert.match(pPr, /^<w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"\/><w:tab w:val="right" w:leader="dot" w:pos="5760"\/><\/w:tabs><w:ind w:firstLine="360"\/><w:jc w:val="left"\/><\/w:pPr>$/);
  assert.deepEqual(view.doc.getParagraphProps(0).tabs, [{ align: 'left', posTwips: 1440 }, { align: 'right', posTwips: 5760, leader: 'dot' }]);
  assert.deepEqual(view.render({ pages: false }).blocks[0].tabs, [{ align: 'left', posPx: 96, leader: null }, { align: 'right', posPx: 384, leader: 'dot' }]);

  // A stop moved: the list is rewritten whole.
  view.setParagraphFormat({ tabs: [{ align: 'left', posTwips: 2160 }] });
  assert.deepEqual(view.render({ pages: false }).blocks[0].tabs, [{ align: 'left', posPx: 144, leader: null }]);

  // Only clears: the paragraph has no stops, and says so with an empty list.
  view.setParagraphFormat({ tabs: [{ align: 'clear', posTwips: 1440 }] });
  assert.match(pPrOf(view, 0), /<w:tabs><w:tab w:val="clear" w:pos="1440"\/><\/w:tabs>/);
  assert.deepEqual(view.doc.getParagraphProps(0).tabs, []);
  assert.deepEqual(view.render({ pages: false }).blocks[0].tabs, []);

  // None at all: the element goes.
  view.setParagraphFormat({ tabs: [] });
  assert.doesNotMatch(pPrOf(view, 0) ?? '', /<w:tabs\b/);
  assert.equal(view.doc.getParagraphProps(0).tabs, null);
  assert.equal(view.render({ pages: false }).blocks[0].tabs, undefined);

  // The paragraph that was never touched has no tabs either.
  assert.equal(view.doc.getParagraphProps(1).tabs, null);
});

test('a table\'s grid, width, spans and row heights ride its blocks; columns and rows resize', () => {
  const view = doc([
    { text: 'before' },
    { table: { rows: [['a', 'b', 'c'], ['d', 'e', 'f']], columns: [2400, 3600, 3360], header: true } },
    { text: 'after' },
  ]);
  let blocks = view.render({ pages: false }).blocks;
  assert.equal(blocks[0].gridPx, undefined);
  assert.deepEqual(blocks[1].gridPx, [160, 240, 224]);
  assert.deepEqual(blocks[1].tableWidth, { type: 'dxa', value: 9360 });
  assert.equal(blocks[1].cellSpan, undefined);
  assert.equal(blocks[1].rowHeightPx, undefined);
  assert.deepEqual(blocks[6].gridPx, [160, 240, 224]);
  assert.equal(blocks[7].gridPx, undefined);
  const table = Number(/^t(\d+):/.exec(blocks[1].container)[1]);

  // Two columns either side of a dragged border: the sum, and so the fixed width, stays.
  view.setTableColumnWidths({ table, widths: { 0: 3000, 1: 3000 } });
  blocks = view.render({ pages: false }).blocks;
  assert.deepEqual(blocks[1].gridPx, [200, 200, 224]);
  assert.deepEqual(blocks[1].tableWidth, { type: 'dxa', value: 9360 });
  assert.equal((xmlOf(view).match(/<w:tcW w:w="3000" w:type="dxa"\/>/g) || []).length, 4);
  assert.match(xmlOf(view), /<w:tblGrid><w:gridCol w:w="3000"\/><w:gridCol w:w="3000"\/><w:gridCol w:w="3360"\/><\/w:tblGrid>/);

  // The right edge dragged: only the last column changes, and a fixed table width follows the grid.
  view.setTableColumnWidths({ table, widths: { 2: 4000 } });
  blocks = view.render({ pages: false }).blocks;
  assert.deepEqual(blocks[1].gridPx, [200, 200, 266.6666666666667]);
  assert.deepEqual(blocks[1].tableWidth, { type: 'dxa', value: 10000 });
  assert.match(xmlOf(view), /<w:tblW w:w="10000" w:type="dxa"\/>/);

  // A row's height, written before tblHeader as the schema orders trPr, and read back on its cells.
  view.setTableRowHeight({ table, row: 0, twips: 600 });
  blocks = view.render({ pages: false }).blocks;
  assert.equal(blocks[1].rowHeightPx, 40);
  assert.equal(blocks[1].rowRule, 'atLeast');
  assert.equal(blocks[4].rowHeightPx, undefined);
  assert.match(xmlOf(view), /<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="atLeast"\/><w:tblHeader\/><\/w:trPr>/);
  view.setTableRowHeight({ table, row: 1, twips: 900 });
  assert.match(xmlOf(view), /<w:tr><w:trPr><w:trHeight w:val="900" w:hRule="atLeast"\/><\/w:trPr><w:tc>/);
  assert.equal(view.render({ pages: false }).blocks[4].rowHeightPx, 60);

  // Cleared: the row takes its content's height again and the header mark survives.
  view.setTableRowHeight({ table, row: 0, twips: null });
  assert.match(xmlOf(view), /<w:tr><w:trPr><w:tblHeader\/><\/w:trPr>/);
  assert.equal(view.render({ pages: false }).blocks[1].rowHeightPx, undefined);

  // All of it survives a save.
  const again = openDocx(view.save()).render({ pages: false }).blocks;
  assert.deepEqual(again[1].gridPx, [200, 200, 266.6666666666667]);
  assert.deepEqual(again[1].tableWidth, { type: 'dxa', value: 10000 });
  assert.equal(again[4].rowHeightPx, 60);

  // A merged cell carries its span; its table's columns refuse, as the engine says.
  view.setSelection({ block: 1, offset: 0 });
  view.tableOp('mergeRight');
  const merged = view.render({ pages: false }).blocks;
  assert.equal(merged[1].cellSpan, 2);
  assert.throws(() => view.setTableColumnWidths({ table, widths: { 0: 2000 } }), /merged/);
});

test('a table that takes a share of the text width says so', () => {
  const view = doc([{ table: { rows: [['a', 'b'], ['c', 'd']] } }]);
  const b = view.render({ pages: false }).blocks[0];
  assert.deepEqual(b.tableWidth, { type: 'pct', value: 5000 });
  assert.deepEqual(b.gridPx, [312, 312]);
});
