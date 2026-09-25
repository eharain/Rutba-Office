// Mark Entry and Insert Index — the engine's half.
//
// Mark Entry writes an XE field after the marked words, hidden text as Word
// writes it, and the paragraph stays editable; Insert Index writes an INDEX
// field whose result is every entry, sorted, under its letter, with the
// pages the window laid them on — subentries indented or run in, page
// numbers after a comma or right-aligned behind a leader, a range from a
// bookmark, a cross-reference in place of a page.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { xeInstr, parseXeInstr, indexInstr, parseIndexInstr, buildIndex, lineTail, letterOf } from '@rutba/ooxml/wordindex';

test('XE and INDEX field codes are written and read as Word writes them', () => {
  assert.equal(xeInstr({ main: 'Apple' }), ' XE "Apple" ');
  assert.equal(xeInstr({ main: 'Apple', sub: 'Red' , bold: true }), ' XE "Apple:Red" \\b ');
  assert.equal(xeInstr({ main: 'Cats', crossRef: 'See Pets' }), ' XE "Cats" \\t "See Pets" ');
  assert.equal(xeInstr({ main: 'Ratio 3:1', bookmark: 'Ratios', italic: true }), ' XE "Ratio 3\\:1" \\r "Ratios" \\i ');
  assert.deepEqual(parseXeInstr(' XE "Ratio 3\\:1:Low" \\r "Ratios" \\i '), { main: 'Ratio 3:1', subs: ['Low'], crossRef: '', bookmark: 'Ratios', bold: false, italic: true });
  assert.deepEqual(parseXeInstr(' XE "Cats" \\t "See Pets" '), { main: 'Cats', subs: [], crossRef: 'See Pets', bookmark: '', bold: false, italic: false });
  assert.equal(parseXeInstr(' SEQ Figure '), null);
  assert.equal(indexInstr({ columns: 2 }), ' INDEX \\h "A" \\c "2" \\z "1033" ');
  assert.equal(indexInstr({ columns: 1, rightAlign: true, lcid: 2057 }), ' INDEX \\e "\t" \\h "A" \\c "1" \\z "2057" ');
  assert.equal(indexInstr({ columns: 2, runIn: true }), ' INDEX \\h "A" \\c "2" \\r \\z "1033" ');
  assert.deepEqual(parseIndexInstr(' INDEX \\e "\t" \\h "A" \\c "3" \\z "2057" '), { columns: 3, rightAlign: true, runIn: false, headings: true, lcid: 2057 });
});

test('entries sort under their letters, pages once each and in order, a range from a bookmark, a cross-reference', () => {
  const groups = buildIndex([
    { main: 'banana', page: 4 },
    { main: 'Apple', page: 7 },
    { main: 'apple', page: 3, bold: true },
    { main: 'Apple', page: 3 },
    { main: 'Apple', subs: ['red'], page: 5 },
    { main: 'Apple', subs: ['green'], page: 9, pageEnd: 11 },
    { main: 'Cats', crossRef: 'See Pets' },
    { main: '3D printing', page: 2 },
  ]);
  assert.deepEqual(groups.map((g) => g.letter), ['Symbols', 'A', 'B', 'C']);
  const a = groups[1].lines;
  assert.deepEqual(a.map((l) => [l.level, l.text]), [[1, 'Apple'], [2, 'green'], [2, 'red']]);
  assert.deepEqual(a[0].pages, [{ text: '3', bold: true, italic: false }, { text: '7', bold: false, italic: false }]);
  assert.deepEqual(a[1].pages.map((p) => p.text), ['9–11']);
  const text = (line, opts) => line.text + lineTail(line, opts).map((s) => s.text).join('');
  assert.equal(text(a[0]), 'Apple, 3, 7');
  assert.equal(text(a[0], { rightAlign: true }), 'Apple\t3, 7');
  assert.equal(text(groups[3].lines[0]), 'Cats. See Pets');
  assert.deepEqual(lineTail(groups[3].lines[0]).find((s) => s.text === 'See'), { text: 'See', italic: true });
  assert.equal(letterOf('éclair'), 'É');
  const runIn = buildIndex([{ main: 'Apple', page: 3 }, { main: 'Apple', subs: ['red'], page: 5 }], { runIn: true });
  assert.deepEqual(runIn[0].lines[0].runIn.map((s) => s.text), ['red']);
});

const paper = () => openDocx(buildDocx({ styles: true, paragraphs: [
  { text: 'Apples grow on trees.' },
  { text: 'Bananas grow in bunches, as apples do not.' },
  { text: 'Cherries are small.' },
  { text: '' },
] }));

test('Mark Entry writes a hidden XE field after the words, and the paragraph stays editable', () => {
  const view = paper();
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 6 });
  view.markIndexEntry({ main: 'Apples' });
  const engine = view.doc.doc;
  const p = engine.editParagraph(0);
  assert.equal(p.structural, false);
  assert.ok(p.xml.includes('<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:instrText xml:space="preserve"> XE "Apples" </w:instrText></w:r>'));
  assert.equal(view.blocks[0].text, 'Apples grow on trees.', 'an XE adds no words');
  const xe = view.blocks[0].runs.find((r) => r.field?.kind === 'xe');
  assert.ok(xe && xe.text === '');
  // Typing elsewhere in the paragraph keeps the field.
  view.setSelection({ block: 0, offset: 21 });
  view.insertText(' Tall ones.');
  assert.ok(engine.editParagraph(0).xml.includes(' XE "Apples" '));
  // Deleting the words round it keeps it too; Backspace over a space right after it does not take it.
  view.setSelection({ block: 0, offset: 7 });
  view.deleteBackward();
  assert.equal(view.blocks[0].text, 'Applesgrow on trees. Tall ones.');
  assert.ok(engine.editParagraph(0).xml.includes(' XE "Apples" '));
  assert.deepEqual(engine.indexEntries().map((e) => [e.block, e.main]), [[0, 'Apples']]);
});

test('Mark All marks the first time the words appear in each paragraph, matching case', () => {
  const view = paper();
  view.setSelection({ block: 1, offset: 28 }, { block: 1, offset: 34 });
  view.markIndexEntry({ main: 'apples', all: true });
  const entries = view.doc.doc.indexEntries();
  assert.deepEqual(entries.map((e) => e.block), [1]);
  const view2 = paper();
  view2.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 6 });
  view2.markIndexEntry({ main: 'Apples', all: true, text: 'grow' });
  assert.deepEqual(view2.doc.doc.indexEntries().map((e) => e.block), [0, 1]);
});

test('Insert Index writes the INDEX field with every entry and its page, and Update Index follows new entries', () => {
  const view = paper();
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 6 });
  view.markIndexEntry({ main: 'Apple' });
  view.setSelection({ block: 1, offset: 0 }, { block: 1, offset: 7 });
  view.markIndexEntry({ main: 'Banana', sub: 'bunches' });
  view.setSelection({ block: 3, offset: 0 });
  view.insertIndex({ columns: 2, rightAlign: true, leader: 'dot', pages: { 0: 1, 1: 2, 2: 2, 3: 3 } });
  const engine = view.doc.doc;
  const idx = engine.indexResult();
  assert.equal(idx.columns, 2);
  assert.equal(idx.rightAlign, true);
  assert.equal(idx.leader, 'dot');
  assert.deepEqual(idx.lines.map((l) => [l.style, l.text]), [
    ['IndexHeading', 'A'], ['Index1', 'Apple\t1'],
    ['IndexHeading', 'B'], ['Index1', 'Banana'], ['Index2', 'bunches\t2'],
  ]);
  const body = engine._body().body;
  assert.ok(body.includes(' INDEX \\e "\t" \\h "A" \\c "2" \\z "1033" '));
  assert.ok(engine.pkg.text('word/styles.xml').includes('w:styleId="IndexHeading"'));
  // The index's own paragraphs are fields: an XE never lands in them, and a mark after them still counts.
  view.setSelection({ block: 2, offset: 0 }, { block: 2, offset: 8 });
  view.markIndexEntry({ main: 'Cherry', italic: true });
  view.updateIndex({ pages: { 0: 1, 1: 2, 2: 4 } });
  assert.deepEqual(engine.indexResult().lines.map((l) => l.text), ['A', 'Apple\t1', 'B', 'Banana', 'bunches\t2', 'C', 'Cherry\t4']);
  assert.ok(/<w:i\/><w:iCs\/><\/w:rPr><w:t xml:space="preserve">4<\/w:t>/.test(engine._body().body), 'an italic entry\'s page is italic');
  // Without fresh pages, Update Index keeps the numbers it had.
  view.updateIndex({});
  assert.deepEqual(engine.indexResult().lines.map((l) => l.text).slice(-1), ['Cherry\t4']);
  // Saved and opened again, the index and the marks are still fields.
  const again = openDocx(view.save()).doc.doc;
  assert.equal(again.indexEntries().length, 3);
  assert.equal(again.indexResult().lines.length, 7);
});

test('run-in and page ranges from a bookmark', () => {
  const view = paper();
  view.setSelection({ block: 0, offset: 0 }, { block: 2, offset: 5 });
  view.addBookmark('Fruit');
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 6 });
  view.markIndexEntry({ main: 'Fruit', bookmark: 'Fruit' });
  view.markIndexEntry({ main: 'Fruit', sub: 'apples' });
  view.setSelection({ block: 3, offset: 0 });
  view.insertIndex({ columns: 1, runIn: true, pages: { 0: 1, 1: 1, 2: 3, 3: 3 } });
  assert.deepEqual(view.doc.doc.indexResult().lines.map((l) => l.text), ['F', 'Fruit, 1–3: apples, 1']);
});
