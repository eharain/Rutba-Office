// Insert Table of Figures, and fields in the body — the engine's half.
//
// A table of figures is Word's TOC field with \c "Figure": an entry per
// caption, a hyperlink to a _Toc bookmark on it with a PAGEREF for its page;
// Update Table follows captions added since. PAGE, NUMPAGES, DATE, TIME,
// FILENAME, AUTHOR and TITLE go in the body as the complex fields Word's
// Field dialog writes, their results worked out and refreshed by F9.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { docFieldInstr, parseDocFieldInstr, formatDatePicture, formatNumber, evaluateDocField } from '@rutba/ooxml/docfields';

const report = () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Figures', style: 'Heading1' },
    { text: 'A chart of sales.' },
    { text: 'A map of the region.' },
    { text: 'The end.' },
  ] }));
  view.setSelection({ block: 1, offset: 0 });
  view.insertCaption({ label: 'Figure', text: 'Sales by month' });
  view.setSelection({ block: 3, offset: 0 });
  view.insertCaption({ label: 'Figure', text: 'The region' });
  view.setSelection({ block: 3, offset: 0 });
  view.insertCaption({ label: 'Table', text: 'Prices' });
  return view;
};

test('Insert Table of Figures writes a TOC \\c field with an entry per caption, linked to its bookmark, with its page', () => {
  const view = report();
  const figures = view.blocks.map((b, i) => [i, b.text]).filter(([, t]) => /^Figure \d/.test(t));
  assert.deepEqual(figures.map(([, t]) => t), ['Figure 1: Sales by month', 'Figure 2: The region']);
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  const pages = Object.fromEntries(view.blocks.map((_, i) => [i, i < 4 ? 1 : 2]));
  view.insertTableOfFigures({ label: 'Figure', pages });
  const engine = view.doc.doc;
  const [tof] = engine.tablesOfFigures();
  assert.equal(tof.label, 'Figure');
  assert.deepEqual(tof.entries.map((e) => [e.text, e.page]), [['Figure 1: Sales by month', '1'], ['Figure 2: The region', '2']]);
  const body = engine._body().body;
  assert.ok(body.includes(' TOC \\h \\z \\c &quot;Figure&quot; ') || body.includes(' TOC \\h \\z \\c "Figure" '));
  for (const e of tof.entries) {
    assert.ok(new RegExp('<w:bookmarkStart\\b[^>]*w:name="' + e.anchor + '"').test(body), 'each entry has its bookmark on the caption');
    assert.ok(body.includes(' PAGEREF ' + e.anchor + ' \\h '));
  }
  assert.ok(engine.pkg.text('word/styles.xml').includes('w:styleId="TableofFigures"'));
  // Not taken for a table of contents: that one is built from headings.
  assert.equal(engine.hasTableOfContents(), false);
});

test('Update Table follows a caption added since; the table of contents and the table of figures keep their own bookmarks', () => {
  const view = report();
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  view.insertTableOfFigures({ label: 'Figure', pages: {} });
  view.setSelection({ block: 0, offset: 0 });
  view.insertTableOfContents({ pages: ['1'] });
  const engine = view.doc.doc;
  const last = view.blocks.length - 1;
  view.setSelection({ block: last, offset: view.blocks[last].text.length });
  view.insertCaption({ label: 'Figure', text: 'Later' });
  view.updateTablesOfFigures({ pages: Object.fromEntries(view.blocks.map((_, i) => [i, 3])) });
  assert.deepEqual(engine.tablesOfFigures()[0].entries.map((e) => [e.text, e.page]), [['Figure 1: Sales by month', '3'], ['Figure 2: The region', '3'], ['Figure 3: Later', '3']]);
  view.updateTableOfContents({});
  const body = engine._body().body;
  for (const e of engine.tablesOfFigures()[0].entries) assert.ok(body.includes('w:name="' + e.anchor + '"'), 'Update Table of the contents kept the figures\' bookmarks');
  assert.equal(engine.tableOfContents().entries.length, 1);
  // Without the label and number: \a, and the caption's own words.
  view.insertTableOfFigures({ label: 'Figure', includeLabel: false, pages: {} });
  assert.deepEqual(engine.tablesOfFigures()[0].entries.map((e) => e.text), ['Sales by month', 'The region', 'Later']);
  assert.equal(engine.tablesOfFigures().length, 1, 'inserting again replaces the one for that label');
});

test('field codes and results: dates in Word\'s pictures, page numbers in their formats', () => {
  assert.equal(docFieldInstr({ name: 'PAGE' }), ' PAGE \\* MERGEFORMAT ');
  assert.equal(docFieldInstr({ name: 'DATE', picture: 'd MMMM yyyy' }), ' DATE \\@ "d MMMM yyyy" \\* MERGEFORMAT ');
  assert.equal(docFieldInstr({ name: 'FILENAME', path: true, format: 'Upper' }), ' FILENAME \\p \\* Upper \\* MERGEFORMAT ');
  assert.deepEqual(parseDocFieldInstr(' DATE \\@ "dddd, d MMMM yyyy" \\* MERGEFORMAT '), { name: 'DATE', picture: 'dddd, d MMMM yyyy', format: '', path: false });
  const when = new Date(2026, 8, 5, 14, 7, 9);
  assert.equal(formatDatePicture(when, 'dddd, d MMMM yyyy'), 'Saturday, 5 September 2026');
  assert.equal(formatDatePicture(when, 'dd/MM/yy'), '05/09/26');
  assert.equal(formatDatePicture(when, 'MMM-yy'), 'Sep-26');
  assert.equal(formatDatePicture(when, 'h:mm am/pm'), '2:07 pm');
  assert.equal(formatDatePicture(when, 'HH:mm:ss'), '14:07:09');
  assert.equal(formatNumber(4, 'roman'), 'iv');
  assert.equal(formatNumber(14, 'ROMAN'), 'XIV');
  assert.equal(formatNumber(28, 'alphabetic'), 'bb');
  assert.equal(evaluateDocField(' NUMPAGES \\* MERGEFORMAT ', { pages: 12 }), '12');
  assert.equal(evaluateDocField(' FILENAME \\* Upper ', { fileName: 'report.docx' }), 'REPORT.DOCX');
  assert.equal(evaluateDocField(' AUTHOR ', { author: 'Ada Lovelace' }), 'Ada Lovelace');
});

test('a field in the body is a complex field, its paragraph editable, and F9 works out its page and the count', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Page  of ' }, { text: 'Second' }] }));
  view.setSelection({ block: 0, offset: 5 });
  view.insertDocField({ name: 'PAGE', page: 1 });
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  view.insertDocField({ name: 'NUMPAGES', pageCount: 1 });
  view.setSelection({ block: 1, offset: 0 });
  view.insertDocField({ name: 'DATE', picture: 'yyyy-MM-dd', now: new Date(2026, 0, 2).toISOString() });
  view.insertText(' ');
  const engine = view.doc.doc;
  assert.equal(view.blocks[0].text, 'Page 1 of 1');
  assert.equal(view.blocks[1].text, '2026-01-02 Second');
  const p = engine.editParagraph(0);
  assert.equal(p.structural, false);
  assert.ok(p.xml.includes('<w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText>'));
  assert.ok(p.xml.includes('<w:fldChar w:fldCharType="separate"/>'));
  const n = view.refreshReferences({ pages: { 0: 3, 1: 4 }, pageCount: 9, now: new Date(2027, 5, 6).toISOString() });
  assert.equal(n, 3);
  assert.equal(view.blocks[0].text, 'Page 3 of 9');
  assert.equal(view.blocks[1].text, '2027-06-06 Second');
  const again = openDocx(view.save());
  assert.equal(again.blocks[0].text, 'Page 3 of 9');
  assert.equal(again.blocks[0].structural, false);
});

test('Table of Figures without page numbers, or with them after the words, as its dialog asks; Update Table keeps the choice', () => {
  const view = report();
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  view.insertTableOfFigures({ label: 'Figure', pageNumbers: true, rightAlign: false, pages: { 2: 1 } });
  const engine = view.doc.doc;
  let body = engine._body().body;
  assert.ok(body.includes('\\p " " \\c "Figure"'), 'a separator in place of the tab');
  assert.ok(!/TableofFigures"\/><w:tabs>/.test(body), 'no right tab stop');
  view.updateTablesOfFigures({ pages: {} });
  assert.ok(engine._body().body.includes('\\p " " \\c "Figure"'), 'Update Table keeps it');
  view.insertTableOfFigures({ label: 'Figure', pageNumbers: false });
  body = engine._body().body;
  assert.ok(body.includes(' TOC \\h \\z \\n \\c "Figure" '));
  assert.ok(!body.slice(body.indexOf('TOC \\h')).split('</w:hyperlink>')[0].includes('PAGEREF'), 'no page numbers');
  assert.deepEqual(engine.tablesOfFigures()[0].entries.map((e) => e.page), [null, null]);
});
