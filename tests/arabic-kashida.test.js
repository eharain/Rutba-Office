// Kashida: justified Arabic stretched by lengthening the joins between its
// letters, as Word justifies a paragraph aligned Justify Low, Medium or High
// (w:jc lowKashida, mediumKashida, highKashida). The places a kashida may go,
// the alignment kept in the file, and the PDF drawing the tatweels.
import test from 'node:test';
import assert from 'node:assert/strict';
import { stretchArabic, kashidaPlaces, KASHIDA_SHARE } from '@rutba/pdf';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';

const TATWEEL = '\u0640';

test('a kashida goes only where two letters join, never inside lam-alef, a word\'s last join first', () => {
  assert.deepEqual(kashidaPlaces('كتب'), [[2, 1]], 'between kaf and teh, and teh and beh; the last first');
  assert.deepEqual(kashidaPlaces('دار'), [], 'dal and reh join nothing after them');
  assert.deepEqual(kashidaPlaces('لا'), [], 'lam-alef is one ligature');
  assert.deepEqual(kashidaPlaces('السلام').flat().sort(), [2, 3], 'never into the lam-alef of salam');
  assert.equal(kashidaPlaces('Hello world').length, 0, 'Latin has no joins to stretch');
  // Marks stay with their letter: the kashida goes after them.
  assert.deepEqual(kashidaPlaces('كَتَبَ'), [[4, 2]]);
});

test('the stretch is shared among the words a round at a time, and nothing is lost', () => {
  const line = 'كتب محمد';
  const stretched = stretchArabic(line, 4);
  assert.equal([...stretched].filter((c) => c === TATWEEL).length, 4);
  assert.equal(stretched.split(TATWEEL).join(''), line, 'only tatweels were added');
  assert.equal(stretchArabic('كتب', 0), 'كتب');
  assert.equal(stretchArabic('دار', 5), 'دار', 'nowhere to stretch, the word as it came');
  assert.deepEqual(KASHIDA_SHARE, { lowKashida: 1 / 3, mediumKashida: 2 / 3, highKashida: 1 });
});

const LONG = Array.from({ length: 30 }, () => 'كتب محمد الكتاب في المكتبة').join(' ');

test('Justify Low, Medium and High are kept in the file as Word writes them', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: LONG }] }));
  view.setSelection({ block: 0, offset: 0 });
  view.setParagraphFormat({ rtl: true, align: 'mediumKashida' });
  const again = openDocx(view.save());
  assert.match(OoxmlPackage.read(again.save()).text('word/document.xml'), /<w:jc w:val="mediumKashida"\/>/);
  assert.equal(again.blocks[0].align, 'mediumKashida');
});

test('a paragraph justified with kashida prints its lines stretched with tatweels, one justified by spaces prints none', async (t) => {
  const { unicodeFont } = await import('../apps/desktop/main/system-fonts.js');
  const font = unicodeFont();
  if (!font) return t.skip('no font with Arabic on this computer');
  const printed = (align) => {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: LONG }] }));
    view.setSelection({ block: 0, offset: 0 });
    view.setParagraphFormat({ rtl: true, align });
    return renderPdf(view, { created: '2026-09-03T00:00:00Z', unicodeFont: font }).buffer.toString('latin1');
  };
  // The ToUnicode map lists every glyph drawn; a tatweel's maps back to U+0640.
  assert.match(printed('highKashida'), /> <0640>/, 'tatweels drawn into the joins');
  assert.match(printed('lowKashida'), /> <0640>/);
  assert.doesNotMatch(printed('both'), /> <0640>/, 'plain justification widens the spaces only');
});
