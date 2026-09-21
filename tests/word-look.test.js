// Paragraph shading and borders — the engine's half.
//
// The two Home tab controls that were greyed out: a colour behind the
// paragraph (w:shd) and lines round it (w:pBdr), written where the schema
// wants them, read back the same, and kept through a save.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const doc = (paragraphs) => openDocx(buildDocx({ styles: true, paragraphs }));
const pPrOf = (view, i) => view.doc.doc.editParagraph(i).pPr;
const LINE = { style: 'single', widthPx: 1, colour: null, spacePt: 1 };

test('shading and borders are written on the paragraph in schema order, and read back', () => {
  const view = doc([{ text: 'Heading', style: 'Heading1' }, { text: 'A boxed, shaded paragraph.' }, { text: 'after' }]);
  view.setSelection({ block: 1, offset: 0 });
  view.setParagraphFormat({ align: 'center', shading: '#FFF2CC', borders: { bottom: LINE, top: { ...LINE, colour: '#1F5F8B', widthPx: 2 } } });

  const pPr = pPrOf(view, 1);
  assert.match(pPr, /<w:pBdr><w:top w:val="single" w:sz="12" w:space="1" w:color="1F5F8B"\/><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"\/><\/w:pBdr>/, 'the borders, top before bottom, in eighths of a point');
  assert.match(pPr, /<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"\/>/, 'the shading');
  assert.ok(pPr.indexOf('<w:pBdr>') < pPr.indexOf('<w:shd') && pPr.indexOf('<w:shd') < pPr.indexOf('<w:jc'), 'pBdr, then shd, then jc — the schema order');

  const back = openDocx(view.save());
  const block = back.render({ pages: false }).blocks[1];
  assert.equal(String(block.shading).toUpperCase(), '#FFF2CC', 'the colour comes back');
  assert.equal(block.borders.bottom.style, 'single');
  assert.equal(block.borders.top.colour.toUpperCase(), '#1F5F8B');
  assert.equal(block.borders.top.widthPx, 2);
  assert.equal(back.render({ pages: false }).blocks[0].shading, undefined, 'the heading was not touched');
});

test('no colour and no border take the elements off again', () => {
  const view = doc([{ text: 'one' }, { text: 'two' }]);
  view.setSelection({ block: 0, offset: 0 });
  view.setParagraphFormat({ shading: '#DEEBF7', borders: { top: LINE, bottom: LINE, left: LINE, right: LINE } });
  assert.match(pPrOf(view, 0), /<w:pBdr>.*<\/w:pBdr>/);
  view.setParagraphFormat({ shading: null, borders: null });
  assert.doesNotMatch(pPrOf(view, 0) || '', /w:shd|w:pBdr/, 'both gone: the paragraph has no properties left at all');
  view.undo();
  assert.match(pPrOf(view, 0) || '', /w:shd/, 'undo brings the shading back');
});
