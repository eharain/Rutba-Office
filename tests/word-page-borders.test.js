// Page borders — the engine's half.
//
// Design → Page Borders writes `<w:pgBorders>` into the section after the
// margins, where the schema puts it; the section carries the sides back for
// the window and the PDF, which draws the box on every sheet.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { buildDocx } from '@rutba/ooxml';

const LINE = { style: 'single', widthPx: 1, colour: null, spacePt: 24 };
const BOX = { offsetFrom: 'page', top: LINE, left: LINE, bottom: LINE, right: LINE };
const sectPrOf = (view) => /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(view.doc.doc.xml)?.[0] || '';

test('a box is written into the section after the margins, read back, kept through a save, printed, cleared and undone', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }] }));
  assert.equal(view.section.pageBorders, null, 'a fresh document has none');

  view.setPageBorders(BOX);
  const borders = view.section.pageBorders;
  assert.equal(borders.offsetFrom, 'page');
  assert.deepEqual(borders.top, { style: 'single', widthPx: 1, colour: '#000000', spacePt: 24 }, 'the side comes back as the reader gives one');
  assert.ok(borders.left && borders.bottom && borders.right, 'all four sides');
  const sectPr = sectPrOf(view);
  assert.match(sectPr, /<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="6" w:space="24" w:color="auto"\/><w:left [^>]*\/><w:bottom [^>]*\/><w:right [^>]*\/><\/w:pgBorders>/, 'written as Word writes it');
  assert.ok(sectPr.indexOf('<w:pgMar') < sectPr.indexOf('<w:pgBorders'), 'after the margins');
  assert.ok(sectPr.indexOf('<w:pgBorders') < sectPr.indexOf('</w:sectPr>'));
  assert.equal(openDocx(view.save()).section.pageBorders.right.widthPx, 1, 'and it survives a save');

  // Printed: four lines, 24 pt in from each edge — the left one runs down x = 24.
  const { buffer } = renderPdf(view, { title: 'framed', created: '2026-09-21T00:00:00Z' });
  const body = buffer.toString('latin1');
  assert.match(body, /\n24 [0-9.]+ m\n24 [0-9.]+ l\nS/, 'the left side');
  assert.match(body, /\n24 [0-9.]+ m\n[0-9.]+ [0-9.]+ l\nS/, 'the top side starts 24 pt in');
  assert.match(body, /\n[0-9.]+ 24 m\n[0-9.]+ 24 l\nS/, 'the bottom side sits 24 pt up');

  view.setPageBorders(null);
  assert.equal(view.section.pageBorders, null, 'cleared');
  assert.doesNotMatch(view.doc.doc.xml, /<w:pgBorders/, 'the element is gone');
  view.undo();
  assert.equal(view.section.pageBorders?.top.style, 'single', 'one undo step brings it back');
});

test('a border measured from the text keeps its space, a heavier double one its weight, and a nil side is no side', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }] }));
  const heavy = { style: 'double', widthPx: 3, colour: '#1F4E79', spacePt: 4 };
  view.setPageBorders({ offsetFrom: 'text', top: heavy, bottom: heavy });
  const sectPr = sectPrOf(view);
  assert.match(sectPr, /<w:pgBorders w:offsetFrom="text"><w:top w:val="double" w:sz="18" w:space="4" w:color="1F4E79"\/><w:bottom [^>]*\/><\/w:pgBorders>/);
  const borders = view.section.pageBorders;
  assert.equal(borders.offsetFrom, 'text');
  assert.deepEqual(borders.top, { style: 'double', widthPx: 3, colour: '#1F4E79', spacePt: 4 });
  assert.equal(borders.left, undefined, 'no left side');

  // The printed frame sits at the margin less the space: 96 px - 4 pt = 90.67 px = 68 pt from the top.
  const { buffer } = renderPdf(view, { title: 'framed', created: '2026-09-21T00:00:00Z' });
  const body = buffer.toString('latin1');
  const top = Math.round((view.section.heightPx - (view.section.margins.top - 4 * (96 / 72))) * 0.75 * 1000) / 1000;
  assert.ok(body.includes(` ${top} m`), 'the top line at the margin less the space');

  // A side written nil by another program is no side at all.
  const doc = view.doc.doc;
  doc.xml = doc.xml.replace('<w:pgBorders w:offsetFrom="text">', '<w:pgBorders w:offsetFrom="text"><w:left w:val="nil" w:sz="0" w:space="0" w:color="auto"/>');
  assert.equal(view.doc.doc.section().pageBorders.left, undefined);
  doc.xml = doc.xml.replace(/<w:pgBorders\b[^>]*>[\s\S]*?<\/w:pgBorders>/, '<w:pgBorders w:offsetFrom="page"><w:top w:val="nil" w:sz="0" w:space="0" w:color="auto"/></w:pgBorders>');
  assert.equal(view.doc.doc.section().pageBorders, null, 'only nil sides is none');
});
