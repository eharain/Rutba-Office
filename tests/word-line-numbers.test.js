// Line numbers — the engine's half.
//
// Layout → Line Numbers writes `<w:lnNumType>` into the section after the
// page borders; the section reads it back with Word's own default (a new
// count on every page when the file does not say), and the PDF numbers the
// lines the paginator laid, every countBy-th of them, in the margin.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { buildDocx } from '@rutba/ooxml';

const sectPrOf = (view) => /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(view.doc.doc.xml)?.[0] || '';
const LINE = { style: 'single', widthPx: 1, colour: null, spacePt: 24 };

test('line numbering is written after the page borders, read back, printed in the margin, changed, cleared and undone', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: ['one', 'two', 'three', 'four', 'five', 'six'].map((text) => ({ text })) }));
  assert.equal(view.section.lineNumbers, null, 'a fresh document has none');
  view.setPageBorders({ offsetFrom: 'page', top: LINE, left: LINE, bottom: LINE, right: LINE });

  view.setLineNumbers({ countBy: 1, restart: 'continuous' });
  assert.deepEqual(view.section.lineNumbers, { countBy: 1, restart: 'continuous', start: 1, distancePx: null });
  assert.match(sectPrOf(view), /<\/w:pgBorders><w:lnNumType w:countBy="1" w:restart="continuous"\/><\/w:sectPr>/, 'after the page borders, as Word writes it');
  assert.equal(openDocx(view.save()).section.lineNumbers.restart, 'continuous', 'and it survives a save');

  const printed = renderPdf(view, { title: 'numbered', created: '2026-09-21T00:00:00Z' }).buffer.toString('latin1');
  for (const n of [1, 2, 3, 6]) assert.ok(printed.includes(`(${n}) Tj`), `line ${n} is numbered on the page`);

  view.setLineNumbers({ countBy: 5, restart: 'newPage', distancePx: 32 });
  assert.match(sectPrOf(view), /<w:lnNumType w:countBy="5" w:distance="480" w:restart="newPage"\/>/, 'the gap in twips');
  assert.deepEqual(view.section.lineNumbers, { countBy: 5, restart: 'newPage', start: 1, distancePx: 32 });
  const fifth = renderPdf(view, { title: 'numbered', created: '2026-09-21T00:00:00Z' }).buffer.toString('latin1');
  assert.ok(fifth.includes('(5) Tj') && !fifth.includes('(4) Tj') && !fifth.includes('(1) Tj'), 'every fifth line only');

  view.setLineNumbers(null);
  assert.equal(view.section.lineNumbers, null, 'cleared');
  assert.doesNotMatch(view.doc.doc.xml, /<w:lnNumType/);
  view.undo();
  assert.equal(view.section.lineNumbers?.countBy, 5, 'one undo step brings it back');

  // Word's default when the file says nothing about restarting: every page.
  const doc = view.doc.doc;
  doc.xml = doc.xml.replace(/<w:lnNumType\b[^>]*\/>/, '<w:lnNumType w:countBy="2"/>');
  assert.deepEqual(doc.section().lineNumbers, { countBy: 2, restart: 'newPage', start: 1, distancePx: null });
});
