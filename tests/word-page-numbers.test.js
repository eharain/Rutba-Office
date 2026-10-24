/**
 * Page numbers from the header and footer box (2.18.0, A3.10 and A3.8).
 *
 * In 2.17.1 "Page number" and "Of total" wrote "{PAGE} of {PAGES}" into the
 * file as words, so no page, print, PDF or Word ever showed a number. Now
 * {PAGE} and {PAGES} are PAGE and NUMPAGES fields, drawn on each page and in
 * the PDF, read by Word, and read back by the box as the braces that wrote
 * them. A tab in a line moves to the centre, a second to the right.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

const LONG = Array.from({ length: 140 }, (_, i) => `Clause ${i + 1}. The shop opens at nine and closes at half past five, every day but Sunday, and the stock is counted on the last Friday of each month.`);
const footerXml = (view) => {
  const pkg = OoxmlPackage.read(view.save());
  const name = pkg.partNames().find((n) => /^word\/footer\d+\.xml$/.test(n));
  return name ? pkg.text(name) : '';
};

test('Page {PAGE} of {PAGES} is written as the PAGE and NUMPAGES fields, not as words', () => {
  const view = openDocx(buildDocx({ paragraphs: LONG }));
  view.setBand('footer', ['Page {PAGE} of {PAGES}']);
  const xml = footerXml(view);
  assert.match(xml, /<w:fldSimple w:instr=" PAGE "><w:r><w:t>1<\/w:t><\/w:r><\/w:fldSimple>/);
  assert.match(xml, /<w:fldSimple w:instr=" NUMPAGES ">/);
  assert.ok(!xml.includes('{PAGE'), 'no braces in the file');
  assert.match(xml, /<w:t xml:space="preserve">Page <\/w:t>[\s\S]*<w:t xml:space="preserve"> of <\/w:t>/);
});

test('each page draws its own number and the count, and so does the PDF', () => {
  const view = openDocx(buildDocx({ paragraphs: LONG }));
  view.setBand('footer', ['Page {PAGE} of {PAGES}']);
  const laid = view.pages;
  assert.ok(laid.pages.length >= 3, 'several pages: ' + laid.pages.length);
  const count = laid.pages.length;
  assert.equal(laid.pages[0].footer[0].text, `Page 1 of ${count}`);
  assert.equal(laid.pages[1].footer[0].text, `Page 2 of ${count}`);
  // Reopened from the saved file, as Word or a later session reads it.
  const again = openDocx(view.save());
  assert.equal(again.pages.pages[count - 1].footer[0].text, `Page ${count} of ${count}`);
  const pdf = renderPdf(again, { created: '2026-10-25T00:00:00Z' }).buffer.toString('latin1');
  assert.match(pdf, new RegExp('\\(Page \\) Tj[^(]*\\(2\\) Tj[^(]*\\( of \\) Tj[^(]*\\(' + count + '\\) Tj'), 'the PDF draws the second page\'s number');
});

test('a footer with page numbers is still edited from the box, which reads the fields back as braces', () => {
  const view = openDocx(buildDocx({ paragraphs: ['One page'] }));
  view.setBand('footer', ['{PAGE} of {PAGES}']);
  const info = view.doc.doc.bandInfo();
  assert.equal(info.footer.structural, false, 'not locked by its own fields');
  const band = view.doc.headerFooters().footers.default;
  assert.deepEqual(band.paragraphs[0].runs.map((r) => r.field), ['PAGE', null, 'NUMPAGES']);
  view.setBand('footer', ['Page {PAGE}']);
  assert.match(footerXml(view), /Page <\/w:t><\/w:r><w:fldSimple w:instr=" PAGE ">/);
  assert.ok(!/NUMPAGES/.test(footerXml(view)), 'the old line is gone');
});

test('a tab in a header line moves to the centre, a second to the right, and the stops are in the file', () => {
  const view = openDocx(buildDocx({ paragraphs: ['One page'] }));
  view.setBand('header', ['Larkspur\tStock list\tPage {PAGE}']);
  const pkg = OoxmlPackage.read(view.save());
  const xml = pkg.text(pkg.partNames().find((n) => /^word\/header\d+\.xml$/.test(n)));
  assert.match(xml, /<w:tabs><w:tab w:val="center" w:pos="\d+"\/><w:tab w:val="right" w:pos="\d+"\/><\/w:tabs>/);
  const band = view.doc.headerFooters().headers.default;
  assert.equal(band.paragraphs[0].text, 'Larkspur\tStock list\tPage 1');
  // A line that starts with a tab keeps it: the tab-only run is read.
  view.setBand('header', ['\t{PAGE}']);
  assert.equal(view.doc.headerFooters().headers.default.paragraphs[0].text, '\t1');
});
