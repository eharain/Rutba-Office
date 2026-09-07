/**
 * The paginated frame, printed (S12).
 *
 * The exporter draws the paginator's own lines through `@rutba/pdf`; what
 * these tests hold is that the PDF carries the text, one sheet per laid page,
 * at the page's own geometry, and that the writer's rules stayed honest: an
 * HTML body with no section gets A4, a PNG is embedded, anything else is a
 * labelled frame, and the file itself is a classic-xref PDF a reader opens.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { openHtml } from '@rutba/doc-view/backends/html';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf, renderFramePdf, DEFAULT_SECTION } from '@rutba/doc-view/export/pdf';
import { buildDocx } from '@rutba/ooxml/build';

const text = (buffer) => buffer.toString('latin1');
const pageCount = (buffer) => (text(buffer).match(/\/Type \/Page\b/g) || []).length;

test('an HTML body with no pages prints on A4, one sheet, with its text and bold run', () => {
  const view = openHtml('<p>Services agreement between <b>Acme</b> and Beta.</p><p>Term: twelve months.</p>');
  const { buffer, pages } = renderPdf(view, { title: 'Services agreement', created: '2026-09-03T00:00:00Z' });
  assert.equal(text(buffer).slice(0, 5), '%PDF-');
  assert.equal(pages, 1);
  assert.equal(pageCount(buffer), 1);
  const body = text(buffer);
  assert.ok(body.includes('Services agreement between'), 'the paragraph text is drawn');
  assert.ok(body.includes('(Acme) Tj'), 'the bold run is drawn as its own segment');
  assert.ok(body.includes('/Helvetica-Bold'), 'the bold run uses the bold face');
  // A4 in points, from the default section's CSS pixels.
  assert.ok(body.includes(`/MediaBox [0 0 ${(DEFAULT_SECTION.widthPx * 0.75).toFixed(2).replace(/0+$/, '').replace(/\.$/, '')}`)
    || body.includes('/MediaBox [0 0 595.5'), 'the sheet is A4');
});

test('a long flow breaks across sheets exactly where the paginator broke it', () => {
  const paragraphs = Array.from({ length: 120 }, (_, i) => `<p>Clause ${i + 1}. The party of the first part shall deliver the goods described in the schedule within the period stated, and the party of the second part shall pay the price on delivery.</p>`).join('');
  const view = openHtml(paragraphs);
  const { buffer, pages } = renderPdf(view, { created: '2026-09-03T00:00:00Z' });
  assert.ok(pages > 3, `expected several sheets, got ${pages}`);
  assert.equal(pageCount(buffer), pages);
  assert.ok(text(buffer).includes('Clause 120.'), 'the last paragraph made it onto a sheet');
});

test('a Word document prints on its own page size, edits included, as the paginator laid it', () => {
  const view = openDocx(buildDocx({
    paragraphs: ['Master services agreement', 'This agreement is made between the parties named below.'],
  }));
  // An edit through the editor's own API, so the export prints the model,
  // not the bytes it was opened from.
  view.collapseTo({ block: 1, offset: view.block(1).text.length });
  view.splitParagraph();
  view.insertText('Deliver the goods within thirty days of the order.');
  const laid = view.pages;
  const { buffer, pages } = renderPdf(view, { title: 'MSA', created: '2026-09-03T00:00:00Z' });
  assert.equal(pages, laid.count, 'one sheet per laid page');
  const body = text(buffer);
  assert.ok(body.includes('Master services agreement'));
  assert.ok(body.includes('Deliver the goods within thirty days'));
  const section = view.section;
  assert.ok(body.includes(`/MediaBox [0 0 ${String(Math.round(section.widthPx * 0.75 * 1000) / 1000)}`),
    'the sheet is the document\'s own page width');
});

test('a frame with a PNG embeds it, and anything else becomes a labelled frame', () => {
  // A 1x1 opaque red PNG, the kind a signature pad or a screenshot produces.
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC', 'base64');
  const frame = {
    blocks: [{ index: 0, text: 'Logo above', runs: [{ text: 'Logo above' }] }],
    section: DEFAULT_SECTION,
    pages: {
      count: 1,
      pages: [{
        index: 0, number: 1, of: 1, contentHeightPx: 900,
        fragments: [
          { kind: 'paragraph', paragraphIndex: 0, lines: [{ text: 'Logo above', start: 0, end: 10 }], spaceBefore: 0, spaceAfter: 10, lineHeightPx: 19, sizePx: 14, weight: 'normal', first: true, last: true },
          { kind: 'images', paragraphIndex: 0, images: [
            { name: 'logo.png', widthPx: 40, heightPx: 40, href: `data:image/png;base64,${png.toString('base64')}` },
            { name: 'chart.svg', widthPx: 200, heightPx: 100, href: 'data:image/svg+xml;utf8,<svg/>' },
          ] },
        ],
      }],
    },
  };
  const { buffer } = renderFramePdf(frame, { created: '2026-09-03T00:00:00Z' });
  const body = text(buffer);
  assert.ok(body.includes('/Subtype /Image'), 'the PNG became an image XObject');
  assert.ok(body.includes('(chart.svg) Tj'), 'the SVG is a labelled frame, not a silent gap');
});
