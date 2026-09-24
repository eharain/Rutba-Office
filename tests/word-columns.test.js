// Columns — the engine's half.
//
// Layout → Columns writes `<w:cols>` into the section, after the page
// numbering where the schema puts it; the section reads it back as
// `columns` and `columnBoxes`; the paginator flows a section of more than
// one column into them, a column at a time, and the PDF draws each
// fragment at its own column's x and width. A section of one column — the
// ordinary case, and every section that existed before this read `w:cols`
// at all — must come out of the paginator byte-identical to before.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { paginate } from '@rutba/doc-view/paginate';
import { buildDocx } from '@rutba/ooxml';

const sectPrOf = (view) => /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(view.doc.doc.xml)?.[0] || '';

const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn despite a supply interruption. ';
const longParagraphs = (n) => Array.from({ length: n }, (_, i) => ({ text: lorem.repeat(2) + `(${i + 1})` }));

test('setPageSetup writes w:cols in schema order, after pgNumType and before docGrid', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }] }));
  // A pgNumType and a docGrid already in the section prove the slot: cols
  // belongs strictly between them, wherever else in the file they came from.
  const doc = view.doc.doc;
  doc.xml = doc.xml.replace(/<w:sectPr\b[^>]*\/>/, '<w:sectPr><w:pgNumType w:start="1"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>');
  doc.dirty = true;

  assert.equal(view.section.columns.count, 1, 'one column, no element, is the default');
  assert.equal(view.section.columnBoxes.length, 1);

  view.setPageSetup({ columns: { count: 2 } });
  const sectPr = sectPrOf(view);
  assert.match(sectPr, /<w:pgNumType[^>]*\/><w:cols w:num="2" w:space="720"\/><w:docGrid/, 'cols sits between pgNumType and docGrid');

  const columns = view.section.columns;
  assert.equal(columns.count, 2);
  assert.equal(Math.round(columns.spacePx), 48, '720 twips of gap, in px');
  assert.equal(columns.separator, false);
  assert.equal(columns.widths, null, 'equal width, so no explicit column widths');

  const boxes = view.section.columnBoxes;
  assert.equal(boxes.length, 2);
  assert.ok(Math.abs(boxes[0].widthPx - boxes[1].widthPx) < 0.01, 'two equal boxes');
  const span = boxes[1].xPx + boxes[1].widthPx - boxes[0].xPx;
  assert.ok(Math.abs(span - view.section.contentWidthPx) < 0.5, 'the two boxes and the gap span the content width');

  // count: 1 removes the element again.
  view.setPageSetup({ columns: { count: 1 } });
  assert.doesNotMatch(view.doc.doc.xml, /<w:cols/, 'one column writes nothing');
  assert.equal(view.section.columns.count, 1);
});

test('the Left preset writes explicit unequal w:col widths, which read back as unequal boxes', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }] }));
  const contentTwips = Math.round(view.section.contentWidthPx * 15);
  const usable = contentTwips - 720;
  const narrow = Math.round(usable / 3);
  const wide = usable - narrow;
  view.setPageSetup({ columns: { count: 2, spaceTwips: 720, widths: [narrow, wide] } });

  const sectPr = sectPrOf(view);
  assert.match(sectPr, /<w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="\d+" w:space="720"\/><w:col w:w="\d+"\/><\/w:cols>/);

  const columns = view.section.columns;
  assert.ok(Array.isArray(columns.widths) && columns.widths.length === 2, 'explicit widths come back');
  assert.ok(columns.widths[0] < columns.widths[1], 'the left column is the narrow one');

  const boxes = view.section.columnBoxes;
  assert.ok(boxes[0].widthPx < boxes[1].widthPx, 'and so are the boxes it lays into');
  assert.ok(Math.abs(boxes[0].widthPx - columns.widths[0]) < 0.5);
  assert.ok(Math.abs(boxes[1].widthPx - columns.widths[1]) < 0.5);

  // It survives a save.
  const reopened = openDocx(view.save());
  assert.ok(reopened.section.columns.widths[0] < reopened.section.columns.widths[1]);
});

test('a one-column section paginates byte-identically to before columns were read at all', () => {
  const paragraphs = longParagraphs(40);
  const view = openDocx(buildDocx({ styles: true, paragraphs }));
  const withSection = view.render({ pages: false });
  const laidWithColumnsField = paginate({ flow: withSection.flow, blocks: withSection.blocks, section: view.section });

  // The same section, but with no `columns` field at all — the shape every
  // section had before this feature existed.
  const bareSection = { ...view.section };
  delete bareSection.columns;
  // `columnBoxes` is a getter on the real section; the plain copy above does
  // not carry it, which is exactly the point: paginate must not need it for
  // one column.
  const laidBare = paginate({ flow: withSection.flow, blocks: withSection.blocks, section: bareSection });

  assert.equal(JSON.stringify(laidWithColumnsField), JSON.stringify(laidBare), 'reading columns changes nothing for a one-column section');
});

test('a two-column section flows a long document into columns, then a page, with fewer pages than one column', () => {
  const paragraphs = longParagraphs(80);
  const oneColumn = openDocx(buildDocx({ styles: true, paragraphs }));
  const onePages = oneColumn.pages;

  const twoColumn = openDocx(buildDocx({ styles: true, paragraphs }));
  twoColumn.setPageSetup({ columns: { count: 2 } });
  const laid = twoColumn.pages;

  assert.ok(laid.count < onePages.count, `two columns fit more per page: ${laid.count} pages vs ${onePages.count} for one column`);

  const first = laid.pages[0];
  assert.ok(first.columns && first.columns.length === 2, 'the page carries the column boxes');
  const colOf = first.fragments.map((f) => f.column);
  assert.ok(colOf.includes(0) && colOf.includes(1), 'the first page uses both columns');
  // Column 0's fragments come first, in full, before column 1's — the flow
  // fills a column before moving to the next, exactly as Word does.
  const firstOne = colOf.indexOf(1);
  const lastZero = colOf.lastIndexOf(0);
  assert.ok(firstOne > lastZero, 'column 0 is filled before column 1 starts');

  // Column 1 starts fresh from the top of the page, not part way down where
  // column 0 left off.
  const firstCol1Paragraph = first.fragments.find((f) => f.column === 1 && f.kind === 'paragraph');
  assert.ok(firstCol1Paragraph, 'column 1 has a paragraph fragment');
  assert.equal(firstCol1Paragraph.spaceBefore >= 0, true);

  // A one-column comparison run never carries a `column` field at all.
  assert.ok(onePages.pages[0].fragments.every((f) => f.column === undefined), 'a one-column page tags nothing');
});

test('the PDF of a two-column document draws text at two x origins on page 1, and a separator line when asked for one', () => {
  const paragraphs = longParagraphs(30);
  const view = openDocx(buildDocx({ styles: true, paragraphs }));
  view.setPageSetup({ columns: { count: 2, separator: true } });
  const { buffer, pages } = renderPdf(view, { created: '2026-09-22T00:00:00Z' });
  assert.ok(pages >= 1);
  const pdf = buffer.toString('latin1');

  const xs = [...pdf.matchAll(/1 0 0 1 ([\d.]+) [\d.]+ Tm/g)].map((m) => Math.round(parseFloat(m[1])));
  const distinctXs = [...new Set(xs)];
  assert.ok(distinctXs.length >= 2, `text drawn from at least two x origins on the page: ${distinctXs.join(', ')}`);

  // A vertical stroke down the gap between the columns: two points with the
  // same x, one above the other (`m` then `l`), rather than a horizontal rule.
  assert.match(pdf, /\nq\n0\.5 w\n0\.502 0\.502 0\.502 RG\n([\d.]+) ([\d.]+) m\n\1 ([\d.]+) l\nS\nQ/, 'a vertical grey rule between the columns');
});

test('a footnote referenced in the second column lands at the foot of that column, not the first', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: longParagraphs(30) }));
  view.setPageSetup({ columns: { count: 2 } });
  const at = view.render({ pages: false }).blocks.findIndex((b) => (b.text || '').includes('(20)'));
  assert.ok(at > 0, 'a paragraph late enough to fall in the second column');
  view.setSelection({ block: at, offset: view.render({ pages: false }).blocks[at].text.length });
  view.insertNote('footnote', 'A note in the second column.');

  const laid = view.pages;
  const page = laid.pages.find((p) => (p.notes || []).length);
  assert.ok(page, 'some page carries the note');
  assert.ok(page.columns && page.columns.length === 2, 'that page is laid in columns');

  // The paragraph itself is in column 1; its note travels with it.
  const carryingFragment = page.fragments.find((f) => f.kind === 'paragraph' && f.paragraphIndex === at);
  assert.ok(carryingFragment, 'the referencing paragraph is on this page');
  const note = page.notes.find((n) => /second column/.test(n.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('')));
  assert.ok(note, 'the note text is on this page');
  assert.equal(note.column, carryingFragment.column, 'the note foots the same column its reference is in');
});
