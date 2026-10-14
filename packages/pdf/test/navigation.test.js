'use strict';

// Bookmarks, page labels, named destinations, links and form fields carried
// through a reorder, an extract, a split and a merge (assemble.js with
// navigation.js). A five-page file is given an outline (nested, one item
// through a named destination, one through a GoTo action, one a web link),
// page labels (i, ii, then A-1 onward), a link through a named destination
// and a form field on two pages; every new file is read back here and by
// pdf.js.

const test = require('node:test');
const assert = require('node:assert');
const { PdfDocument } = require('../src/document');
const { openPdf, Ref } = require('../src/reader');
const { assemble, splitEvery } = require('../src/assemble');
const { readOutline, pageLabels, labelText, namedDestinations } = require('../src/navigation');
const { readFields } = require('../src/forms');
const { book } = require('./nav-fixture');

const created = '2026-10-14T09:00:00Z';

const firstWords = (file) => file.pages().map((_, i) => file.pageText(i).split('\n')[0]);
/** The outline as "title->page words" lines, indented by level. */
function outlineOf(file) {
  const pageOf = (dest) => {
    if (!dest) return '';
    const idx = file.pages().findIndex((p) => p.ref && dest[0] instanceof Ref && p.ref.num === dest[0].num);
    return idx < 0 ? '?' : file.pageText(idx).split('\n')[0];
  };
  const lines = [];
  const walk = (nodes, depth) => nodes.forEach((n) => { lines.push(`${'  '.repeat(depth)}${n.title}->${n.action ? 'link' : pageOf(n.dest)}`); walk(n.kids, depth + 1); });
  walk(readOutline(file), 0);
  return lines;
}
const labelsOf = (file) => (pageLabels(file) || []).map(labelText);

test('the source reads: outline, named destinations, labels', () => {
  const file = openPdf(book());
  assert.deepEqual(outlineOf(file), ['Contents->Page 1', '  Intro->Page 1', '  Body->Page 3', '    Detail->Page 4', '    End->Page 5', 'Web->link']);
  assert.deepEqual(labelsOf(file), ['i', 'ii', 'A-1', 'A-2', 'A-3']);
  assert.ok(namedDestinations(file).has('chapter-two'));
  const body = readOutline(file)[0].kids[1];
  assert.equal(body.bold, true);
  assert.deepEqual(body.color, [0.8, 0, 0]);
});

test('a reorder keeps every bookmark on its page where the page now stands, and each page its label', () => {
  const src = openPdf(book());
  const out = openPdf(assemble([4, 2, 0, 1, 3].map((page) => ({ file: src, page }))));
  assert.deepEqual(firstWords(out), ['Page 5', 'Page 3', 'Page 1', 'Page 2', 'Page 4']);
  assert.deepEqual(outlineOf(out), ['Contents->Page 1', '  Intro->Page 1', '  Body->Page 3', '    Detail->Page 4', '    End->Page 5', 'Web->link']);
  assert.deepEqual(labelsOf(out), ['A-3', 'A-1', 'i', 'ii', 'A-2']);
  // The link through a named destination now names its page outright.
  const link = out.resolve(out.resolve(out.pages()[2].dict.get('Annots'))[0]);
  const dest = link.get('Dest');
  assert.ok(Array.isArray(dest) && dest[0] instanceof Ref);
  assert.equal(dest[0].num, out.pages()[4].ref.num);
  // The bookmark keeps its look.
  const body = readOutline(out)[0].kids[1];
  assert.equal(body.bold, true);
});

test('an extract keeps the bookmarks of the pages taken, a bookmark to a page left out giving way to its own', () => {
  const src = openPdf(book());
  const out = openPdf(assemble([3, 4].map((page) => ({ file: src, page }))));
  // Contents (page 1) and Body (page 3) are gone; Detail and End move up.
  assert.deepEqual(outlineOf(out), ['Detail->Page 4', 'End->Page 5', 'Web->link']);
  assert.deepEqual(labelsOf(out), ['A-2', 'A-3']);
  // The field: one of its two widgets came, so the field did.
  const fields = readFields(out);
  assert.equal(fields.length, 1);
  assert.equal(fields[0].name, 'reviewer');
  assert.equal(fields[0].value, 'Page reviewer');
  assert.deepEqual(fields[0].widgets.map((w) => w.page), [0]);
});

test('a split into files of two gives each part its bookmarks and labels', () => {
  const src = openPdf(book());
  const parts = splitEvery(5, 2).map((group) => openPdf(assemble(group.map((page) => ({ file: src, page })))));
  assert.deepEqual(parts.map(outlineOf), [
    ['Contents->Page 1', '  Intro->Page 1', 'Web->link'],
    ['Body->Page 3', '  Detail->Page 4', 'Web->link'],
    ['End->Page 5', 'Web->link'],
  ]);
  assert.deepEqual(parts.map(labelsOf), [['i', 'ii'], ['A-1', 'A-2'], ['A-3']]);
  assert.deepEqual(parts.map((p) => readFields(p).length), [1, 1, 0]);
});

test('a merge follows one file\'s bookmarks with the other\'s, keeps both labels, and keeps two fields of one name apart', () => {
  const a = openPdf(book('Alpha'));
  const b = openPdf(book('Beta'));
  const items = [...[0, 1, 2, 3, 4].map((page) => ({ file: a, page })), ...[0, 1, 2, 3, 4].map((page) => ({ file: b, page }))];
  const out = openPdf(assemble(items));
  const lines = outlineOf(out);
  assert.deepEqual(lines, [
    'Contents->Alpha 1', '  Intro->Alpha 1', '  Body->Alpha 3', '    Detail->Alpha 4', '    End->Alpha 5', 'Web->link',
    'Contents->Beta 1', '  Intro->Beta 1', '  Body->Beta 3', '    Detail->Beta 4', '    End->Beta 5', 'Web->link',
  ]);
  assert.deepEqual(labelsOf(out), ['i', 'ii', 'A-1', 'A-2', 'A-3', 'i', 'ii', 'A-1', 'A-2', 'A-3']);
  const fields = readFields(out);
  assert.deepEqual(fields.map((f) => `${f.name}=${f.value}`), ['reviewer=Alpha reviewer', 'reviewer_2=Beta reviewer']);
  // The second file's link goes to the second file's page, through its own named destination.
  const link = out.resolve(out.resolve(out.pages()[5].dict.get('Annots'))[0]);
  assert.equal(link.get('Dest')[0].num, out.pages()[8].ref.num);
});

test('a file without bookmarks or labels gets none, and one with them merged after it keeps its own', () => {
  const plain = new PdfDocument({ created });
  plain.addPage().text('Plain', 72, 100, { font: 'Helvetica', size: 24 });
  const p = openPdf(plain.toBuffer());
  const alone = openPdf(assemble([{ file: p, page: 0 }]));
  assert.equal(alone.catalog().get('Outlines'), undefined);
  assert.equal(alone.catalog().get('PageLabels'), undefined);
  const b = openPdf(book('Beta'));
  const out = openPdf(assemble([{ file: p, page: 0 }, { file: b, page: 0 }, { file: b, page: 2 }]));
  assert.deepEqual(labelsOf(out), ['1', 'i', 'A-1']);
  assert.deepEqual(outlineOf(out), ['Contents->Beta 1', '  Intro->Beta 1', '  Body->Beta 3', 'Web->link']);
});

test('pdf.js reads the carried bookmarks, their pages and the labels', async (t) => {
  const src = openPdf(book());
  const bytes = assemble([4, 2, 0, 1, 3].map((page) => ({ file: src, page })));
  const lib = await require('./pdfjs').loadPdfjs();
  if (!lib) { t.skip('pdf.js is not vendored'); return; }
  const task = lib.getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  const doc = await task.promise;
  try {
    const outline = await doc.getOutline();
    assert.deepEqual(outline.map((o) => o.title), ['Contents', 'Web']);
    assert.deepEqual(outline[0].items.map((o) => o.title), ['Intro', 'Body']);
    const pageOf = async (dest) => doc.getPageIndex(dest[0]);
    assert.equal(await pageOf(outline[0].items[1].dest), 1, 'Body now on the second page');
    assert.equal(await pageOf(outline[0].items[1].items[0].dest), 4, 'Detail, through its named destination, on the fifth');
    assert.equal(outline[1].url, 'https://rutba.io/');
    assert.deepEqual(await doc.getPageLabels(), ['A-3', 'A-1', 'i', 'ii', 'A-2']);
  } finally {
    await task.destroy();
  }
});
