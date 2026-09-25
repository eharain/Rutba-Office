// Table of contents — the engine's half.
//
// References → Table of Contents used to write plain text; it now writes a
// real field: a `w:sdt` building block (`docPartGallery` "Table of
// Contents") holding TOC1..TOCn paragraphs, each a hyperlink to a `_Toc`
// bookmark minted on its heading with a PAGEREF field for the page. Update
// Table rebuilds the entries from the current headings and bookmarks,
// keeping the field where it was; F9 does the same alongside REF/SEQ.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const fixture = () => openDocx(buildDocx({
  styles: true,
  paragraphs: [
    { text: 'Title', style: 'Title' },
    { text: 'Introduction', style: 'Heading1' },
    { text: 'Some body text here.' },
    { text: 'Background', style: 'Heading2' },
    { text: 'More body text.' },
    { text: 'Method', style: 'Heading1' },
    { text: 'Even more text.' },
  ],
}));

/** Inserts a TOC right after the title paragraph, pages given for all three headings. */
const withToc = (pages = ['2', '2', '3']) => {
  const view = fixture();
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  view.insertTableOfContents({ pages });
  return view;
};

test('insertTableOfContents writes the field the way Word does: sdt building block, TOC field, hyperlink + PAGEREF per entry', () => {
  const view = withToc();
  const engine = view.doc.doc;
  const span = engine._tocSpan();
  assert.ok(span, 'a table of contents field is in the body');
  const xml = engine._body().body.slice(span.start, span.end);

  assert.ok(xml.startsWith('<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>'));
  assert.ok(xml.endsWith('</w:sdtContent></w:sdt>'));
  assert.ok(xml.includes('<w:pStyle w:val="TOC1"/>'));
  assert.ok(xml.includes('<w:pStyle w:val="TOC2"/>'));
  assert.ok(xml.includes('<w:tab w:val="right" w:leader="dot" w:pos="'));
  assert.ok(xml.includes(' TOC \\o "1-3" \\h \\z \\u '));
  assert.ok(xml.includes('<w:fldChar w:fldCharType="begin"/>'));
  assert.ok(xml.includes('<w:fldChar w:fldCharType="separate"/>'));
  assert.ok(xml.includes('<w:fldChar w:fldCharType="end"/>'));
  assert.ok(/<w:hyperlink w:anchor="_Toc\d+" w:history="1">/.test(xml));
  assert.ok(xml.includes('PAGEREF _Toc'));
  // The outer field's begin sits before the FIRST entry's hyperlink...
  assert.ok(xml.indexOf('<w:fldChar w:fldCharType="begin"/>') < xml.indexOf('<w:hyperlink'));
  // ...and its end is the very last thing in the block, after the last entry's hyperlink closes.
  assert.ok(xml.lastIndexOf('</w:hyperlink>') < xml.lastIndexOf('<w:fldChar w:fldCharType="end"/>'));
});

test('bookmarks are placed on the headings, one per entry, hidden from the ordinary bookmark list', () => {
  const view = withToc();
  const engine = view.doc.doc;

  // _Toc bookmarks are Word's own hidden convention (a leading underscore) —
  // bookmarks() leaves every one of those out, same as _GoBack.
  assert.deepEqual(engine.bookmarks(), []);

  const starts = [...engine.xml.matchAll(/<w:bookmarkStart\b[^>]*w:name="(_Toc\d+)"/g)].map((m) => m[1]);
  assert.equal(starts.length, 3);
  assert.equal(new Set(starts).size, 3, 'three distinct bookmark names');

  const toc = engine.tableOfContents();
  assert.deepEqual(toc.entries.map((e) => e.anchor), starts);
  // Each bookmark actually sits on its heading paragraph, not off in the TOC block itself.
  const introduction = engine.paragraph(1).xml;
  assert.ok(introduction.includes('<w:bookmarkStart') && introduction.includes(starts[0]));
});

test('only levels 1-3 are listed; a Heading4 is left out', () => {
  const view = openDocx(buildDocx({
    styles: true,
    paragraphs: [
      { text: 'A', style: 'Heading1' },
      { text: 'B', style: 'Heading4' },
      { text: 'C', style: 'Heading2' },
    ],
  }));
  view.setSelection({ block: 0, offset: 0 });
  view.insertTableOfContents({});
  const toc = view.doc.doc.tableOfContents();
  assert.deepEqual(toc.entries.map((e) => e.text), ['A', 'C']);
});

test('Update Table adds an entry for a heading added afterwards', () => {
  const view = withToc();
  view.setSelection({ block: view.blocks.length - 1, offset: view.blocks[view.blocks.length - 1].text.length });
  view.splitParagraph();
  view.setParagraphFormat({ styleId: 'Heading1' });
  view.insertText('Conclusion');

  view.doc.doc.updateTableOfContents({});
  const toc = view.doc.doc.tableOfContents();
  assert.deepEqual(toc.entries.map((e) => e.text), ['Introduction', 'Background', 'Method', 'Conclusion']);
});

test('Update Table drops an entry for a heading that was removed', () => {
  const view = withToc();
  const engine = view.doc.doc;
  // "Background" is Heading2 at paragraphs() index 3 — demote it out of the
  // way by stripping its style directly (it carries a `_Toc` bookmark now,
  // which `setParagraphText`'s structural guard still refuses to touch).
  const p = engine.paragraph(3);
  assert.equal(p.text, 'Background');
  const withoutStyle = p.xml.replace('<w:pStyle w:val="Heading2"/>', '');
  engine._spliceBody(p.start, p.end, withoutStyle);

  engine.updateTableOfContents({});
  const toc = engine.tableOfContents();
  assert.deepEqual(toc.entries.map((e) => e.text), ['Introduction', 'Method']);
});

test('Update Table picks up a renamed heading\'s new words', () => {
  const view = withToc();
  const engine = view.doc.doc;
  // Same reason as above: rewrite the run directly rather than through
  // setParagraphText, which still refuses a bookmarked paragraph.
  engine.setParagraphRuns(1, [{ rPr: null, text: 'Overview' }]);
  engine.updateTableOfContents({});
  const toc = engine.tableOfContents();
  assert.equal(toc.entries[0].text, 'Overview');
});

test('the page map supplied on insert is honoured, one page per heading in document order', () => {
  const view = withToc(['4', '4', '5']);
  const toc = view.doc.doc.tableOfContents();
  assert.deepEqual(toc.entries.map((e) => e.page), ['4', '4', '5']);
});

test('with no page map, entries are inserted blank; Update Table with one fills them; a later plain Update Table keeps them', () => {
  const view = fixture();
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  view.insertTableOfContents({});
  assert.deepEqual(view.doc.doc.tableOfContents().entries.map((e) => e.page), [null, null, null]);

  view.doc.doc.updateTableOfContents({ pages: ['7', '7', '8'] });
  assert.deepEqual(view.doc.doc.tableOfContents().entries.map((e) => e.page), ['7', '7', '8']);

  view.doc.doc.updateTableOfContents({});
  assert.deepEqual(view.doc.doc.tableOfContents().entries.map((e) => e.page), ['7', '7', '8']);
});

test('a table of contents survives a save and reopen, and F9 refreshes it there too', () => {
  const view = withToc();
  const reopened = openDocx(view.save());
  assert.ok(reopened.doc.doc.hasTableOfContents());
  assert.deepEqual(reopened.doc.doc.tableOfContents().entries.map((e) => e.text), ['Introduction', 'Background', 'Method']);

  reopened.doc.doc.setParagraphRuns(1, [{ rPr: null, text: 'Overview' }]);
  reopened.updateFields();
  assert.equal(reopened.doc.doc.tableOfContents().entries[0].text, 'Overview');
});

test('a Word-authored table of contents fixture (no docPartUnique, ordinary hand-built shape) is read and refreshed by Update Table', () => {
  const view = fixture();
  // A hand-built TOC in the shape Word itself writes, dropped in ahead of the
  // headings — no insertTableOfContents call, so nothing here rode this
  // engine's own writer.
  const engine = view.doc.doc;
  const heading = engine.paragraph(1);
  engine._mintBookmark('_Toc555000001', 1, 1);
  const handBuilt =
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent>' +
    '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:hyperlink w:anchor="_Toc555000001" w:history="1">' +
    '<w:r><w:t xml:space="preserve">Stale Title</w:t></w:r><w:r><w:tab/></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc555000001 \\h </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>9</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '</w:hyperlink>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '</w:p></w:sdtContent></w:sdt>';
  const p0 = engine.paragraph(0);
  engine._spliceBody(p0.end, p0.end, handBuilt);

  assert.ok(engine.hasTableOfContents());
  const before = engine.tableOfContents();
  assert.deepEqual(before.entries.map((e) => e.text), ['Stale Title']);

  engine.updateTableOfContents({});
  const after = engine.tableOfContents();
  assert.deepEqual(after.entries.map((e) => e.text), ['Introduction', 'Background', 'Method']);
});

test('undo after insertTableOfContents takes the field off again', () => {
  const view = withToc();
  assert.ok(view.doc.doc.hasTableOfContents());

  view.undo();
  assert.equal(view.doc.doc.hasTableOfContents(), false);
});

test('Remove Table of Contents takes the field off; false the second time', () => {
  const view = withToc();
  assert.equal(view.removeTableOfContents(), true);
  assert.equal(view.doc.doc.hasTableOfContents(), false);
  assert.equal(view.removeTableOfContents(), false);
});
