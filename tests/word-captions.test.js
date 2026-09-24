// Captions — the engine's half.
//
// Insert → Caption writes a new paragraph, styled "Caption", after the
// block the caret sits in: the label, a space, a SEQ field for the running
// number, then `: ` and the caller's own words — `Figure 3: a diagram of
// it`, the way Word's own dialog writes one. The number is a COMPLEX field
// (`w:fldChar` begin/separate/end around ` SEQ Figure \* ARABIC `), the
// shape Word itself writes for a caption; a cross-reference's `w:fldSimple`
// is a different, simpler thing. Update Fields renumbers every label's
// captions from where they now sit, the same way it refreshes a REF.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const fixture = () => openDocx(buildDocx({
  styles: true,
  paragraphs: [
    { text: 'First paragraph text', style: 'Heading1' },
    { text: 'Second paragraph text' },
    { text: 'Third paragraph text' },
  ],
}));

const pXml = (view, i) => view.doc.doc.paragraph(i).xml;

test('addCaption inserts a new paragraph, styled Caption, right after `at`', () => {
  const view = fixture();
  const engine = view.doc.doc;

  engine.addCaption({ label: 'Figure', text: 'a diagram of the pipeline', at: 1 });

  assert.equal(engine.paragraphCount(), 4, 'one paragraph added');
  const p = engine.paragraph(2);
  assert.equal(p.style, 'Caption');
  assert.equal(p.text, 'Figure 1: a diagram of the pipeline');
  // The paragraph after it is untouched, still where it was.
  assert.equal(engine.paragraph(3).text, 'Third paragraph text');
});

test('numbering climbs across three captions of the same label, in document order', () => {
  const view = fixture();
  const engine = view.doc.doc;

  engine.addCaption({ label: 'Figure', text: 'one', at: 0 });
  engine.addCaption({ label: 'Figure', text: 'two', at: 1 });
  engine.addCaption({ label: 'Figure', text: 'three', at: 2 });

  assert.equal(engine.paragraph(1).text, 'Figure 1: one');
  assert.equal(engine.paragraph(2).text, 'Figure 2: two');
  assert.equal(engine.paragraph(3).text, 'Figure 3: three');
});

test('two labels are numbered independently', () => {
  const view = fixture();
  const engine = view.doc.doc;

  engine.addCaption({ label: 'Figure', text: 'f1', at: 0 });
  engine.addCaption({ label: 'Table', text: 't1', at: 1 });
  engine.addCaption({ label: 'Figure', text: 'f2', at: 2 });
  engine.addCaption({ label: 'Table', text: 't2', at: 3 });

  assert.equal(engine.paragraph(1).text, 'Figure 1: f1');
  assert.equal(engine.paragraph(2).text, 'Table 1: t1');
  assert.equal(engine.paragraph(3).text, 'Figure 2: f2');
  assert.equal(engine.paragraph(4).text, 'Table 2: t2');
});

test('deleting the middle caption and updating fields renumbers what is left', () => {
  const view = fixture();
  const engine = view.doc.doc;

  engine.addCaption({ label: 'Figure', text: 'one', at: 0 });
  engine.addCaption({ label: 'Figure', text: 'two', at: 1 });
  engine.addCaption({ label: 'Figure', text: 'three', at: 2 });
  assert.equal(engine.paragraph(2).text, 'Figure 2: two');

  // The middle caption taken out directly — a caption paragraph carries a
  // complex field and so stays structural (see word-xref.test.js's own
  // complex-field test), the same reason the ordinary delete path refuses
  // it; this is the paragraph gone the way Word's own Delete key leaves it,
  // for the renumbering this test is actually about.
  const middle = engine.paragraph(2);
  engine._spliceBody(middle.start, middle.end, '');
  assert.equal(engine.paragraph(2).text, 'Figure 3: three', 'the third caption slid up, its number not yet fixed');

  const changed = engine.refreshRefFields();
  assert.ok(changed >= 1, 'at least the surviving caption\'s number changed');
  assert.equal(engine.paragraph(1).text, 'Figure 1: one');
  assert.equal(engine.paragraph(2).text, 'Figure 2: three', 'renumbered to its new place, not its old one');
});

test('the SEQ field\'s XML is exactly the shape Word writes', () => {
  const view = fixture();
  const engine = view.doc.doc;
  engine.addCaption({ label: 'Figure', text: 'a diagram', at: 1 });

  assert.equal(
    pXml(view, 2),
    '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>'
    + '<w:r><w:t xml:space="preserve">Figure </w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> SEQ Figure \\* ARABIC </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + '<w:r><w:t xml:space="preserve">1</w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    + '<w:r><w:t xml:space="preserve">: a diagram</w:t></w:r>'
    + '</w:p>',
  );

  // Read back through the general field-parsing path too — the same one a
  // REF field uses — so the page can shade it exactly as it shades a REF.
  const field = engine.paragraph(2).runs.find((r) => r.field);
  assert.deepEqual(field.field, { instr: ' SEQ Figure \\* ARABIC ', kind: 'seq', name: 'Figure' });
  assert.equal(field.text, '1');
  assert.equal(engine.paragraph(2).structural, true, 'a complex field keeps the paragraph structural');
  assert.ok(engine.paragraph(2).structuralTags.includes('w:fldChar'));
});

test('the Caption style is written once, whatever number of captions the document gets', () => {
  const view = fixture();
  const engine = view.doc.doc;
  engine.addCaption({ label: 'Figure', text: 'one', at: 0 });
  engine.addCaption({ label: 'Table', text: 'two', at: 1 });

  const stylesXml = engine.pkg.text('word/styles.xml');
  const count = (stylesXml.match(/<w:style\b[^>]*\bw:styleId="Caption"/g) || []).length;
  assert.equal(count, 1, 'one Caption style, not one per caption');
  assert.match(stylesXml, /<w:style w:type="paragraph" w:styleId="Caption">[\s\S]*?<w:i\/>[\s\S]*?<w:sz w:val="18"/, 'italic, 9pt');
});

test('a caption survives a save and reopen, its field kept and its number renumbered on update', () => {
  const view = fixture();
  const engine = view.doc.doc;
  engine.addCaption({ label: 'Figure', text: 'one', at: 0 });
  engine.addCaption({ label: 'Figure', text: 'two', at: 1 });

  const reopened = openDocx(view.save());
  const doc2 = reopened.doc.doc;
  assert.equal(doc2.paragraph(1).text, 'Figure 1: one');
  assert.equal(doc2.paragraph(2).text, 'Figure 2: two');

  // Delete the first caption in the REOPENED copy, then update fields —
  // proof the round trip kept a real, renumberable SEQ field and not a
  // caption frozen as inert text.
  const first = doc2.paragraph(1);
  doc2._spliceBody(first.start, first.end, '');
  doc2.refreshRefFields();
  assert.equal(doc2.paragraph(1).text, 'Figure 1: two');
});

test('a Word-authored SEQ caption round-trips: read as a field, renumbered by Update Fields', () => {
  // Built by hand, the way word-xref.test.js builds its own complex-field
  // fixture — standing in for a document that came from Word itself rather
  // than from `addCaption`.
  const view = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Figure 1: existing', style: 'Heading1' },
  ] }));
  const engine = view.doc.doc;
  const p = engine.paragraph(0);
  const caption = '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>'
    + '<w:r><w:t xml:space="preserve">Figure </w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> SEQ Figure \\* ARABIC </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + '<w:r><w:t xml:space="preserve">1</w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    + '<w:r><w:t xml:space="preserve">: a scan from Word itself</w:t></w:r>'
    + '</w:p>';
  engine._spliceBody(p.start, p.end, caption);

  assert.equal(engine.paragraph(0).text, 'Figure 1: a scan from Word itself');
  assert.equal(engine.fields()[0]?.kind, 'seq');

  // A second caption ahead of it — Update Fields must still read the first
  // as "1", since it is first in the (now two-caption) document.
  engine.addCaption({ label: 'Figure', text: 'a diagram after it', at: 0 });
  assert.equal(engine.paragraph(1).text, 'Figure 2: a diagram after it');
  assert.equal(engine.paragraph(0).text, 'Figure 1: a scan from Word itself');
});

test('undo after insertCaption takes the paragraph off again', () => {
  const view = fixture();
  view.setSelection({ block: 1, offset: 0 });

  const before = view.doc.doc.paragraphCount();
  view.insertCaption({ label: 'Figure', text: 'a diagram' });
  assert.equal(view.doc.doc.paragraphCount(), before + 1);
  assert.equal(view.block(2).text, 'Figure 1: a diagram');

  view.undo();
  assert.equal(view.doc.doc.paragraphCount(), before);
  assert.equal(view.doc.doc.fields().length, 0);
});
