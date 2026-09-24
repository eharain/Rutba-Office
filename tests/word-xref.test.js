// Cross-references — the engine's half.
//
// Insert → Cross-reference writes a REF field: `<w:fldSimple w:instr=" REF
// Summary \h ">` wrapping a run of the bookmark's own words, sitting among
// the paragraph's other runs exactly as Word writes one. The paragraph that
// carries it stays editable — the field is one atomic run, kept whole under
// typing and deletion — and Update Fields refreshes its words from the
// bookmark on demand, the way Word's own F9 does.
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
    { text: 'Fourth paragraph text' },
  ],
}));

const pXml = (view, i) => view.doc.doc.paragraph(i).xml;

/** Bookmarks paragraph 1 and inserts a REF to it at the end of paragraph 2. */
const withReference = () => {
  const view = fixture();
  view.doc.doc.addBookmark('Summary', 1, 1);
  view.setSelection({ block: 2, offset: view.blocks[2].text.length });
  view.insertCrossReference('Summary');
  return view;
};

test('insertCrossReference writes the fldSimple exactly as Word does, with the bookmark\'s words, and the paragraph stays editable', () => {
  const view = withReference();

  assert.equal(view.blocks[2].text, 'Third paragraph textSecond paragraph text');
  assert.equal(view.blocks[2].structural, false, 'a field keeps the paragraph editable');

  assert.equal(
    pXml(view, 2),
    '<w:p><w:r><w:t xml:space="preserve">Third paragraph text</w:t></w:r>'
    + '<w:fldSimple w:instr=" REF Summary \\h "><w:r><w:t xml:space="preserve">Second paragraph text</w:t></w:r></w:fldSimple>'
    + '</w:p>',
  );
});

test('fields() lists every simple field in the body', () => {
  const view = withReference();
  assert.deepEqual(view.doc.doc.fields(), [
    { paragraph: 2, index: 1, instr: ' REF Summary \\h ', kind: 'ref', name: 'Summary', text: 'Second paragraph text' },
  ]);
});

test('the paragraph carrying a field stays editable: typing before or after the field leaves it whole', () => {
  const view = withReference();
  const fieldStart = 'Third paragraph text'.length;

  view.setSelection({ block: 2, offset: fieldStart });
  view.insertText('>> ');
  assert.equal(view.blocks[2].text, 'Third paragraph text>> Second paragraph text');
  assert.ok(
    view.blocks[2].runs.some((r) => r.field?.name === 'Summary' && r.text === 'Second paragraph text'),
    'the field survived, whole',
  );

  view.setSelection({ block: 2, offset: view.blocks[2].text.length });
  view.insertText(' <<');
  assert.equal(view.blocks[2].text, 'Third paragraph text>> Second paragraph text <<');
});

test('Backspace at the field\'s end removes the whole field in one press', () => {
  const view = withReference();
  view.setSelection({ block: 2, offset: view.blocks[2].text.length });
  view.deleteBackward();

  assert.equal(view.blocks[2].text, 'Third paragraph text');
  assert.equal(view.blocks[2].runs.some((r) => r.field), false, 'the field is gone, not shortened by one character');
});

test('a caret placed inside a field cannot type into it — the letter lands after the whole field', () => {
  const view = withReference();
  const fieldStart = 'Third paragraph text'.length;
  // Never reached by a click or an arrow key (see positions.js), but forced
  // here to prove the rule holds even so.
  view.setSelection({ block: 2, offset: fieldStart + 3 });
  view.insertText('Z');

  assert.equal(view.blocks[2].text, 'Third paragraph textSecond paragraph textZ');
  const field = view.blocks[2].runs.find((r) => r.field);
  assert.equal(field.text, 'Second paragraph text', 'the field itself is untouched');
});

test('updateFields refreshes a REF to its bookmark\'s current words, and reports how many changed', () => {
  const view = withReference();

  view.setSelection({ block: 1, offset: 0 });
  view.insertText('>> ');

  const n = view.updateFields();
  assert.equal(n, 1);
  assert.equal(view.blocks[2].runs.find((r) => r.field)?.text, '>> Second paragraph text');

  assert.equal(view.updateFields(), 0, 'nothing left to change');
});

test('updateFields writes Word\'s own error text once the bookmark is gone', () => {
  const view = withReference();
  view.removeBookmark('Summary');

  const n = view.updateFields();
  assert.equal(n, 1);
  assert.equal(view.blocks[2].runs.find((r) => r.field)?.text, 'Error! Reference source not found.');
});

test('a complex field (w:fldChar) stays structural — only a simple field became editable', () => {
  const view = fixture();
  const engine = view.doc.doc;
  const p = engine.paragraph(3);
  const complex = p.open + (p.pPr ?? '')
    + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> REF Summary \\h </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + '<w:r><w:t>Second paragraph text</w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    + '</w:p>';
  engine._spliceBody(p.start, p.end, complex);

  assert.equal(engine.paragraph(3).structural, true);
  assert.ok(engine.paragraph(3).structuralTags.includes('w:fldChar'));
});

test('a PAGE simple field written by hand reads as kind: "page", and round-trips untouched', () => {
  const view = fixture();
  const engine = view.doc.doc;
  const p = engine.paragraph(3);
  const withPage = p.open + (p.pPr ?? '')
    + '<w:r><w:t xml:space="preserve">Page </w:t></w:r>'
    + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>'
    + '</w:p>';
  engine._spliceBody(p.start, p.end, withPage);

  const field = engine.paragraph(3).runs.find((r) => r.field);
  assert.equal(field.field.kind, 'page');
  assert.equal(field.field.name, null);
  assert.equal(field.text, '1');
  assert.equal(engine.paragraph(3).structural, false);

  // Untouched — nothing edited this paragraph, so it never gets rebuilt.
  assert.equal(engine.paragraph(3).xml, withPage);
  const reopened = openDocx(view.save());
  assert.equal(reopened.doc.doc.paragraph(3).xml, withPage);
});

test('a cross-reference survives a save and reopen', () => {
  const view = withReference();
  const reopened = openDocx(view.save());

  const field = reopened.doc.doc.fields()[0];
  assert.equal(field.kind, 'ref');
  assert.equal(field.name, 'Summary');
  assert.equal(field.text, 'Second paragraph text');
  assert.equal(reopened.block(2).structural, false);
});

test('undo after insertCrossReference takes the field off again', () => {
  const view = withReference();
  assert.ok(view.blocks[2].runs.some((r) => r.field));

  view.undo();
  assert.equal(view.blocks[2].text, 'Third paragraph text');
  assert.equal(view.blocks[2].runs.some((r) => r.field), false);
});
