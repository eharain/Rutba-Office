/**
 * The same fidelity claim, for `.docx`.
 *
 * The document case is harder to get right than the spreadsheet case, because a
 * real template's intelligence lives in structures a text editor would flatten:
 * field codes that recompute, bookmarks, comment anchors, content controls, and
 * the section properties that carry the letterhead.
 *
 * Shaped around a bank balance-confirmation letter, because that is the kind of
 * document the fidelity requirement actually exists for.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { OoxmlPackage, Document, fidelityOf, comparePackages, preservedButUnsupported } from '@rutba/ooxml';
import { buildComplexDocument, FRAGILE_DOC_PARTS, FRAGILE_DOC_ELEMENTS } from './fixtures/complex-document.js';

const LETTER = buildComplexDocument();

test('the fixture is a readable word package', () => {
  const pkg = OoxmlPackage.read(LETTER);
  assert.equal(pkg.kind(), 'document');
  assert.equal(pkg.mainDocument(), 'word/document.xml');
  for (const part of FRAGILE_DOC_PARTS) assert.ok(pkg.has(part), 'fixture missing ' + part);
});

test('open and save with no edits is byte-identical', async () => {
  const result = await fidelityOf(LETTER);
  assert.equal(result.changed.length, 0);
  assert.equal(result.score, 1);
  assert.deepEqual(result.output, LETTER);
});

test('reading the document sees paragraphs, controls and bookmarks', () => {
  const doc = Document.open(LETTER);
  const text = doc.text();
  assert.match(text, /Confirmation of balance/);
  assert.match(text, /Dear Sir or Madam/);

  const controls = Object.fromEntries(doc.contentControls().map((c) => [c.tag, c]));
  assert.equal(controls.RUTBA_AR_TOTAL.alias, 'Outstanding balance');
  assert.equal(controls.RUTBA_AR_TOTAL.text, '0.00');
  assert.equal(controls.RUTBA_AS_OF.text, '1970-01-01');

  assert.deepEqual(doc.bookmarks().map((b) => b.name), ['Salutation']);
});

test('filling a content control rewrites exactly one part', async () => {
  const result = await fidelityOf(LETTER, (pkg) => {
    const doc = new Document(pkg);
    doc.fillContentControls({ RUTBA_AR_TOTAL: '412,000.00', RUTBA_AS_OF: '2026-08-20' });
    doc.save();
  });

  assert.deepEqual(result.intendedEdits, ['word/document.xml']);
  assert.deepEqual(result.unintendedChanges, []);
  assert.equal(result.removed.length, 0);
  for (const part of FRAGILE_DOC_PARTS) {
    assert.ok(result.identical.includes(part), part + ' was disturbed');
  }
});

test('everything structural in the letter survives the fill', () => {
  const doc = Document.open(LETTER);
  doc.fillContentControls({ RUTBA_AR_TOTAL: '412,000.00' });
  const after = OoxmlPackage.read(doc.save()).text('word/document.xml');

  for (const element of FRAGILE_DOC_ELEMENTS) {
    assert.ok(after.includes('<' + element), element + ' lost');
  }
  // the specific ones that would break the template or the letterhead
  assert.ok(after.includes('DOCPROPERTY'), 'field code lost — the reference number stops updating');
  assert.ok(after.includes('w:name="Salutation"'), 'bookmark lost');
  assert.ok(after.includes('<w:pgMar w:top="1440"'), 'page margins lost');
  assert.ok(after.includes('w:type="default" r:id="rIdHdr"'), 'header wiring lost');
  assert.ok(after.includes('<w:commentReference w:id="0"/>'), 'comment anchor lost');
  assert.ok(after.includes('w:numId w:val="1"'), 'list numbering lost');
});

test('a filled control keeps its own formatting and its properties', () => {
  const doc = Document.open(LETTER);
  doc.setContentControlText('RUTBA_AR_TOTAL', '412,000.00');
  const after = OoxmlPackage.read(doc.save()).text('word/document.xml');

  // the value was bold before; it must still be bold
  assert.match(after, /<w:sdtContent><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">412,000\.00<\/w:t><\/w:r><\/w:sdtContent>/);
  // and the control's identity is intact, so Word still recognises it
  assert.ok(after.includes('<w:tag w:val="RUTBA_AR_TOTAL"/>'));
  assert.ok(after.includes('<w:alias w:val="Outstanding balance"/>'));
  assert.ok(after.includes('<w:id w:val="123456789"/>'));

  assert.equal(Document.open(doc.save()).contentControls()[0].text, '412,000.00');
});

test('filling one control leaves the other alone', () => {
  const doc = Document.open(LETTER);
  doc.setContentControlText('RUTBA_AS_OF', '2026-08-20');
  const reread = Document.open(doc.save());
  const byTag = Object.fromEntries(reread.contentControls().map((c) => [c.tag, c.text]));
  assert.equal(byTag.RUTBA_AS_OF, '2026-08-20');
  assert.equal(byTag.RUTBA_AR_TOTAL, '0.00', 'the untouched control was modified');
});

test('an unknown tag is an error naming what the document does have', () => {
  const doc = Document.open(LETTER);
  assert.throws(
    () => doc.fillContentControls({ NOT_A_TAG: 'x' }),
    /no content control tagged: NOT_A_TAG; document has: RUTBA_AR_TOTAL, RUTBA_AS_OF/,
  );
});

test('editing plain paragraph text preserves its style and run formatting', () => {
  const doc = Document.open(LETTER);
  const paragraphs = doc.paragraphs();
  const heading = paragraphs.findIndex((p) => p.text === 'Confirmation of balance');
  assert.ok(heading >= 0);

  doc.setParagraphText(heading, 'Confirmation of balance — August 2026');
  const after = OoxmlPackage.read(doc.save()).text('word/document.xml');

  assert.ok(after.includes('<w:pStyle w:val="Heading1"/>'), 'paragraph style lost');
  assert.ok(after.includes('<w:spacing w:before="240" w:after="120"/>'), 'paragraph spacing lost');
  assert.match(after, /<w:rPr><w:b\/><w:sz w:val="32"\/><\/w:rPr><w:t xml:space="preserve">Confirmation of balance — August 2026<\/w:t>/);
});

test('a paragraph carrying structure refuses to be flattened', () => {
  const doc = Document.open(LETTER);
  const paragraphs = doc.paragraphs();

  const withField = paragraphs.findIndex((p) => p.xml.includes('w:fldSimple'));
  assert.throws(() => doc.setParagraphText(withField, 'Ref: something'), /carries w:fldSimple/);

  const withBookmark = paragraphs.findIndex((p) => p.xml.includes('w:bookmarkStart'));
  assert.throws(() => doc.setParagraphText(withBookmark, 'Dear customer,'), /carries w:bookmarkStart/);

  const withControl = paragraphs.findIndex((p) => p.xml.includes('<w:sdt'));
  assert.throws(() => doc.setParagraphText(withControl, 'flattened'), /carries w:sdt/);
});

test('paragraphs inside tables and controls are not addressed by top-level index', () => {
  const doc = Document.open(LETTER);
  const texts = doc.paragraphs().map((p) => p.text);
  assert.ok(!texts.includes('Description'), 'table cell paragraph leaked into the top-level list');
  assert.ok(texts.includes('Invoices raised in the period'), 'list paragraph missing');
});

test('ten fills accumulate no damage', () => {
  let buf = LETTER;
  for (let i = 0; i < 10; i++) {
    const doc = Document.open(buf);
    doc.fillContentControls({ RUTBA_AR_TOTAL: String(400000 + i), RUTBA_AS_OF: '2026-08-' + (10 + i) });
    buf = doc.save();
  }
  const diff = comparePackages(LETTER, buf);
  assert.deepEqual(diff.changed.map((c) => c.name), ['word/document.xml']);
  assert.equal(diff.removed.length, 0);

  const after = OoxmlPackage.read(buf).text('word/document.xml');
  for (const element of FRAGILE_DOC_ELEMENTS) assert.ok(after.includes('<' + element), element + ' eroded');
  const byTag = Object.fromEntries(Document.open(buf).contentControls().map((c) => [c.tag, c.text]));
  assert.equal(byTag.RUTBA_AR_TOTAL, '400009');
});

test('we report which document features we preserve but cannot render', () => {
  assert.deepEqual(preservedButUnsupported(LETTER), ['footers', 'footnotes', 'headers']);
});

test('the workbook view refuses a word package and vice versa', async () => {
  const { Workbook } = await import('@rutba/ooxml');
  assert.throws(() => Workbook.open(LETTER), /not a spreadsheet package/);
  const { buildComplexWorkbook } = await import('./fixtures/complex-workbook.js');
  assert.throws(() => Document.open(buildComplexWorkbook()), /not a word processing package/);
});
