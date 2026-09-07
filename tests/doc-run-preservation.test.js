/**
 * A paragraph rebuild must not throw away what it cannot express.
 *
 * The defect this pins was quiet and expensive: the rebuilders reassembled an
 * edited paragraph from its TEXT runs, and a drawing run has no `<w:t>` — so
 * one keystroke in the paragraph carrying the company logo deleted the logo
 * from the saved file, with every test green. Now every rebuild carries the
 * kept fragments: whole non-text runs (a drawing, an explicit page-break
 * `w:br`), drawings lifted out of mixed runs, and any other non-run child.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

const DRAWING = '<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/></wp:inline></w:drawing>';

/** A letter whose first paragraph carries `extra` after its text run. */
function withExtra(extra) {
  const bytes = buildDocx({ paragraphs: ['Our letterhead', 'Body text'] });
  const pkg = OoxmlPackage.read(bytes);
  const xml = pkg.text('word/document.xml');
  const marker = 'Our letterhead</w:t></w:r>';
  pkg.write_('word/document.xml', xml.replace(marker, marker + extra));
  return openDocx(pkg.write());
}

const savedXml = (view) => OoxmlPackage.read(view.save()).text('word/document.xml');

test('typing into a paragraph with an inline image keeps the image', () => {
  const view = withExtra('<w:r>' + DRAWING + '</w:r>');
  view.collapseTo({ block: 0, offset: 3 });
  view.insertText('X');
  assert.equal(view.block(0).text, 'OurX letterhead', 'the keystroke landed');
  assert.ok(savedXml(view).includes('<w:drawing>'), 'and the logo is still in the file');
});

test('bolding, deleting and clearing formatting all keep the image too', () => {
  const view = withExtra('<w:r>' + DRAWING + '</w:r>');
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 3 });
  view.toggleFormat('b');
  view.clearFormat();
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 4 });
  view.deleteSelection();
  assert.equal(view.block(0).text, 'letterhead');
  assert.ok(savedXml(view).includes('<w:drawing>'));
});

test('Enter keeps the image with the first half of the split', () => {
  const view = withExtra('<w:r>' + DRAWING + '</w:r>');
  view.collapseTo({ block: 0, offset: 'Our'.length });
  view.splitParagraph();
  const xml = savedXml(view);
  const firstParagraphEnd = xml.indexOf('</w:p>');
  assert.ok(xml.slice(0, firstParagraphEnd).includes('<w:drawing>'),
    'the image stays where it was, not on the new line the caret moved to');
});

test('backspace-merge keeps the images of BOTH paragraphs', () => {
  const view = withExtra('<w:r>' + DRAWING + '</w:r>');
  // give the second paragraph an image as well
  const pkg = OoxmlPackage.read(view.save());
  const xml = pkg.text('word/document.xml');
  pkg.write_('word/document.xml', xml.replace('Body text</w:t></w:r>', 'Body text</w:t></w:r><w:r>' + DRAWING + '</w:r>'));
  const both = openDocx(pkg.write());

  both.collapseTo({ block: 1, offset: 0 });
  both.deleteBackward();
  assert.equal(both.blocks.length, 1, 'the paragraphs merged');
  const out = savedXml(both);
  assert.equal((out.match(/<w:drawing>/g) ?? []).length, 2, 'no image paid for the merge');
});

test('an explicit page-break run survives typing in its paragraph', () => {
  const view = withExtra('<w:r><w:br w:type="page"/></w:r>');
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('Re: ');
  assert.ok(savedXml(view).includes('<w:br w:type="page"/>'),
    'the author\'s break is not for a keystroke to delete');
});

test('a drawing nested INSIDE a text run is lifted out, not lost with the rewrite', () => {
  // Word sometimes puts text and a drawing in one run. The text half is being
  // rewritten; the picture half must not go with it.
  const bytes = buildDocx({ paragraphs: ['Signed,', 'Body'] });
  const pkg = OoxmlPackage.read(bytes);
  const xml = pkg.text('word/document.xml');
  pkg.write_('word/document.xml', xml.replace(
    '<w:t xml:space="preserve">Signed,</w:t>',
    '<w:t xml:space="preserve">Signed,</w:t>' + DRAWING,
  ));
  const view = openDocx(pkg.write());
  view.collapseTo({ block: 0, offset: 'Signed,'.length });
  view.insertText(' J. Smith');
  assert.equal(view.block(0).text, 'Signed, J. Smith');
  assert.ok(savedXml(view).includes('<w:drawing>'), 'the signature image survived the mixed run');
});

test('save, undo, save — the second save matches the document, not the first save', () => {
  // The bug this pins predates the feature that found it: a snapshot from
  // before an edit says dirty:false, and a save in between had already
  // written the edited xml into the package — so an undo-then-save emitted
  // the typing the undo had removed, silently.
  const view = openDocx(buildDocx({ paragraphs: ['Dear customer', ''] }));
  const pristine = savedXml(view);
  view.collapseTo({ block: 0, offset: 'Dear'.length });
  view.insertText(' valued');
  assert.ok(savedXml(view).includes('Dear valued customer'), 'the first save carries the typing');
  view.undo();
  assert.equal(savedXml(view), pristine, 'the save after undo carries the document, not the stale part');
});

test('a plain paragraph rebuilds exactly as before — no new residue', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Only text here', 'And here'] }));
  view.collapseTo({ block: 0, offset: 4 });
  view.insertText(' plain');
  const xml = savedXml(view);
  assert.ok(!xml.includes('<w:drawing'), 'nothing invented');
  assert.equal(view.block(0).text, 'Only plain text here');
});
