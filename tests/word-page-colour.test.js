// Page colour — the engine's half.
//
// Design → Page Colour writes `<w:background>` as the document's first child
// and `<w:displayBackgroundShape/>` in the settings, which is what makes Word
// show it; the section carries it back for the page and the PDF.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';

test('a page colour is written before the body with the setting that shows it, read back, cleared and undone', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }] }));
  assert.equal(view.section.background, null, 'a fresh document has none');

  view.setPageColour('#DEEBF7');
  assert.equal(view.section.background, '#DEEBF7', 'the section carries it');
  const xml = view.doc.doc.xml;
  const documentOpen = /<w:document\b[^>]*>/.exec(xml);
  assert.ok(xml.slice(documentOpen.index + documentOpen[0].length).startsWith('<w:background w:color="DEEBF7"/>'), 'the first child of w:document');

  const pkg = OoxmlPackage.read(view.save());
  assert.ok(pkg.has('word/settings.xml'), 'a settings part exists');
  assert.match(pkg.text('word/settings.xml'), /<w:displayBackgroundShape\/>/, 'and says to show the background');
  assert.match(pkg.text('word/_rels/document.xml.rels'), /relationships\/settings/, 'the main part points at it');
  assert.equal(openDocx(view.save()).section.background, '#DEEBF7', 'and it survives a save');

  view.setPageColour(null);
  assert.equal(view.section.background, null, 'cleared');
  assert.doesNotMatch(view.doc.doc.xml, /<w:background/, 'the element is gone');

  view.undo();
  assert.equal(view.section.background, '#DEEBF7', 'one undo step brings it back');
  assert.throws(() => view.setPageColour('blue'), /six hex digits/);
});
