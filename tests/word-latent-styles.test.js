/**
 * Word's built-in styles in a file that has not used them (2.18.0, A3.3).
 *
 * A document whose styles part holds only Normal offered only Normal: no
 * Title, no Heading 1 to 3, and References > Add Text's levels failed with
 * "no paragraph style". Word offers its latent built-in styles there and
 * writes one in the first time it is applied; so does this now, one style
 * element added and the file's own styles left as they were.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

const ONLY_NORMAL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:rPr><w:i/></w:rPr></w:style>'
  + '</w:styles>';

function onlyNormal() {
  const view = openDocx(buildDocx({ paragraphs: ['Opening jobs', 'Unlock the door'], styles: true }));
  view.doc.doc.pkg.write_('word/styles.xml', ONLY_NORMAL);
  view._stylesWritten();
  return view;
}

test('a styles part with only Normal offers Title and Heading 1 to 3 after its own styles', () => {
  const view = onlyNormal();
  const ids = view.paragraphStyles.map((s) => s.id);
  assert.deepEqual(ids, ['Normal', 'Quote', 'Title', 'Heading1', 'Heading2', 'Heading3']);
  assert.equal(view.paragraphStyles.find((s) => s.id === 'Heading1').latent, true);
  assert.equal(view.paragraphStyles.find((s) => s.id === 'Quote').latent, undefined, 'the file\'s own are not latent');
});

test('Add Text > Level 1 writes Heading 1 in and applies it; the file\'s own styles are kept as they were', () => {
  const view = onlyNormal();
  view.collapseTo({ block: 0, offset: 0 });
  view.setParagraphFormat({ styleId: 'Heading1' });
  const xml = view.doc.doc.pkg.text('word/styles.xml');
  assert.match(xml, /<w:style w:type="paragraph" w:styleId="Heading1">/);
  assert.ok(xml.includes('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>'), 'Normal untouched');
  assert.ok(xml.includes('<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:rPr><w:i/></w:rPr></w:style>'), 'Quote untouched');
  assert.equal(view.docStyles.Heading1.bold, true, 'the heading looks like one');
  // Heading 1 is the file's own now, offered once.
  assert.equal(view.paragraphStyles.filter((s) => s.id === 'Heading1').length, 1);
  assert.equal(view.paragraphStyles.find((s) => s.id === 'Heading1').latent, undefined);
  // Saved and opened again: the paragraph is a heading.
  const again = openDocx(view.save());
  assert.equal(again.doc.getParagraphProps(0).style, 'Heading1');
  // A second heading adds no second definition.
  view.collapseTo({ block: 1, offset: 0 });
  view.setParagraphFormat({ styleId: 'Heading1' });
  assert.equal((view.doc.doc.pkg.text('word/styles.xml').match(/w:styleId="Heading1"/g) || []).length, 1);
});

test('a style the file has under another id, by its name, is not offered twice; an unknown id is still refused', () => {
  const view = onlyNormal();
  const xml = view.doc.doc.pkg.text('word/styles.xml').replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="berschrift1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style></w:styles>');
  view.doc.doc.pkg.write_('word/styles.xml', xml);
  view._stylesWritten();
  const ids = view.paragraphStyles.map((s) => s.id);
  assert.ok(ids.includes('berschrift1'));
  assert.ok(!ids.includes('Heading1'), 'its own heading 1 stands for Heading 1');
  view.collapseTo({ block: 0, offset: 0 });
  assert.throws(() => view.doc.setParagraphProp(0, 'style', 'NoSuchStyle'), /no paragraph style/);
});
