// Watermark — the engine's half.
//
// Design → Watermark writes Word's own WordArt shape — a VML text path,
// rotated, in a fill colour — as the default header's first paragraph;
// the reader in headers.js gives it back as `bands.watermark`, apart from
// the header's words, for the window and the PDF to draw.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';

const bands = (view) => view.doc.doc.headerFooters();
const headerXml = (view) => { const b = bands(view).headers.default; return b ? view.doc.doc.pkg.text(b.part) : ''; };

test('a watermark on a document with no header makes one, is read back, kept through a save, replaced, removed and undone', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }] }));
  assert.ok(!bands(view).watermark, 'a fresh document has none');

  view.setWatermark('DRAFT');
  assert.deepEqual(bands(view).watermark, { text: 'DRAFT', colour: 'silver', rotation: 315, band: 'default' }, 'read back as the reader gives one');
  const xml = headerXml(view);
  assert.match(xml, /<w:hdr xmlns:w="[^"]*" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">/, 'the root declares VML');
  assert.match(xml, /<w:hdr[^>]*><w:p><w:r><w:rPr><w:noProof\/><\/w:rPr><w:pict><v:shapetype id="_x0000_t136"/, 'the shape type, first in the header');
  assert.match(xml, /<v:shape id="PowerPlusWaterMarkObject1"[^>]*type="#_x0000_t136"[^>]*rotation:315;[^>]*fillcolor="silver" stroked="f"><v:fill opacity=".5"\/><v:textpath style="[^"]*" string="DRAFT"\/><\/v:shape><\/w:pict><\/w:r><\/w:p>/, 'the shape as Word writes it');
  assert.equal(bands(view).headers.default.paragraphs.length, 1, 'the shape is not a line of the header');

  const pkg = OoxmlPackage.read(view.save());
  assert.match(pkg.text('word/document.xml'), /<w:headerReference w:type="default" r:id="[^"]+"/, 'the section points at the header');
  assert.match(pkg.text('[Content_Types].xml'), /header\+xml/);
  assert.equal(openDocx(view.save()).doc.doc.headerFooters().watermark.text, 'DRAFT', 'and it survives a save');

  view.setWatermark('CONFIDENTIAL', { colour: '#C00000', rotation: 0 });
  assert.deepEqual(bands(view).watermark, { text: 'CONFIDENTIAL', colour: '#c00000', rotation: 0, band: 'default' });
  assert.equal((headerXml(view).match(/<v:textpath\b[^>]*string=/g) || []).length, 1, 'one shape: replaced, not added');

  view.setWatermark('R&D');
  assert.match(headerXml(view), /string="R&amp;D"/, 'the words are escaped for the attribute');
  assert.equal(bands(view).watermark.text, 'R&D', 'and come back whole');

  view.setWatermark(null);
  assert.ok(!bands(view).watermark, 'removed');
  assert.doesNotMatch(headerXml(view), /<v:shape\b/, 'the shape is gone');
  assert.match(headerXml(view), /<w:p\/>/, 'the header keeps a paragraph');
  view.undo();
  assert.equal(bands(view).watermark?.text, 'R&D', 'one undo step brings it back');
});

test('a header with words keeps them under the watermark, and its words can change without losing the shape', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }] }));
  view.setBand('header', ['Acme Ltd']);
  view.setWatermark('SAMPLE');
  const xml = headerXml(view);
  assert.ok(xml.indexOf('<v:textpath') < xml.indexOf('Acme Ltd'), 'the shape comes first');
  assert.match(xml, /<w:hdr xmlns:w="[^"]*" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">/, 'the root gained the namespaces');
  const lines = bands(view).headers.default.paragraphs;
  assert.equal(lines.length, 1, 'one line of words');
  assert.ok(JSON.stringify(lines).includes('Acme Ltd'));

  view.setBand('header', ['Acme Ltd', 'Quarterly']);
  assert.equal(bands(view).watermark?.text, 'SAMPLE', 'the words changed and the shape stayed');
  assert.equal(bands(view).headers.default.paragraphs.length, 2);
  assert.ok(JSON.stringify(bands(view).headers.default.paragraphs).includes('Quarterly'));
  assert.equal((headerXml(view).match(/<v:textpath\b[^>]*string=/g) || []).length, 1, 'still one shape');
  assert.equal(openDocx(view.save()).doc.doc.headerFooters().watermark.text, 'SAMPLE', 'through a save');
});
