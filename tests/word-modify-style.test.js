/**
 * Modify Style, and Update Heading 1 to Match Selection (2.18.0, A3.4).
 *
 * In 2.17.1 a style could be picked but never changed: a person restyling
 * every heading did it one heading at a time. Now the style itself changes
 * in styles.xml, so every paragraph with it follows, the saved file carries
 * it, and one undo puts it back.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

const styleXml = (view, id) => {
  const xml = view.doc.doc.pkg.text('word/styles.xml');
  return new RegExp('<w:style\\b[^>]*w:styleId="' + id + '"[\\s\\S]*?</w:style>').exec(xml)?.[0] ?? null;
};

function headed() {
  const view = openDocx(buildDocx({ paragraphs: ['Opening jobs', 'Unlock the door', 'Closing jobs', 'Lock the door'], styles: true }));
  view.collapseTo({ block: 0, offset: 0 });
  view.setParagraphFormat({ styleId: 'Heading1' });
  view.collapseTo({ block: 2, offset: 0 });
  view.setParagraphFormat({ styleId: 'Heading1' });
  return view;
}

test('Modify Style changes the style itself: every heading follows, and styles.xml keeps it', () => {
  const view = headed();
  view.modifyStyle('Heading1', { fontName: 'Georgia', sizePt: 18, italic: true, bold: false, colour: '#AA3300', spaceBeforePt: 18, spaceAfterPt: 6 });
  const s = view.docStyles.Heading1;
  assert.equal(s.fontName, 'Georgia');
  assert.equal(s.sizePx, 24, '18 pt');
  assert.equal(s.italic, true);
  assert.equal(Boolean(s.bold), false);
  assert.equal(s.colour, '#aa3300');
  assert.equal(s.spaceBeforePx, 24);
  assert.equal(s.spaceAfterPx, 8);
  assert.equal(view.stylesChanged, true, 'the window takes the whole model');
  const xml = styleXml(view, 'Heading1');
  // In the schema's order: pPr before rPr, and rFonts, i, color, sz inside it.
  assert.match(xml, /<w:pPr><w:spacing w:before="360" w:after="120"\/><\/w:pPr><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia"\/><w:i\/><w:iCs\/><w:color w:val="AA3300"\/><w:sz w:val="36"\/><w:szCs w:val="36"\/><\/w:rPr>/);
  // Both headings point at the style, and neither carries the look itself.
  const body = view.doc.doc.xml;
  assert.equal((body.match(/<w:pStyle w:val="Heading1"\/>/g) || []).length, 2);
  assert.doesNotMatch(body, /Georgia/);
  // Saved, and read again.
  const again = openDocx(view.save());
  assert.equal(again.docStyles.Heading1.fontName, 'Georgia');
});

test('one undo puts the style back, and a redo takes it again', () => {
  const view = headed();
  const before = styleXml(view, 'Heading1');
  view.modifyStyle('Heading1', { sizePt: 20 });
  assert.ok(Math.abs(view.docStyles.Heading1.sizePx - 80 / 3) < 1e-9);
  view.undo();
  assert.equal(styleXml(view, 'Heading1'), before);
  assert.ok(Math.abs(view.docStyles.Heading1.sizePx - 64 / 3) < 1e-9, 'back to 16 pt');
  view.redo();
  assert.ok(Math.abs(view.docStyles.Heading1.sizePx - 80 / 3) < 1e-9);
});

test('Update Heading 1 to Match Selection takes the words\' own look into the style', () => {
  const view = headed();
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 12 });
  view.setRunFormat({ fontName: 'Verdana', fontColour: '2E7D32' });
  view.toggleFormat('i');
  view.setParagraphFormat({ spaceAfter: 3 });
  view.collapseTo({ block: 0, offset: 4 });
  view.updateStyleToMatch('Heading1');
  const s = view.docStyles.Heading1;
  assert.equal(s.italic, true);
  assert.equal(s.spaceAfterPx, 4, '3 pt after');
  assert.equal(s.bold, true, 'what the words take from the style stays');
  assert.ok(Math.abs(s.sizePx - 64 / 3) < 1e-9, 'and so does its size');
  assert.equal(s.fontName, 'Verdana');
  assert.equal(s.colour, '#2e7d32');
});

test('a style the file lacks is added from the standard set; a made-up one is refused', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Plain words'] }));
  view.modifyStyle('Heading2', { colour: '336699' });
  assert.equal(view.docStyles.Heading2.colour, '#336699');
  assert.equal(view.docStyles.Heading2.bold, true, 'the rest of Heading 2 as standard');
  assert.throws(() => view.modifyStyle('NoSuchStyle', { bold: true }), /no paragraph style/);
});

test('bold taken off a style based on a bold one is written as off', () => {
  const pkgView = openDocx(buildDocx({ paragraphs: ['Words'], styles: true }));
  const pkg = pkgView.doc.doc.pkg;
  const xml = pkg.text('word/styles.xml').replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="Sub"><w:name w:val="Sub"/><w:basedOn w:val="Heading1"/></w:style></w:styles>');
  pkg.write_('word/styles.xml', xml);
  pkgView._stylesWritten();
  pkgView.modifyStyle('Sub', { bold: false });
  assert.match(styleXml(pkgView, 'Sub'), /<w:b w:val="0"\/><w:bCs w:val="0"\/>/);
  assert.equal(Boolean(pkgView.docStyles.Sub.bold), false);
  assert.equal(OoxmlPackage.read(pkgView.save()).text('word/styles.xml').includes('w:styleId="Sub"'), true);
});
