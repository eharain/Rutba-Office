// Draw in a document — the engine's half.
//
// A stroke floats in front of the words where it was drawn on its page:
// Word's wps freeform in mc:AlternateContent, positioned from the page's
// corner, no wrap, named "Ink N", boxed to hold its width, anchored in a
// paragraph on that page; the page draws it; Ink to Shape's shape floats
// the same way; a stroke is removed like any drawing.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const paper = () => openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'First.' }, { text: 'Second.' }, { text: 'Third.' }] }));

test('a stroke floats where it was drawn, as Word writes a freeform in front of the words', () => {
  const view = paper();
  view.insertInk({ block: 1, strokes: [
    { points: [[200, 300], [240, 320], [300, 330]], color: '#0078D4', width: 1, alpha: 1, cap: 'rnd' },
    { points: [[100, 500], [500, 500]], color: '#FFFF00', width: 8, alpha: 0.5, cap: 'sq' },
  ] });
  const xml = view.doc.doc.xml;
  assert.match(xml, /<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor\b[^>]*behindDoc="0"/);
  assert.match(xml, /<wp:positionH relativeFrom="page"><wp:posOffset>\d+<\/wp:posOffset><\/wp:positionH><wp:positionV relativeFrom="page">/);
  assert.match(xml, /<wp:wrapNone\/><wp:docPr id="\d+" name="Ink 1"\/>/);
  assert.match(xml, /<wps:wsp><wps:cNvSpPr\/><wps:spPr>[\s\S]*?<a:custGeom>[\s\S]*?<a:quadBezTo>/);
  assert.match(xml, /<a:ln w="\d+" cap="sq"><a:solidFill><a:srgbClr val="FFFF00"><a:alpha val="50000"\/><\/a:srgbClr><\/a:solidFill><a:round\/><\/a:ln>/);
  // Anchored in the paragraph named, before its words.
  const p = view.doc.doc.editParagraph(1);
  assert.ok(p.xml.indexOf('name="Ink 1"') < p.xml.indexOf('Second.'));
  assert.equal(view.blocks[1].text, 'Second.', 'ink adds no words');
  // The page draws it.
  const ink = view.doc.doc.drawings().filter((d) => /^Ink \d+$/.test(d.name));
  assert.equal(ink.length, 2);
  const frame = view.render();
  const images = (frame.blocks || []).flatMap((b) => b.images || []).filter((im) => /^Ink/.test(im.name || ''));
  assert.equal(images.length, 2);
  assert.match(Buffer.from(images[0].href.split(',')[1], 'base64').toString(), /<path d="M[^"]*Q/);
  // And it goes like any drawing.
  view.removeDrawing(ink[0].id);
  assert.equal(view.doc.doc.drawings().filter((d) => /^Ink/.test(d.name)).length, 1);
});

test('Ink to Shape floats the shape where it was drawn, outlined in the pen\'s colour, and the file keeps both', () => {
  const view = paper();
  view.insertFloatingShape({ block: 0, preset: 'ellipse', x: 150, y: 200, width: 220, height: 120, colour: '#E81123', widthPt: 2 });
  view.insertInk({ block: 0, strokes: [{ points: [[50, 60], [90, 80]], color: '#000000', width: 1 }] });
  const back = openDocx(view.save());
  const names = back.doc.doc.drawings().map((d) => d.name);
  assert.ok(names.some((n) => /^Shape \d+$/.test(n)) && names.includes('Ink 1'), JSON.stringify(names));
  assert.match(back.doc.doc.xml, /<a:prstGeom prst="ellipse"><a:avLst\/><\/a:prstGeom><a:noFill\/><a:ln w="\d+"><a:solidFill><a:srgbClr val="E81123"\/>/);
});
