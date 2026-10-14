// Ink as Office keeps it: InkML written as Word and PowerPoint write it and
// read as they write it (explicit values, first and second differences),
// a document's strokes saved as w14 content parts and drawn again from
// their InkML when the file is opened again, a stroke Word drew included.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';
import { inkmlXml, readInkml, readTrace, inkPaths } from '@rutba/ooxml/inkml';

test('a trace is read with its packed differences undone', () => {
  // 10 20, then first differences +5 +1 (x 15 y 21), the same again by second difference 0 0 (20 22), then explicit 0 0.
  assert.deepEqual(readTrace("10 20, '5 '1, \"0 \"0, !0 !0"), [[10, 20], [15, 21], [20, 22], [0, 0]]);
  // Office packs signs against the value before them and keeps the mode until it changes.
  assert.deepEqual(readTrace("100 100,'10'-5,10-5,\"1\"1"), [[100, 100], [110, 95], [120, 90], [131, 86]]);
});

test('InkML written for a stroke reads back the same: points, colour, width, a highlighter\'s tip', () => {
  const xml = inkmlXml([
    { points: [[0, 0], [96, 48]], color: '#e71224', width: 2.25 },
    { points: [[10, 10], [200, 10]], color: '#FFFF00', width: 12, alpha: 0.5, cap: 'sq', highlighter: true },
  ]);
  assert.match(xml, /<inkml:ink xmlns:inkml="http:\/\/www\.w3\.org\/2003\/InkML">/);
  assert.match(xml, /<inkml:channelProperty channel="X" name="resolution" value="1000" units="1\/cm"\/>/);
  const ink = readInkml(xml);
  assert.equal(ink.strokes.length, 2);
  assert.deepEqual(ink.strokes[0].points, [[0, 0], [2540, 1270]], 'a pixel is 2540/96 hundredths of a millimetre');
  assert.equal(ink.strokes[0].color, '#E71224');
  assert.ok(Math.abs(ink.strokes[0].widthCm - (2.25 / 72) * 2.54) < 0.001);
  assert.equal(ink.strokes[1].highlighter, true);
  const paths = inkPaths(xml, 914400, 457200);
  assert.equal(paths.length, 2);
  assert.match(paths[0].path, /^<a:path w="914400" h="457200" fill="none" extrusionOk="0"><a:moveTo>/);
});

test('a document\'s ink is saved as Word saves it and drawn again from its InkML', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Sign here.' }] }));
  view.insertInk({ block: 0, strokes: [{ points: [[150, 300], [300, 340], [420, 320]], color: '#0000FF', width: 2, alpha: 1, cap: 'rnd' }] });
  const bytes = view.save();
  const pkg = OoxmlPackage.read(bytes);
  const xml = pkg.text('word/document.xml');
  assert.match(xml, /xmlns:wpi="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordprocessingInk"/);
  assert.match(xml, /<a:graphicData uri="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordprocessingInk"><w14:contentPart bwMode="auto" r:id="(rId\d+)"><w14:nvContentPartPr><w14:cNvContentPartPr\/><\/w14:nvContentPartPr><w14:xfrm>/);
  const rId = /<w14:contentPart bwMode="auto" r:id="(rId\d+)"/.exec(xml)[1];
  assert.match(pkg.text('word/_rels/document.xml.rels'), new RegExp(`Id="${rId}" Type="http://schemas.microsoft.com/office/2011/relationships/ink" Target="ink/ink1.xml"`));
  assert.match(pkg.text('[Content_Types].xml'), /PartName="\/word\/ink\/ink1\.xml" ContentType="application\/inkml\+xml"/);
  const ink = readInkml(pkg.text('word/ink/ink1.xml'));
  assert.deepEqual(ink.strokes[0].points[0], [Math.round(150 * 2540 / 96), Math.round(300 * 2540 / 96)], 'its points where they are on the page');
  // Opened again, the ink is drawn once, from the InkML, not again from the fallback.
  const again = openDocx(bytes);
  const shapes = again.render().blocks.flatMap((b) => b.images || []).filter((im) => /^Ink/.test(im.name || ''));
  assert.equal(shapes.length, 1);
  assert.match(Buffer.from(shapes[0].href.split(',')[1], 'base64').toString(), /<path d="M[^"]*Q/);
  assert.equal(again.doc.doc.drawings().filter((d) => d.name === 'Ink 1').length, 1);
});

test('a highlighter stroke is drawn in its own see-through yellow', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'x' }] }));
  view.insertInk({ block: 0, strokes: [{ points: [[100, 500], [500, 500]], color: '#FFFF00', width: 12, alpha: 0.5, cap: 'sq' }] });
  const svg = view.render().blocks.flatMap((b) => b.images || []).map((im) => Buffer.from(im.href.split(',')[1], 'base64').toString())[0];
  assert.match(svg, /stroke="#ffff00"/i, 'the colour, not its alpha, is the stroke');
});
