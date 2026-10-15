// A turned picture, shape or text box keeps its angle through save and
// open in each of the three formats: `rot` on its own a:xfrm, in 60,000ths
// of a degree, read back as the same angle, with its box unmoved, so the
// frame the window draws on it stands where it stood.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { readZip } from '@rutba/ooxml/zip';

/** A saved file's parts, by name, as text. */
const partsOf = (bytes) => new Map(readZip(Buffer.from(bytes)).entries.map((e) => [e.name, e.data.toString('utf8')]));

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');

test('presentations: a shape, a picture and a group turned to 30, 90 and 135 keep their angles and boxes as a:xfrm rot', () => {
  const deck = Deck.open(buildPptx({ title: 'Turned', slides: [{ layout: 'blank' }] }));
  const box = deck.addShape(0, { preset: 'rect', x: 460, y: 300, w: 340, h: 90, name: 'Box', text: 'Turn me' });
  const added = deck.addPicture(0, { data: PNG, contentType: 'image/png', name: 'Pic', x: 120, y: 470, w: 160, h: 100 });
  const pic = typeof added === 'object' ? added.id : added;
  const a = deck.addShape(0, { preset: 'rect', x: 900, y: 100, w: 100, h: 50 });
  const b = deck.addShape(0, { preset: 'ellipse', x: 1020, y: 100, w: 100, h: 50 });
  const group = deck.groupShapes(0, [a, b]);
  deck.rotateShapes(0, [box], 135);
  deck.rotateShapes(0, [pic], 90);
  deck.rotateShapes(0, [group], 30);
  const before = Object.fromEntries(deck.slide(0).shapes.map((s) => [String(s.id), s.geometry]));
  const again = Deck.open(deck.save());
  const xml = again.pkg.text('ppt/slides/slide1.xml');
  for (const [id, deg] of [[box, 135], [pic, 90], [group, 30]]) {
    const g = again.slide(0).shapes.find((s) => String(s.id) === String(id)).geometry;
    assert.equal(g.rot, deg, `shape ${id} angle`);
    for (const k of ['x', 'y', 'w', 'h']) assert.ok(Math.abs(g[k] - before[String(id)][k]) <= 1, `shape ${id} ${k}`);
  }
  assert.match(xml, /<a:xfrm\b[^>]*\brot="8100000"/, '135 degrees as 60,000ths');
  assert.match(xml, /<a:xfrm\b[^>]*\brot="5400000"/, '90 degrees');
  assert.match(xml, /<a:xfrm\b[^>]*\brot="1800000"/, '30 degrees, on the group');
});

test('presentations: resizing a turned shape keeps its angle', () => {
  const deck = Deck.open(buildPptx({ title: 'Turned', slides: [{ layout: 'blank' }] }));
  const id = deck.addShape(0, { preset: 'rect', x: 100, y: 100, w: 200, h: 80 });
  deck.rotateShapes(0, [id], 30);
  deck.setGeometry(0, id, { x: 90, y: 95, w: 260, h: 110 });
  const g = Deck.open(deck.save()).slide(0).shapes.find((s) => String(s.id) === String(id)).geometry;
  assert.equal(g.rot, 30);
  assert.equal(Math.round(g.w), 260);
});

test('documents: a floating picture, a text box and a picture in line turned keep their angles through save and open', () => {
  const view = openDocx(buildDocx({ paragraphs: [{ text: 'Turned' }, { text: 'Words around the drawings.' }, { text: 'In line: ' }] }));
  view.setSelection({ block: 2, offset: 9 });
  view.insertImage({ name: 'Inline', contentType: 'image/png', data: PNG, widthPx: 160, heightPx: 60 });
  view.setSelection({ block: 1, offset: 0 });
  view.insertImage({ name: 'Banner', contentType: 'image/png', data: PNG, widthPx: 280, heightPx: 80 });
  const banner = view.drawings().find((d) => d.name === 'Banner').id;
  view.updateDrawings({ id: banner, wrap: 'none', behind: false, h: { rel: 'column', offsetPx: 170 }, v: { rel: 'paragraph', offsetPx: 60 } });
  view.setSelection({ block: 1, offset: 0 });
  view.insertTextBox({ name: 'Note', widthPx: 220, heightPx: 70, h: { rel: 'column', offsetPx: 20 }, v: { rel: 'paragraph', offsetPx: 40 }, paragraphs: [{ text: 'Turned words' }] });
  const byName = (v, n) => v.drawings().find((d) => d.name === n);
  for (const [name, rot] of [['Banner', 135], ['Note', 30], ['Inline', 90]]) view.updateDrawings({ id: byName(view, name).id, rot });
  const bytes = view.save();
  const again = openDocx(bytes);
  assert.equal(byName(again, 'Banner').rot, 135);
  assert.equal(byName(again, 'Note').rot, 30);
  assert.equal(byName(again, 'Inline').rot, 90);
  assert.equal(Math.round(byName(again, 'Banner').widthPx), 280, 'the box unmoved by the turn');
  const xml = partsOf(bytes).get('word/document.xml');
  assert.match(xml, /<a:xfrm\b[^>]*\brot="8100000"/);
  assert.match(xml, /\brot="1800000"/);
  assert.match(xml, /<a:xfrm\b[^>]*\brot="5400000"/);
});

test('worksheets: a shape and a picture turned keep their angles and boxes through save and open', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Plan', rows: [['Quarter', 'Sales'], ['Q1', 120]] }] }));
  view.select(5, 3);
  view.select(7, 7, { extend: true });
  view.insertShape({ geometry: 'rect', text: 'Turn me' });
  view.select(12, 2);
  view.insertPicture({ name: 'Logo', contentType: 'image/png', data: PNG, widthPx: 160, heightPx: 60 });
  const [shape, picture] = ['shape', 'image'].map((k) => view.render().drawings.find((d) => d.kind === k));
  view.rotateDrawings({ ids: [shape.id], to: 135 });
  view.rotateDrawings({ ids: [picture.id], to: 90 });
  const bytes = view.save();
  const again = new SheetView(bytes).render().drawings;
  const s = again.find((d) => d.kind === 'shape');
  const p = again.find((d) => d.kind === 'image');
  assert.equal(s.rot, 135);
  assert.equal(p.rot, 90);
  for (const k of ['x', 'y', 'width', 'height']) {
    assert.equal(s[k], shape[k], `the shape's ${k}`);
    assert.equal(p[k], picture[k], `the picture's ${k}`);
  }
  const xml = [...partsOf(bytes)].find(([n]) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))?.[1] || '';
  assert.match(xml, /<a:xfrm\b[^>]*\brot="8100000"/);
  assert.match(xml, /<a:xfrm\b[^>]*\brot="5400000"/);
});
