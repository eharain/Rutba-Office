// Insert → Action: what a click on a shape does in the show, kept on the
// shape as PowerPoint keeps it — a jump, a slide of the deck by its
// relationship, or a web address — read back as the show wants it, and
// taken off again.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const deck = () => Deck.open(buildPptx({ slides: [1, 2, 3, 4].map((n) => ({ layout: 'title', title: `Slide ${n}`, body: 'Words' })) }));
const shapeOf = (d, i = 0) => d.slide(i).shapes.find((s) => s.placeholder?.type === 'title' || s.placeholder?.type === 'ctrTitle');
const xmlOf = (d, i = 0) => d.pkg.text(d.slideParts[i].part);

test('a jump is written first in the shape\'s cNvPr and read back by its name', () => {
  const d = deck();
  const id = shapeOf(d).id;
  assert.deepEqual(d.setAction(0, id, { kind: 'next' }), { kind: 'next' });
  assert.match(xmlOf(d), new RegExp(`<p:cNvPr id="${id}"[^>]*><a:hlinkClick r:id="" action="ppaction://hlinkshowjump\\?jump=nextslide"/></p:cNvPr>`));
  for (const kind of ['previous', 'first', 'last', 'end']) assert.equal(d.setAction(0, id, { kind }).kind, kind);
  assert.equal((xmlOf(d).match(/<a:hlinkClick\b/g) || []).length, 1, 'one action at a time');
});

test('a slide of the deck is kept as a relationship, and the action follows the slide when slides move', () => {
  const d = deck();
  const id = shapeOf(d).id;
  assert.deepEqual(d.setAction(0, id, { kind: 'slide', slide: 2 }), { kind: 'slide', slide: 2 });
  assert.match(d.pkg.text('ppt/slides/_rels/slide1.xml.rels'), /relationships\/slide" Target="slide3\.xml"/);
  d.moveSlide(2, 3);
  assert.deepEqual(shapeOf(d).action, { kind: 'slide', slide: 3 }, 'slide 3 is now fourth, and the action still goes to it');
  const back = Deck.open(d.save());
  assert.deepEqual(shapeOf(back).action, { kind: 'slide', slide: 3 });
});

test('a web address is an external relationship, and anything but http, https or mailto is refused', () => {
  const d = deck();
  const id = shapeOf(d).id;
  assert.deepEqual(d.setAction(0, id, { kind: 'url', url: 'https://rutba.io/office' }), { kind: 'url', url: 'https://rutba.io/office' });
  assert.match(xmlOf(d), /<a:hlinkClick r:id="rId\d+"\/>/);
  assert.throws(() => d.setAction(0, id, { kind: 'url', url: 'javascript:alert(1)' }), /web address/);
  assert.throws(() => d.setAction(0, id, { kind: 'url', url: 'file:///C:/Windows/System32/calc.exe' }), /web address/);
});

test('null takes the action off and leaves the shape as it was; an action the suite does not run is kept and named', () => {
  const d = deck();
  const id = shapeOf(d).id;
  const before = xmlOf(d);
  d.setAction(0, id, { kind: 'last' });
  assert.equal(d.setAction(0, id, null), null);
  assert.equal(xmlOf(d), before, 'the slide is back to the bytes it had');
  const part = d.slideParts[0].part;
  d.pkg.write_(part, Buffer.from(xmlOf(d).replace(new RegExp(`(<p:cNvPr id="${id}"[^>]*?)/>`), (m, open) => `${open}><a:hlinkClick r:id="" action="ppaction://macro?name=Run"/></p:cNvPr>`), 'utf8'));
  assert.deepEqual(shapeOf(d).action, { kind: 'other', action: 'ppaction://macro?name=Run' });
});

test('the document service: an action is one op, and the window\'s model carries it', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-action-')), 'actions.pptx');
  fs.writeFileSync(file, buildPptx({ slides: [1, 2, 3].map((n) => ({ layout: 'title', title: `Slide ${n}` })) }));
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = await docs.open({ path: file });
  const titleIn = (m) => m.slide.shapes.find((s) => s.placeholder?.type === 'title' || s.placeholder?.type === 'ctrTitle');
  const shape = titleIn(docs.model({ id, slide: 0 })).id;
  assert.equal(titleIn(docs.model({ id, slide: 0 })).action, null);
  docs.apply({ id, ops: [{ op: 'setAction', slide: 0, shape, action: { kind: 'url', url: 'https://rutba.io' } }] });
  assert.deepEqual(titleIn(docs.model({ id, slide: 0 })).action, { kind: 'url', url: 'https://rutba.io' });
  docs.undo({ id });
  assert.equal(titleIn(docs.model({ id, slide: 0 })).action, null, 'one undo takes it off');
});
