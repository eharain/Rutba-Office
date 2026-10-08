/**
 * Design → Design Ideas: layouts a slide's own content suits, worked out by
 * the deck — the picture beside the words with a crop that keeps its
 * proportions, the picture across the slide under a title band, the title on
 * an accent band, a side bar, the title centred, light words on the dark
 * colour — previewed as a changed scene and applied by the deck's methods.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { Deck, buildPptx, previewIdea, renderThumbnail } from '@rutba/presentation';

function png(w, h) {
  const raw = Buffer.alloc((w * 3 + 1) * h, 120);
  for (let y = 0; y < h; y++) raw[y * (w * 3 + 1)] = 0;
  const chunk = (t, d) => { const len = Buffer.alloc(4); len.writeUInt32BE(d.length); const td = Buffer.concat([Buffer.from(t), d]); const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(td) >>> 0); return Buffer.concat([len, td, crc]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}
const words = () => Deck.open(buildPptx({ slides: [{ title: 'Quarterly results', body: ['Sales rose 12%', 'Costs fell'] }] }));
const pictured = () => {
  const d = Deck.open(buildPptx({ slides: [{ title: 'Our new office', body: ['Opened in May'] }] }));
  d.addPicture(0, { data: png(40, 30), contentType: 'image/png', x: 700, y: 300, w: 400, h: 300 });
  return d;
};
const xml = (d) => d.pkg.text('ppt/slides/slide1.xml');

test('A slide of words is offered a title band, a side bar, the title centred and light words on the dark colour', () => {
  const d = words();
  assert.deepEqual(d.designIdeas(0).map((i) => i.id), ['title-band', 'side-bar', 'centred', 'deep-ground']);
  d.applyDesignIdea(0, 'title-band');
  const s = d.slide(0);
  const band = s.shapes.find((x) => x.name === 'Design idea');
  assert.ok(band, 'the band added');
  assert.equal(s.shapes[0], band, 'behind everything else');
  assert.equal(band.geometry.y, 0);
  assert.equal(band.geometry.w, s.size.width);
  assert.match(xml(d), /<p:sp>(?:(?!<\/p:sp>).)*name="Design idea"(?:(?!<\/p:sp>).)*<a:schemeClr val="accent1"\/>/s, 'in the theme\'s accent, so a new theme recolours it');
  const title = s.shapes.find((x) => x.placeholder?.type === 'title');
  assert.ok(title.text.paragraphs[0].runs.every((r) => /^#?FFFFFF$/i.test(r.color)), 'the title in white on the band');
  assert.ok(title.geometry.y < band.geometry.h);
});

test('A slide with a picture puts it beside the words, cropped to keep its proportions, or across the slide under a title band', () => {
  const d = pictured();
  assert.deepEqual(d.designIdeas(0).map((i) => i.id), ['picture-right', 'picture-left', 'picture-band']);
  d.applyDesignIdea(0, 'picture-right');
  const pic = d.slide(0).shapes.find((x) => x.kind === 'picture');
  assert.equal(Math.round(pic.geometry.x + pic.geometry.w), Math.round(d.slide(0).size.width), 'to the right edge');
  assert.equal(pic.geometry.y, 0);
  assert.equal(Math.round(pic.geometry.h), Math.round(d.slide(0).size.height));
  assert.match(xml(d), /<a:blip\b[^>]*\/><a:srcRect l="\d+" r="\d+"\/>/, 'the sides cropped equally, after the blip');
  assert.ok(Math.abs(pic.crop.l - pic.crop.r) < 1e-9 && pic.crop.l > 0.1);

  const e = pictured();
  e.applyDesignIdea(0, 'picture-band');
  const order = e.slide(0).shapes.map((x) => x.kind === 'picture' ? 'picture' : x.name === 'Design idea' ? 'band' : x.placeholder?.type || 'other');
  assert.deepEqual(order.slice(0, 2), ['picture', 'band'], 'the picture at the back, the band in front of it');
  assert.ok(order.indexOf('title') > 1, 'the title in front of both');
  assert.match(xml(e), /<a:srcRect t="\d+" b="\d+"\/>/, 'the picture cropped top and bottom to fill the slide');
});

test('A preview is the scene changed in memory, drawn by the renderer, the deck untouched', () => {
  const d = words();
  const before = xml(d);
  const [idea] = d.designIdeas(0);
  const scene = previewIdea(d.slide(0), idea);
  assert.equal(scene.shapes.length, d.slide(0).shapes.length + 1);
  assert.match(renderThumbnail(scene, 200), /^<svg/);
  assert.equal(xml(d), before);
  assert.throws(() => d.applyDesignIdea(0, 'no-such-idea'), /no longer fits/);
  assert.deepEqual(Deck.open(buildPptx({ slides: [{ layout: 'blank' }] })).designIdeas(0), [], 'nothing to arrange, no ideas');
});
