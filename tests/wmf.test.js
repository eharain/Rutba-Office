// A WMF drawn as an SVG: Word 97's and PowerPoint 97's drawings and clip art.
//
// The 16-bit metafile is read here record by record onto the same device
// an EMF is drawn on: its placeable header gives the picture's size, its
// window maps onto that size, its objects take the first free slot, and
// its words are 8-bit text in their font's character set.
import test from 'node:test';
import assert from 'node:assert/strict';
import { wmfToSvg, isWmf } from '../packages/office-formats/src/wmf.js';
import { placeableWmf } from '../packages/office-formats/src/msdoc-old.js';
import { WmfWriter, showcaseWmf, dib24, rgb } from './fixtures/metafiles.js';
import { decodePng } from '../packages/ocr/src/png.js';

function elements(svg, tag) {
  const out = [];
  const re = new RegExp(`<${tag}\\b([^>]*?)(/?)>`, 'g');
  let m;
  while ((m = re.exec(svg))) {
    const attrs = {};
    for (const a of m[1].matchAll(/([\w:-]+)="([^"]*)"/g)) attrs[a[1]] = a[2];
    if (!m[2] && tag === 'text') attrs.$text = svg.slice(re.lastIndex, svg.indexOf('</text>', re.lastIndex));
    out.push(attrs);
  }
  return out;
}

test('a Windows 3.1 drawing in twips: its box, shapes, hatch, hole, arc and words', () => {
  const bytes = showcaseWmf();
  assert.ok(isWmf(bytes));
  const out = wmfToSvg(bytes);
  assert.ok(out);
  assert.deepEqual([out.width, out.height], [288, 192], 'three inches by two, at 96 pixels an inch');
  assert.match(out.svg, /viewBox="0 0 4320 2880"/);
  assert.doesNotMatch(out.svg, /<script|\son\w+=/);
  const paths = elements(out.svg, 'path');
  const box = paths.find((p) => p.fill === '#dae3f3');
  assert.match(box.d, /^M307\.5 127\.5L4012\.5 127\.5C/, 'a rounded rectangle, its corners a quarter inch round');
  assert.equal(box.stroke, '#1f3864');
  assert.equal(box['stroke-width'], '15', 'fifteen twips: a pixel');
  assert.equal(box['stroke-linecap'], undefined, 'a pen a pixel wide ends square, on the pixels it lights');
  const sun = paths.find((p) => p.fill === '#ffc000');
  assert.match(sun.d, /^M1552\.5 1080C/, 'an ellipse from its right-hand point');
  const arc = paths.find((p) => p.stroke === '#c00000');
  assert.equal(arc.fill, 'none');
  assert.equal(arc['stroke-width'], '45');
  assert.match(arc.d, /^M3000 1080C.* 1800 1080$/, 'from three o\'clock round the top to nine');
  assert.match(arc.d, /C[^C]* 2400 480C/, 'through the top of the circle');
  const frame = paths.find((p) => /^url\(#h\d+\)$/.test(p.fill));
  assert.equal(frame['fill-rule'], 'evenodd', 'alternate filling leaves the hole');
  assert.equal((frame.d.match(/M/g) || []).length, 2, 'two parts');
  assert.match(out.svg, /<pattern id="h\d+" patternUnits="userSpaceOnUse" width="120" height="120"><g transform="scale\(15\)"><path d="M-1 9L9 -1"/, 'the hatch eight pixels a side, whatever the units');
  const texts = elements(out.svg, 'text');
  const title = texts.find((t) => t.$text === 'Rutba drawing');
  assert.match(title['font-family'], /^'Times New Roman', serif$/);
  assert.equal(title['font-size'], '300');
  assert.equal(title['text-anchor'], 'middle', 'centred on its point, no advances given');
  assert.equal(title.transform, 'translate(2160 2160)');
  const line = texts.find((t) => t.$text === 'a WMF made at test time');
  assert.equal(line['font-style'], 'italic');
  assert.equal(line.transform, 'translate(360 2340)');
});

test('objects take the first free slot, as Windows numbers them', () => {
  const w = new WmfWriter(1440, 1440);
  w.windowOrg(0, 0).windowExt(1440, 1440);
  w.brush(0, rgb('#ff0000')); // slot 0
  w.brush(0, rgb('#00ff00')); // slot 1
  w.remove(0);
  w.brush(0, rgb('#0000ff')); // slot 0 again
  w.pen(5, 0, 0); // slot 2
  w.select(2).select(0).rectangle(0, 0, 100, 100).select(1).rectangle(200, 0, 300, 100);
  const fills = elements(wmfToSvg(w.build()).svg, 'path').map((p) => p.fill);
  assert.deepEqual(fills, ['#0000ff', '#00ff00']);
});

test('without a placeable header the window the records set is the picture', () => {
  const w = new WmfWriter(0, 0);
  w.windowOrg(100, 100).windowExt(1000, 500).brush(0, rgb('#336699')).select(0).pen(5, 0, 0).select(1).rectangle(100, 100, 600, 350);
  const out = wmfToSvg(w.build({ placeable: false }));
  assert.ok(out);
  assert.match(out.svg, /viewBox="0 0 1000 500"/);
  assert.equal(elements(out.svg, 'path')[0].d, 'M0 0L500 0L500 250L0 250Z');
});

test('a window measured upwards is drawn the right way up', () => {
  const w = new WmfWriter(1000, 1000, 1000);
  w.windowOrg(0, 1000).windowExt(1000, -1000).brush(0, rgb('#000000')).select(0).pen(5, 0, 0).select(1).rectangle(0, 0, 500, 250);
  const d = elements(wmfToSvg(w.build()).svg, 'path')[0].d;
  assert.equal(d, 'M0 1000L500 1000L500 750L0 750Z', 'y = 0 is the bottom');
});

test('a bitmap stretched into the picture, and the old Word way of wrapping a metafile', () => {
  const w = new WmfWriter(200, 100, 96);
  w.windowOrg(0, 0).windowExt(200, 100).stretchDib(10, 10, 80, 40, dib24(4, 2, (x) => (x < 2 ? 0xff0000 : 0x00ff00)));
  const raw = w.build({ placeable: false });
  // Word 6 kept a metafile without a placeable header; the reader gives it one from the picture's size in twips.
  const wrapped = placeableWmf(raw, 3000, 1500);
  assert.equal(wrapped.readUInt32LE?.(0) ?? Buffer.from(wrapped).readUInt32LE(0), 0x9ac6cdd7);
  for (const bytes of [w.build(), Buffer.from(wrapped)]) {
    const images = elements(wmfToSvg(bytes).svg, 'image');
    assert.equal(images.length, 1);
    const img = decodePng(Buffer.from(images[0].href.split(',')[1], 'base64'));
    assert.deepEqual([img.width, img.height], [4, 2]);
    assert.deepEqual([...img.data.subarray(0, 4)], [255, 0, 0, 255]);
  }
});

test('8-bit words in their font\'s character set, and a symbol font\'s letters where Windows maps them', () => {
  const w = new WmfWriter(2000, 500);
  w.windowOrg(0, 0).windowExt(2000, 500).bkMode(1);
  w.font({ height: -200, face: 'Arial', charset: 204 }).select(0).textOut(0, 0, '\xcf\xf0\xe8');
  w.font({ height: -200, face: 'Symbol', charset: 2 }).select(1).textOut(0, 250, 'ab');
  const texts = elements(wmfToSvg(w.build()).svg, 'text');
  assert.equal(texts[0].$text, 'При', 'Cyrillic, from Windows-1251');
  assert.equal(texts[1].$text, '');
});

test('a damaged WMF stops where it breaks, and nothing throws', () => {
  const good = showcaseWmf();
  let seed = 11;
  const rand = (n) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
  for (let i = 0; i < 300; i++) {
    const b = Buffer.from(good);
    for (let k = 0; k < 6; k++) b[40 + rand(b.length - 40)] = rand(256);
    if (i % 4 === 0) b.writeUInt16LE(0xffff, 40 + 2 * rand(Math.floor((b.length - 42) / 2)));
    const out = wmfToSvg(i % 5 === 0 ? b.subarray(0, 40 + rand(b.length - 40)) : b);
    if (out) assert.doesNotMatch(out.svg, /<script|\son\w+=/);
  }
  assert.equal(wmfToSvg(Buffer.alloc(30)), null);
  assert.equal(isWmf(Buffer.from('plain words, no metafile here at all')), false);
  assert.equal(wmfToSvg(new WmfWriter(100, 100).build()), null, 'nothing drawn');
});
