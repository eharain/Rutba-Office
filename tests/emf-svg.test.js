// An EMF drawn as an SVG: what Windows draws, where it draws it.
//
// A chart pasted from another program into Word or PowerPoint is usually
// kept as an EMF of lines, shapes and words, and Documents, Worksheets and
// Presentations showed its empty frame. These metafiles are written here
// as Windows' own metafile device writes them, and each test reads back
// what the SVG puts where: words at their advances and alignment, pens at
// their widths, fills, arcs turning the right way, the map modes, the world
// transform, clips, bitmaps, and a damaged file stopping where it breaks.
import test from 'node:test';
import assert from 'node:assert/strict';
import { emfToSvg, emfToPng, isEmf } from '../packages/office-formats/src/emf.js';
import { EmfWriter, showcaseEmf, dib24, rgb, advances } from './fixtures/metafiles.js';
import { decodePng } from '../packages/ocr/src/png.js';

/** Every element of a kind, with its attributes. */
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

/** Nothing in it can run, and nothing in it reaches outside it. */
function assertInert(svg) {
  assert.doesNotMatch(svg, /<script|<foreignObject|\son\w+=|javascript:/i, 'nothing scriptable');
  for (const m of svg.matchAll(/(?:href|xlink:href)="([^"]*)"/g)) assert.match(m[1], /^data:image\/(png|jpeg);base64,/, 'every picture is inside the SVG');
  for (const m of svg.matchAll(/url\(([^)]*)\)/g)) assert.match(m[1], /^#[a-z]\d+$/, 'every url() points inside');
}

test('the chart a program pasted is drawn: its words, bars, line, pie, legend, turned title and logo', () => {
  const bytes = showcaseEmf();
  assert.ok(isEmf(bytes));
  const out = emfToSvg(bytes);
  assert.ok(out, 'drawn');
  assert.deepEqual([out.width, out.height], [480, 300], 'its size from the frame, at 96 pixels an inch');
  assert.match(out.svg, /^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg" viewBox="0 0 480 300(\.\d+)?"/);
  assertInert(out.svg);

  const texts = elements(out.svg, 'text');
  const title = texts.find((t) => t.$text === 'Quarterly sales');
  assert.ok(title, 'the title is words, not a picture of them');
  assert.equal(title['font-weight'], '700');
  assert.equal(title['font-size'], '20');
  assert.match(title['font-family'], /^'Arial', sans-serif$/);
  assert.equal(title.fill, '#1f3864');
  // Centred on x = 240: the advances add up to the title's width, half of it to each side.
  const xs = title.x.split(' ').map(Number);
  const dx = advances('Quarterly sales', 20, { bold: true });
  const width = dx.reduce((a, b) => a + b, 0);
  assert.equal(xs.length, 'Quarterly sales'.length, 'one position a letter, as ExtTextOut gave them');
  assert.equal(xs[0], -width / 2);
  assert.equal(xs[1] - xs[0], dx[0]);
  assert.equal(title.transform, 'translate(240 10)');
  assert.equal(Number(title.y), 18.1, 'aligned by its top: the baseline an ascent below');

  const label = texts.find((t) => t.$text === '300');
  assert.equal(Number(label.y), 0, 'aligned by its baseline');
  assert.equal(Number(label.x.split(' ').pop()) + advances('300', 12).pop(), 0, 'right-aligned: the last letter ends at the point');

  const turned = texts.find((t) => t.$text === 'Sales (k)');
  assert.match(turned.transform, /rotate\(-90\)$/, 'escapement 900 reads upwards');
  assert.equal(texts.find((t) => t.$text === 'North')['font-style'], 'italic');

  const paths = elements(out.svg, 'path');
  const bar = paths.find((p) => p.fill === '#4472c4' && /^M80\.5 140\.5/.test(p.d));
  assert.ok(bar, 'the solid bar, its one-pixel outline on its edge pixels');
  assert.equal(bar.stroke, '#2f5597');
  const line = paths.find((p) => p.stroke === '#c00000' && p['stroke-width'] === '3');
  assert.ok(line, 'the series is three pixels wide');
  assert.equal(line['stroke-linecap'], 'round');
  assert.equal(line.fill, 'none');
  const grid = paths.find((p) => p.stroke === '#bfbfbf');
  assert.equal(grid['stroke-dasharray'], '3 3', 'a dotted hairline');
  assert.equal(grid['vector-effect'], 'non-scaling-stroke', 'one pixel at any size');
  assert.equal(grid.d, 'M60.5 80.5L300.5 80.5', 'on the pixels it lights');
  assert.ok(paths.some((p) => /^url\(#g\d+\)$/.test(p.fill)), 'the gradient bar');
  assert.ok(paths.some((p) => /^url\(#h\d+\)$/.test(p.fill)), 'the hatched bar');
  assert.match(out.svg, /<linearGradient[^>]*x1="150" y1="100" x2="150" y2="260"><stop offset="0" stop-color="#ed7d31"\/><stop offset="1" stop-color="#ffc000"\/>/, 'shaded downwards');
  const slices = paths.filter((p) => ['#4472c4', '#ed7d31', '#a5a5a5'].includes(p.fill) && /^M390 120L/.test(p.d));
  assert.equal(slices.length, 3, 'three slices from the centre');
  // The first slice runs anticlockwise from three o'clock to twelve: its end is at the top.
  assert.match(slices[0].d, /C[^C]* 390 60Z$/);
  const legend = paths.find((p) => p.fill === '#f2f2f2');
  assert.match(legend.d, /C/, 'a rounded box');
  const images = elements(out.svg, 'image');
  assert.equal(images.length, 1);
  assert.deepEqual([images[0].x, images[0].y, images[0].width, images[0].height], ['446', '6', '24', '24']);
  const logo = decodePng(Buffer.from(images[0].href.split(',')[1], 'base64'));
  assert.deepEqual([logo.width, logo.height], [16, 16], 'the bitmap at its own size, stretched by the image');
});

test('map modes: an anisotropic window onto a viewport, and hundredths of a millimetre measured upwards', () => {
  const e = new EmfWriter(200, 100);
  e.mapMode(8).windowOrg(0, 0).windowExt(1000, 1000).viewportOrg(0, 0).viewportExt(200, 100);
  e.pen(1, 5, 0, 0).select(1).brush(2, 0, rgb('#ff0000')).select(2).rectangle(500, 500, 1000, 1000);
  e.mapMode(3).windowOrg(0, 0).viewportOrg(0, 100);
  e.brush(3, 0, rgb('#0000ff')).select(3).rectangle(0, 1000, 1000, 0);
  const svg = emfToSvg(e.build()).svg;
  const paths = elements(svg, 'path');
  assert.equal(paths.find((p) => p.fill === '#ff0000').d, 'M100 50L200 50L200 100L100 100Z');
  // 10 mm across and up from the bottom-left, at the reference device's 3.7795 pixels a millimetre.
  const blue = paths.find((p) => p.fill === '#0000ff').d;
  assert.equal(blue, 'M0 100L37.8 100L37.8 62.2L0 62.2Z');
});

test('a world transform turns what follows, and setting it back to identity stops', () => {
  const e = new EmfWriter(100, 100);
  e.pen(1, 5, 0, 0).select(1).brush(2, 0, rgb('#00ff00')).select(2);
  const c = Math.cos(Math.PI / 2);
  e.world([c, 1, -1, c, 50, 0]).rectangle(0, 0, 20, 10);
  e.modifyWorld([1, 0, 0, 1, 0, 0], 1).rectangle(0, 0, 20, 10);
  const paths = elements(emfToSvg(e.build()).svg, 'path').filter((p) => p.fill === '#00ff00');
  assert.equal(paths[0].d, 'M50 0L50 20L40 20L40 0Z', 'a quarter turn about the origin, then moved');
  assert.equal(paths[1].d, 'M0 0L20 0L20 10L0 10Z');
});

test('a clip holds what is drawn inside it, and restoring the saved state lets it go', () => {
  const e = new EmfWriter(100, 100);
  e.pen(1, 5, 0, 0).select(1).brush(2, 0, rgb('#123456')).select(2);
  e.save().intersectClip(10, 10, 50, 50).intersectClip(20, 0, 100, 40).ellipse(0, 0, 100, 100).restore(-1).rectangle(0, 0, 5, 5);
  const svg = emfToSvg(e.build()).svg;
  assert.match(svg, /<clipPath id="c\d+" clipPathUnits="userSpaceOnUse"><path d="M20 10H50V40H20Z"\/><\/clipPath>/, 'two rectangles meet as one');
  assert.match(svg, /<g clip-path="url\(#c\d+\)"><path d="M100 50C/, 'the ellipse inside it');
  assert.match(svg, /<\/g><path d="M0 0L5 0L5 5L0 5Z"/, 'the rectangle after the restore is outside it');
});

test('paths: a figure built, closed and filled, and a pen\'s dashes and joins', () => {
  const e = new EmfWriter(100, 100);
  e.extPen(1, 0x2000 | 0x200 | 1, 4, rgb('#000080')).select(1).brush(2, 0, rgb('#ffff00')).select(2);
  e.beginPath().moveTo(10, 10).lineTo(90, 10).lineTo(50, 80).closeFigure().endPath().strokeAndFillPath();
  const p = elements(emfToSvg(e.build()).svg, 'path')[0];
  assert.equal(p.d, 'M10 10L90 10L50 80Z');
  assert.equal(p.fill, '#ffff00');
  assert.equal(p['stroke-width'], '4');
  assert.equal(p['stroke-dasharray'], '12 4', 'a geometric dash is three widths on, one off');
  assert.equal(p['stroke-linejoin'], undefined, 'mitred joins are the SVG default');
  assert.equal(p['stroke-linecap'], undefined, 'flat ends');
});

test('a mask and its picture are one picture with a see-through background', () => {
  const inside = (x, y) => x >= 1 && x <= 2 && y >= 1 && y <= 2;
  const mask = dib24(4, 4, (x, y) => (inside(x, y) ? 0 : 0xffffff));
  const picture = dib24(4, 4, (x, y) => (inside(x, y) ? 0x208040 : 0));
  const e = new EmfWriter(4, 4);
  e.stretchDib(0, 0, 4, 4, mask, 0x8800c6).stretchDib(0, 0, 4, 4, picture, 0xee0086);
  const images = elements(emfToSvg(e.build()).svg, 'image');
  assert.equal(images.length, 1, 'played together');
  const img = decodePng(Buffer.from(images[0].href.split(',')[1], 'base64'));
  const at = (x, y) => [...img.data.subarray((y * img.width + x) * 4, (y * img.width + x) * 4 + 4)];
  assert.deepEqual(at(1, 1), [0x20, 0x80, 0x40, 255]);
  assert.equal(at(0, 0)[3], 0);
  // The PNG path still reads the same file.
  assert.ok(emfToPng(e.build()));
});

test('words are kept as words, their letters escaped and a font name kept to a font name', () => {
  const e = new EmfWriter(200, 50);
  e.font(1, { height: -14, face: 'Evil"><script>x' }).select(1).bkMode(1);
  e.text(5, 5, 'a < b & "c"\u0001');
  const svg = emfToSvg(e.build()).svg;
  assertInert(svg);
  const t = elements(svg, 'text')[0];
  assert.equal(t.$text, 'a &lt; b &amp; &quot;c&quot; ');
  assert.match(t['font-family'], /^'Evilscriptx', sans-serif$/);
  assert.equal(t.x, '0', 'no advances given: the font\'s own');
});

test('words on an opaque background sit on a box of the background colour', () => {
  const e = new EmfWriter(200, 50);
  e.font(1, { height: -10 }).select(1).bkColor(rgb('#ffff00'));
  e.text(10, 10, 'Hi', { dx: [7, 3] });
  const svg = emfToSvg(e.build()).svg;
  assert.match(svg, /<rect transform="translate\(10 10\)" x="0" y="0" width="10" height="11\.17" fill="#ffff00"\/><text/);
});

test('a damaged metafile stops where it breaks: what came before is drawn, and nothing throws', () => {
  const good = showcaseEmf();
  const full = emfToSvg(good).svg;
  for (const cut of [120, 400, 1000, 2000, good.length - 30]) {
    const out = emfToSvg(good.subarray(0, cut));
    if (out) assert.ok(out.svg.length <= full.length + 200);
  }
  let seed = 7;
  const rand = (n) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
  const started = Date.now();
  for (let i = 0; i < 300; i++) {
    const b = Buffer.from(good);
    for (let k = 0; k < 6; k++) b[108 + rand(b.length - 108)] = rand(256);
    if (i % 3 === 0) b.writeUInt32LE(0xffffffff, 108 + 4 * rand(Math.floor((b.length - 112) / 4)));
    const out = emfToSvg(b);
    if (out) assertInert(out.svg);
  }
  assert.ok(Date.now() - started < 20000, 'quickly');
  // A count of points far past the record's end reads only what the record holds.
  const e = new EmfWriter(10, 10);
  e.add(4, Buffer.alloc(16), Buffer.from([0xff, 0xff, 0xff, 0x7f]), Buffer.alloc(16));
  assert.doesNotThrow(() => emfToSvg(e.build()));
  // Not a metafile, and a header alone.
  assert.equal(emfToSvg(Buffer.from('a few words that are no metafile at all, padded out to more than a header')), null);
  assert.equal(emfToSvg(new EmfWriter(10, 10).build()), null, 'nothing drawn');
});

test('a dual EMF+ file draws from its ordinary records; its EMF+ comments are passed over', () => {
  const e = new EmfWriter(50, 50);
  e.comment(Buffer.concat([Buffer.from('EMF+'), Buffer.alloc(28, 0xab)]));
  e.pen(1, 0, 0, rgb('#ff0000')).select(1).moveTo(0, 0).lineTo(49, 49);
  const p = elements(emfToSvg(e.build()).svg, 'path');
  assert.equal(p.length, 1);
  assert.equal(p[0].stroke, '#ff0000');
});
