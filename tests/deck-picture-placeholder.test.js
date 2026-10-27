// A picture that fills a layout's photo placeholder.
//
// A title layout with a hexagon "Click photo icon to insert photo": the
// slide's picture states no place and no outline of its own, so PowerPoint
// draws it where the layout's placeholder is, cut to the hexagon.
// Presentations drew it nowhere: the picture had no place, and a picture
// was only ever drawn as a rectangle.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { pngOf } from './fixtures/cube-glb.js';

const LAYOUT = 'ppt/slideLayouts/slideLayout1.xml';
const SLIDE = 'ppt/slides/slide1.xml';
const EMU = 9525;

/** A deck whose first slide fills the title layout's hexagon photo placeholder. */
function deckWithPhotoPlaceholder() {
  const deck = Deck.open(buildPptx({ title: 'Photo', slides: [{ layout: 'title', title: 'Tidal model', body: 'Overview' }] }));
  deck.addPicture(0, { data: pngOf([[[200, 40, 40, 255]]]), contentType: 'image/png', name: 'Photo', x: 10, y: 10, w: 20, h: 20 });
  const pkg = OoxmlPackage.read(Buffer.from(deck.save()));
  const put = (name, xml) => { pkg.part(name).data = Buffer.from(xml, 'utf8'); };
  const hexagon =
    '<p:sp><p:nvSpPr><p:cNvPr id="90" name="Picture Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr>' +
    `<p:spPr><a:xfrm><a:off x="${500 * EMU}" y="${40 * EMU}"/><a:ext cx="${300 * EMU}" cy="${260 * EMU}"/></a:xfrm><a:prstGeom prst="hexagon"><a:avLst/></a:prstGeom></p:spPr>` +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click photo icon to insert photo</a:t></a:r></a:p></p:txBody></p:sp>';
  put(LAYOUT, pkg.text(LAYOUT).replace('</p:spTree>', hexagon + '</p:spTree>'));
  // The slide's picture as PowerPoint writes one in a placeholder: no place, no outline.
  put(SLIDE, pkg.text(SLIDE)
    .replace(/(<p:cNvPr[^>]*name="Photo"[\s\S]*?)<p:nvPr\/>/, '$1<p:nvPr><p:ph type="pic" sz="quarter" idx="11"/></p:nvPr>')
    .replace(/(<p:blipFill>[\s\S]*?<\/p:blipFill>)<p:spPr>[\s\S]*?<\/p:spPr>/, '$1<p:spPr/>'));
  return Deck.open(Buffer.from(pkg.write()));
}

test('a picture in a photo placeholder goes where the layout puts it, cut to its outline', () => {
  const deck = deckWithPhotoPlaceholder();
  const picture = deck.slide(0).shapes.find((s) => s.kind === 'picture');
  assert.deepEqual([picture.placeholder.type, picture.placeholder.idx], ['pic', '11'], 'it fills the placeholder');
  assert.deepEqual([picture.geometry.x, picture.geometry.y, picture.geometry.w, picture.geometry.h].map(Math.round), [500, 40, 300, 260], 'the layout\'s place and size');
  assert.equal(picture.preset, 'hexagon', 'the layout\'s outline');

  const svg = renderSlide(deck.slide(0), { resolveImage: (s) => (s.kind === 'picture' ? 'data:image/png;base64,AAAA' : null) });
  const clip = /<clipPath id="([^"]+)"><polygon points="([^"]+)"\/><\/clipPath>/.exec(svg);
  assert.ok(clip, 'drawn inside its outline');
  assert.match(svg, new RegExp(`<image[^>]*href="data:image/png;base64,AAAA"[^>]*clip-path="url\\(#${clip[1]}\\)"`), 'the picture is the one clipped');
});

test('the hexagon is PowerPoint\'s: flat along its top and bottom, its points at the sides', () => {
  const deck = deckWithPhotoPlaceholder();
  const svg = renderSlide(deck.slide(0), { resolveImage: () => 'data:image/png;base64,AAAA' });
  const points = /<clipPath[^>]*><polygon points="([^"]+)"/.exec(svg)[1].trim().split(/\s+/).map((p) => p.split(',').map((v) => Math.round(Number(v))));
  // 300 x 260 at (500, 40): a quarter of the shorter side, 65, cut in at each side.
  assert.deepEqual(points, [[500, 170], [565, 40], [735, 40], [800, 170], [735, 300], [565, 300]]);
});
