// Design → Background Styles — the engine's half: a slide's own background is
// PowerPoint's p:bg, the first child of p:cSld, written and taken off the way
// PowerPoint writes it, read back into the scene the renderer draws, and
// distinct from the layout's or master's, which shows through when the slide
// states none of its own.
//
// `background(index)` gives back the slide's own spec — colours upper case,
// as the ribbon writes one. `slide(index).background` (and the renderer's
// view of it) is the resolved scene, whose colours the theme machinery
// always lower-cases, the same as any other fill on a shape.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide, renderThumbnail } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({
  title: 'Backgrounds',
  slides: [
    { layout: 'title', title: 'Opening', body: 'A deck with a background' },
    { layout: 'obj', title: 'Second slide', body: ['One'] },
  ],
});

test('a solid colour is written as PowerPoint writes it, read back, drawn, and taken off again', () => {
  const deck = Deck.open(DECK);
  assert.equal(deck.background(0), null, 'a built slide has none of its own');
  // Every built deck's master carries its own bg1 background, which shows
  // through until a slide (or its layout) states one of its own.
  const inherited = deck.slide(0).background;
  assert.deepEqual(inherited, { type: 'solid', color: '#ffffff', alpha: 1 }, "the master's own background shows through");

  assert.equal(deck.setBackground(0, { colour: '1F3864' }), true);
  assert.deepEqual(deck.background(0), { colour: '1F3864' });
  assert.deepEqual(deck.slide(0).background, { type: 'solid', color: '#1f3864', alpha: 1 });
  assert.ok(deck.dirty);

  // The first child of cSld, before spTree, exactly as PowerPoint writes one.
  const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part);
  assert.match(xml, /<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="1F3864"\/><\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg><p:spTree>/);
  assert.equal((xml.match(/<p:bg>/g) || []).length, 1);

  // Drawn: the stage's background rect and the thumbnail's both carry it.
  const svg = renderSlide(deck.slide(0));
  assert.match(svg, /<rect x="0" y="0"[^>]*fill="#1f3864"/);
  const thumb = renderThumbnail(deck.slide(0));
  assert.match(thumb, /<rect x="0" y="0"[^>]*fill="#1f3864"/);

  // The other slide is untouched.
  assert.equal(deck.background(1), null);
  assert.doesNotMatch(OoxmlPackage.read(deck.save()).text(deck.slideParts[1].part), /<p:bg>/);

  // Survives a round trip.
  const reopened = Deck.open(deck.save());
  assert.deepEqual(reopened.background(0), { colour: '1F3864' });
  assert.deepEqual(reopened.slide(0).background, { type: 'solid', color: '#1f3864', alpha: 1 });

  // null takes the element off again, and the master's background shows through.
  assert.equal(deck.setBackground(0, null), true);
  assert.equal(deck.background(0), null);
  assert.deepEqual(deck.slide(0).background, inherited);
  assert.doesNotMatch(OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part), /<p:bg>/);
  assert.equal(deck.setBackground(0, null), false, 'nothing left to take off');
});

test('a theme colour with a luminance modifier is written as a:schemeClr and read back', () => {
  const deck = Deck.open(DECK);
  assert.equal(deck.setBackground(0, { scheme: 'accent1', lumMod: 75 }), true);
  assert.deepEqual(deck.background(0), { scheme: 'accent1', lumMod: 75 });

  const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part);
  assert.match(xml, /<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"\/><\/a:schemeClr><\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg>/);

  // Resolved through the theme in the scene, the way a shape's fill is.
  const resolved = deck.slide(0).background;
  assert.equal(resolved.type, 'solid');
  assert.match(resolved.color, /^#[0-9a-f]{6}$/);

  // lumOff too, in schema order after lumMod.
  assert.equal(deck.setBackground(0, { scheme: 'bg2', lumOff: 20 }), true);
  assert.deepEqual(deck.background(0), { scheme: 'bg2', lumOff: 20 });
  assert.match(OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part), /<a:schemeClr val="bg2"><a:lumOff val="20000"\/><\/a:schemeClr>/);

  const reopened = Deck.open(deck.save());
  assert.deepEqual(reopened.background(0), { scheme: 'bg2', lumOff: 20 });
});

test('a two-stop gradient is written with a:gsLst and a:lin, read back, and drawn with a linearGradient', () => {
  const deck = Deck.open(DECK);
  assert.equal(
    deck.setBackground(0, { gradient: { from: { scheme: 'accent1' }, to: { scheme: 'accent1', lumMod: 75 } } }),
    true
  );
  assert.deepEqual(deck.background(0), { gradient: { from: { scheme: 'accent1' }, to: { scheme: 'accent1', lumMod: 75 }, angle: 90 } }, 'the default angle, top to bottom, comes back stated');

  const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part);
  assert.match(
    xml,
    /<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"\/><\/a:gs><a:gs pos="100000"><a:schemeClr val="accent1"><a:lumMod val="75000"\/><\/a:schemeClr><\/a:gs><\/a:gsLst><a:lin ang="5400000"\/><\/a:gradFill><a:effectLst\/><\/p:bgPr><\/p:bg>/
  );

  const scene = deck.slide(0).background;
  assert.equal(scene.type, 'gradient');
  assert.equal(scene.stops.length, 2);
  assert.equal(scene.stops[0].offset, 0);
  assert.equal(scene.stops[1].offset, 1);
  assert.equal(scene.angle, 90);

  const svg = renderSlide(deck.slide(0));
  assert.match(svg, /<linearGradient id="g1"/);
  assert.match(svg, /fill="url\(#g1\)"/);

  // A stated angle round-trips too, and a colour spec may be a bare hex string.
  assert.equal(deck.setBackground(0, { gradient: { from: 'FF0000', to: '0000FF', angle: 45 } }), true);
  assert.equal(deck.background(0).gradient.angle, 45);
  assert.match(OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part), /<a:lin ang="2700000"\/>/);
  const reopened = Deck.open(deck.save());
  assert.deepEqual(reopened.background(0), { gradient: { from: { colour: 'FF0000' }, to: { colour: '0000FF' }, angle: 45 } });
});

test("setBackground(null) takes the element off and the slide shows the layout's background again", () => {
  const deck = Deck.open(DECK);
  // The layout carries its own background — set here by hand, the way a
  // deck built elsewhere might carry one — which beats the master's when
  // the slide states none of its own, and shows through again once the
  // slide's own is taken off.
  const layoutPart = deck.layoutOf(0);
  const layoutXml = deck.pkg.text(layoutPart);
  const withBg = layoutXml.replace(
    /(<p:cSld[^>]*>)/,
    '$1<p:bg><p:bgPr><a:solidFill><a:srgbClr val="009900"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
  );
  deck.pkg.write_(layoutPart, Buffer.from(withBg, 'utf8'));
  deck._scenes.clear();
  deck.layouts.clear();

  assert.equal(deck.background(0), null, 'the slide states none of its own');
  assert.deepEqual(deck.slide(0).background, { type: 'solid', color: '#009900', alpha: 1 }, 'inherited from the layout');

  assert.equal(deck.setBackground(0, { colour: 'FFFFFF' }), true);
  assert.deepEqual(deck.slide(0).background, { type: 'solid', color: '#ffffff', alpha: 1 }, "the slide's own hides the layout's");

  assert.equal(deck.setBackground(0, null), true);
  assert.equal(deck.background(0), null);
  assert.deepEqual(deck.slide(0).background, { type: 'solid', color: '#009900', alpha: 1 }, "the layout's shows through again");
});

test('a slide already carrying a p:bgRef is replaced cleanly', () => {
  const pkg = OoxmlPackage.read(DECK);
  const deck = new Deck(pkg);
  const part = deck.slideParts[0].part;
  const xml = deck.pkg.text(part);
  const withRef = xml.replace(
    '<p:cSld><p:spTree>',
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="accent2"/></p:bgRef></p:bg><p:spTree>'
  );
  deck.pkg.write_(part, Buffer.from(withRef, 'utf8'));
  deck._scenes.clear();

  assert.deepEqual(deck.background(0), { scheme: 'accent2' }, 'the bgRef form is read too');

  assert.equal(deck.setBackground(0, { colour: '112233' }), true);
  const next = OoxmlPackage.read(deck.save()).text(part);
  assert.doesNotMatch(next, /bgRef/);
  assert.match(next, /<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"\/>/);
  assert.equal((next.match(/<p:bg>/g) || []).length, 1);
});

test('other slides are byte-identical, the deck survives a save-and-reopen, and a bad index throws', () => {
  const deck = Deck.open(DECK);
  const before1 = deck.pkg.text(deck.slideParts[1].part);
  assert.equal(deck.setBackground(0, { colour: 'ABCDEF' }), true);
  assert.equal(deck.pkg.text(deck.slideParts[1].part), before1);

  const reopened = Deck.open(deck.save());
  assert.deepEqual(reopened.background(0), { colour: 'ABCDEF' });
  assert.equal(reopened.background(1), null);
  assert.equal(reopened.slideCount, 2);

  assert.throws(() => deck.setBackground(9, { colour: '000000' }), /no slide/);
  assert.throws(() => deck.background(9), /no slide/);
});
