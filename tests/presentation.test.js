// The presentation engine, end to end: build a deck, read it back, edit it,
// and prove the edit changed only what it should have.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide, sceneText, TEMPLATES } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({
  title: 'Quarterly Review',
  author: 'Rutba Office',
  slides: [
    { layout: 'title', title: 'Quarterly Review', body: 'Q3 2026' },
    {
      layout: 'obj',
      title: 'Highlights',
      body: [
        { runs: [{ text: 'Revenue up ' }, { text: '18%', bold: true }] },
        { level: 1, runs: [{ text: 'Driven by the new region' }] },
        'Churn down to 2.1%',
      ],
    },
    {
      layout: 'obj',
      title: 'Detail',
      textBoxes: [
        { x: 60, y: 300, w: 400, h: 80, paragraphs: [{ align: 'center', runs: [{ text: 'A free-standing note', italic: true, size: 14 }] }] },
      ],
    },
  ],
});

test('a built deck is a valid OOXML package with every part a presentation needs', () => {
  const pkg = OoxmlPackage.read(DECK);
  const names = pkg.partNames();
  for (const required of [
    '[Content_Types].xml',
    '_rels/.rels',
    'ppt/presentation.xml',
    'ppt/_rels/presentation.xml.rels',
    'ppt/slideMasters/slideMaster1.xml',
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    'ppt/slideLayouts/slideLayout1.xml',
    'ppt/theme/theme1.xml',
    'ppt/slides/slide1.xml',
    'ppt/slides/_rels/slide1.xml.rels',
    'docProps/core.xml',
  ]) {
    assert.ok(names.includes(required), `missing part: ${required}`);
  }
  // Every slide is declared in the content types, or PowerPoint repairs the file.
  const types = pkg.text('[Content_Types].xml');
  assert.match(types, /slides\/slide3\.xml/);
});

test('the deck reads back with its slides, size and text', () => {
  const deck = Deck.open(DECK);
  assert.equal(deck.slideCount, 3);
  assert.equal(deck.size.cx, 12192000);
  assert.equal(Math.round(deck.size.width), 1280);

  const first = deck.slide(0);
  assert.equal(sceneText(first).split('\n')[0], 'Quarterly Review');

  const second = deck.slide(1);
  const text = sceneText(second);
  assert.match(text, /Highlights/);
  assert.match(text, /Revenue up 18%/);
  assert.match(text, /Churn down to 2\.1%/);
});

test('a run keeps the formatting it was written with', () => {
  const deck = Deck.open(DECK);
  const slide = deck.slide(1);
  const body = slide.shapes.find((s) => s.placeholder?.type === 'body');
  assert.ok(body, 'the body placeholder is on the slide');
  const runs = body.text.paragraphs[0].runs;
  assert.equal(runs.length, 2);
  assert.equal(runs[1].text, '18%');
  assert.equal(runs[1].bold, true);
  assert.equal(body.text.paragraphs[1].level, 1);
});

test('a placeholder with no geometry of its own inherits it from the layout', () => {
  const deck = Deck.open(DECK);
  const slide = deck.slide(0);
  const title = slide.shapes.find((s) => s.placeholder?.type === 'ctrTitle');
  assert.ok(title.geometry, 'geometry came from somewhere');
  assert.ok(title.geometry.w > 100, 'and it is a real size');
});

test('theme colours resolve through the master colour map', () => {
  const deck = Deck.open(DECK);
  const slide = deck.slide(0);
  assert.equal(slide.background.type, 'solid');
  assert.equal(slide.background.color.toLowerCase(), '#ffffff', 'bg1 maps to lt1');
  assert.equal(slide.theme.colors.accent1.toLowerCase(), '#2b5fd9');
});

test('editing text rewrites one slide and leaves the others byte-identical', () => {
  const deck = Deck.open(DECK);
  const before = OoxmlPackage.read(DECK);
  const slide = deck.slide(1);
  const title = slide.shapes.find((s) => s.placeholder?.type === 'title');

  deck.setText(1, title.id, [{ runs: [{ text: 'Revised highlights', bold: true }] }]);
  const out = deck.save();

  const after = OoxmlPackage.read(out);
  assert.notDeepEqual(after.read('ppt/slides/slide2.xml'), before.read('ppt/slides/slide2.xml'));
  for (const part of ['ppt/slides/slide1.xml', 'ppt/slides/slide3.xml', 'ppt/theme/theme1.xml', 'ppt/slideMasters/slideMaster1.xml']) {
    assert.deepEqual(after.read(part), before.read(part), `${part} was rewritten and should not have been`);
  }

  const reopened = Deck.open(out);
  assert.match(sceneText(reopened.slide(1)), /Revised highlights/);
  assert.doesNotMatch(sceneText(reopened.slide(1)), /^Highlights$/m);
});

test('moving a shape writes explicit geometry where it had inherited it', () => {
  const deck = Deck.open(DECK);
  const title = deck.slide(1).shapes.find((s) => s.placeholder?.type === 'title');
  assert.ok(title);
  deck.setGeometry(1, title.id, { x: 100, y: 50, w: 600, h: 90 });
  const reopened = Deck.open(deck.save());
  const moved = reopened.slide(1).shapes.find((s) => s.id === title.id);
  assert.equal(Math.round(moved.geometry.x), 100);
  assert.equal(Math.round(moved.geometry.y), 50);
  assert.equal(Math.round(moved.geometry.w), 600);
});

test('a text box can be added, read back, and removed', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTextBox(0, { x: 40, y: 400, w: 300, h: 60, paragraphs: [{ runs: [{ text: 'Added later' }] }] });
  let reopened = Deck.open(deck.save());
  assert.match(sceneText(reopened.slide(0)), /Added later/);

  const shape = reopened.slide(0).shapes.find((s) => s.id === String(id));
  assert.ok(shape, 'the new shape has the id we were given');
  reopened.removeShape(0, id);
  const again = Deck.open(reopened.save());
  assert.doesNotMatch(sceneText(again.slide(0)), /Added later/);
});

test('slides can be duplicated, reordered and removed', () => {
  const deck = Deck.open(DECK);
  const at = deck.duplicateSlide(1);
  assert.equal(deck.slideCount, 4);
  assert.equal(at, 2, 'the copy lands directly after its original');
  assert.match(sceneText(deck.slide(2)), /Highlights/);

  deck.moveSlide(2, 0);
  assert.match(sceneText(deck.slide(0)), /Highlights/);

  deck.removeSlide(0);
  assert.equal(deck.slideCount, 3);
  assert.match(sceneText(deck.slide(0)), /Quarterly Review/);

  // and it still opens as a package after all that
  const out = Deck.open(deck.save());
  assert.equal(out.slideCount, 3);
});

test('a slide renders to SVG with its text and shapes', () => {
  const deck = Deck.open(DECK);
  const svg = renderSlide(deck.slide(1));
  assert.match(svg, /^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg"/);
  assert.match(svg, /viewBox="0 0 1280 720"/);
  assert.match(svg, /Highlights/);
  assert.match(svg, /font-weight="700"/, 'the bold run is bold in the drawing');
  assert.match(svg, /18%/);
});

test('an outline gives every slide a title without reading the parts twice', () => {
  const deck = Deck.open(DECK);
  const outline = deck.outline();
  assert.equal(outline.length, 3);
  assert.equal(outline[0].title, 'Quarterly Review');
  assert.equal(outline[1].title, 'Highlights');
});

test('the pitch template is a real deck', () => {
  const deck = Deck.open(TEMPLATES.pitch());
  assert.equal(deck.slideCount, 5);
  assert.equal(deck.outline()[4].title, 'The ask');
});

test('an XML-hostile string survives the round trip', () => {
  const deck = Deck.open(DECK);
  const title = deck.slide(0).shapes.find((s) => s.placeholder?.type === 'ctrTitle');
  const nasty = 'R&D <tags> & "quotes" — 100% ünïcode';
  deck.setText(0, title.id, [{ runs: [{ text: nasty }] }]);
  const reopened = Deck.open(deck.save());
  assert.equal(sceneText(reopened.slide(0)).split('\n')[0], nasty);
});

test('leading and trailing spaces are preserved in a run', () => {
  const deck = Deck.open(DECK);
  const title = deck.slide(0).shapes.find((s) => s.placeholder?.type === 'ctrTitle');
  deck.setText(0, title.id, [{ runs: [{ text: '  spaced  ' }] }]);
  const reopened = Deck.open(deck.save());
  assert.equal(sceneText(reopened.slide(0)).split('\n')[0], '  spaced  ');
});

test('a custom geometry draws as its own path, scaled onto the shape, not as a rectangle', async () => {
  const { readSlideScene } = await import('@rutba/presentation');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    // A banner with a diagonal cut: the shape a preset cannot describe.
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Banner"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
    '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100">' +
    '<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="80" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/>' +
    '</a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="2292FD"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>' +
    '</p:spTree></p:cSld></p:sld>';
  const scene = readSlideScene(xml, {});
  const banner = scene.shapes.find((s) => s.name === 'Banner');
  assert.equal(banner.preset, 'custom');
  assert.deepEqual(banner.path, { w: 100, h: 100, d: 'M0 0 L100 0 L80 100 L0 100 Z', filled: true });
  const svg = renderSlide(scene, { width: 960 });
  // 914400 EMU is one inch, 96 px: the path's 100-unit box scales to 96.
  assert.match(svg, /<path d="M0 0L96 0L76\.8 96L0 96Z" fill="#2292fd"/i);
  assert.ok(!/<rect x="0" y="0" width="96" height="96"/.test(svg), 'no rectangle stands in for it');
});

test('an arcTo in a custom geometry becomes an SVG arc from its sweep', async () => {
  const { readSlideScene } = await import('@rutba/presentation');
  const xml = '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Swoosh"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
    '<a:custGeom><a:pathLst><a:path w="200" h="200"><a:moveTo><a:pt x="0" y="100"/></a:moveTo>' +
    // A half turn clockwise from the left of a 100-radius circle centred at (100,100): ends at (200,100).
    '<a:arcTo wR="100" hR="100" stAng="10800000" swAng="10800000"/><a:close/></a:path></a:pathLst></a:custGeom></p:spPr></p:sp>' +
    '</p:spTree></p:cSld></p:sld>';
  const scene = readSlideScene(xml, {});
  const d = scene.shapes[0].path.d;
  assert.match(d, /^M0 100 A100 100 0 0 1 200 100 Z$/);
});
