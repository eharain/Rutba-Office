// The presentation engine, end to end: build a deck, read it back, edit it,
// and prove the edit changed only what it should have.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide, renderThumbnail, sceneText, TEMPLATES } from '@rutba/presentation';
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

test('a picture can be added to a slide: media part, relationship, default content type, and it draws as an image', () => {
  const PNG_1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const png = Buffer.from(PNG_1x1, 'base64');
  const deck = Deck.open(DECK);
  const untouched = OoxmlPackage.read(DECK).text('ppt/slides/slide1.xml');

  const { id, part } = deck.addPicture(1, { data: png, contentType: 'image/png', name: 'dot.png', x: 100, y: 120, w: 200, h: 150 });
  assert.equal(part, 'ppt/media/image1.png');

  const reopened = Deck.open(deck.save());
  const shape = reopened.slide(1).shapes.find((s) => s.id === String(id));
  assert.equal(shape?.kind, 'picture', 'the shape reads back as a picture');
  assert.equal(shape.source?.part, part, 'its blip resolves to the media part');
  assert.deepEqual(
    [shape.geometry.x, shape.geometry.y, shape.geometry.w, shape.geometry.h].map(Math.round),
    [100, 120, 200, 150],
    'placed where it was asked to be'
  );
  assert.ok(Buffer.from(reopened.media(part)).equals(png), 'the bytes come back as they went in');

  const pkg = OoxmlPackage.read(reopened.save());
  assert.equal(pkg.contentTypeOf(part), 'image/png');
  assert.match(pkg.text('[Content_Types].xml'), /<Default Extension="png" ContentType="image\/png"\/>/, 'declared as a default, the way Office does');
  assert.doesNotMatch(pkg.text('[Content_Types].xml'), /Override[^>]*media\/image1/, 'not as a per-picture override');
  assert.equal(pkg.text('ppt/slides/slide1.xml'), untouched, 'the other slides are byte-identical');

  const svg = renderSlide(reopened.slide(1), {
    width: 640,
    resolveImage: (s) => `data:image/png;base64,${Buffer.from(reopened.media(s.source.part)).toString('base64')}`,
  });
  assert.match(svg, /<image [^>]*href="data:image\/png;base64,/, 'drawn as an image, not as the "Picture" placeholder');

  const second = reopened.addPicture(1, { data: png, contentType: 'image/png', w: 10, h: 10 });
  assert.equal(second.part, 'ppt/media/image2.png', 'a second picture gets the next number');
  assert.notEqual(second.id, id);
  assert.throws(() => reopened.addPicture(1, { data: png, contentType: 'image/tiff', w: 10, h: 10 }), /unsupported picture type/);
});

test('a preset shape can be added with the theme accent, a line and centred text, and draws as its geometry', () => {
  const deck = Deck.open(DECK);
  const untouched = OoxmlPackage.read(DECK).text('ppt/slides/slide1.xml');
  const id = deck.addShape(2, { preset: 'ellipse', x: 200, y: 100, w: 240, h: 160, text: 'Hello' });
  const bare = deck.addShape(2, { preset: 'star5', x: 500, y: 100, w: 120, h: 120, fill: '#FFC000', line: 'none' });

  const reopened = Deck.open(deck.save());
  const scene = reopened.slide(2);
  const oval = scene.shapes.find((s) => s.id === String(id));
  assert.equal(oval?.kind, 'shape');
  assert.equal(oval.preset, 'ellipse');
  assert.equal(oval.fill?.type, 'solid');
  assert.equal(oval.fill.color.toUpperCase(), String(scene.theme.colors.accent1).toUpperCase(), 'filled with the theme accent, not a hard-coded colour');
  assert.ok(oval.line?.color, 'it has a line');
  assert.equal(oval.text?.anchor, 'middle', 'the text sits in the middle of the shape');
  assert.equal(oval.text.paragraphs[0].runs.map((r) => r.text).join(''), 'Hello');
  assert.deepEqual([oval.geometry.x, oval.geometry.y, oval.geometry.w, oval.geometry.h].map(Math.round), [200, 100, 240, 160]);

  const star = scene.shapes.find((s) => s.id === String(bare));
  assert.equal(star.preset, 'star5');
  assert.equal(star.fill.color.toUpperCase(), '#FFC000');
  assert.equal(star.line?.type, 'none');

  const svg = renderSlide(scene, { width: 640 });
  assert.match(svg, /<ellipse [^>]*fill="#/, 'the oval draws as an ellipse');
  assert.match(svg, /<polygon [^>]*fill="#FFC000"/i, 'the star draws as its polygon in its own colour');
  assert.match(svg, />Hello</, 'and the text is on it');
  assert.equal(OoxmlPackage.read(reopened.save()).text('ppt/slides/slide1.xml'), untouched, 'the other slides are byte-identical');
  assert.throws(() => reopened.addShape(2, { preset: 'rect"/><x', w: 10, h: 10 }), /not a preset geometry/);
});

test('a bullet stored for Wingdings or Symbol is drawn as the character it looks like', async () => {
  const { bulletGlyph } = await import('@rutba/ooxml/glyphs');
  assert.equal(bulletGlyph('§', 'Wingdings'), '▪', 'Wingdings 0xA7 is a small square');
  assert.equal(bulletGlyph('', 'Wingdings'), '▪', 'the same code point in the private-use area');
  assert.equal(bulletGlyph('ü', 'Wingdings'), '✓');
  assert.equal(bulletGlyph('', 'Symbol'), '•');
  assert.equal(bulletGlyph(''), '•', 'a private-use bullet with no font named is still a bullet');
  assert.equal(bulletGlyph('–', 'Calibri'), '–', 'a plain character in a text font is itself');
  assert.equal(bulletGlyph('§', 'Arial'), '§');
});

test('the outline reads titles, counts and notes off the XML, and agrees with the scenes', () => {
  const deck = Deck.open(DECK);
  const outline = deck.outline();
  assert.equal(outline.length, deck.slideCount);
  for (const o of outline) {
    const scene = deck.slide(o.index);
    const title = scene.shapes.find((s) => s.placeholder?.type === 'title' || s.placeholder?.type === 'ctrTitle');
    const body = title?.text || title?.inheritedText;
    const expected = body ? body.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join(' ').trim() : '';
    assert.equal(o.title, expected, `slide ${o.index + 1} title`);
    assert.equal(o.shapes, scene.shapes.length, `slide ${o.index + 1} shape count`);
    assert.equal(o.notes, scene.notes.replace(/\s+/g, ' ').trim(), `slide ${o.index + 1} notes`);
  }
  deck.setNotes(1, 'Speaker line one\nand two');
  assert.equal(deck.outline()[1].notes, 'Speaker line one and two');
  assert.match(Deck.open(buildPptx({ title: 'Fish &amp; Chips', slides: [{ layout: 'title', title: 'Fish & Chips' }] })).outline()[0].title, /^Fish & Chips$/, 'entities are decoded');
});

test('a thumbnail leaves out a custom path with a hundred thousand points; the slide itself keeps it', () => {
  const deck = Deck.open(DECK);
  const pkg = deck.pkg;
  const part = deck.slideParts[2].part;
  let points = '';
  for (let i = 0; i < 6000; i++) points += `<a:lnTo><a:pt x="${(i * 37) % 1000}" y="${(i * 91) % 1000}"/></a:lnTo>`;
  const sp =
    '<p:sp><p:nvSpPr><p:cNvPr id="99" name="Map"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="3048000"/></a:xfrm>' +
    '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>' +
    `<a:path w="1000" h="1000"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${points}<a:close/></a:path></a:pathLst></a:custGeom>` +
    '<a:solidFill><a:srgbClr val="336699"/></a:solidFill></p:spPr></p:sp>';
  const xml = pkg.text(part).replace('</p:spTree>', sp + '</p:spTree>');
  pkg.write_(part, Buffer.from(xml, 'utf8'));
  const scene = deck.slide(2);
  const full = renderSlide(scene, { width: 640 });
  assert.match(full, /<path d="M/, 'the slide draws the path');
  const thumb = renderThumbnail(scene, 220);
  assert.doesNotMatch(thumb, /<path d="M/, 'the thumbnail leaves it out');
  assert.ok(thumb.length < full.length / 4, `the thumbnail is small (${thumb.length} vs ${full.length})`);
});
