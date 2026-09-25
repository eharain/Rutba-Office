// View → Slide Master — the deck engine's half.
//
// A master and its layouts are parts like a slide: the same verbs edit them,
// and every slide on them follows. The master's title and body placeholders
// carry the text styles every slide inherits; a layout's placeholders carry
// their own; the shapes a master or a layout draws sit under every slide.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const MASTER = 'ppt/slideMasters/slideMaster1.xml';
const TITLE_LAYOUT = 'ppt/slideLayouts/slideLayout1.xml';
const CONTENT_LAYOUT = 'ppt/slideLayouts/slideLayout2.xml';
const DECK = buildPptx({
  title: 'Master',
  slides: [
    { layout: 'title', title: 'Opening', body: 'Subtitle' },
    { layout: 'obj', title: 'First', body: ['One', 'Two'] },
    { layout: 'obj', title: [{ runs: [{ text: 'Own green', color: '#00AA00' }] }], body: ['Three'] },
  ],
});
const ph = (scene, type) => scene.shapes.find((s) => s.placeholder?.type === type);

test('the master and a layout are drawn as scenes of their own, each empty placeholder showing PowerPoint\'s prompt', () => {
  const deck = Deck.open(DECK);
  const master = deck.partScene(MASTER);
  assert.equal(master.kind, 'master');
  assert.equal(ph(master, 'title').text.paragraphs[0].runs[0].text, 'Click to edit Master title style');
  assert.deepEqual(ph(master, 'body').text.paragraphs.map((p) => p.level || 0), [0, 1, 2, 3, 4], 'the body prompts five levels');
  assert.equal(ph(master, 'title').prompt, true);
  assert.equal(ph(master, 'title').textStyle[0].size, 44, 'drawn in the master title style');
  const layout = deck.partScene(TITLE_LAYOUT);
  assert.equal(layout.kind, 'layout');
  assert.equal(layout.name, 'Title Slide');
  assert.equal(ph(layout, 'subTitle').text.paragraphs[0].runs[0].text, 'Click to edit Master subtitle style');
  assert.ok(ph(layout, 'ctrTitle').geometry, 'the layout\'s own placement');
  const svg = renderSlide(master, { placeholderFrames: true });
  assert.match(svg, /stroke-dasharray="6 4"/, 'placeholder frames');
  assert.match(svg, /Click to edit Master title style/);
  assert.equal(deck.dirty, false, 'a prompt is shown, never written');
});

test('a shape on the master is drawn under every slide, and a layout or a slide can hide it', () => {
  const deck = Deck.open(DECK);
  deck.addShape(MASTER, { preset: 'rect', x: 0, y: 700, w: 1280, h: 20, fill: 'C00000', line: 'none' });
  for (let i = 0; i < 3; i++) {
    assert.equal(deck.slide(i).underlay.length, 1, `slide ${i + 1} has the master's band under it`);
    assert.match(renderSlide(deck.slide(i)), /fill="#C00000"/i);
  }
  assert.equal(deck.slide(1).shapes.some((s) => s.underlay), false, 'it is not one of the slide\'s own shapes');
  // The show animates a slide's own shapes: the master's are never tagged as one of them.
  const tagged = renderSlide(deck.slide(2), { tagShapes: true });
  assert.match(tagged, /<rect x="0" y="700"[^>]*fill="#C00000"/i);
  assert.doesNotMatch(tagged, /<g data-shape="[^"]*"><rect x="0" y="700"/);
  assert.equal(deck.setHideBackgroundGraphics(CONTENT_LAYOUT, true), true);
  assert.match(deck.pkg.text(CONTENT_LAYOUT), /<p:sldLayout\b[^>]*showMasterSp="0"/);
  assert.equal(deck.slide(1).underlay.length, 0, 'the content layout hides it');
  assert.equal(deck.slide(0).underlay.length, 1, 'the title layout does not');
  deck.setHideBackgroundGraphics(0, true);
  assert.equal(deck.slide(0).underlay.length, 0, 'and a slide can hide it too');
});

test('formatting the master\'s title placeholder writes the title style, and every slide\'s title follows but its own colour', () => {
  const deck = Deck.open(DECK);
  const title = ph(deck.partScene(MASTER), 'title');
  assert.equal(deck.setTextStyle(MASTER, title.id, { color: '#1F6FB2', font: 'Georgia', size: 36, bold: false, align: 'center' }), true);
  const style = /<p:titleStyle>[\s\S]*?<\/p:titleStyle>/.exec(deck.pkg.text(MASTER))[0];
  assert.equal(style, '<p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr b="0" sz="3600"><a:solidFill><a:srgbClr val="1F6FB2"/></a:solidFill><a:latin typeface="Georgia"/></a:defRPr></a:lvl1pPr></p:titleStyle>', 'children in the schema\'s order');
  for (const i of [0, 1]) {
    const t = deck.slide(i).shapes.find((s) => /title/i.test(s.placeholder?.type));
    assert.equal(t.textStyle[0].color, '#1f6fb2', `slide ${i + 1}'s title`);
    assert.equal(t.textStyle[0].font, 'Georgia');
    assert.equal(t.textStyle[0].size, 36);
  }
  const own = deck.slide(2).shapes.find((s) => s.placeholder?.type === 'title');
  assert.equal(own.text.paragraphs[0].runs[0].color.toLowerCase(), '#00aa00', 'a run\'s own colour stays');
  assert.match(renderSlide(deck.slide(1)), /font-family="Georgia, Georgia, 'Times New Roman'/);
  // The body style's five levels scale together.
  const body = ph(deck.partScene(MASTER), 'body');
  deck.setTextStyle(MASTER, body.id, { size: 24 });
  const sizes = [...deck.pkg.text(MASTER).matchAll(/<a:lvl(\d)pPr[^>]*>(?:(?!<\/a:lvl\dpPr>)[\s\S])*?<a:defRPr sz="(\d+)"/g)].filter((m) => /<p:bodyStyle>/.test(deck.pkg.text(MASTER))).map((m) => Number(m[2]));
  assert.ok(sizes.includes(2400) && sizes.includes(2160), `levels scaled: ${sizes.join(',')}`);
});

test('a layout\'s placeholder is styled in its own list style, and only the slides on that layout change', () => {
  const deck = Deck.open(DECK);
  const sub = ph(deck.partScene(TITLE_LAYOUT), 'subTitle');
  deck.setTextStyle(TITLE_LAYOUT, sub.id, { color: 'AA3300', italic: true });
  assert.match(deck.pkg.text(TITLE_LAYOUT), /<a:lstStyle><a:lvl1pPr><a:defRPr i="1"><a:solidFill><a:srgbClr val="AA3300"\/><\/a:solidFill><\/a:defRPr><\/a:lvl1pPr>/);
  assert.equal(deck.slide(0).shapes.find((s) => s.placeholder?.type === 'subTitle').textStyle[0].color, '#aa3300');
  assert.equal(deck.slide(1).shapes.find((s) => s.placeholder?.type === 'body').textStyle[0].color, '#000000', 'the content layout\'s body is untouched');
});

test('moving the master\'s title placeholder moves every slide title that does not place itself', () => {
  const deck = Deck.open(DECK);
  const title = ph(deck.partScene(MASTER), 'title');
  deck.setGeometry(MASTER, title.id, { x: 40, y: 30, w: 1200, h: 90 });
  const t = deck.slide(1).shapes.find((s) => s.placeholder?.type === 'title');
  assert.deepEqual([Math.round(t.geometry.x), Math.round(t.geometry.y), Math.round(t.geometry.w)], [40, 30, 1200]);
  assert.equal(Math.round(deck.slide(0).shapes.find((s) => s.placeholder?.type === 'ctrTitle').geometry.y), 203, 'the title slide\'s layout places its own');
  // A background on the master reaches slides that state none.
  deck.setBackground(MASTER, { colour: 'F3EFE6' });
  assert.equal(deck.slide(1).background.color.toLowerCase(), '#f3efe6');
});

test('Insert Layout adds a layout PowerPoint can open — its part, both relationships, a fresh id and its content type — and Rename and Delete work on it', () => {
  const deck = Deck.open(DECK);
  const part = deck.insertLayout(MASTER, { name: 'Quote' });
  assert.equal(part, 'ppt/slideLayouts/slideLayout3.xml');
  const master = deck.pkg.text(MASTER);
  const ids = [...master.matchAll(/<p:sldLayoutId id="(\d+)"/g)].map((m) => Number(m[1]));
  assert.equal(new Set(ids).size, 3);
  assert.ok(ids.every((n) => n > 2147483648));
  assert.match(deck.pkg.text('[Content_Types].xml'), /PartName="\/ppt\/slideLayouts\/slideLayout3\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.presentationml\.slideLayout\+xml"/);
  assert.match(deck.pkg.text('ppt/slideLayouts/_rels/slideLayout3.xml.rels'), /slideMaster" Target="\.\.\/slideMasters\/slideMaster1\.xml"/);
  const list = deck.masterList()[0];
  assert.deepEqual(list.layouts.map((l) => l.name), ['Title Slide', 'Title and Content', 'Quote']);
  assert.equal(list.layouts[2].hasTitle, true);
  assert.equal(deck.renamePart(part, 'Pull quote'), true);
  assert.equal(deck.masterList()[0].layouts[2].name, 'Pull quote');
  // Saved and opened again, a slide can be put on it.
  const again = Deck.open(deck.save());
  assert.equal(again.applyLayout(1, part), true);
  assert.equal(again.slide(1).layout, part);
  assert.throws(() => again.removeLayout(part), /A slide uses this layout/);
  // Unused, it goes — part, rels, list entry and content type.
  deck.removeLayout(part);
  assert.equal(deck.pkg.has(part), false);
  assert.equal(deck.pkg.has('ppt/slideLayouts/_rels/slideLayout3.xml.rels'), false);
  assert.doesNotMatch(deck.pkg.text(MASTER), new RegExp(`r:id="${/r:id="(rId\d+)"/.exec(master.split('<p:sldLayoutId').pop())[1]}"`));
  assert.doesNotMatch(deck.pkg.text('[Content_Types].xml'), /slideLayout3\.xml/);
  assert.ok(Deck.open(deck.save()));
});

test('Title and Footers put a layout\'s placeholders on or off, Insert Placeholder adds one, and Preserve marks the master', () => {
  const deck = Deck.open(DECK);
  assert.equal(deck.setMasterPlaceholders(CONTENT_LAYOUT, { title: false }), true);
  assert.equal(deck.masterList()[0].layouts[1].hasTitle, false);
  deck.setMasterPlaceholders(CONTENT_LAYOUT, { title: true, footers: true });
  const l = deck.masterList()[0].layouts[1];
  assert.equal(l.hasTitle, true);
  assert.equal(l.hasFooters, true);
  const scene = deck.partScene(CONTENT_LAYOUT);
  assert.ok(['dt', 'ftr', 'sldNum'].every((t) => ph(scene, t)), 'date, footer and number');
  assert.ok(ph(scene, 'title').geometry, 'the new title takes the master\'s place');
  deck.setMasterPlaceholders(MASTER, { footers: true });
  assert.ok(ph(deck.partScene(MASTER), 'sldNum').geometry.y > deck.size.height * 0.85, 'the master\'s own footers go along the bottom');
  const id = deck.insertPlaceholder(TITLE_LAYOUT, 'picture');
  assert.match(deck.pkg.text(TITLE_LAYOUT), new RegExp(`<p:cNvPr id="${id}" name="Picture Placeholder [^"]*"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" sz="quarter" idx="13"/>`));
  assert.throws(() => deck.insertPlaceholder(MASTER, 'text'), /on a layout/);
  deck.setMasterPreserve(MASTER, true);
  assert.match(deck.pkg.text(MASTER), /<p:sldMaster preserve="1"/);
  assert.equal(deck.masterList()[0].preserve, true);
});

test('a master edit is one undo step, and a master whose title shares the shape tree\'s id is still edited', () => {
  const pkg = OoxmlPackage.read(DECK);
  // An older build of this suite gave the master's title placeholder id 1, the shape tree's own.
  pkg.write_(MASTER, pkg.text(MASTER).replace('<p:cNvPr id="2" name="title Placeholder 1"/>', '<p:cNvPr id="1" name="title Placeholder 1"/>'));
  const deck = Deck.open(pkg.write());
  const before = deck.pkg.text(MASTER);
  const snap = deck.snapshot();
  const title = ph(deck.partScene(MASTER), 'title');
  assert.equal(String(title.id), '1');
  deck.setTextStyle(MASTER, title.id, { color: '336699' });
  deck.pushUndo(snap);
  assert.equal(deck.slide(1).shapes.find((s) => s.placeholder?.type === 'title').textStyle[0].color, '#336699');
  deck.undo();
  assert.equal(deck.pkg.text(MASTER), before);
  assert.equal(deck.slide(1).shapes.find((s) => s.placeholder?.type === 'title').textStyle[0].color, '#000000');
});
