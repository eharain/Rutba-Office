// WordArt's Transform in a deck: a text box's words laid along Arch Up,
// Arch Down, Circle or Button — written as PowerPoint writes a:prstTxWarp,
// read back with its handles, drawn along the preset's path with
// <textPath>, and taken off again.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { renderSlide } from '@rutba/presentation/render';
import { warpPaths, WARP_PRESETS } from '@rutba/drawing/warp';

const deckWithBox = () => {
  const d = Deck.open(buildPptx({ title: 'Art', slides: [{ layout: 'blank' }] }));
  const id = d.addTextBox(0, { x: 100, y: 100, w: 400, h: 200, paragraphs: [{ runs: [{ text: 'Rutba Office' }] }] });
  return { d, id: id?.id ?? id };
};

test('a Transform is written first in the body properties, read back, and taken off', () => {
  const { d, id } = deckWithBox();
  d.setTextWarp(0, id, 'textArchUp');
  const xml = d.pkg.text(d.slideParts[0].part);
  assert.match(xml, /<a:bodyPr\b[^>]*><a:prstTxWarp prst="textArchUp"><a:avLst\/><\/a:prstTxWarp>/);
  const shape = Deck.open(d.save()).slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.deepEqual(shape.text.warp, { preset: 'textArchUp', adj: {} });
  d.setTextWarp(0, id, 'textCircle');
  assert.equal((d.pkg.text(d.slideParts[0].part).match(/prstTxWarp prst=/g) || []).length, 1, 'one preset, replaced');
  d.setTextWarp(0, id, null);
  assert.doesNotMatch(d.pkg.text(d.slideParts[0].part), /prstTxWarp/);
  assert.equal(d.slide(0).shapes.find((s) => String(s.id) === String(id)).text.warp, undefined);
  assert.throws(() => d.setTextWarp(0, id, '<x/>'), /not a Transform preset/);
});

test('a PowerPoint Transform with its handles is read', () => {
  const { d, id } = deckWithBox();
  d.setTextWarp(0, id, 'textArchDown');
  // As PowerPoint writes one, with the handle dragged.
  const part = d.slideParts[0].part;
  const pkg = OoxmlPackage.read(d.save());
  pkg.write_(part, pkg.text(part).replace('<a:prstTxWarp prst="textArchDown"><a:avLst/>', '<a:prstTxWarp prst="textArchDown"><a:avLst><a:gd name="adj" fmla="val 21599999"/></a:avLst>'));
  const shape = Deck.open(pkg.write()).slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.deepEqual(shape.text.warp, { preset: 'textArchDown', adj: { adj: 'val 21599999' } });
});

test('the words are drawn along the preset\'s path, centred on it', () => {
  const { d, id } = deckWithBox();
  d.setTextWarp(0, id, 'textArchUp');
  const svg = renderSlide(d.slide(0), {});
  assert.match(svg, /<path id="(wp[a-z0-9]+)" d="M [\d.]+ [\d.]+ A [\d.]+ [\d.]+ 0 0 1 [\d.]+ [\d.]+" fill="none"\/>/);
  assert.match(svg, /<textPath href="#wp[a-z0-9]+" startOffset="50\.0%" text-anchor="middle"><tspan[^>]*>Rutba Office<\/tspan><\/textPath>/);
  d.setTextWarp(0, id, 'textCircle');
  assert.match(renderSlide(d.slide(0), {}), /textLength="[\d.]+" lengthAdjust="spacing"/, 'round a circle the letters are spaced out to go the whole way');
});

test('each preset\'s paths: one arc, one circle, three lines for the button', () => {
  const box = { x: 0, y: 0, w: 200, h: 100 };
  assert.equal(warpPaths('textArchUp', box).length, 1);
  assert.equal(warpPaths('textCircle', box)[0].length > warpPaths('textArchUp', box)[0].length * 1.9, true);
  assert.equal(warpPaths('textButton', box).length, 3);
  assert.deepEqual(warpPaths('textWave1', box), [], 'not drawn yet: straight');
  assert.deepEqual(WARP_PRESETS.map((p) => p.id), ['textNoShape', 'textArchUp', 'textArchDown', 'textCircle', 'textButton']);
});

test('a sheet shape\'s words take a Transform, drawn along its path and written as Excel writes it', async () => {
  const { buildXlsx } = await import('@rutba/ooxml/build');
  const { SheetView } = await import('@rutba/sheet-view');
  const { buildShape } = await import('@rutba/drawing');
  const { renderSvg, scene } = await import('@rutba/drawing');
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['a']] }] }), { viewportWidth: 800, viewportHeight: 400 });
  view.insertWordArt?.({ text: 'Rutba Office', size: 36 }) ?? view.addWordArt?.({ text: 'Rutba Office' });
  const shape = view.render().objects.find((o) => o.kind === 'shape');
  assert.ok(shape?.hasText, 'a WordArt shape with words');
  view.setShapeWarp({ id: shape.id, preset: 'textArchUp' });
  assert.equal(view.render().objects.find((o) => o.id === shape.id).textWarp, 'textArchUp');
  const drawing = OoxmlPackage.read(Buffer.from(view.serialize())).partNames().find((n) => /xl\/drawings\/drawing\d+\.xml$/.test(n));
  assert.match(OoxmlPackage.read(Buffer.from(view.serialize())).text(drawing), /<a:bodyPr\b[^>]*><a:prstTxWarp prst="textArchUp"><a:avLst\/><\/a:prstTxWarp>/);
  const svg = renderSvg(scene({ width: 400, height: 200, children: [buildShape({ geometry: 'rect', fill: 'none', text: 'Rutba Office', textSize: 24, textWarp: 'textArchUp' }, { x: 10, y: 10, width: 380, height: 180 })] }));
  assert.match(svg, /<textPath href="#wp[a-z0-9]+"[^>]*>.*Rutba Office/);
  view.setShapeWarp({ id: shape.id, preset: null });
  assert.equal(view.render().objects.find((o) => o.id === shape.id).textWarp, null);
});

test('a document\'s text box takes a Transform, written first in its body properties as Word writes it', async () => {
  const { openDocx } = await import('@rutba/doc-view/backends/ooxml');
  const { buildDocx } = await import('@rutba/ooxml');
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'First.' }, { text: 'Second.' }] }));
  view.setSelection({ block: 1, offset: 0 });
  view.insertTextBox({ widthPx: 300, heightPx: 120, autoFit: true, paragraphs: [{ text: 'Rutba Office', sizePt: 36 }] });
  const box = () => view.render({ pages: false }).blocks[1].textBoxes[0];
  const bodyPr = () => /<wps:bodyPr\b[^>]*\/>|<wps:bodyPr\b[^>]*>[\s\S]*?<\/wps:bodyPr>/.exec(view.doc.doc.xml)[0];
  assert.equal(box().warp, null);
  view.updateDrawings({ id: box().id, warp: 'textArchUp' });
  assert.match(bodyPr(), /^<wps:bodyPr\b[^>]*><a:prstTxWarp prst="textArchUp"><a:avLst\/><\/a:prstTxWarp><a:spAutoFit\/><\/wps:bodyPr>$/, 'the Transform first, then the fit, as the schema orders them');
  assert.equal(box().warp, 'textArchUp');
  view.updateDrawings({ id: box().id, autoFit: false });
  assert.match(bodyPr(), /<a:prstTxWarp prst="textArchUp"><a:avLst\/><\/a:prstTxWarp><a:noAutofit\/>/, 'a new fit goes after the Transform');
  view.updateDrawings({ id: box().id, warp: 'textCircle' });
  assert.equal((bodyPr().match(/prstTxWarp prst=/g) || []).length, 1, 'one preset, replaced');
  const reopened = openDocx(view.save()).render({ pages: false }).blocks[1].textBoxes[0];
  assert.equal(reopened.warp, 'textCircle', 'read back from the saved file');
  view.updateDrawings({ id: box().id, warp: null });
  assert.doesNotMatch(bodyPr(), /prstTxWarp/);
  assert.equal(box().warp, null);
  assert.throws(() => view.updateDrawings({ id: box().id, warp: '"/><x' }), /not a Transform preset/);
});
