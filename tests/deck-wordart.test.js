// WordArt on a slide — the engine's half.
//
// Insert → WordArt puts a text box on the slide whose words carry
// PowerPoint's text effects in the run itself: an outline (`a:ln`), no fill
// for an outline alone, and a glow and shadow (`a:effectLst`), in the order
// CT_TextCharacterProperties wants. The reader gives them back, an edit to
// the words keeps them, and the slide draws them: a stroke round each
// letter, and a filter round the line.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';

const blank = () => Deck.open(buildPptx({ title: 'Art', slides: [{ layout: 'blank' }] }));
const shadow = { blurPx: 4, distPx: 4, dir: 45, color: '#000000', alpha: 0.43 };
const glow = { radiusPt: 5, color: '#4472C4', alpha: 0.4 };

test('a WordArt run is written with its outline, fill and effects in schema order, and read back', () => {
  const deck = blank();
  deck.addTextBox(0, { x: 100, y: 100, w: 600, h: 100, name: 'WordArt', paragraphs: [{ align: 'center', runs: [{ text: 'Your text here', size: 54, color: '#FFFFFF', outline: { width: 1, color: '#4472C4' }, textEffects: { glow, shadow } }] }] });
  const xml = deck.pkg.text('ppt/slides/slide1.xml');
  const rPr = /<a:rPr\b[^>]*>([\s\S]*?)<\/a:rPr>/.exec(xml)[1];
  assert.match(rPr, /^<a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"\/><\/a:solidFill><\/a:ln><a:solidFill><a:srgbClr val="FFFFFF"\/><\/a:solidFill><a:effectLst><a:glow rad="63500"><a:srgbClr val="4472C4"><a:alpha val="40000"\/><\/a:srgbClr><\/a:glow><a:outerShdw blurRad="38100" dist="38100" dir="2700000"/);
  const run = deck.slide(0).shapes.find((s) => s.name.startsWith('WordArt')).text.paragraphs[0].runs[0];
  assert.deepEqual(run.outline, { width: 1, color: '#4472c4' });
  assert.equal(run.noFill, undefined);
  assert.equal(run.textEffects.glow.radiusPt, 5);
  assert.equal(Math.round(run.textEffects.shadow.distPx), 4);
  assert.equal(run.textEffects.shadow.dir, 45);
});

test('an outline alone has no fill, and an edit to the words keeps every effect', () => {
  const deck = blank();
  const id = deck.addTextBox(0, { x: 100, y: 100, w: 600, h: 100, name: 'WordArt', paragraphs: [{ runs: [{ text: 'Hollow', size: 54, noFill: true, outline: { width: 1.5, color: '#ED7D31' } }] }] });
  let run = deck.slide(0).shapes.find((s) => s.id === id || s.name.startsWith('WordArt')).text.paragraphs[0].runs[0];
  assert.equal(run.noFill, true);
  assert.deepEqual(run.outline, { width: 1.5, color: '#ed7d31' });
  // The words edited the way the window edits them: the paragraphs back, as read, with new text.
  const shape = deck.slide(0).shapes.find((s) => s.name.startsWith('WordArt'));
  const paragraphs = shape.text.paragraphs.map((p) => ({ ...p, runs: p.runs.map((r) => ({ ...r, text: 'Hollow words' })) }));
  deck.setText(0, shape.id, paragraphs);
  run = deck.slide(0).shapes.find((s) => s.name.startsWith('WordArt')).text.paragraphs[0].runs[0];
  assert.equal(run.text, 'Hollow words');
  assert.equal(run.noFill, true);
  assert.deepEqual(run.outline, { width: 1.5, color: '#ed7d31' });
});

test('the slide draws the outline as a stroke, an outline alone unfilled, and the shadow and glow as a filter round the line', () => {
  const deck = blank();
  deck.addTextBox(0, { x: 100, y: 100, w: 600, h: 100, name: 'WordArt', paragraphs: [{ runs: [{ text: 'Shine', size: 54, color: '#FFC000', textEffects: { glow, shadow } }] }] });
  deck.addTextBox(0, { x: 100, y: 300, w: 600, h: 100, name: 'WordArt', paragraphs: [{ runs: [{ text: 'Hollow', size: 54, noFill: true, outline: { width: 1.5, color: '#ED7D31' } }] }] });
  const svg = renderSlide(deck.slide(0), { width: 1280 });
  assert.match(svg, /<tspan [^>]*fill="none"[^>]*stroke="#ed7d31" stroke-width="2\.00"[^>]*>Hollow<\/tspan>/);
  const shine = /<text\b([^>]*)><tspan [^>]*>Shine<\/tspan>/.exec(svg);
  assert.ok(shine, 'the line is drawn');
  const id = /filter="url\(#([^)]+)\)"/.exec(shine[1])?.[1];
  assert.ok(id, 'the line carries a filter');
  const filter = new RegExp('<filter id="' + id + '"[^>]*>([\\s\\S]*?)</filter>').exec(svg)?.[1] || '';
  assert.match(filter, /feFlood flood-color="#4472c4"/);
  assert.match(filter, /feOffset/);
});
